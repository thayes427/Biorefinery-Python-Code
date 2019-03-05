# -*- coding: utf-8 -*-
"""
Created on Thu Feb 28 14:58:42 2019

@author: MENGstudents


"""

import Illuminate_Simulations as simulations
from os import path, rmdir, makedirs
from psutil import process_iter
from pandas import read_excel
from re import findall
from math import isclose
from shutil import copyfile, rmtree



def compatibility_test(status_queue, excel_input_file, calculator_file, aspen_file, dispatch):
    '''
    Tests for compatibility issues in the user's input files. First tests
    the Excel Calculator file then the Aspen model. The status updates from this
    function are live updated on the GUI in the Main_App.
    '''
    
    aspen_dir = copy_aspen_to_temp_dir(aspen_file)
    
    status_queue.put((False, 'Testing Compatibility of Excel Input File...'))
    excel_input_errors_found = test_excel_input_file(excel_input_file, status_queue)
    if excel_input_errors_found:
        status_queue.put((True, 'Finished Testing Excel Input File, Please Fix Errors'))
    else:
        status_queue.put((False, 'SUCCESS: Excel Input File is Compatible with Illuminate'))
    status_queue.put((False, 'Testing Compatibility of Excel Calculator File...'))
    errors_found = test_calculator_file(calculator_file, aspen_file, status_queue)
    if errors_found:
        status_queue.put((True, 'Finished Testing Excel Calculator File, Please Fix Errors'))
    else:
        status_queue.put((False, 'SUCCESS: Excel Calculator File is Compatible with Illuminate'))
        
    if excel_input_errors_found:
        status_queue.put((True, 'Cannot test compatibility of Aspen model until Excel input file is compatible. Please fix errors with Excel input file and test compatibility again'))
    else:
        status_queue.put((False, 'Testing Compatibility of Aspen Model...'))
        errors_found = test_aspen_file(aspen_dir, excel_input_file, dispatch, status_queue)
        if errors_found:
            status_queue.put((True, 'Finished Testing Aspen Model, Please Fix Errors'))
        else:
            status_queue.put((False, 'SUCCESS: Aspen Model is Compatible with Illuminate'))
    status_queue.put((False, 'Finished with Compatibility Test'))
    
    
def copy_aspen_to_temp_dir(aspen_file):
    '''
    Copies the aspen .apw or .bkp file provided by the user to a temporary
    directory within the 'Output' folder. This is done in order to encapsulate
    all of the extra files that Aspen outputs so that they can be easily removed
    if Aspen crashes or is aborted. It first checks to see if this directory exists,
    and if it does exist, then it deletes the temporary directory and all
    of its contents.
    '''
    
    output_directory = path.join(path.dirname(aspen_file),'Output')
    if not path.exists(output_directory):
        makedirs(output_directory)
    if not path.exists(path.join(output_directory,'Temp')):
        makedirs(path.join(output_directory,'Temp'))
    
    temp_directory = path.join(path.dirname(aspen_file),'Output','Temp')
    
    # delete the directory if it exists
    try:
        rmdir(temp_directory)
    except: 
        pass
    try:
        rmtree(temp_directory)
    except: 
        pass
    makedirs(temp_directory)
    aspen_file_temp = path.join(temp_directory,path.basename(aspen_file))
    copyfile(aspen_file, aspen_file_temp)
    
    return aspen_file_temp
    
def test_excel_input_file(excel_input_file, status_queue):
    
    errors_found = False
    col_types = {'Variable Name': str, 'Variable Aspen Call': str, 'Distribution Parameters': str, 
                 'Bounds': str, 'Fortran Call':str, 'Fortran Value to Change': str, 
                 'Distribution Type': str, 'Toggle': bool}
    try:
        df = read_excel(excel_input_file, sheet_name='Inputs', dtype=col_types)
    except:
        status_queue.put((True, 'There must be a sheet titled "Inputs" in the Excel input file and another sheet titled "Warning Messages"'))
        errors_found = True
        return
    try:
        df = read_excel(excel_input_file, sheet_name='Warning Messages')
    except:
        status_queue.put((True, 'There must be a sheet titled "Inputs" in the Excel input file and another sheet titled "Warning Messages"'))
        errors_found = True
    
    
    required_columns = ['Variable Name', 'Variable Aspen Call', 'Distribution Parameters', 
                 'Bounds', 'Fortran Call', 'Fortran Value to Change', 
                 'Distribution Type', 'Toggle']
    df = read_excel(excel_input_file, sheet_name='Inputs', dtype=col_types)
    user_columns = set(df.columns)
    
    for col in required_columns:
        if col not in user_columns:
            status_queue.put((True, 'Column "' + col + '" must be in the "Inputs" tab of the Excel input file'))
            errors_found = True
    
    return errors_found
    
    
    
def test_aspen_file(aspen_file,excel_input_file, dispatch, status_queue):
    '''
    Makes sure the Aspen model can be opened, tests to make sure that all
    Aspen nodes specified in the Excel input file exist in the Aspen model and 
    are not None. Finally, it makes sure that for any Fortran variables, the value
    to change can be found within the Fortran variable string.
    '''
    
    errors_found = False
    ######### Open Aspen COM and get a handle on that COM so we can terminate it #######
    aspens_to_ignore = set()
    for p in process_iter():
        if 'aspen' in p.name().lower() or 'apwn' in p.name().lower():
            aspens_to_ignore.add(p.pid)       
    status_queue.put((False, 'Opening Aspen Model...'))
    try:
        aspencom, obj = simulations.open_aspenCOMS(aspen_file, dispatch)
    except:
        status_queue.put((True, 'Aspen model cannot be opened'))
        errors_found = True
        return
    for p in process_iter():
        if ('aspen' in p.name().lower() or 'apwn' in p.name().lower()) and p.pid not in aspens_to_ignore:
            aspen_to_delete = p
        
            
    # Make sure that all nodes in the tree exist
    status_queue.put((False, 'Testing Aspen Paths Specified in Input File...'))
    col_types = {'Variable Name': str, 'Variable Aspen Call': str, 'Distribution Parameters': str, 
                 'Bounds': str, 'Fortran Call':str, 'Fortran Value to Change': str, 
                 'Distribution Type': str, 'Toggle': bool}
    df = read_excel(open(excel_input_file,'rb'), dtype=col_types)
    for index, row in df.iterrows():
        if row['Toggle']: 
            try:
                if obj.FindNode(row['Variable Aspen Call']).Value is None:     
                    status_queue.put((True, 'The value at the node "'+ row['Variable Aspen Call'] + \
                                      '" for variable "' + row['Variable Name'] + \
                                      '" is None. Are you sure this is the right path?'))
                    errors_found = True
            except:
                status_queue.put((True, 'Aspen call "'+ row['Variable Aspen Call'] + \
                                  '" for variable "' + row['Variable Name'] + \
                                  '" does not exist in the Aspen model'))
                errors_found = True
    for index, row in df.iterrows():
        if row['Toggle'] and row['Fortran Call']:
            if row['Fortran Value to Change'] not in row['Fortran Call']:
                status_queue.put((True, 'The fortran value to change "' + \
                                  row['Fortran Value to Change'] + '" for variable "' + \
                                  row['Variable Name'] + '" is not in the Fortran call "' +
                                  row['Fortran Call'] + '"'))
                errors_found = True
     
    aspen_to_delete.terminate()
    return errors_found


def test_calculator_file(calculator_file, aspen_file, status_queue):
    '''
    Tests for compatibility issues in the Excel Calculator file. First makes sure
    the Output tab exists and is configured properly. It then makes sure the .bkp
    reference in the setup tab is configured as expected. Finally, it tests
    to make sure macros are named properly and are functional.
    '''
    
    errors_found = False
    ########### Open Excel COM and get a handle on it to terminate it later ########
    excels_to_ignore = set()
    for p in process_iter():
        if 'excel' in p.name().lower():
            excels_to_ignore.add(p.pid)
    excel, book = simulations.open_excelCOMS(calculator_file)
    for p in process_iter():
        if 'excel' in p.name().lower() and p.pid not in excels_to_ignore:
            excel_to_delete = p
    
    ########### Make sure that the output tab exists  ###################
    output_tab_exists = False
    try:
        book.Sheets('Output')
        output_tab_exists = True
    except:
        status_queue.put((True,'"Output" tab missing from Excel calculator .xlsm file. Please add this tab'))
        errors_found = True
        
        
    ########## Make sure output tab is set up as it is supposed to be  ######## 
    if output_tab_exists:
        if any(str(v) != "Variable Name" for v in book.Sheets('Output').Evaluate('B2')):
            status_queue.put((True,'Output tab is not configured properly. The column header for '+\
                              '"Variable Name" should be in B2 so that the first variable name is in B3'))
            errors_found = True
        elif any(str(v) != "Value" for v in book.Sheets('Output').Evaluate('C2')):
            status_queue.put((True,'Output tab is not configured properly. The column header '+\
                              'for "Value" should be in C2 so that the first variable value is in C3'))
            errors_found = True
            
            
    ######### Make sure the bkp file reference is where it should be #########
    try:
        vba_code = book.VBProject.VBComponents("GelAllData").CodeModule.Lines(1,500000)
    except:
        try:
            vba_code = book.VBProject.VBComponents("GetAllData").CodeModule.Lines(1,500000)
        except:
            vba_code = ''
    
    i=0
    get_data_VBA = ""
    while i < len(vba_code):
        if vba_code[i:i+31] == 'Public Sub sub_GetSumData_ASPEN':
            s_ind = i
            while vba_code[i-7:i] != 'End Sub':
                i += 1
            get_data_VBA = vba_code[s_ind:i]
            break
        i +=1
    
    if get_data_VBA:
        bkp_reference_cell = findall(r"RTrim\(Worksheets\(\"Set-up\"\)\.Range\(\"([A-Z]+[0-9]+)\"\)\.VALUE", 
                                              get_data_VBA)[0]
    else:
        bkp_reference_cell = 'B1'
    
    
    
    
    setup_tab_functional = True
    try:
        book.Sheets('Set-up')
    except:
        setup_tab_functional = False
        status_queue.put((True,'"Set-up" tab missing from Excel calculator .xlsm file. '+\
                          'Please rename this tab.'))
        errors_found = True
    try:
        filename, file_extension = path.splitext(book.Sheets('Set-up').Evaluate(bkp_reference_cell).Value)
        if not (file_extension=='.bkp' or file_extension == '.apw'):
            status_queue.put((True,'In the "Set-up" tab, the name of the .apw or .bkp '+\
                              'should be in cell B1. If, however, you have made VBA '+\
                              'accessible to Illuminate, then you can have this bkp '+\
                              'reference in a different location. If it is not in B1, '+\
                              'then the reference in "sub_GetSumData_ASPEN" must be updated. '+\
                              'If the location of this reference needs '+\
                              'to be changed, make sure that you also change it in the "sub_GetSumData" macro'))
            errors_found = True
    except:
        setup_tab_functional = False
        status_queue.put((True,'In the "Set-up" tab, the name of the .apw or .bkp '+\
                              'should be in cell B1. If, however, you have made VBA accessible to '+\
                              'Illuminate, then you can have this bkp reference in a different '+\
                              'location. If it is not in B1, then the reference in "sub_GetSumData_ASPEN" '+\
                              'must be updated. If the location of this reference needs '+\
                              'to be changed, make sure that you also change it in the "sub_GetSumData" macro'))
        errors_found = True
    
        
    ####################  Test all important macros ########################
    
    try:
        clear_load_cell = findall(r"Range\(\"([A-Z]+[0-9]+)\"\)\.End\(xlDown\)\.Row", 
                                  book.VBProject.VBComponents("GelAllData").CodeModule.Lines(1,500000))[0]
    except:
        try:
            clear_load_cell = findall(r"Range\(\"([A-Z]+[0-9]+)\"\)\.End\(xlDown\)\.Row", 
                                      book.VBProject.VBComponents("GelAllData").CodeModule.Lines(1,500000))[0]
        except:
            clear_load_cell = 'C7'
    try:
        excel.Run('sub_ClearSumData_ASPEN')
        try:
            if book.Sheets('aspen').Evaluate(clear_load_cell).Value is not None:
                status_queue.put((True,'Excel macro sub_ClearSumData_ASPEN does not appear to '+\
                                  'be working. Values in column C of sheet "aspen" are not being cleared.'))
                errors_found = True
        except:
            pass
    except:
        status_queue.put((True, 'Macro with name "sub_ClearSumData_ASPEN" does '+\
                          'not exist or is broken'))
        errors_found = True
    

    
    
    if setup_tab_functional:
        try:
            book.Sheets('Set-up').Evaluate(bkp_reference_cell).Value = aspen_file
            excel.Run('sub_GetSumData_ASPEN')
            if book.Sheets('aspen').Evaluate(clear_load_cell).Value is None:
                status_queue.put((True,'"sub_GetSumData_ASPEN" does not appear to be '+\
                                  'working. Values should be populated in column C of sheet "aspen"'))
                errors_found = True
                
        except:
            status_queue.put((True,'Macro with name "sub_GetSumData_ASPEN" does not '+\
                              'exist or is broken'))
            errors_found = True
        
    
    try:
        excel.Run('solvedcfror')
    except:
        status_queue.put((True, 'Macro with name "solvedcfror" does not exist or is broken.'))
        errors_found = True
        return

    
    try:
        module1_VBA = book.VBProject.VBComponents("Module1").CodeModule.Lines(1,50000000)
        vba_code_access = True
    except:
        status_queue.put((True, 'Unable to access "solvedcfror" VBA code and therefore cannot test'+\
                          '"solvedcfror" functionality. ' +\
                          'If you would like Illuminate to be able to test this, you must enable access'+\
                          'by opening the .xlsm file and going to' +\
                          'File -> Options -> Trust Center -> Trust Center Settings -> '+\
                          'Macro Settings -> Trust Access to VBA project object model'))
        errors_found = True
        vba_code_access = False
    if vba_code_access:
        i=0
        DCFROR_VBA = ""
        while i < len(module1_VBA):
            if module1_VBA[i:i+15] == 'Sub solvedcfror':
                s_ind = i
                while module1_VBA[i-7:i] != 'End Sub':
                    i += 1
                DCFROR_VBA = module1_VBA[s_ind:i]
                break
            i +=1
            
        DCFROR_cells = findall(
                r"Range\(\"([A-Z]+[0-9]+)\"\)\.GoalSeek Goal\:\=0\, ChangingCell\:\=Range\(\"([A-Z]+[0-9]+)\"\)", 
                DCFROR_VBA)
        DCFROR_sheetname = findall(r"Sheets\(\"(.*)\"\).Select",DCFROR_VBA)[0]
        
        if not DCFROR_cells:
            status_queue.put((True, 'Cannot find VBA code for "solvedcfror" in Module1. This is not critical, '+\
                              ', but it means that Illuminate cannot test the functionality of this macro. If you are confident' +\
                              'that it is working, then disregard this message.'))
            errors_found = True
        else:
            goal_seek, change_cell = DCFROR_cells[0]
            seek_val = float(book.Sheets(DCFROR_sheetname).Evaluate(goal_seek).Value)
            book.Sheets(DCFROR_sheetname).Evaluate(change_cell).Value = 5.0
            
            
            
            if isclose(book.Sheets(DCFROR_sheetname).Evaluate(goal_seek).Value, seek_val):
                status_queue.put((True, 'The "goal seek" and "change cell" cells indicated '+\
                                  'in the "solvedcfror" code do not appear to be linked.'+\
                                  'Make sure these are the correct cells referenced in the macro' ))
                errors_found = True
            
            excel.Run('solvedcfror')
            if not isclose(float(book.Sheets(DCFROR_sheetname).Evaluate(goal_seek).Value), 0.0):
                status_queue.put((True, 'The "solvedcfror" function is not minimizing the '+\
                                  '"goal seek" cell to 0 as it should be.' ))
                errors_found = True

        
        
    excel_to_delete.terminate()
    return errors_found
    
    
    