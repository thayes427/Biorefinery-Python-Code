# -*- coding: utf-8 -*-
"""
Created on Thu Feb 28 14:58:42 2019

@author: MENGstudents


"""

import Illuminate_Simulations as simulations
from os import path
from psutil import process_iter
from pandas import read_excel


def compatibility_test(status_queue, excel_input_file, calculator_file, aspen_file, dispatch):
    '''
    Tests for compatibility issues in the user's input files. First tests
    the Excel Calculator file then the Aspen model. The status updates from this
    function are live updated on the GUI in the Main_App.
    '''
    
    status_queue.put((False, 'Testing Compatibility of Excel Calculator File...'))
    test_calculator_file(calculator_file, aspen_file, status_queue)
    status_queue.put((False, 'Testing Compatibility of Aspen Model...'))
    test_aspen_file(aspen_file, excel_input_file, dispatch, status_queue)
    status_queue.put((False, 'Finished with Compatibility Test'))
    
    
def test_aspen_file(aspen_file,excel_input_file, dispatch, status_queue):
    '''
    Makes sure the Aspen model can be opened, tests to make sure that all
    Aspen nodes specified in the Excel input file exist in the Aspen model and 
    are not None. Finally, it makes sure that for any Fortran variables, the value
    to change can be found within the Fortran variable string.
    '''
    
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
            except:
                status_queue.put((True, 'Aspen call "'+ row['Variable Aspen Call'] + \
                                  '" for variable "' + row['Variable Name'] + \
                                  '" does not exist in the Aspen model'))
    for index, row in df.iterrows():
        if row['Toggle'] and row['Fortran Call']:
            if row['Fortran Value to Change'] not in row['Fortran Call']:
                status_queue.put((True, 'The fortran value to change "' + \
                                  row['Fortran Value to Change'] + '" for variable "' + \
                                  row['Variable Name'] + '" is not in the Fortran call "' +
                                  row['Fortran Call'] + '"'))
     
    aspen_to_delete.terminate()


def test_calculator_file(calculator_file, aspen_file, status_queue):
    '''
    Tests for compatibility issues in the Excel Calculator file. First makes sure
    the Output tab exists and is configured properly. It then makes sure the .bkp
    reference in the setup tab is configured as expected. Finally, it tests
    to make sure macros are named properly and are functional.
    '''
    
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
        
        
    ########## Make sure output tab is set up as it is supposed to be  ######## 
    if output_tab_exists:
        if any(str(v) != "Variable Name" for v in book.Sheets('Output').Evaluate('B2')):
            status_queue.put((True,'Output tab is not configured properly. The column header for '+\
                              '"Variable Name" should be in B3 so that the first variable name is in B4'))
        elif any(str(v) != "Value" for v in book.Sheets('Output').Evaluate('C2')):
            status_queue.put((True,'Output tab is not configured properly. The column header '+\
                              'for "Value" should be in C3 so that the first variable value is in C4'))
            
            
    ######### Make sure the bkp file reference is where it should be #########
    setup_tab_functional = True
    try:
        book.Sheets('Set-up')
    except:
        setup_tab_functional = False
        status_queue.put((True,'"Set-up" tab missing from Excel calculator .xlsm file. '+\
                          'Please rename this tab.'))
    try:
        filename, file_extension = path.splitext(book.Sheets('Set-up').Evaluate('B1').Value)
        if not (file_extension=='.bkp' or file_extension == '.apw'):
            status_queue.put((True,'In the "Set-up" tab, the name of the .apw or .bkp '+\
                              'should be in cell B1. If the location of this reference needs '+\
                              'to be changed, make sure that you also change it in the "sub_GetSumData" macro'))
    except:
        setup_tab_functional = False
        status_queue.put((True,'In the "Set-up" tab, the name of the .apw or .bkp should '+\
                          'be in cell B1. If the location of this reference needs to be changed,'+\
                          'make sure that you also change it in the "sub_GetSumData" macro'))
    
        
    ####################  Test all important macros ########################
    try:
        excel.Run('sub_ClearSumData_ASPEN')
    except:
        status_queue.put((True, 'Macro with name "sub_ClearSumData_ASPEN" does '+\
                          'not exist in the Excel calculator .xlsm file'))
    
    try:
        if not all(str(v)=='None' for v in book.Sheets('aspen').Evaluate('C7:C20')):
            status_queue.put((True,'Excel macro sub_ClearSumData_ASPEN does not appear to '+\
                              'be working. Values in column C of sheet "aspen" are not being cleared.'))
    except:
        pass
    
    
    if setup_tab_functional:
        try:
            book.Sheets('Set-up').Evaluate('B1').Value = aspen_file
            excel.Run('sub_GetSumData_ASPEN')
            if not all(str(v)!='None' for v in book.Sheets('aspen').Evaluate('C8')):
                status_queue.put((True,'"sub_GetSumData_ASPEN" does not appear to be '+\
                                  'working. Values should be populated in column C of sheet "aspen"'))
                
        except:
            status_queue.put((True,'Macro with name "sub_GetSumData_ASPEN" does not '+\
                              'exist in the Excel calculator .xlsm file or calls another'+\
                              'function that does not exist'))
        
    
    
    #excel.Run('solvedcfror')
        
        
    excel_to_delete.terminate()
    
    
    