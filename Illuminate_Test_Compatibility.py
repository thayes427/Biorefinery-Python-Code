# -*- coding: utf-8 -*-
"""
Created on Thu Feb 28 14:58:42 2019

@author: MENGstudents

output tab: 
that it exists
distinguish between value and string to see if they've put it in the right spot
instead could just make sure they've put the column headers in the right spot
VBA functions
test to make sure they all exist and work
clear: make sure data is empty
fill: make sure data has appeared in the cell
solve: try to find the NPV - goes to 0 when DCFROR is called. Break the code and see what happens when it fails. 
Aspen
make sure that all of the nodes they want to change are not nonetype
make sure that the fortran value to change is in the fortran value

"""

import Illuminate_Simulations as simulations
from os import path



def compatibility_test(error_queue, excel_input_file, calculator_file, aspen_file):
    error_queue.put((False, 'Testing Compatibility of Excel Calculator File'))
    test_calculator_file(calculator_file, aspen_file, error_queue)
    

def test_calculator_file(calculator_file, aspen_file, error_queue):
    
    excel, book = simulations.open_excelCOMS(calculator_file)
    
    ########### Make sure that the output tab exists  ###################
    output_tab_exists = False
    try:
        book.Sheets('Output')
        output_tab_exists = True
    except:
        error_queue.put((True, '"Output" tab missing from Excel calculator .xlsm file. Please add this tab'))
        
        
        
    ########## Make sure output tab is set up as it is supposed to be  ######## 
    if output_tab_exists:
        if any(str(v) != "Variable Name" for v in book.Sheets('Output').Evaluate('B2')):
            error_queue.put((True,'Output tab is not configured properly. The column header for "Variable Name"\nshould be in B3 so that the first variable name is in B4'))
        elif any(str(v) != "Variable Value" for v in book.Sheets('Output').Evaluate('C2')):
            error_queue.put((True,'Output tab is not configured properly. The column header for "Variable Value" \
                                    should be in C3 so that the first variable value is in C4'))
            
            
    ######### Make sure the bkp file reference is where it should be #########
    setup_tab_functional = True
    try:
        book.Sheets('Set-up')
    except:
        setup_tab_functional = False
        error_queue.put((True,'"Set-up" tab missing from Excel calculator .xlsm file. Please rename this tab.'))
    try:
        filename, file_extension = path.splitext(book.Sheets('Set-up').Evaluate('B1').Value)
        if not (file_extension=='.bkp' or file_extension == '.apw'):
            error_queue.put((True,'In the "Set-up" tab, the name of the .apw or .bkp should be in cell B1. If the location of this \
                                    reference needs to be changed, make sure that you also change it in the "sub_GetSumData" macro'))
    except:
        setup_tab_functional = False
        error_queue.put((True,'In the "Set-up" tab, the name of the .apw or .bkp should be in cell B1. If the location of this \
                                    reference needs to be changed, make sure that you also change it in the "sub_GetSumData" macro'))
    
        
    
        
    ####################  Test all important macros ########################
    try:
        excel.Run('sub_ClearSumData_ASPEN')
    except:
        error_queue.put((True,'Macro with name "sub_ClearSumData_ASPEN" does not exist in the Excel calculator .xlsm file'))
    
    try:
        if not all(str(v)=='None' for v in book.Sheets('aspen').Evaluate('C7:C20')):
            error_queue.put((True,'Excel macro sub_ClearSumData_ASPEN does not appear to be working. Values in column C of sheet\n"aspen" are not being cleared.'))
    except:
        pass
    
    
    if setup_tab_functional:
        try:
            book.Sheets('Set-up').Evaluate('B1').Value = aspen_file
            excel.Run('sub_GetSumData_ASPEN')
            if not all(str(v)!='None' for v in book.Sheets('aspen').Evaluate('C8')):
                error_queue.put((True,'"sub_GetSumData_ASPEN" does not appear to be working. Values should be populated in column C of sheet "aspen"'))
                
        except:
            error_queue.put((True,'Macro with name "sub_GetSumData_ASPEN" does not exist in the Excel calculator .xlsm file or calls\nanother function that does not exist'))
        
    
    #excel.Run('solvedcfror')
        
        
        
    
    # 
    
    