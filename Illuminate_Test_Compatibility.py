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


calculator_file_name = "C:/Users/MENGstudents/Desktop/Biorefinery-Design-Project/BC1707A_sugars_V10_mod-2.xlsm"



def compatibility_test(excel_input_file, calculator_file, aspen_file):
    return

def test_calculator_file(calculator_file_name, error_statements):
    
    excel, book = simulations.open_excelCOMS(calculator_file_name)
    return excel, book
    
    
    ########### Make sure that the output tab exists  ###################
    output_tab_exists = False
    try:
        book.Sheets('Output')
        output_tab_exists = True
    except:
        error_statements.append('"Output" tab missing from Excel calculator .xlsm file. Please add this tab')
        
        
        
    ########## Make sure output tab is set up as it is supposed to be  ######## 
    if output_tab_exists:
        if any(str(v) != "Variable Name" for v in book.Sheets('Output').Evaluate('B2')):
            error_statements.append('Output tab is not configured properly. The column header for "Variable Name" \
                                    should be in B3 so that the first variable name is in B4')
        elif any(str(v) != "Variable Value" for v in book.Sheets('Output').Evaluate('C2')):
            error_statements.append('Output tab is not configured properly. The column header for "Variable Value" \
                                    should be in C3 so that the first variable value is in C4')
            
            
    ######### Make sure the bkp file reference is where it should be #########
    try:
        book.Sheets('Set-up')
    except:
        error_statements.append('"Set-up" tab missing from Excel calculator .xlsm file.')
    try:
        filename, file_extension = path.splitext(book.Sheets('Set-up').Evaluate('B1').Value)
        if not (file_extension=='.bkp' or file_extension == '.apw'):
            error_statements.append(')
    
        
        
    ############  Test all important macros ###################
    try:
        excel.Run('sub_ClearSumData_ASPEN')
    except:
        error_statements.append('Macro with name "sub_ClearSumData_ASPEN" does not exist in the Excel calculator .xlsm file')
    
    try:
        if not all(str(v)=='None' for v in book.Sheets('aspen').Evaluate('C8:C20')):
            error_statements.append('Excel macro sub_ClearSumData_ASPEN does not appear to be working. Values in column C of sheet "aspen" are not being cleared.')
    except:
        pass
    
    
    try:
        excel.Run('sub_GetSumData_ASPEN')
    except:
        error_statements.append('Macro with name "sub_GetSumData_ASPEN" does not exist in the Excel calculator .xlsm file or calls another function that does not exist')
    
    excel.Run('sub_GetSumData_ASPEN')
    excel.Run('solvedcfror')
        
        
    return error_statements
        
    
    # 
    
    