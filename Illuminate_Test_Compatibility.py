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


calculator_file_name = "C:/Users/MENGstudents/Desktop/Biorefinery-Design-Project/BC1707A_sugars_V10_mod-2.xlsm"

def test_output_tab(calculator_file_name, error_statements):
    
    excel, book = simulations.open_excelCOMS(calculator_file_name)
    return book
    
    
    # a) make sure that the output tab exists
    output_tab_exists = False
    try:
        book.Sheets('Output')
        output_tab_exists = True
    except:
        error_statements.append('"Output" tab missing from Excel calculator .xlsm file. Please add this tab')
        
    # b) Make sure output tab is set up as it is supposed to be    
    if output_tab_exists:
        if any(str(v) != "Variable Name" for v in book.Sheets('Output').Evaluate('B2')):
            error_statements.append('Output tab is not configured properly. The column header for "Variable Name" \
                                    should be in B3 so that the first variable name is in B4')
        
        
        
        
    
        
    return error_statements
        
    
    # 
    
    