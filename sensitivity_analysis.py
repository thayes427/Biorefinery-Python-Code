# -*- coding: utf-8 -*-
"""
Created on Thu Feb 22 10:59:41 2018

@author: Group D

This library performs objective 2. 
"""

import win32com.client as win32
import os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from time import time
from math import ceil
import random
import csv
import GUI_multivariate as GUI
import psutil

#aspenfilename =  'BC1508F-BC_FY17Target._Final_5ptoC5_updated022618.bkp'
#excelfilename = 'DESIGN_OBJ2_test_MFSP-updated.xlsm' 


def open_COMS(aspenfilename, excelfilename):
    
    
    print('Initializing Aspen COM...')
    aspen = win32.Dispatch('Apwn.Document')
    print('Aspen COM Initialized')
    aspen.InitFromArchive(os.path.abspath(aspenfilename))
    print('Aspen File Open.')
    obj = aspen.Tree
    
    print('Initializing Excel COM...')
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    print('Excel COM Initialized')
    book = excel.Workbooks.Open(os.path.abspath(excelfilename))
    print('Excel File Open')
    
    
    return aspen,obj,excel,book

def get_distributions(gui_excel_input, ntrials=1):
    '''
    Given the excel input from the user in the GUI, produce a list_of_variables
    the user wants to change as well as their distributions that should be 
    randomly sampled from. 
    '''
    
    with open(gui_excel_input) as f:
        reader = csv.DictReader(f)# Skip the header row
        simulation_vars = {}
        simulation_dist = {}
        for row in reader:
            if row['Toggle'].lower().strip() == 'true':
                dist_type = row['Format of Range'].lower()
                aspen_variable = row['Variable Name']
                aspen_call = row['Variable Aspen Call']
                bounds = row['Bounds'].split(',')
                lb = float(bounds[0].strip())
                ub = float(bounds[1].strip())
                if 'normal' in dist_type or 'gaussian' in dist_type:
                    dist_variables = row['Range of Values'].split(',')
                    distribution = sample_gauss(float(dist_variables[0].strip()),
                              float(dist_variables[1].strip()), lb, ub, ntrials)
                if 'linspace' in dist_type:
                    linspace_vars = row['Range of Values'].split(',')
                    distribution = np.linspace(float(linspace_vars[0].strip()), 
                                               float(linspace_vars[1].strip()),
                                               float(linspace_vars[2].strip()))
                if 'poisson' in dist_type:
                    lambda_p = float(row['Range of Values'].strip())
                    distribution = sample_poisson(lambda_p, lb, ub, ntrials)
                if 'pareto' in dist_type:
                    pareto_vals = row['Range of Values'].split(',')
                    shape = float(pareto_vals[0].strip())
                    scale = float(pareto_vals[1].strip())
                    distribution = sample_pareto(shape, scale, lb, ub, ntrials)
                if 'list' in dist_type:
                    lst = row['Range of Values'].split(',')
                    distribution = []
                    for l in lst:
                        distribution.append(float(l.strip()))                
                if 'uniform' in dist_type:
                    lb_ub = row['Range of Values'].split(',')
                    lb_uniform, ub_uniform = float(lb_ub[0].strip()), float(lb_ub[1].strip())
                    distribution = sample_uniform(lb_uniform, ub_uniform, lb, ub, ntrials)
                simulation_dist[aspen_variable] = distribution
                fortran_index = (0,0)
                if row['Fortran Call'].strip() != "":
                    is_fortran = True
                    fortran_call = row['Fortran Call']
                    value_to_change = row['Fortran Value to Change'].strip()
                    len_val = len(value_to_change)
                    for i in range(len(fortran_call)):
                        if fortran_call[i:i+len_val] == value_to_change:
                            fortran_index = (i, i+len_val) #NOT INCLUSIVE
                    for i, v in enumerate(distribution):
                        distribution[i] = make_fortran(fortran_call, fortran_index, v)
                simulation_vars[(aspen_variable, aspen_call, fortran_index)] = distribution
    
    return simulation_vars, simulation_dist
    
def sample_gauss(mean, std, lb, ub, ntrials):
    d = []
    for i in range(ntrials):
        rand_sample = np.random.normal(mean,std)
        while(rand_sample < lb or rand_sample > ub):
            rand_sample = np.random.normal(mean,std)
        d.append(rand_sample)
    return d

def sample_uniform(lb_uniform, ub_uniform, lb, ub, ntrials):
    d = []
    for i in range(ntrials):
        rand_sample = np.random.uniform(lb_uniform, ub_uniform)
        while(rand_sample < lb or rand_sample > ub):
            rand_sample = np.random.uniform(lb_uniform, ub_uniform)
        d.append(rand_sample)
    return d


def sample_poisson(lambda_p, lb, ub, ntrials):
    d = []
    for i in range(ntrials):
        rand_sample = np.random.poisson(10000*lambda_p)/10000
        while(rand_sample < lb or rand_sample > ub):
            rand_sample = np.random.poisson(10000*lambda_p)/10000
        d.append(rand_sample)
    return d

def sample_pareto(shape, scale, lb, ub, ntrials):
    d = []
    for i in range(ntrials):
        rand_sample = (np.random.pareto(shape) + 1) * scale
        while(rand_sample < lb or rand_sample > ub):
            rand_sample = (np.random.pareto(shape) + 1) * scale
        d.append(rand_sample)
    return d

def make_fortran(fortran_call, fortran_index, val):
    return fortran_call[:fortran_index[0]] + str(val) + fortran_call[fortran_index[1]:]


def multivariate_sensitivity_analysis(aspenfilename, excelfilename, 
    gui_excel_input, num_trials, output_file_name, simulation_vars, disp_graphs=True):
    global dfstreams
    aspen,obj,excel,book = open_COMS(aspenfilename,excelfilename)
    
    
    SUC_LOC = r"\Data\Blocks\A300\Data\Blocks\B1\Input\FRAC\TOC5"
    
    vars_to_change = []
    with open(gui_excel_input) as f:
        reader = csv.DictReader(f)# Skip the header row
        for row in reader:
            if row['Toggle'].lower().strip() == 'true':
                vars_to_change.append(row["Variable Name"])
    variable_values = {} # a dictionary to store the values each variable takes for each simulation
    
    columns = vars_to_change + ['Biofuel Output', 'Succinic Acid Output', 'Fixed Op Costs',\
              'Var OpCosts', ' Capital Costs', 'MFSP','Fixed Capital Investment',\
              'Capital Investment with Interest','Loan Payment per Year','Depreciation','Cash on Hand',\
              'Steam Plant Value','Bag Cost', 'Aspen Errors']
    
    dfstreams = pd.DataFrame(columns=columns)
    obj.FindNode(SUC_LOC).Value = 0.4
    
    ########## RUN SIMULATION #########
    old_time = time()
    start_time = time()
    for trial in range(num_trials):
        ####### UPDATE ASPEN VARIABLES  ########
        for (aspen_variable, aspen_call, fortran_index), dist in simulation_vars.items():
            obj.FindNode(aspen_call).Value = dist[trial]
            if type(dist[trial]) == str:
                variable_values[aspen_variable] = float(dist[trial][fortran_index[0]:fortran_index[1]])
            else:
                variable_values[aspen_variable] = dist[trial]
        
        ########## STORE THE RANDOMLY SAMPLED VARIABLE VALUES  ##########
        case_values = []
        for v in vars_to_change:
            case_values.append(variable_values[v])
            
        print(variable_values)
        ######## RUN ASPEN SIMULATION WITH RANDOMLY SAMPLED VARIABLES #######
        aspen.Reinit()
        aspen.Engine.Run2()
        stop = CheckConverge(aspen)
        errors = FindErrors(aspen)
        for e in errors:
            print(e)
        errors = ' ; '.join(errors)
        
        if stop:
            writer = pd.ExcelWriter(output_file_name)
            dfstreams.to_excel(writer,'Sheet1')
            writer.save()
            return dfstreams
        
        column = [x for x in book.Sheets('Aspen_Streams').Evaluate("D1:D100") if x.Value != None] 
        
        if obj.FindNode(column[0]) == None:
                print('ERROR in Trial Number '+ str(trial))
                continue

        stream_values = []

        for index,stream in enumerate(column):
            stream_value = obj.FindNode(stream).Value
            stream_values.append((stream_value,))
        
        cell_string = "C1:C" + str(len(column))
        book.Sheets('ASPEN_Streams').Evaluate(cell_string).Value = stream_values
 
        excel.Calculate()
        excel.Run('SOLVE_DCFROR')
        
        dfstreams.loc[trial] = case_values + [x.Value for x in book.Sheets('Output').Evaluate("C3:C15")] + [errors]
        if disp_graphs:
            GUI.plot_on_GUI(dfstreams, vars_to_change)
        
        ######### KEEP TRACK OF RUN TIME PER TRIAL ########
        print('Elapsed Time: ', time() - old_time)
        old_time = time()
        elapsed_time = start_time - time()
        time_remaining = (num_trials - trial - 1)*(elapsed_time / (trial + 1))
        GUI.display_time_remaining(time_remaining)
        aspen.Engine.ConnectionDialog()
        #aspen.Close()
        #aspen.Quit()
        
        ############### CHECK TO SEE IF USER WANTS TO ABORT ##########
        #abort = GUI.check_abort()
        #if abort:
        #    break
        
    writer = pd.ExcelWriter(output_file_name + '.xlsx')
    dfstreams.to_excel(writer, sheet_name ='MFSP')
    stats = dfstreams['MFSP'].describe()
    stats.to_excel(writer, sheet_name = 'Summary Stats')
    writer.save()
    
    if disp_graphs:
        plt.savefig(output_file_name + '.png')
        plt.show()
        
    close_aspen_instances()
    print("-----------FINISHED-----------")
    return dfstreams

def close_aspen_instances():
    for p in psutil.process_iter():
        if p.name() == 'AspenPlus.exe':
            p.terminate()

def univariate_analysis(aspenfilename, excelfilename, aspencall, aspen_var_name, values, fortran_index, output_file_name):
    '''
    THIS FUNCTION ONLY NEEDS TO BE RUN ONCE
    
    Function fills a dataframe with information needed to perform
    a monte carlo simulation on profitability.
    This function interfaces with an ASPEN file for an
    integrated biorefinery and the NREL TEA file. 
    
    Inputs:
        aspenfilename: string
        excelfilename: string
    Outputs:
        dfstreams
            index is the SA fractionalization
            columns hold info from the TEA calcs
        ***function also outputs an excel file with the same info 
        in the dataframe
    '''
    
    aspen,obj,excel,book = open_COMS(aspenfilename,excelfilename)
    v = aspen_var_name
    
    
    
    columns= ['Biofuel Output', 'Succinic Acid Output', 'Fixed Op Costs',\
              'Var OpCosts', ' Capital Costs', 'MFSP','Fixed Capital Investment',\
              'Capital Investment with Interest','Loan Payment per Year','Depreciation','Cash on Hand',\
              'Steam Plant Value','Bag Cost']
    
    dfstreams = pd.DataFrame(columns=columns)
    
    SUC_LOC = r"\Data\Blocks\A300\Data\Blocks\B1\Input\FRAC\TOC5"
    obj.FindNode(SUC_LOC).Value = 0.4
    
    old_time = time()
    start_time = time()
    trial_counter = 1
    counter = 1
    for case in values:
        print(v + " = " + str(case))
        obj.FindNode(aspencall).Value = case
        
        aspen.Reinit()
        aspen.Engine.Run2()
        stop = CheckConverge(aspen)
        errors = FindErrors(aspen)
        for e in errors:
            print(e)
        
        if stop:
            writer = pd.ExcelWriter('3-7-2018_df_final.xlsx')
            dfstreams.to_excel(writer,'Sheet1')
            writer.save()
            return dfstreams

        column = [x for x in book.Sheets('Aspen_Streams').Evaluate("D1:D100") if x.Value != None] 
        
        if obj.FindNode(column[0]) == None:
                print('ERROR in Aspen for fraction '+ str(case))
                continue
        stream_values = []

        for index,stream in enumerate(column):
            stream_value = obj.FindNode(stream).Value   
            stream_values.append((stream_value,))
        
        cell_string = "C1:C" + str(len(column))
        book.Sheets('ASPEN_Streams').Evaluate(cell_string).Value = stream_values
 
        excel.Calculate()
        excel.Run('SOLVE_DCFROR')
        
        if type(case) == str:
            case = float(case[fortran_index[0]:fortran_index[1]])
        
        dfstreams.loc[case] = [x.Value for x in book.Sheets('Output').Evaluate("C3:C15")]
        print('Elapsed Time: ', time() - old_time)
        old_time = time()
        elapsed_time = start_time - time()
        time_remaining = (len(values) - trial_counter)*(elapsed_time / (trial_counter))
        GUI.display_time_remaining(time_remaining)
        trial_counter += 1
        GUI.plot_univ_on_GUI(dfstreams, v, counter, type(values[0]) == str)
        counter += 1
        aspen.Engine.ConnectionDialog()
        #aspen.Close()
        #aspen.Quit()
        
        
        ############### CHECK TO SEE IF USER WANTS TO ABORT ##########
        #abort = GUI.check_next_analysis()
        #if abort:
        #    break
    
    
    writer = pd.ExcelWriter(output_file_name + '_' + v + '.xlsx')
    dfstreams.to_excel(writer,'Sheet1')
    writer.save()
    close_aspen_instances()
    
    return dfstreams

def FindErrors(aspen):
    obj = aspen.Tree
    error = r'\Data\Results Summary\Run-Status\Output\PER_ERROR'
    not_done = True
    counter = 1
    error_number = 0
    error_statements = []
    while not_done:
        try:
            check_for_errors = obj.FindNode(error + '\\' +  str(counter)).Value
            if "error" in check_for_errors.lower():
                error_statements.append(check_for_errors)
                scan_errors = True
                counter += 1
                while scan_errors:
                    if len(obj.FindNode(error + '\\' + str(counter)).Value.lower()) > 0:
                        error_statements[error_number] = error_statements[error_number] + obj.FindNode(error + '\\' + str(counter)).Value
                        counter += 1
                    else:
                        scan_errors = False
                        error_number += 1
                        counter += 1
            else:
                counter += 1
        except:
            not_done = False
    return error_statements

def CheckConverge(aspen):
    
    obj = aspen.Tree
    error = r'\Data\Blocks\REFINE\Data\Blocks\FRAC\Output\PER_ERROR\1'
    stage = r'\Data\Blocks\REFINE\Data\Blocks\FRAC\Input\NSTAGE'
    fracstm = r'\Data\Blocks\REFINE\Data\Blocks\FRAC\Input\FEED_STAGE\FRACSTM'
    fracfd = r'\Data\Blocks\REFINE\Data\Blocks\FRAC\Input\FEED_STAGE\FRACFD' 
    stm_stage = r'\Data\Blocks\REFINE\Data\Blocks\FRAC\Input\FEED_CONVEN\FRACSTM'
    #fd_stage = r'\Data\Blocks\REFINE\Data\Blocks\FRAC\Input\FEED_CONVEN\FRACFD'
    nstage = obj.FindNode(stage)
    
    #init_stage = obj.FindNode(stage).Value
    #init_fracstm = obj.FindNode(fracstm).Value
    #init_fracfd = obj.FindNode(fracfd).Value
    #init_stm_stage = obj.FindNode(stm_stage).Value
    
    while obj.FindNode(error) != None:
        
        nstage = obj.FindNode(stage)
        
        obj.FindNode(stm_stage).Value = "ABOVE-STAGE"
        nstage.Value -= 1
        obj.FindNode(fracstm).Value -= 1
        obj.FindNode(stm_stage).Value = "ON-STAGE"
        obj.FindNode(fracfd).Value = ceil(nstage.Value/2)
        
        print('Failed to Converge, Adjusting stages and Feed Stage #')
        print('Number of Stages: ', nstage.Value)
        print('Feed Stage: ', obj.FindNode(fracfd).Value)
        
        if nstage.Value < 2:
            return True
        
        aspen.Reinit()
        aspen.Engine.Run2()
        
    print("Converged with " + str(nstage.Value) + ' stages')
    print('Feed Stage: ', obj.FindNode(fracfd).Value)
    #obj.FindNode(stage).Value = init_stage
    #obj.FindNode(fracstm).Value = init_fracstm
    #obj.FindNode(fracfd).Value = init_fracfd
    #obj.FindNode(stm_stage).Value = init_stm_stage
    return False

    
if __name__ == "__main__":
    pass