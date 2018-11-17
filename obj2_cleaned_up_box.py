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
import model_fin as model
import matplotlib.pyplot as plt
from time import time
from math import ceil
import random
import csv
import GUI_multivariate as GUI

aspenfilename =  'BC1508F-BC_FY17Target._Final_5ptoC5_updated022618.bkp'
excelfilename = 'DESIGN_OBJ2_test_MFSP-updated.xlsm' 


def open_COMS(aspenfilename, excelfilename):
    
    
    print('Initializing Aspen COM...')
    aspen = win32.Dispatch('Apwn.Document')
    print('Aspen COM Initialized.')
    aspen.InitFromArchive(os.path.abspath(aspenfilename))
    print('Aspen File Open.')
    obj = aspen.Tree
    
    print('Initializing Excel COM...')
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    print('Excel COM Initialized')
    book = excel.Workbooks.Open(os.path.abspath(excelfilename))
    print('Excel File Open')
    
    
    return aspen,obj,excel,book

def get_distributions(gui_excel_input):
    '''
    Given the excel input from the user in the GUI, produce a list_of_variables
    the user wants to change as well as their distributions that should be 
    randomly sampled from. 
    '''
    
    with open(gui_excel_input) as f:
        reader = csv.DictReader(f)# Skip the header row
        gauss_vars = {}
        other_dist_vars = {}
        uniform_vars = {}
        poisson_vars = {}
        pareto_vars = {}
        for row in reader:
            if row['Toggle'].lower().strip() == 'true':
                dist_type = row['Format of Range'].lower()
                aspen_variable = row['Variable Name']
                aspen_call = row['Variable Aspen Call']
                bounds = row['Bounds'].split(',')
                lb = float(bounds[0].strip())
                ub = float(bounds[1].strip())
                if row['Fortran Call'].strip() == "":
                    is_fortran = False
                    fortran_call = None
                    change_index = None
                if row['Fortran Call'].strip() != "":
                    is_fortran = True
                    fortran_call = row['Fortran Call']
                    value_to_change = row['Fortran Value to Change'].strip()
                    len_val = len(value_to_change)
                    change_index = (0,0)
                    for i in range(len(fortran_call)):
                        if fortran_call[i:i+len_val] == value_to_change:
                            change_index = (i, i+len_val) #NOT INCLUSIVE     
                if 'normal' in dist_type or 'gaussian' in dist_type:
                    dist_variables = row['Range of Values'].split(',')
                    gauss_vars[(aspen_variable, aspen_call, (is_fortran, fortran_call, change_index))] = (float(dist_variables[0].strip()),
                              float(dist_variables[1].strip()), lb, ub)
                if 'linspace' in dist_type:
                    linspace_vars = row['Range of Values'].split(',')
                    distribution = np.linspace(float(linspace_vars[0].strip()), 
                                               float(linspace_vars[1].strip()),
                                               float(linspace_vars[2].strip()))
                    other_dist_vars[(aspen_variable, aspen_call, (is_fortran, fortran_call, change_index))] = (distribution, lb, ub)
                if 'poisson' in dist_type:
                    lambda_p = float(row['Range of Values'].strip())
                    poisson_vars[(aspen_variable, aspen_call, (is_fortran, fortran_call, change_index))] = (lambda_p, lb, ub)
                if 'pareto' in dist_type:
                    pareto_vals = row['Range of Values'].split(',')
                    shape = float(pareto_vals[0].strip())
                    scale = float(pareto_vals[1].strip())
                    pareto_vars[(aspen_variable, aspen_call, (is_fortran, fortran_call, change_index))] = (shape, scale, lb, ub)
                if 'list' in dist_type:
                    lst = row['Range of Values'].split(',')
                    distribution = []
                    for l in lst:
                        distribution.append(float(l.strip()))
                    other_dist_vars[(aspen_variable, aspen_call, (is_fortran, fortran_call, change_index))] = (distribution, lb, ub)
                    
                if 'uniform' in dist_type:
                    lb_ub = row['Range of Values'].split(',')
                    lb_uniform, ub_uniform = float(lb_ub[0].strip()), float(lb_ub[1].strip())
                    uniform_vars[(aspen_variable, aspen_call, 
                                  (is_fortran, fortran_call, 
                                   change_index))] = (lb_uniform, ub_uniform, lb, ub)
    return gauss_vars, uniform_vars, poisson_vars, pareto_vars, other_dist_vars
    
def sample_gauss(mean, std, lb, ub):
    rand_sample = np.random.normal(mean,std)
    while(rand_sample < lb or rand_sample > ub):
        rand_sample = np.random.normal(mean,std)
    return rand_sample

def sample_uniform(lb_uniform, ub_uniform, lb, ub):
    rand_sample = np.random.uniform(lb_uniform, ub_uniform)
    while(rand_sample < lb or rand_sample > ub):
        rand_sample = np.random.uniform(lb_uniform, ub_uniform)
    return rand_sample

def sample_list(dist, lb, ub):
    rand_sample = random.choice(dist)
    while(rand_sample < lb or rand_sample > ub):
        rand_sample = random.choice(dist)
    return rand_sample

def sample_poisson(lambda_p, lb, ub):
    rand_sample = np.random.poisson(1000*lambda_p)/1000
    while(rand_sample < lb or rand_sample > ub):
        rand_sample = np.random.poisson(1000*lambda_p)/1000
    return rand_sample

def sample_pareto(shape, scale, lb, ub):
    rand_sample = (np.random.pareto(shape) + 1) * scale
    while(rand_sample < lb or rand_sample > ub):
        rand_sample = (np.random.pareto(shape) + 1) * scale
    return rand_sample

def make_fortran(is_fortran, fortran_call, change_index, val):
    return fortran_call[:change_index[0] - 1] + str(val) + fortran_call[change_index[1] + 1:]

def multivariate_sensitivity_analysis(aspenfilename, excelfilename, 
    gui_excel_input, num_trials, output_file_name, graph_plot):
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
              'Steam Plant Value','Bag Cost']
    
    gauss_vars, uniform_vars, poisson_vars, pareto_vars, other_dist_vars = get_distributions(gui_excel_input)
    
    dfstreams = pd.DataFrame(columns=columns)
    obj.FindNode(SUC_LOC).Value = 0.4
    
    ###### RUN SIMULATION   #########
    old_time = time()
    for trial in range(num_trials):
        
        ####### DRAW RANDOMLY FROM GAUSSIAN DIST VARIABLES ########
        for (aspen_variable, aspen_call,  (is_fortran, fortran_call, change_index)), (mean, std, lb, ub) in gauss_vars.items():
            rand_sample = sample_gauss(mean, std, lb, ub)
            if is_fortran:
                modified_var = make_fortran(is_fortran, fortran_call, change_index, rand_sample)
                obj.FindNode(aspen_call).Value = modified_var
            else:
                obj.FindNode(aspen_call).Value = rand_sample
            variable_values[aspen_variable] = rand_sample
            
        ####### DRAW RANDOMLY FROM UNIFORM DIST VARIABLES ########
        for (aspen_variable, aspen_call,  (is_fortran, fortran_call, change_index)), (lb_uniform, ub_uniform, lb, ub) in uniform_vars.items():
            rand_sample = sample_uniform(lb_uniform, ub_uniform, lb, ub)
            if is_fortran:
                modified_var = make_fortran(is_fortran, fortran_call, change_index, rand_sample)
                obj.FindNode(aspen_call).Value = modified_var
            else:
                obj.FindNode(aspen_call).Value = rand_sample
            variable_values[aspen_variable] = rand_sample
            
        ####### DRAW RANDOMLY FROM OTHER VARIABLE DISTRIBUTIONS ##########
        for (aspen_variable, aspen_call,  (is_fortran, fortran_call, change_index)), (dist, lb, ub) in other_dist_vars.items():
            rand_sample = sample_list(dist, lb, ub)
            if is_fortran:
                modified_var = make_fortran(is_fortran, fortran_call, change_index, rand_sample)
                obj.FindNode(aspen_call).Value = modified_var
            else:
                obj.FindNode(aspen_call).Value = rand_sample
            variable_values[aspen_variable] = rand_sample
            
        ############ DRAW RANDOMLY FROM PARETO DISTRIBUTIONS ###########
        for (aspen_variable, aspen_call,  (is_fortran, fortran_call, change_index)), (shape, scale, lb, ub) in pareto_vars.items():
            rand_sample = sample_pareto(shape, scale, lb, ub)
            if is_fortran:
                modified_var = make_fortran(is_fortran, fortran_call, change_index, rand_sample)
                obj.FindNode(aspen_call).Value = modified_var
            else:
                obj.FindNode(aspen_call).Value = rand_sample
            variable_values[aspen_variable] = rand_sample
            
        ########### DRAW RANDOMLY FROM POISSON DISTRIBUTIONS ###########
        for (aspen_variable, aspen_call,  (is_fortran, fortran_call, change_index)), (lambda_p, lb, ub) in poisson_vars.items():
            rand_sample = sample_poisson(lambda_p, lb, ub)
            if is_fortran:
                modified_var = make_fortran(is_fortran, fortran_call, change_index, rand_sample)
                obj.FindNode(aspen_call).Value = modified_var
            else:
                obj.FindNode(aspen_call).Value = rand_sample
            variable_values[aspen_variable] = rand_sample
        
        ########## STORE THE RANDOMLY SAMPLED VARIABLE VALUES  ##########
        case_values = []
        for v in vars_to_change:
            case_values.append(variable_values[v])
            
        
        ######## RUN ASPEN SIMULATION WITH RANDOMLY SAMPLED VARIABLES #######
        aspen.Reinit()
        aspen.Engine.Run2()
        stop = CheckConverge(aspen)
        errors = FindErrors(aspen)
        print(errors)
        for e in errors:
            print(e)
        
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
        
        dfstreams.loc[trial] = case_values + [x.Value for x in book.Sheets('Output').Evaluate("C3:C15")]
        if graph_plot == 1:
            GUI.plot_on_GUI(dfstreams)
        
        ######### KEEP TRACK OF RUN TIME PER TRIAL ########
        print(time() - old_time)
        old_time = time()
    
    writer = pd.ExcelWriter(output_file_name + '.xlsx')
    dfstreams.to_excel(writer,'Sheet1')
    writer.save()
    
    if graph_plot == 1:
        total_MFSP = dfstreams["MFSP"]
    
        num_bins = 100
        n, bins, patches = plt.hist(total_MFSP, num_bins, facecolor='blue', alpha=0.5)
        plt.xlabel('MFSP Price ($)')
        plt.ylabel('Count of simulations')
        plt.title('Historgram of MFSP prices based on simulations')
        plt.savefig(output_file_name + '.png')
        plt.show()
    
    aspen.Close()
    print("-----------FINISHED-----------")
    return dfstreams
        
        

def univariate_analysis(aspenfilename, excelfilename, aspencall, aspen_var_name, values, output_file_name):
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
    
    
    
    columns= ['Biofuel Output', 'Succinic Acid Output', 'Fixed Op Costs',\
              'Var OpCosts', ' Capital Costs', 'MFSP','Fixed Capital Investment',\
              'Capital Investment with Interest','Loan Payment per Year','Depreciation','Cash on Hand',\
              'Steam Plant Value','Bag Cost']
    
    dfstreams = pd.DataFrame(columns=columns)
    
   

    #succ_fracs = np.linspace(0,.5,51)
    #succ_fracs = [25]
    SUC_LOC = r"\Data\Blocks\A300\Data\Blocks\B1\Input\FRAC\TOC5"
    obj.FindNode(SUC_LOC).Value = 0.4
    
    for case in values:
        
        print("variable value: " +str(case))
        print(time() - old_time)
        old_time = time()
        #succ_frac = case
        obj.FindNode(aspencall).Value = case
        
        #stream splitting
        #obj.FindNode(SUC_LOC).Value = succ_frac
        
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
            #print(stream,obj.FindNode(stream))
            


            stream_value = obj.FindNode(stream).Value
            
            stream_values.append((stream_value,))
        
        cell_string = "C1:C" + str(len(column))
        book.Sheets('ASPEN_Streams').Evaluate(cell_string).Value = stream_values
 
        excel.Calculate()
        excel.Run('SOLVE_DCFROR')
        
        dfstreams.loc[case] = [x.Value for x in book.Sheets('Output').Evaluate("C3:C15")]
    
    writer = pd.ExcelWriter(output_file_name + 'xlsx')
    dfstreams.to_excel(writer,'Sheet1')
    writer.save()
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
    
    while obj.FindNode(error) != None:
        
        nstage = obj.FindNode(stage)
        
        obj.FindNode(stm_stage).Value = "ABOVE-STAGE"
        nstage.Value -= 1
        obj.FindNode(fracstm).Value -= 1
        obj.FindNode(stm_stage).Value = "ON-STAGE"
        obj.FindNode(fracfd).Value = ceil(nstage.Value/2)
        
        print(nstage.Value)
        print(obj.FindNode(fracfd).Value)
        
        if nstage.Value < 2:
            return True
        
        aspen.Reinit()
        aspen.Engine.Run2()
        
    print("Converged: " + str(nstage.Value))
    print(obj.FindNode(fracfd).Value)
    return False

def get_price_preds(file,data_col):
    '''
    Interface with price predictor model to fill an array with SA and biodiesel
    prices over a future period.
    
    Inputs:
    [str]file: .csv filename with oil predictions
    [int]data_col: column containing data (barrels)

    Outputs:
    [np.array]price preds:
        price_preds[0] = biofuel price
        price_preds[1] = SA price
    '''
    BARRELL_TO_MT = 7.33
    #predict the biofuel prices
    ind = model.upload(file,data_col)
    ind = model.gen_linspace(ind,span = 'year')
    
    #predict the biofuel prices
    biofuel_model = get_model(ind,"biofuel")

    ind *= BARRELL_TO_MT
    
    MA_model = get_model(ind,"ma")
    SA_model = get_model(MA_model,"sa")

    price_preds = np.array((biofuel_model, SA_model))
    
    return price_preds

def get_model(ind,dep_name):
    """
    Call model.predict with params 
    matching input name
    
    Inputs:
    [np.array]ind: independent data
    [str]dep_name: name of dependent variable

    Outputs:
    [np.array]model: model of dependent data
    """
    
    if dep_name == "biofuel":
        abt = (1.0653, 69.8492, 1)
        mu = -0.04769
        std = 21.6212
        
    if dep_name == "ma":
        abt = (1.378, 648.5, 3)
        mu = 0.0002883
        std = 114.8
        
    if dep_name == "sa":    
        abt = (0.7465, 1050, 3)
        mu = -1.619
        std = 144.7

    preds = model.predict(abt, ind, mu, std)
    
    return preds

def get_case_profit(case_index, dfstreams, price_preds):
    '''
    Calculates the total profit for a refinery over the 
    number of months in the price_preds array.
    Does not factor in capital costs
    
    Inputs:
        case_index: integer
        dfstreams: data frame with TEA and ASPEN outputs
        price_preds: 2D array with Biofuel prices and SA prices
            on a monthly basis
    Outputs:
        total_profit: float
    '''
    streams_for_case = dfstreams.iloc[case_index]
    price_preds[1] /= 1000
    total_profit = \
        streams_for_case['Biofuel Output']*price_preds[0].sum() + \
        streams_for_case['Succinic Acid Output']*(price_preds[1]).sum() +\
        streams_for_case['Var OpCosts']*len(price_preds) +\
        streams_for_case['Fixed Op Costs']*len(price_preds)
        
    return total_profit

    
#The following functions relate to monte carlo analysis
    #and visualization
def monte_carlo(dfstreams,num_lives,file = 'oil.csv',data_col = 1):
    '''
    For each variation of fractionalization stored in dfstreams
    this function will calculate the profitability over a time period 
    many different times. In each simulation for each case, the price 
    predictions will be varied.
    Inputs: 
        dfstreams
        num_lives
        file
        data_col
    Outputs:
        case_profits: 2D array
            case_profits[0] is the SA fractionalization
            case_profits[1] list of potential profits associated 
                that fractionalization
    '''
    case_profits = []
    for case in range(len(dfstreams.index)):
        
        profits = []
        for lifetime in range(num_lives):
            
            price_preds = get_price_preds(file,data_col)
            profit = get_case_profit(case, dfstreams, price_preds)
            profits.append(profit)
           
        case_profits.append((dfstreams.iloc[case].name, profits))
        
    return case_profits

def plot_histograms(case_profits,number = None):
    '''
    Used to visualize the outputs of the monte carlo simulation
    Inputs: 
        case_profits: 2D array
        number: integer, the number of cases distributions
            to display
    '''
    ax = plt.axes()
    
    if number == 1:
        step = len(case_profits[0][1])
        for case,dist in case_profits[::step]:
            
            case_name = "SA%: " + str(case)
            ax.hist(dist,alpha = .5,label = case_name)
            
    elif number == None:
        for case,dist in case_profits[::1]:
            
            case_name = "SA%: " + str(case)
            ax.hist(dist,alpha = .5,label = case_name)
            
    else:
        step = round(len(case_profits)/number)
        for case,dist in case_profits[::step]:
            
            case_name = "SA%: " + str(case)
            ax.hist(dist,alpha = .5,label = case_name)
    
    ax.set_xlabel("Profit")
    ax.set_ylabel("Frequency")
    ax.set_title("Profitability Distributions")
    ax.legend(loc = 0, fontsize = 'xx-small')
    plt.show()
    
    
if __name__ == "__main__":
    pass