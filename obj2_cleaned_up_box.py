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
#import model
import matplotlib.pyplot as plt
from time import time
from math import ceil

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


def fill_streams_dataframe(aspenfilename, excelfilename):
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
    
    SUC_LOC = r"\Data\Blocks\A300\Data\Blocks\B1\Input\FRAC\TOC5"
    
    columns= ['Biofuel Output', 'Succinic Acid Output', 'Fixed Op Costs',\
              'Var OpCosts', ' Capital Costs', 'MFSP','Fixed Capital Investment',\
              'Capital Investment with Interest','Loan Payment per Year','Depreciation','Cash on Hand',\
              'Steam Plant Value','Bag Cost']
    
    dfstreams = pd.DataFrame(columns=columns)

    succ_fracs = np.linspace(0,.5,51)
    old_time = time()
    for case in succ_fracs:
        
        print("succinic acid Fraction: " +str(case))
        print(time() - old_time)
        old_time = time()
        succ_frac = case
        
        #stream splitting
        obj.FindNode(SUC_LOC).Value = succ_frac
        
        aspen.Reinit()
        aspen.Engine.Run2()
        stop = CheckConverge(aspen)
        
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
            print(stream,obj.FindNode(stream))
            


            stream_value = obj.FindNode(stream).Value
            
            stream_values.append((stream_value,))
        
        cell_string = "C1:C" + str(len(column))
        book.Sheets('ASPEN_Streams').Evaluate(cell_string).Value = stream_values
 
        excel.Calculate()
        excel.Run('SOLVE_DCFROR')
        
        dfstreams.loc[case] = [x.Value for x in book.Sheets('Output').Evaluate("C3:C15")]
    
    writer = pd.ExcelWriter('3-7-2018_df_final.xlsx')
    dfstreams.to_excel(writer,'Sheet1')
    writer.save()
    return dfstreams

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