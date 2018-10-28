# -*- coding: utf-8 -*-
"""
Created on Thu Mar  8 18:50:13 2018

@author: owner
"""

# -*- coding: utf-8 -*-
"""
Created on Thu Mar  1 11:05:27 2018

@author: Group D
This file assumes that a dataframe has already been calculated. The dataframe contains
information from an aspen simulation. 

Functionality to output price predictions for biofuel and SA. Requires an input of
a csv file that has historical gas prices in 2014 dollar terms ($/gallon). 
Requires an xlsx file with crude oil prices in 2014 dollar terms ($/barrel).
"""

import pandas as pd
import numpy as np
import model as mod
import matplotlib.pyplot as plt
#from time import time
#from math import ceil


def get_price_preds(file_oil = "hist_oil.csv",data_col = 3, a_error = 0,randomness = True):
    '''
    Interface with price predictor model to fill an array with SA and biodiesel
    prices over a future period.
    
    Inputs:
    [str]file: .csv filename with oil predictions
    filename is 'hist_oil.csv" with the prices in 2014 dollar terms ($/barrel)
    data_column = 3
    [int]data_col: column containing data (barrels)

    Outputs:
    [np.array]price preds:
        price_preds[0] = biofuel price
        price_preds[1] = SA price
    '''
    #read in gas prices
    oildf = pd.read_excel('hist_gas.xlsx')
    #['Month', 'CPI', 'Oil (Real)', 'Oil (Nominal 01/2014 CPI Basis)']
    biofuel_prices = oildf['gas_nominal'].tolist()
    
    
    BARRELL_TO_MT = 7.33 
    ind = mod.upload(file_oil,data_col)
    ind = mod.gen_linspace(ind,span = 'month')
    ind *= BARRELL_TO_MT

    MA_model = get_model(ind,"ma", a_error,randomness)
    SA_model = get_model(MA_model,"sa", a_error,randomness)
    SA_model /=1000
    
    price_preds = np.array((biofuel_prices, SA_model))
    #price_preds = SA_model
    return price_preds

def get_model(ind,dep_name, a_error = 0,randomness = True):
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
        abt = [1.0653, 69.8492, 1]
        abt[0] += a_error
        mu = -0.04769
        std = 21.6212
        
    if dep_name == "ma":
        abt = [1.133, 700, 2]
        abt[0] += a_error
        mu = -0.001
        std = 116.9

        preds = mod.predict(abt, ind, mu, std,randomness = randomness,ind_dep = "oil_ma")
        
    if dep_name == "sa":    
        abt = [0.7683, 840, 2]
        abt[0] += a_error
        mu = -2.137
        std = 114.51 
        preds = mod.predict(abt, ind, mu, std,randomness = randomness,ind_dep = "ma_sa")

    return preds

def get_monthly_profit(streams_for_case, bio_price, succ_price):
    '''
    calculates the net revenue in a given month
    '''
    
    month_profit = \
        streams_for_case['Biofuel Output']*bio_price + \
        streams_for_case['Succinic Acid Output']*succ_price +\
        streams_for_case['Var OpCosts'] +\
        streams_for_case['Fixed Op Costs']
   
    return month_profit


def monthly_to_yearly(price_preds):
    """
    Takes montly predictions for (bio,sa) and averages them
    to get yearly predictions 
    """    
    predictions = []
    bio_prices,SA_prices = price_preds
    i = 0
    j = 0
    #print(bio_prices)
    while i < len(bio_prices) and j < len(SA_prices):
         #print(bio_prices[i:i+12])
         #print(np.mean(bio_prices[i:i+12]))
         yearly_bio = np.mean(bio_prices[i:i+12])
         yearly_sa = np.mean(SA_prices[i:i+12])
    
         predictions.append((yearly_bio,yearly_sa))
         i+=12
            
    return predictions



###everything below this line is used for outputting the results automatically
#oilfilename = 'hist_oil.csv'
#datacolumn = 3
#streamsfilename = 'trial_lifetime_run2.xlsx'
#dfstreams = pd.read_excel(streamsfilename, 'trial_lifetime_run2')

#profits, profits_high, profits_low = possible_profits(dfstreams,oilfilename, datacolumn)

#generating historgrams
#plot with 3 fractions
#plot_histograms(profits, 2)
#plot with all fractions
#plot_histograms(profits)
#generating average 10 year profits
#plot_histograms(profits, 10)

#generating average 10 year profit with error
if __name__ == "__main__":

    JAN_FIRST_1988 = 144
    DEC_FIRST_2017 = 504
    
############################MA RANDOM WALK GRAPH################################## 
    
    '''
    num_walks = 2
    BARRELL_TO_MT = 7.33
    oil = mod.upload('hist_oil.csv',3)
    oil = mod.gen_linspace(oil,span = 'month')[144:504]
    oil *= BARRELL_TO_MT
    data = np.loadtxt('MA_OIL_PRICE.csv',skiprows = 1,delimiter = ',',usecols = [1,2])
    ma = data[:,0]/1000
    fig =plt.figure(1)
    ax = plt.axes()
    ax.grid()
    years = np.linspace(1988,2017,360)
    ax.plot(years[:-4],oil[4:]/1000,'k-', label ="Historical Oil")
    ax.plot(years[-132:-2],ma[2:],'r-', label = "Historical MA")
    ma_no_randomness = get_model(oil,'ma', randomness = False)
    ax.plot(years,ma_no_randomness/1000, label = "Modeled MA")
    nolabel = 1
    for walk in range(num_walks):
        ma_with_randomness = get_model(oil,'ma')
        if nolabel:
            ax.plot(years,ma_with_randomness/1000,'--',label = "Predictions",alpha = .5)
            nolabel = 0
            
        ax.plot(years,ma_with_randomness/1000,'--',alpha = .5)
    ax.set_xlim([1988,2017])
    ax.legend()
    ax.set_xlabel("Year")
    ax.set_ylabel("Price (USD/kg)")
    #plt.savefig("final_MA_random_walks.png",dpi = 300)
    
    
    JAN_FIRST_2017 = 144
    DEC_FIRST_2017 = 504
    JAN_FIRST_2014 = 312
    
    num_walks = 2
    BARRELL_TO_MT = 7.33
    oil = mod.upload('hist_oil.csv',3)
    oil = mod.gen_linspace(oil,span = 'month')[JAN_FIRST_1988:DEC_FIRST_2017]
    oil *= BARRELL_TO_MT
    #data = np.loadtxt('MA_OIL_PRICE.csv',skiprows = 1,delimiter = ',',usecols = [1,2])
    data2 = np.loadtxt('nom_data2.csv',skiprows = 1,delimiter = ',',usecols = [1,2])
    ma = data[:,0]/1000
    SA = data2[:,1]/1000
    
    fig =plt.figure(1)
    ax = plt.axes()
    ax.grid()
    
    ma_with_randomness = get_model(oil,'ma')
    years = np.linspace(1988,2017,360)
    ax.plot(years[:-2],ma_with_randomness[2:]/1000,'k-', label ="Predicted MA")
    ax.plot(years[317:350],SA[2:],'r-', label = "Historical SA")
    sa_no_randomness = get_model(ma_with_randomness,'sa', randomness = False)
    
    std_to_match = SA[2:] - sa_no_randomness[311:346-2]
    
    
    ax.plot(years,sa_no_randomness/1000, label = "Modeled SA")
    nolabel = 1
    for walk in range(num_walks): 
        sa_with_randomness = get_model(ma_with_randomness,'sa')
        if nolabel:
            ax.plot(years[2:],sa_with_randomness[:-2]/1000,'--',label = "SA Predictions",alpha = .5)
            nolabel = 0
            
        ax.plot(years,sa_with_randomness/1000,'--',alpha = .5)
    ax.set_xlim([1988,2017])
    ax.legend()
    ax.set_xlabel("Year")
    ax.set_ylabel("Price (USD/kg)")
    plt.savefig("final_SA_random_walks.png",dpi = 300)
    '''
    
#######################GENERATE NVP FROM RANDOM SA WALKS##########################################
    '''
    oilfilename = 'hist_oil.csv'
    #oilfilename = 'future_oil.csv'
    streamsfilename = '3-7-2018_df_final.xlsx'
    dfstreams = pd.read_excel(streamsfilename, 'Sheet1')
    N = 1000
    lifetime_predicts = []
    
    for lifetime in range(N):
        yearly_price_preds =  monthly_to_yearly(get_price_preds(oilfilename,3)[:,JAN_FIRST_1988:DEC_FIRST_2017])
        lifetime_predicts.append(yearly_price_preds)

    
    
    apv_by_case, case_npv = get_profits_by_case(dfstreams, lifetime_predicts)
    avg_NPVs = get_each_case_avg_NPV(apv_by_case)
    #####################automated plotting##########################################
    #autplot.plot_average_pv_30years(avg_NPVs) 
    #autplot.one_case_hist(case_npv, 0.50)       
    frac1,frac2,frac3,frac4 = (0.25,.35,.40,.5)
    lab1,lab2,lab3,lab4 = 0,0,0,0
    fig = plt.figure()
    for case,lifetimes in apv_by_case:
        num_bins = 20
        rwidth = .75
        lifetimes = np.asarray(lifetimes)
        lifetimes/=(10**6)
        lifetime_NPVs = []
        for lifetime in lifetimes:
            lifetime = lifetime.sum()
            lifetime_NPVs.append(lifetime)
        trials = len(lifetime_NPVs)*1.0
        avg_npv = sum(lifetime_NPVs)/trials
            
        if case == frac1:
            ax1 = fig.add_subplot(141)
            
            weights = np.ones_like(lifetime_NPVs)/float(len(lifetime_NPVs))
    
            ax1.hist(lifetime_NPVs,weights = weights,bins = num_bins,rwidth = rwidth)
            
            ax1.plot([avg_npv,avg_npv],[0,1],'r--')
            ax1.plot([0,0],[0,1],'k--')
            print(avg_npv)
            ax1.set_ylabel('Probability')
            ax1.set_ylim([0,.3])
            ax1.set_title('SA Production: '+str(frac1*100) + '%')
    
        if case == frac2:
            ax2 = fig.add_subplot(142)
            
            #ax2.grid()
            weights = np.ones_like(lifetime_NPVs)/float(len(lifetime_NPVs))
    
            ax2.hist(lifetime_NPVs,weights = weights,bins = num_bins,rwidth = rwidth)
            ax2.plot([avg_npv,avg_npv],[0,1],'r--')
            ax2.plot([0,0],[0,1],'k--')
            print(avg_npv)
            ax2.set_ylim([0,.3])
            ax2.set_xlabel('                                                 Lifetime NPV')
            ax2.set_title('SA Production: '+str(frac2*100) + '%')
                
        if case == frac3:
            ax3 = fig.add_subplot(143)
            
            #ax3.grid()
            
            weights = np.ones_like(lifetime_NPVs)/float(len(lifetime_NPVs))
    
            ax3.hist(lifetime_NPVs,weights = weights,bins = num_bins,rwidth = rwidth)
            ax3.plot([avg_npv,avg_npv],[0,1],'r--')
            ax3.plot([0,0],[0,1],'k--')
            print(avg_npv)
            #ax3.set_ylabel('Probability')
            #ax3.set_xlabel('Lifetime NPV')
            ax3.set_ylim([0,.3])
            ax3.set_title('SA Production: '+str(frac3*100) + '%')
            
    
             
        if case == frac4:
            ax4 = fig.add_subplot(144)
            
            weights = np.ones_like(lifetime_NPVs)/float(len(lifetime_NPVs))
    
            ax4.hist(lifetime_NPVs,weights = weights,bins = num_bins,rwidth = rwidth)
            ax4.plot([avg_npv,avg_npv],[0,1],'r--')
            ax4.plot([0,0],[0,1],'k--')
            print(avg_npv)
            ax4.set_ylim([0,.3])
            #ax4.set_xlabel('Lifetime NPV')
            ax4.set_title('SA Production: '+str(frac4*100) + '%')
    ax1.tight_layout()
    ax2.tight_layout()
    ax3.tight_layout()
    ax4.tight_layout()
    
    plt.savefig("final_random_walk_npvs.png",dpi =300)
    
    
#all apvs by case
    frac1,frac2,frac3,frac4 = (0.25,.35,.4,.5)
    trials = [5,20,35,40]
    lab1,lab2,lab3,lab4 = 0,0,0,0
    for case,lifetimes in apv_by_case:
        lifetimes = np.asarray(lifetimes)
        lifetimes/=(10**6)
        print(case)
        print(lifetimes)
        average_lifetime = np.mean(lifetimes,0)
        print(average_lifetime)
        print(len(average_lifetime))
        print()
        
        if case == frac1:
            figA = plt.subplot(221)
            ax = figA.axes
            
            for trial,lifetime in enumerate(lifetimes):
                if not lab1 and trial == 5:
                    #ax.plot(lifetime,label = str(frac1))
                    lab1 = 1
                ax.set_title("SA Production: " + str(frac1*100) + "%")
                ax.plot(average_lifetime,'b-')
                ax.set_ylabel('APV (2014 USD, mil.)')
                ax.grid()
                ax.legend()
                ax.set_xlim([0,29])
                
        if case == frac2:
            figB = plt.subplot(222)
            ax = figB.axes
            
            for trial,lifetime in enumerate(lifetimes):
                if not lab2 and trial == 5:
                    #ax.plot(lifetime,label = str(frac2))
                    lab1 = 1
                ax.set_title("SA Production: " + str(frac2*100) + "%")
                ax.plot(average_lifetime,'b-')
                ax.grid()
                ax.legend()
                ax.set_xlim([0,29])
                
        if case == frac3:
            figC = plt.subplot(223)
            ax = figC.axes
            
            for trial,lifetime in enumerate(lifetimes):
                if not lab3 and trial == 5:
                    #ax.plot(lifetime,label = str(frac3))
                    lab1 = 1
                    
                ax.set_title("SA Production: " + str(frac3*100) + "%")
                ax.plot(average_lifetime,'b-')
                ax.set_ylabel('APV (2014 USD, mil.)')
                ax.set_xlabel('Year')
                ax.grid()
                ax.legend()
                ax.set_xlim([0,29])
        
             
        if case == frac4:
            figD = plt.subplot(224)
            ax = figD.axes
            
            for trial,lifetime in enumerate(lifetimes):
                if not lab4 and trial == 5:
                    #ax.plot(lifetime,label = str(frac4))
                    lab4 = 1
                    
                ax.set_title("SA Production: " + str(frac4*100) + "%")
                ax.plot(average_lifetime,'b-')
                ax.set_xlabel('Year')
                ax.grid()
                ax.legend()
                ax.set_xlim([0,29])
            
 ##############################CHECK RANDOM WALKS################################   
    
 
    BARRELL_TO_MT = 7.33 
    oil = mod.upload('hist_oil.csv',3)
    oil = mod.gen_linspace(oil,span = 'month')
    oil *= BARRELL_TO_MT

    
    ##CHECK RANDOM WALK AGAINST REAL MA DATA
    #1988-2017 = [264:504]
    #2014-1016 = oildf[456 -144:492 - 144]
    
    data = np.loadtxt('nominalized_data.csv',skiprows = 1,delimiter = ',',usecols = [1,2])
    ma_no_randomness = get_model(oil,'ma', randomness = False) 
    ma = data[:,1]
    plt.plot(ma_no_randomness)
    i = 0
    N= 1000
    resids = []
    while i < N:
        ma_with_randomness = get_model(oil,'ma')
        #print(ma[2:] - ma_with_randomness[372:502])
        resids.append(ma_no_randomness[372:504] - ma_with_randomness[372:504])
        plt.plot(ma_with_randomness,'--')
        i+=1
    resids_dist = list(itertools.chain.from_iterable(resids))
    #plt.hist(resids_dist)
    print(np.std(resids_dist))
    print(np.mean(resids_dist))
    
    
    ##CHECK RANDOM WALK AGAINST REAL SA DATA
    ma_with_randomness = get_model(oil,'ma')
    sa_no_randomness = get_model(ma_with_randomness,'sa',randomness = False)
    i = 0
    N= 1000
    resids = []
    plt.plot(ma_with_randomness)
    plt.plot(sa_no_randomness)
    while i < N:
        sa_with_randomness = get_model(ma_with_randomness,'sa')
        #print(ma - ma_with_randomness[372:504])
        resids.append(sa_no_randomness[456:492] - sa_with_randomness[456:492])
        plt.plot(sa_with_randomness,'--',alpha = .5)
        i+=1
    resids_dist = list(itertools.chain.from_iterable(resids))
    #plt.hist(resids_dist)
    print(np.std(resids_dist))
    print(np.mean(resids_dist))
    '''
#^^^^^^^^^^^^^^^^^^^^^^^^^^CHECK RANDOM WALKS^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^#
##############################check_default_random_walks########################
    '''
    N = 1000
    lifetime_predicts = []
    for lifetime in range(N):
        yearly_price_preds =  monthly_to_yearly(get_price_preds(oilfilename,3)[:,JAN_FIRST_1988:DEC_FIRST_2017])
        lifetime_predicts.append(yearly_price_preds)
    lifetime_predicts = [lifetime[-10:] for lifetime in lifetime_predicts]
    profits_by_case = get_profits_by_case(dfstreams,lifetime_predicts)
    plot_success_probs(dfstreams,profits_by_case)
    '''
#########################  AVG_NPV RANDOM WALKS ################################  
'''
    oilfilename = 'hist_oil.csv'
    #oilfilename = 'future_oil.csv'
    streamsfilename = '3-7-2018_df_final.xlsx'
    dfstreams = pd.read_excel(streamsfilename, 'Sheet1')
    N = 1000
    lifetime_predicts = []
    
    for lifetime in range(N):
        yearly_price_preds =  monthly_to_yearly(get_price_preds(oilfilename,3)[:,JAN_FIRST_1988:DEC_FIRST_2017])
        lifetime_predicts.append(yearly_price_preds)

    fig = plt.figure()
    for case,lifetimes in apv_by_case:
        num_bins = 20
        rwidth = .75
        lifetimes = np.asarray(lifetimes)
        lifetimes/=(10**6)
        lifetime_NPVs = []
        for lifetime in lifetimes:
            lifetime = lifetime.sum()
            lifetime_NPVs.append(lifetime)
        trials = len(lifetime_NPVs)*1.0
        avg_npv = sum(lifetime_NPVs)/trials
        ax = plt.axes()
        ax.plot(case*100,avg_npv,'b.')
        ax.plot([0,50],[0,0],'k-')
        ax.set_xlim([0,50])
        ax.set_ylabel("Average 30-Year Lifetime NPV (mil. USD)")
        ax.set_xlabel("SA Production (%)")
        plt.savefig("real_avg_npv_rand_walks.png")
        
        



######################## PROFITABILITY INDEX ######################################  
    oilfilename = 'hist_oil.csv'
    #oilfilename = 'future_oil.csv'
    streamsfilename = '3-7-2018_df_final.xlsx'
    dfstreams = pd.read_excel(streamsfilename, 'Sheet1')
    N = 1000
    lifetime_predicts = []
    for lifetime in range(N):
        yearly_price_preds =  monthly_to_yearly(get_price_preds(oilfilename,3)[:,JAN_FIRST_1988:DEC_FIRST_2017])
        lifetime_predicts.append(yearly_price_preds)

    fig = plt.figure()
    for case,lifetimes in apv_by_case:
        num_bins = 20
        rwidth = .75
        lifetimes = np.asarray(lifetimes)
        lifetimes/=(10**6)
        lifetime_NPVs = []
        for lifetime in lifetimes:
            lifetime = lifetime.sum()
            lifetime_NPVs.append(lifetime)
        trials = len(lifetime_NPVs)*1.0
        avg_npv = sum(lifetime_NPVs)/trials
        PI = avg_npv/(dfstreams.loc[case]["Fixed Capital Investment"]/(10**6)) + 1
        ax = plt.axes()
        ax.plot(case*100,PI,'b.')
        ax.plot([0,50],[1,1],'k-')
        ax.set_xlabel("SA Production (%)")
        ax.set_ylabel("Profit Index")
        ax.set_xlim(0,50)
        plt.savefig("real_Profit_Index.png",dpi = 300)
'''            
            
        
    