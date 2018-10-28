# -*- coding: utf-8 -*-
"""
Created on Fri Mar  2 11:11:40 2018

@author: owner
"""

"""
NOTE: This file is command-line executable. Unfortunately, the plot it outputs
is not very informative, and almost certainly wrong.
"""




import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import historical_prices_plant_analysis as hist
import projected_prices_plant_analysis as proj



def get_success_prob(case, cash_on_hand, loan):
    
    frac = case[0]
    income_series_list = case[1]
    trials = len(income_series_list)
    
    failures = 0
    
    for income_series in income_series_list:
        for income in income_series:
            if income > loan:
                continue
            
            else:
              cash_on_hand -= loan - income
              
              if cash_on_hand <= 0:
                  failures += 1
                  cash_on_hand = dfstreams.loc[frac]["Cash on Hand"]
                  break
    

    return (frac,failures)

def plot_success_probs(dfstreams,profits_by_case,plot = True):
    

    
    success_tally = []
    
    for case in profits_by_case:
        
        frac = case[0]
        trials = len(case[1])
        print(trials)
        #print(frac)
        cash_on_hand = dfstreams.loc[frac]["Cash on Hand"]
        #print(cash_on_hand)
        loan = dfstreams.loc[frac]["Loan Payment per Year"]
        #print(dfstreams.loc[frac]["Loan Payment per Year"])
        frac, failures = get_success_prob(case,cash_on_hand,loan)
        success_tally.append((frac,failures/trials))
    
    if plot:
        fig = plt.figure()
        ax = plt.axes()
        frac,success = zip(*success_tally)
        for life,trial in enumerate(success):
            if (trial >=.5):
                if life == 0:
                    ax.plot(frac[life]*100,trial,'r^',label = "P(success)< 50%")
                ax.plot(frac[life]*100,trial,'r^')
            else:
                if life == 48:
                    ax.plot(frac[life]*100,trial,'g^',label = "P(success)> 50%")
                ax.plot(frac[life]*100,trial,'g^')
            ax.set_xlim([0,50])
            ax.set_ylim([0,1])
            print('here')
        ax.set_ylabel('Probability of Loan Default')
        ax.set_xlabel('Succinic Acid Production (%)')
        #ax.set_title('10-year Probability of Loan Defaulting')
        #ax.set_xlim([0,50])
        #ax.set_ylim([-.05,1.05])
        ax.legend(loc = 3)
        ax.grid(True)
        plt.show()
    return success_tally
    
def get_yearly_predictions(price_preds):
    """
    Takes montly predictions for (bio,sa) and averages them
    to get yearly predictions 
    """    
    predictions = []
    bio_prices,SA_prices = price_preds
    i = 0
    j = 0
    while i < len(bio_prices[:-12]) and j < len(SA_prices):

         yearly_bio = np.mean(bio_prices[i:i+12])
         if len(SA_prices) == len(bio_prices):
             yearly_sa = np.mean(SA_prices[i:i+12])
         else:
             yearly_sa = SA_prices[j]
             j+=1
         predictions.append((yearly_bio,yearly_sa))
         i+=12
         
         
    return predictions

def slice_predictions(predictions):
    years = 10
    j = 0
    lifetime_preds = []
    print(len(predictions))
    #print(predictions[j:j+years])
    while len(predictions[j:j+years]) == 10:
        print(predictions[j:j+years])
        print(len(predictions[j:j+years]),j)
        lifetime_preds.append(predictions[j:j+years])
        j+=1
    print(len(lifetime_preds))
    return lifetime_preds   

def get_profits_by_case(dfstreams,lifetime_preds):
    profits_by_case = []
    for case_index in range(len(dfstreams)):  
        streams_for_case = dfstreams.iloc[case_index]
        case_yearly_profits = []
        for lifetime in lifetime_preds:
            lifetime_yearly_profits = []
            for year in lifetime:
                bio_price = year[0] 
                sa_price = year[1]
                lifetime_yearly_profits.append(get_yearly_profit(streams_for_case, bio_price, sa_price))
            case_yearly_profits.append(lifetime_yearly_profits)
        profits_by_case.append((dfstreams.index[case_index],case_yearly_profits))  
    
    return profits_by_case

def get_yearly_profit(streams_for_case, bio_price, succ_price):
    '''
    calculates the net revenue in a given month
    '''
    months_per_year = 12
    profit = \
        streams_for_case['Biofuel Output']*months_per_year*bio_price + \
        streams_for_case['Succinic Acid Output']*months_per_year*succ_price -\
        streams_for_case['Var OpCosts']*months_per_year -\
        streams_for_case['Fixed Op Costs']*months_per_year
   
    return profit

if __name__ == "__main__":
    oilfilename = 'hist_oil.csv'
    #oilfilename = 'future_oil.csv'
    streamsfilename = '3-7-2018_df_final.xlsx'
    dfstreams = pd.read_excel(streamsfilename, 'Sheet1')
    price_preds = proj.get_price_preds(oilfilename,3,randomness= False)
    #price_preds[1]*=1.3
    #price_preds = hist.get_price_preds(oilfilename,1)
    lifetime_preds = slice_predictions(get_yearly_predictions(price_preds))
    
    profits_by_case = get_profits_by_case(dfstreams,lifetime_preds)
    plot_success_probs(dfstreams,profits_by_case)
    
    
    
