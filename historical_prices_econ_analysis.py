# -*- coding: utf-8 -*-
"""
Created on Mon Mar  5 15:19:07 2018

@author: owner
"""

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
import automated_plotting as autplot
#####################################funcs to run###########################################

def get_profits_by_case(dfstreams,lifetime_preds):
    apv_list_by_case_by_lifetime = []
    net_pv_by_case_by_lifetime = []
    
    for case_index in range(len(dfstreams)):  
        streams_for_case = dfstreams.iloc[case_index]
        lifetime_npv = []
        lifetime_yearly_apv = []
        
        for lifetime in lifetime_preds:
            yearly_apv_list = []
            net_revenue = 0
            net_pv = 0
            
            LP = streams_for_case['Fixed Capital Investment']*.6
            FCI = streams_for_case['Fixed Capital Investment']
            stm_val = streams_for_case['Steam Plant Value']
            
            for year_num,year_prices in enumerate(lifetime):
                
                bio_price = year_prices[0] 
                sa_price = year_prices[1]
                
                print_arg = False
                #if year_num == 0:
                    #print_arg = True
                #print(case_index)
                
                APV,LP,net_revenue = get_yearly_profit(streams_for_case, bio_price, sa_price,LP,year_num,FCI,net_revenue, stm_val, print_arg)
                
                yearly_apv_list.append(APV)
                net_pv += APV
                #print(year_num, APV)
            lifetime_yearly_apv.append(yearly_apv_list)
            lifetime_npv.append(net_pv)
            
        apv_list_by_case_by_lifetime.append((dfstreams.index[case_index],lifetime_yearly_apv))   
        net_pv_by_case_by_lifetime.append((dfstreams.index[case_index],lifetime_npv)) 
    return apv_list_by_case_by_lifetime, net_pv_by_case_by_lifetime

def get_each_case_avg_NPV(net_pv_by_case_by_lifetime):
    avg_NPV = []
    cases = []
    for case, NPV in net_pv_by_case_by_lifetime:
        avg_NPV.append(np.mean(NPV))
        cases.append(case)
    avg_NPVs= [cases, avg_NPV]
    return avg_NPVs

##############################################functions called######################################################
def get_yearly_profit(streams_for_case, bio_price, sa_price,LP,year_num,FCI,net_revenue, stm_val, print_arg):
    '''
    calculates the net revenue in a given month
    '''           
    months_per_year = 12
    
    revenue = \
        streams_for_case['Biofuel Output']*months_per_year*bio_price + \
        streams_for_case['Succinic Acid Output']*months_per_year*sa_price -\
        streams_for_case['Var OpCosts']*months_per_year -\
        streams_for_case['Fixed Op Costs']*months_per_year 
    
    net_revenue,LP,taxes = get_taxes(revenue,LP,year_num,FCI,streams_for_case,net_revenue, stm_val, print_arg)

    #if print_arg == True:
        #print('yearly vals',year_num, net_revenue, LP, taxes)
    
    profit = revenue - taxes
    #print(year_num) 
    if LP != 0:
        profit -= streams_for_case["Loan Payment per Year"]
    #print(profit)
    bag_years = [0,5,10,15,20,25]
    if year_num in bag_years:
        profit -= streams_for_case['Bag Cost']
    #print(profit)
     
    APV = profit/((1+0.1)**(year_num+1))
   
    return APV,LP,net_revenue

def get_taxes(revenue,LP,year_num,FCI,streams_for_case,old_net_revenue, stm_val, print_arg):
    
    interest = LP*.08

    if LP > 0 :
        LP -= (streams_for_case["Loan Payment per Year"] - interest)
    
    else:
        LP = 0
    #print(year_num, LP, interest)
    gen_dep_charges,stm_dep_charges = get_depreciations(FCI, stm_val)
    
    
    net_revenue = (revenue - interest - gen_dep_charges[year_num] - stm_dep_charges[year_num]) 
    bag_years = [0,5,10,15,20,25]
    
    if year_num in bag_years:
        net_revenue -= streams_for_case['Bag Cost']
    
    if old_net_revenue < 0:
        net_revenue += old_net_revenue
    else:
        pass
    
    if net_revenue > 0:
        tax_incur = 0.35*net_revenue
    else:
        tax_incur = 0
    #if streams_for_case["Loan Payment per Year"]  -  9.369985*(10**7) < 1000 :
        #print(tax_incur)
    #print(net_revenue)
    return net_revenue,LP,tax_incur


def get_depreciations(FCI, stm_val):
    
    gen_plant_depreciation = [.1429,.2449,.1749,.1249,.0893,.0892,.0893,.0446,0,0,0\
                              ,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0]
    steam_plant_depreciation = [.0375,.07219,.06677,.06177,.05713,.05285,.04888,.04522,\
                            .04462,.04461,.04462,.04461,.04462,.04461,.04462,.04461,\
                            .04462,.04461,.04462,.04461,.02231,0,0,0,0,0,0,0,0,0]
    
    charge_list = []
    
    gen_val = FCI - stm_val
    
    for year,yearly_dep in enumerate(gen_plant_depreciation):
        #print('start', gen_val, )
        gen_dep_charge = gen_val*gen_plant_depreciation[year]
        #print(gen_val)
        #gen_val -= gen_dep_charge
        #print('gen_val end', gen_val)
        stm_dep_charge = stm_val*steam_plant_depreciation[year]
        
        #stm_val -= stm_dep_charge
        
        charge_list.append((gen_dep_charge,stm_dep_charge))
    
    gen_dep_charges,stm_dep_charges = zip(*charge_list)
    return gen_dep_charges,stm_dep_charges
#######################################used in generating price####################################

def get_yearly_predictions(price_preds):
    """
    Takes montly predictions for (bio,sa) and averages them
    to get yearly predictions 
    """    
    predictions = []
    bio_prices,SA_prices = price_preds
    i = 0
    j = 0
    #print(bio_prices)
    while i < len(bio_prices[:-2]) and j < len(SA_prices):
         #print(bio_prices[i:i+12])
         #print(np.mean(bio_prices[i:i+12]))
         yearly_bio = np.mean(bio_prices[i:i+12])
         if len(SA_prices) == len(bio_prices):
             yearly_sa = np.mean(SA_prices[i:i+12])
         else:
             yearly_sa = SA_prices[j]
             j+=1
         
         predictions.append((yearly_bio,yearly_sa))
         i+=12
         
    #print(len(predictions))
    #print(predictions)    
    return predictions

def slice_predictions(predictions):
    years = 30
    j = 0
    lifetime_preds = []
    #print(len(predictions))
    #print(predictions[j:j+years])
    while len(predictions[j:j+years]) == 30:
        #print(predictions[j:j+years])
        #print(len(predictions[j:j+years]),j)
        lifetime_preds.append(predictions[j:j+years])
        j+=1
    #print(len(lifetime_preds))
    return lifetime_preds
###############################################plotting funcs##########################################


if __name__ == "__main__":
    oilfilename = 'hist_oil.csv'
    #oilfilename = 'future_oil.csv'
    streamsfilename = '3-7-2018_df_final.xlsx'
    dfstreams = pd.read_excel(streamsfilename, 'Sheet1')
    shortdf = dfstreams.iloc[46:49]
    price_preds = hist.get_price_preds(oilfilename,3)
    #price_preds = hist.get_price_preds(oilfilename,1)
    lifetime_preds = slice_predictions(get_yearly_predictions(price_preds))
    short_preds = lifetime_preds[0:2]
    
    #apv_by_case, case_npv = get_profits_by_case(shortdf, short_preds)
    #apv_by_case, case_npv = get_profits_by_case(shortdf, lifetime_preds)
    #apv_by_case, case_npv = get_profits_by_case(dfstreams, short_preds)
    apv_by_case, case_npv = get_profits_by_case(dfstreams, lifetime_preds)
    avg_NPVs = get_each_case_avg_NPV(apv_by_case)
    #####################automated plotting##########################################
    autplot.plot_average_pv_30years(avg_NPVs) 
    autplot.one_case_hist(case_npv, 0.50)
    #autplot.plot_multiple_histograms(case_npv, 2)
    
'''
if __name__ == "__main__":
    oilfilename = 'hist_oil.csv'
    #oilfilename = 'future_oil.csv'
    streamsfilename = '3-6-2018_trial_lifetime_run.xlsx'
    dfstreams = pd.read_excel(streamsfilename, 'Sheet1')
    price_preds = hist.get_price_preds(oilfilename,3)
    #price_preds = hist.get_price_preds(oilfilename,1)
    lifetime_preds = slice_predictions(get_yearly_predictions(price_preds))
    profits_by_case = get_profits_by_case(dfstreams,lifetime_preds)
    #print(profits_by_case[19])
    #print(profits_by_case[20])
    summed_lifetimes_by_case = []
    for case,lifetimes in profits_by_case:
        summed_lifetimes = []
        for lifetime in lifetimes:
            lifetime = sum(lifetime)
            summed_lifetimes.append(lifetime)
        summed_lifetimes_by_case.append((case,summed_lifetimes))
    
    for case,lifetime in summed_lifetimes_by_case:
        plt.hist(lifetime,alpha = .5)
        if case == 0.2:
            print(lifetime)
        if case == 0.19:
            print(lifetime)
    plt.show()



if __name__ == "__main__":
    oilfilename = 'hist_oil.csv'
    #oilfilename = 'future_oil.csv'
    streamsfilename = '3-6-2018_trial_lifetime_run.xlsx'
    dfstreams = pd.read_excel(streamsfilename, 'Sheet1')
    price_preds = hist.get_price_preds(oilfilename,3)
    #price_preds = hist.get_price_preds(oilfilename,1)
    lifetime_preds = slice_predictions(get_yearly_predictions(price_preds))
    profits_by_case = get_profits_by_case(dfstreams,lifetime_preds)
    summed_lifetimes_by_case = []
    for case,lifetimes in profits_by_case:
        summed_lifetimes = []
        for lifetime in lifetimes:
            lifetime = sum(lifetime)
            summed_lifetimes.append(lifetime)
        summed_lifetimes_by_case.append((case,summed_lifetimes))
    
    for case,lifetime in summed_lifetimes_by_case:
        plt.hist(lifetime,alpha = .5)
    plt.show()
    #avg_profits_by_case = get_lifetime_profits_by_case(dfstreams,lifetime_preds)
    #x,y = zip(*avg_profits_by_case)

    fig = plt.figure()
    ax = plt.axes()
    x = np.asarray(x)
    y = np.asarray(y)
    ax.plot([0,55],[0,0],'k-')
    for i,yval in enumerate(y):
        #print(yval)
        if yval < 0:
            print((x[i]*100,yval/(10**9)))
            ax.plot(x[i]*100,yval/(10**9),'r^')
        else:
            print((x[i]*100,yval/(10**9)))
            ax.plot(x[i]*100,yval/(10**9),'g^')
    ax.grid(True)
    ax.set_ylabel("Average 10yr Profit (Billions of USD)")
    ax.set_xlabel("Succinic Acid Production (%)")
    ax.set_xlim([0,50])
    plt.show()
'''
    
    
    
