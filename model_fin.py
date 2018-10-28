import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import math
from scipy import stats

import random
from operator import add, sub, mul

######################
#MODEL FOR PRICE DATA#
######################  

def upload(filename, ind_col, dep_col='', date_col=''):
    data = np.loadtxt(filename,dtype = np.str, delimiter = ",",unpack=True,skiprows = 1)
    if len(np.shape(data)) == 1:
        
        data = [float(x) for x in data]
        print(data)
        return np.asarray(data)
        
    ind = data[int(ind_col)]
    print(ind)
    ind = [float(x) for x in ind]
    ind = np.asarray(ind)

    
    if date_col != '':
        dates = data[int(date_col)]
        dates = [pd.to_datetime(date) for date in dates]
        if dep_col == '':
            return([ind,dates])
   
    if dep_col != '':
        dep = data[int(dep_col)]
        print(dep)
        dep = [float(x) for x in dep]
        dep = np.asarray(dep)
        
        #best_abt, best_abt_lsqr = mod.overall_func(dep, ind)
        return([ind,dep,dates])
              
    return(ind) 


def get_vectors(a_min = 0,a_max = 50,b_min = -600,b_max = -400,\
    n = 100,t_min = -3,t_max = 3):
    '''
    INPUTS: bounds for linear space of guesses for fit
    parameters and time shift. t's MUST BE INTEGERS!
    
    OUTPUTS: equally spaced numpy arrays between guess extrema.
             i.e a_s = [-2.0,-1.5,0.,.... etc... 6.0]
    '''
    
    t_span = (t_max - t_min) + 1
    
    a_s = np.linspace(a_min,a_max,n)
    b_s = np.linspace(b_min,b_max,n)
    t_s = np.linspace(t_min,t_max,t_span, dtype = int)
    
    return a_s, b_s, t_s,n
    
    
#outputs the best model parameters given the input guesses
def output_model(dep,ind, a_s, b_s, t_s):
    '''
    INPUTS: range of guesses for a, b, and time shift.
            array of dependent parameter data, array of
            independent parameter data.
            
    OUTPUTS: Best fit parameters which generate a model for the
             dependent data using the formula:
                 
             MODELED DEPENDENT(t) = a*(INDEPENDENT(t - shift) + b
             
             and best lsqr.
    '''
    #initialize lists to collect the fit coefficients and lsqrs

    best_coeffs = []
    best_lsqrs = []
    
    for shift in t_s:  
        
        #normalizes the lqsrs of different shifts for comparison
        norm_factor = len(ind[:-shift])
        
        for j,a in enumerate(a_s):
            for i,b in enumerate(b_s):
                
                #generate model from selected a and b
                dep_model = a*(ind) + b
                
                #implement time shift and calculate sum of least squares
                if shift > 0:
                    
                    lsqr = ((dep_model[:-shift] - dep[shift:])**2).sum()
                    lsqr = lsqr/norm_factor
                    
                    
                if shift == 0:
                    lsqr = ((dep_model - dep)**2).sum()
                    
                if shift < 0:
                    lsqr = ((dep_model[-shift:] - dep[:shift])**2).sum()
                    lsqr = lsqr/norm_factor
                
                
                #collect best b, for each a and t
                if i == 0:
                    best_lsqr = lsqr
                    best_b = b
                    
                else:
                    if lsqr < best_lsqr:
        
                        best_lsqr = lsqr
                        best_b = b
                        
            #collect best a,b pair for each t           
            if j == 0:
                
                best_ab_lsqr = best_lsqr
                best_a = a
                best_ab = (best_a,best_b)
                
                
            else:
                if best_lsqr < best_ab_lsqr:
                    
                    best_ab_lsqr = best_lsqr
                    best_a = a
                    best_ab = (best_a,best_b)
        
        best_coeffs.append([best_ab,shift,best_ab_lsqr])
        best_lsqrs.append(best_ab_lsqr)
        
    #IF CODE IS BROKEN TAB THE LINE BELOW BACK IN!!!!!!
    
    #select the overall lowest least square
    best_lsqr = min(best_lsqrs)
    
        
    #loop through the list of paramters and select 
    #the ones corresponding to the best lsqr            
    for coeffs in best_coeffs:
        
        best_ab = coeffs[0]
        #best_a = best_ab[0]
        #best_b = best_ab[1]
        best_a,best_b = best_ab
        
        shift = coeffs[1]
        lsqr = coeffs[2]
        
        if lsqr == best_lsqr:
            best_abt = (best_a,best_b,shift)
            best_abt_lsqr = lsqr
            
    return best_abt, best_abt_lsqr
      
      
    #brute force check 
    #is the best prarmeters in the original guess range
    #resets the range if not
def range_moving(best_abt, a_s, b_s, n, range_is_correct = False):
    '''
    This function checks that the a and b corresponding to the best
    lsqr are not on the edge of the search region in parameter space. If 
    they are on the edge, it shifts the search range.
     
    INPUTS: The best a,b, and t from output_model(), the search ranges used 
    for a and b to find these params and the step number (n), the best lsqr
    for these params.
    
    OUTPUTS: IF the range is correct, returns the a,b, shift, and lsqr already
    found. IF the range is incorrect, returns a new range for use in
    output_model().  
    '''
    best_a,best_b,shift = best_abt
    
    #best_a = float(best_abt[0])

    #best_b = best_abt[1]
    
    #collect mins and maxxes for comparison
    a_min = a_s[0]
    a_max = a_s[-1]
    b_min = b_s[0]
    b_max = b_s[-1]
    
    if (best_a != a_min and best_a != a_max) and \
       (best_b != b_min and best_b != b_max): 
        return True, a_s, b_s
        
        
    #endpoints are best a or b, readjust range of potential parameters
    else: 
        #a is the lowest endpoint
        if best_a == a_min:
            new_a_s = np.linspace(a_min-(a_max-a_min),a_max,n)
        #a is the highest possible endpoint
        elif best_a == a_max:
            new_a_s = np.linspace(a_max,a_max+(a_max-a_min),n)
        else:
            new_a_s = a_s
        #b is the lowest endpoint
        if best_b == b_min:
            new_b_s = np.linspace(b_min-(b_max-b_min),b_min,n)
        #b is the highest endpoint
        elif best_b == b_max:
            new_b_s = np.linspace(b_max,b_max+(b_max-b_min),n)
        else:
            new_b_s = b_s
    return False, new_a_s, new_b_s


#check if paramater range is narror enough
#narrows range to specific degree of precision
def range_narrow_precise(best_abt, a_s, b_s, t_s,n, best_abt_lsqr,dep,ind,\
    a_tol = .001, b_tol = .001):
    '''
    This function is only called when the search range is determined to contain
    the loswest lsqr. It narrows the search range, keeping the step number until
    the a,b, and t are found to a desired precision.
    
    INPUTS: The best a,b, and t from output_model(), the search ranges used 
    for a and b to find these params and the step number (n), the best lsqr
    for these params. Independent and dependent data. Tolerances for a and b.
    '''
    
    #collect mins and maxxes for comparison
    a_min = a_s[0]
    a_max = a_s[-1]
    b_min = b_s[0]
    b_max = b_s[-1]
    #define tolerance for check

    best_a,best_b,shift = best_abt
    
    a_array_midpoint = (a_min+a_max)/2.0

    b_array_midpoint = (b_min+b_max)/2.0

    #redefine vectors with smaller ranges for better confidence
    if best_a < a_array_midpoint:
        new_a_s = np.linspace(a_min,a_array_midpoint,n)
        
    elif best_a >= a_array_midpoint:
        new_a_s = np.linspace(a_array_midpoint, a_max,n)
        
    if best_b < b_array_midpoint:
        new_b_s = np.linspace(b_min,b_array_midpoint,n)
        
    elif best_b >= b_array_midpoint:
        new_b_s = np.linspace(b_array_midpoint, b_max,n)
        
    #run parameter search with new, tighter ranges
    new_best_abt, new_best_lsqr = output_model(dep, ind, new_a_s, new_b_s, t_s)
    
    new_best_a,new_best_b,shift = new_best_abt
    #tolerance check
    a_precision = abs(new_best_a - best_a)
    b_precision = abs(new_best_b - best_b)
    
    #return new vectors that have a more narrow range if the old 
    #model was not precise enough
    if a_precision <= a_tol and b_precision <= b_tol:
      return True, a_s, b_s, best_abt_lsqr
      
    return False, new_a_s, new_b_s, new_best_lsqr


def plot_as_time_series(dep,ind,dates,abt,lqsr,show_ind = False):
    '''
    plots the data output by the model. optional parameter to add in randomness if 
    the used wants to add randomness to the model
    '''
    a,b,t = abt
    model = gen_model_data(abt,ind)
    
    ax = plt.axes()
    
    if show_ind:
        ax.plot(dates,ind,'r-',label= 'Independent')
        
    ax.plot(dates,dep,'b-',label = 'Dependent')
    
    if t > 0:
        
        ax.plot(dates[t:],model[:-t],'g-',label = 'Model')
    if t < 0:
        ax.plot(dates[:-t],model[:-t],'g-',label = 'Model')
    if t == 0:
        ax.plot(dates,model,'g-',label = 'Model')
    plt.legend()
    plt.xlabel('year')
    plt.ylabel('Unit Price')
    plt.savefig('Model.jpg')
    plt.show()

#given a best_abt and best_abt_lsqr can add randomness to some model generated data 
def add_randomness(mu, std, model,vol):

    new_model = []
    
    #scale up volatility by increasing std
    std *= (1+vol)
    
    #add in CAGR
    #mu *= (1+growth_pre)
    #print(model)
    for i,val in enumerate(model):
        
        resid = np.random.normal(mu, std)
        new_model.append(val + resid)
    
    return np.array(new_model)

def get_residuals(best_abt,ind,dep):
    '''
    Upon aligning the time frame of the model and dependent variable we are determining the residuals
    of the match
	input: best_abt: best coefficients
	inp: read in the oil price data
	dep : read in the maleic anhydride values
	Outputs: residuals
    '''
    a,b,shift = best_abt
    model = gen_model_data(best_abt,ind)    
    
    if shift > 0:
        model = model[:-shift] #we are shifting model to match time frame of dependent variable
        dep = dep[shift:] #we are excluding data points that reflect oil data we don't have

    elif shift < 0:
        model = model[-shift:]
        dep = dep[:shift]

    residuals = model - dep
    return residuals
    

def get_summary_stats(best_abt, ind, dep, conf = 95):
    '''
    This function will generate summary statistics for the model
    
    input: best_abt: best coefficients
	inp: read in the oil price data
	dep : read in the maleic anhydride values
	Outputs: Mean of the residuals, standard of error on the residuals,
        standard error, margin of error on the slope, margin of error intercept
    '''
    residuals = get_residuals(best_abt,ind,dep)
        
    mean_res = np.mean(residuals)
    std_res = np.std(residuals)
    #gets information about the standard deviation
    n = len(residuals)
    #sum of squared errors
    ssr = math.sqrt((np.sum(np.square(residuals))/(n-2)))
    #explained sum of squares by ind variables
    ess = math.sqrt(np.sum(np.square(np.subtract(ind, np.mean(ind)))))
    #standard error
    std_err = ssr/ess
    #gets tcrit
    alpha = 1 - (conf/100.0)
    p = 1 - (alpha/2.0)
    tcrit = stats.t.ppf(p, n - 1)
    #margin of error in regression slope
    margin_err_slope = std_err*tcrit
    return mean_res, std_res, std_err, margin_err_slope
        
#used to generate data if fed a csv of hisorical data    
def predict(abt, ind, mu, std, vol = 0,span = 'month',cutoff = None, plot=False, time='Month'):
    '''
    This function would only be used if one wanted to generate future data
    of price/MT for a chemical. The function outputs an array of $/MT points 
    based on historical data or model generated data. 
    
    INPUTS: 
        growth_pre: float greater than 0 and less than 1, annual growth rate 
            prediction
        vol: float greater than 0 and less than 1, prediction for increase in
            volatility
        abt: currently input from best model, but could build in presets
        ind: independent data
        dep: historical dependent data but could be skipped and use only abt
    '''
    ind = gen_linspace(ind,span,cutoff) # user written function to parse data into months
    #print('got here')
    print(ind)
    model = gen_model_data(abt,ind) # generate data from model

    #mu, std, std_err, margin_err_slope = get_summary_stats(abt, ind, dep, conf = 95)
    new_model = add_randomness(mu,std,model,vol) # add randomness to data
    if plot:
        #x_ax = np.arange(len(new_model))
        plt.plot(ind,label = 'Independent')
        plt.plot(new_model,label = 'Dependent')
        plt.xlabel('Time (' + str(time) + 's)')
        plt.ylabel('Unit Price')
        plt.legend()
        plt.savefig('Prediction.jpg')
        plt.show()
        print('And you ask yourself, How did I get here')
    return new_model

def single_point_predict(value,mu,std,vol,CAGR,num_months,plot= False):

    print(plot,'SAY SOMETHING')
    CMGR = (CAGR + 1.0)**(1/12.0) - 1.0 # convert CAGR to monthly
    time_span = range(num_months - 1)
    monthly_preds = [value]
    monthly_preds_with_resid = [value]

    std = std*(1 + vol)
    
    for month in time_span:
        pred = (value)*(1 + CMGR) # predict new point
        monthly_preds.append(pred)
        value = pred 
        point_std = std*(1 + CMGR)**month 
        
        print(point_std)
        resid = np.random.normal(mu, point_std)# determine a residual
        monthly_preds_with_resid.append(pred + resid)

    if plot:
        plt.plot(time_span,monthly_preds_with_resid[1:])
        plt.ylabel('Unit Price')
        plt.xlabel('Time (months)')
        plt.savefig('Single_Point_Prediction.jpg')
        plt.show()
    return (monthly_preds_with_resid)
    
def gen_model_data(best_abt,ind):
    '''
    generate model data from oil prices and best parameters
    return: model values
    '''
    print(best_abt,ind)
    a,b,t = best_abt
    
    model = a*ind + b
    
    return model

def gen_linspace(ind,span = 'month', cutoff= None):
    print('SPAN: '+ span)
    if span == 'month':
        return ind[:cutoff]
        
    if span == 'year': # if data is read in with increments of a year
        print('ici')
        interp_ind  = []
        
        for i,val in enumerate(ind[:-1]):
            months = np.linspace(ind[i],ind[i+1],13) # split data into months

            months = months[:-1]
            for j,item in enumerate(months): # IS THIS FOR LOOP NECESSARY?
                interp_ind.append(months[j])
            #print(interp_ind) 
            
        return np.array(interp_ind[:cutoff])
        
    if span == 'quarter': # if data is read in in increments of a quarter
        interp_ind  = []
        
        for i,val in enumerate(ind[:-1]):
            months = np.linspace(ind[i],ind[i+1],4)

            months = months[:-1]
            for j,item in enumerate(months):
                interp_ind.append(months[j])
                
        return np.array(interp_ind[:cutoff]) # return data in fomrated fashion
            
        
                                 
def overall_func(filename,ind_col, dep_col='', date_col='',a_min = 0,a_max = 2,\
                    b_min = 40,b_max = 50,n = 200,t_min = -3,t_max = 3,\
                    plot = False, get_sum = False):
    
    a_s, b_s, t_s, n = get_vectors(a_min,a_max,b_min,b_max,n,t_min,t_max)
    
    data = upload(filename,ind_col, dep_col, date_col)
    
    if len(data) == 3:
        
        ind  = np.asarray(data[0])
        dep =  np.asarray(data[1])
        dates = np.asarray(data[2])
        
        #initialize booleans for while loops
        range_correct = False
        range_precise = False
        range_search_iter = 0
        range_narrowing_iter = 0
        
        # search for the best parameters of model until they  are not end points
        while (not range_correct) and range_search_iter < 1000:
            best_abt, best_abt_lsqr = output_model(dep, ind, a_s, b_s, t_s)
            range_correct, a_s, b_s = range_moving(best_abt, a_s, b_s,\
                n, best_abt_lsqr)
            range_search_iter += 1
        # narrow in on the search of the best model parameters until tolerance is met
        while (not range_precise) and range_narrowing_iter < 1000:
            best_abt, best_abt_lsqr = output_model(dep, ind, a_s, b_s, t_s)
            range_precise, a_s, b_s, best_abt_lsqr = range_narrow_precise(best_abt, \
                a_s, b_s, t_s,n, best_abt_lsqr,dep,ind)
            range_narrowing_iter += 1
        # plot the oil data, MA data, and best model over a given time frame
        if plot:
            plot_as_time_series(dep,ind,dates,best_abt,best_abt_lsqr,True) 
        # If the user asks for a summary of the statisical data of the model retunr data
        if get_sum:
            mean_res, std_res, std_err, margin_err_slope = get_summary_stats(best_abt,\
            ind, dep, conf = 95)
            return best_abt, best_abt_lsqr, mean_res, std_res, std_err, margin_err_slope
            
        return best_abt, best_abt_lsqr
    
    return print('Not enough data supplied')
  

