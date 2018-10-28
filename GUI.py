from tkinter import *
from tkinter import messagebox
import numpy as np
import model_fin as mod
from tkinter import ttk
from tkinter.filedialog import askopenfilename
#reload(mod)

################Global Constants###############
OIL_TO_MA = (1.378,648.5,3)
MA_TO_SA = (.7465,1050,3)
OIL_TO_SA = (1.1387,1468,3)
TO_PERCENT = .01
 
def quit():
    global root
    root.destroy()
 
############TAB 2 FUNCTIONS##################### 

def run_predict():

    
    indpt = mod.upload(data.get(),data_col.get())
    cutoff = len(indpt)
    #print(len(indpt))
    print(indpt)
    
    volatility = get_slider_val()
    abt = get_preset_params()
    
    mu = float(mu_res.get())
    std = float(sig_res.get())
    print(make_plot.get())
    time_type = get_span()
    gen_data = mod.predict(abt, 
                            np.asarray(indpt),
                            mu,
                            std,
                            volatility*.01,
                            time_type,
                            cutoff,
                            make_plot.get(),time_type)
    
    np.savetxt(str(data.get())[:-4]+'_out.csv', gen_data, delimiter=',')  
 
def run_single_point_predict():
    #print('here')
    pt = float(single.get())
    volatility,growth = get_single_slider_vals()
    get_mean_std()

    mu = float(mu_res.get())
    std = float(sig_res.get())
        
    print(make_plot.get(),'LOOK')
    pred_data = mod.single_point_predict(pt,
                            mu,
                            std,
                            volatility*.01,
                            growth*.01,
                            int(years.get()),
                            make_plot.get())
    #print('heresies')
    print(pred_data,'HI')
    np.savetxt('single_pred_out.csv', pred_data, delimiter=',')       
    
def get_slider_val():
    return (vol.get())

def get_single_slider_vals():
    return (vol2.get(), grow.get())

def get_preset_params():
    
    if a.get() != '' and b.get() != '' and shift.get() != '':
        abt = (float(a.get()),float(b.get()),float(shift.get()))
    
    if v.get() == 1:
        abt = OIL_TO_MA
        mu_res.delete(0,END)
        mu_res.insert(0,2.883*(10**-4))
        
        sig_res.delete(0,END)
        sig_res.insert(0,114.8)
        
    if v.get() == 2: 
        abt = MA_TO_SA
        
        mu_res.delete(0,END)
        mu_res.insert(0,-1.619)
        
        sig_res.delete(0,END)
        sig_res.insert(0,144.7)
        
    if v.get() == 3: 
        abt = OIL_TO_SA
        
        mu_res.delete(0,END)
        mu_res.insert(0,3.470*(10**-3))
        
        sig_res.delete(0,END)
        sig_res.insert(0,116.0)
        
    return abt
    
def get_mean_std():
    
    if v.get() == 1:

        mu_res.delete(0,END)
        mu_res.insert(0,2.883*(10**-4))
        
        sig_res.delete(0,END)
        sig_res.insert(0,114.8)
        
    if v.get() == 2: 

        mu_res.delete(0,END)
        mu_res.insert(0,-1.619)
        
        sig_res.delete(0,END)
        sig_res.insert(0,144.7)
        
    if v.get() == 2: 
        
        mu_res.delete(0,END)
        mu_res.insert(0,3.470*(10**-3))
        
        sig_res.delete(0,END)
        sig_res.insert(0,116.0)
    
    
def get_span():    
    if spacing.get() == 1:
        return 'month'
    elif spacing.get() == 2:
        return 'quarter'
    elif spacing.get() == 3:
        return 'year'
    return 'month'
                
def open_pred_file():
    root.filename = askopenfilename(initialdir = "/",
                                                title = "Select file",
                                                filetypes = (("csv files","*.csv"),("all files","*.*")))
    data.insert(0,root.filename)

def open_reg_file():
    root.filename = askopenfilename(initialdir = "/",
                                                title = "Select file",
                                                filetypes = (("csv files","*.csv"),("all files","*.*")))
    name.insert(0,root.filename)
        
###############TAB 1 FUNCTIONS#####################                 
                                                             
def run_regression():
    if show_stats.get():
        abt,lsqr, mean_res, std_res, std_err, moe_slope = mod.overall_func(name.get(),
                                    indpt.get(),dep.get(),days.get(),
                                    a_min = 0.5,a_max = 2,\
                                    b_min = 600,b_max = 700,n = 200,t_min = -2,
                                    t_max = 2,plot = show_plot.get(),get_sum = True)
                                    
        a_val,b_val, shift_val = abt
        mu = mean_res
        std = std_res
        '''                                                        
        showinfo('Output',"\
                 a: %s\n\
                 b: %s\n \
                 time shift: +%s\n\
                 lsqr: %s\n\
                 mean residuals: %s\n\
                 std dev residuals: %s\n\
                 SE residuals: %s\n\
                 slope margin of error: %s\
                 " % (str(a_val),
                str(b_val),str(shift_val),\
                str(lsqr),str(mean_res),str(std_res),str(std_err),str(moe_slope)))
        '''
        Label(tab1,text = "\
                 a: %s\n\
                 b: %s\n \
                 time shift: +%s\n\
                 lsqr: %s\n\
                 mean residuals: %s\n\
                 std dev residuals: %s\n\
                 SE residuals: %s\n\
                 slope margin of error: %s\
                 " % (str(a_val),
                str(b_val),str(shift_val),\
                str(lsqr),str(mean_res),str(std_res),str(std_err)
                ,str(moe_slope))).grid(row = 9,sticky = E)
    
    else:
        abt,lsqr = mod.overall_func(name.get(),indpt.get(),dep.get(),days.get(),
                                    a_min = 0.5,a_max = 2,\
                                    b_min = 600,b_max = 700,n = 200,t_min = -2,
                                    t_max = 2,plot = show_plot.get())
    
        a_val,b_val, shift_val = abt 
    
        showinfo('Output',"a: %s\nb: %s\n time shift: +%s\nlsqr: %s" % (str(a_val),
                str(b_val),str(shift_val),\
                str(lsqr)))
    
            
    a.insert(0,str(a_val))
    b.insert(0,str(b_val))
    shift.insert(0,str(shift_val))
    sig_res.insert(0,str(std))
    mu_res.insert(0,str(mu))

############3#INITIALIZE ROOT AND TABS###############
root = Tk()

note = ttk.Notebook(root)
note.grid()

tab1 = ttk.Frame(note)
note.add(tab1,text = "Regression")

tab2 = ttk.Frame(note)
note.add(tab2,text = "Linear Model Price Generator")

tab3 = ttk.Frame(note)
note.add(tab3,text = "CAGR Price Prediction")


###############TAB 1 LABELS#################
Button(tab1, 
        text='Upload Data File',
        command=open_reg_file).grid(row=0,
        column=0,
        sticky = E,  
        pady = 5,padx = 5)

name = Entry(tab1)
name.grid(row=0, column=1)

Label(tab1, 
      text="Independent Data Column:").grid(row=2,sticky = E,pady = 5,padx = 5)
indpt = Entry(tab1)
indpt.grid(row=2, column=1,pady = 5,padx = 5)

Label(tab1, 
      text="Dependent Data Column:").grid(row=3,sticky = E,pady = 5,padx = 5)
dep = Entry(tab1)
dep.grid(row=3, column=1,pady = 5,padx = 5)

Label(tab1, 
      text="Date Column:").grid(row=1,sticky = E,pady = 5,padx = 5)
days = Entry(tab1)
days.grid(row=1, column=1,pady = 5,padx = 5)


###############TAB 1 BUTTONS#################
Button(tab1, 
       text='Exit Application', 
       command=quit).grid(row=6,
       column=0, 
       sticky=W, 
       pady=4)
       
Button(tab1,
       text='Run Regression',
       command=run_regression).grid(row=6,
       column=1, 
       sticky=W, 
       pady=4)

show_stats = IntVar()
Checkbutton(tab1, text="Show Summary Stats", variable=show_stats).grid(row=4, sticky=W)

show_plot = IntVar()
Checkbutton(tab1, text="Show Plot", variable=show_plot).grid(row=5, sticky=W)


############# TAB 3 ################

Button(tab3, 
       text='Gen Model from Point', 
       command=run_single_point_predict
       ).grid(row = 7,padx =5, pady = 5)

Label(tab3, 
      text="Initial Price:").grid(row=2,sticky = E,pady = 5,padx = 5)
single = Entry(tab3)
single.grid(row=1, column=1,pady = 5,padx = 5)   

Label(tab3, 
      text="Prediction Time:").grid(row=0,sticky = E,pady = 5,padx = 5)
years = Entry(tab3)
years.grid(row=0, column=1,pady = 5,padx = 5)    

Label(tab3, 
      text="CAGR[%]:").grid(row=3,sticky = E,pady = 5,padx = 5)
grow = Entry(tab3)
grow.grid(row=2, column=1,pady = 5,padx = 5)  

Label(tab3, 
      text="mean residuals:").grid(row=4,sticky = E,pady = 5,padx = 5)
mu_res = Entry(tab3)
mu_res.grid(row=3, column=1,pady = 5,padx = 5) 

Label(tab3, 
      text="std residuals:").grid(row=5,sticky = E,pady = 5,padx = 5)
sig_res = Entry(tab3)
sig_res.grid(row=4, column=1,pady = 5,padx = 5)

Label(tab3).grid(row = 1, column = 0, sticky = E)
datatype_lf = ttk.Labelframe(tab3, text='Data Type:')
datatype_lf.grid(row = 1,column = 1,sticky = W,pady = 10,padx = 20)

spacing = IntVar()
Radiobutton(datatype_lf, text="Monthly", variable=spacing,value = 1).grid(row=0,column=0 , sticky=W)
Radiobutton(datatype_lf, text="Quarterly", variable=spacing,value = 2).grid(row=0,column = 1, sticky=W)
Radiobutton(datatype_lf, text="Yearly", variable=spacing,value = 3).grid(row=0, column = 2, sticky=W)


######### BUILTIN PARAM BUTTONS ################

v = IntVar()

chem_lf = ttk.Labelframe(tab2, text='Choose a Model:')
chem_lf.grid(row = 3,column = 1,sticky = W)
oil = Radiobutton(chem_lf, 
              text="Oil -> MA",
              padx = 20, 
              variable=v, 
              value=1).grid(row = 9,column = 0)
              
MA = Radiobutton(chem_lf, 
              text="MA -> SA",
              padx = 20, 
              variable=v, 
              value=2).grid(row = 10, column = 0)
              
SA = Radiobutton(chem_lf, 
              text="Oil -> SA",
              padx = 20, 
              variable=v, 
              value=3).grid(row = 9, column = 1)
                   
SA = Radiobutton(chem_lf, 
              text="Use Custom",
              padx = 20, 
              variable=v, 
              value=4).grid(row = 10, column = 1)         

              

###########TAB 2 ENTRY LABELS###############

Button(tab2, 
        text='Upload Data File',
        command=open_pred_file).grid(row=0,
        column=0, sticky = E,padx = 5,pady = 5)
        
data = Entry(tab2)
data.grid(row=0,column = 1, sticky = W)


Label(tab2, text="Use Column:").grid(row=1,sticky = E)
data_col = Entry(tab2)
data_col.grid(row=1, column=1,sticky = W,pady = 10)

ttk.Labelframe(tab2).grid(row = 2, column=0, sticky = W)
datatype_lf = ttk.Labelframe(tab2, text='Data Sampling Space')
datatype_lf.grid(row = 3,column = 0,sticky = W,pady = 10,padx = 20)

spacing = IntVar()
Radiobutton(datatype_lf, text="Month", variable=spacing,value = 1).grid(row=0,column=0 , sticky=W)
Radiobutton(datatype_lf, text="Quarter", variable=spacing,value = 2).grid(row=0,column = 1, sticky=W)
Radiobutton(datatype_lf, text="Year", variable=spacing,value = 3).grid(row=0, column = 2, sticky=W)


#####################CUSTOM PARAMS####################################
custom_lf = ttk.Labelframe(tab2, text='Custom Parameters:')
custom_lf.grid(row = 5,column = 0, columnspan = 2,sticky = N)

Label(custom_lf, text="a:").grid(row = 0,sticky = E)
a = Entry(custom_lf)
a.grid(row=0, column=1)

Label(custom_lf, text="b:").grid(row = 1,sticky = E,pady = 5)
b = Entry(custom_lf)
b.grid(row=1, column=1)

Label(custom_lf, text="shift:").grid(row = 2,sticky = E,pady = 5)
shift = Entry(custom_lf)
shift.grid(row=2, column=1)

Label(custom_lf, text="mean res:").grid(row = 3,sticky = E,pady = 5,padx = 5)
mu_res = Entry(custom_lf)
mu_res.grid(row = 3, column = 1)

Label(custom_lf, text="std res:").grid(row = 4, sticky = E,pady = 5,padx = 5)
sig_res = Entry(custom_lf)
sig_res.grid(row = 4, column = 1)


############### DATATYPE ######################



###############TAB 2 BUTTONS#################

make_plot = IntVar()
Checkbutton(tab2, text="Show Plot", variable=make_plot).grid(row=19, column = 0, sticky=W)


Button(tab2,
       text='Exit Application',
       command=quit).grid(row=20, 
       column=1, 
       sticky=E, 
       pady=4)
       
Button(tab2, 
       text='Gen Model Data',
       command=run_predict).grid(row=20, 
       column=0, 
       sticky=W)

mainloop()
  
