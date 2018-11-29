from tkinter import *
from tkinter import messagebox
import numpy as np
import obj2_cleaned_up_box as msens
from tkinter import ttk
from tkinter.filedialog import askopenfilename
import matplotlib
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2TkAgg
from matplotlib import pyplot as pplt
import csv

###################GOBALS####################
single_point_var_val= {}
univar_var_num_sim = {}
simulation_dist = {}
simulation_vars = {}


################Tab 1 Functions###############
 
def quit():
    global root
    root.destroy()

def open_excel_file():
    root.filename = askopenfilename(initialdir = "/",
                                                title = "Select file")
                                        
    excel.insert(0,'C:/Users/MENGstudents/Desktop/Biorefinery Design Project/Variable_Call_Excel.csv')
    
def open_aspen_file():
    root.filename = askopenfilename(initialdir = "/",
                                                title = "Select file")
    aspen.insert(0,'C:/Users/MENGstudents/Desktop/Biorefinery Design Project/BC1508F-BC_FY17Target._Final_5ptoC5_updated022618.apw')

def open_solver_file():
    root.filename = askopenfilename(initialdir = "/",
                                                title = "Select file")
    solver.insert(0,'C:/Users/MENGstudents/Desktop/Biorefinery Design Project/DESIGN_OBJ2_test_MFSP-updated.xlsm')
    
def plot_on_GUI(d_f_output, vars_to_change = []):
    '''
    This function will autoupdate the GUI to display the histogram distribution of MFSP
    and and histogram distributions of all other variables that were passed to the 
    function. 

    Inputs: 
        d_f_outputs: dictionary with output of a single simulation, where the 
            key is the variable and the value is the output value of the variable
            after the simulation
        vars_to_change: list of variables that were input
    
    '''
    columns = 5
    num_rows= ((len(vars_to_change) + 1) % columns) + 1
    counter = 1
    fig = pplt.figure(figsize = (15,7))
    a = fig.add_subplot(num_rows,columns,counter)
    counter += 1
    total_MFSP = d_f_output["MFSP"]
    num_bins = 100
    n, bins, patches = pplt.hist(total_MFSP, num_bins, facecolor='blue', alpha=0.5)
    a.set_title ("MFSP Distribution")
    a.set_xlabel("MFSP ($)")
    if len(vars_to_change) != 0:
        for var in vars_to_change:
            a = fig.add_subplot(num_rows,columns,counter)
            counter += 1
            total_data = d_f_output[var]
            num_bins = 100
            n, bins, patches = pplt.hist(total_data, num_bins, facecolor='blue', alpha=0.5)
            a.set_title(var)
    canvas = FigureCanvasTkAgg(fig)
    canvas.get_tk_widget().grid(row=8, column = 0,columnspan = 10, rowspan = 10, sticky= W+E+N+S, pady = 5,padx = 5,)
        
    root.update_idletasks()
    
def get_distributions(is_univar):
    global simulation_vars, simulation_dist
    if is_univar:
        max_num_sim = max(int(slot.get()) for slot in univar_var_num_sim.values())
        simulation_vars, simulation_dist = msens.get_distributions(str(excel.get()), max_num_sim)
        for (aspen_variable, aspen_call, fortran_index), dist in simulation_vars.items():
            if aspen_variable in univar_var_num_sim:
                num_trials_per_var = int(univar_var_num_sim[aspen_variable].get())
                dist = dist[:num_trials_per_var]
    else:
        simulation_vars, simulation_dist = msens.get_distributions(str(excel.get()), numtrial= int(sim.get()))
        
    
    
def plot_init_dist():
    '''
    This function will plot the distribution of variable calls prior to running
    the simulation. This will enable users to see whether the distributions are as they expected.
    
    Inputs:
        input_var_dist: dictionary where the key is the variable name, and values are
            lists of the values that will be used in the function.    
    
    '''
    
    global simulation_dist
    columns = 5
    num_rows= ((len(simulation_dist) + 1) % columns) + 1
    counter = 1
    fig = pplt.figure(figsize = (15,7))
    for var, values in simulation_dist.items():
        a = fig.add_subplot(num_rows,columns,counter)
        counter += 1
        num_bins = 100
        n, bins, patches = pplt.hist(values, num_bins, facecolor='blue', alpha=0.5)
        a.set_title(var)
    canvas = FigureCanvasTkAgg(fig)
    canvas.get_tk_widget().grid(row=8, column = 0,columnspan = 10, rowspan = 10, sticky= W+E+N+S, pady = 5,padx = 5,)
        
    root.update_idletasks()

sp_row_num = None
univar_row_num = None
    
def load_variables_into_GUI(tab_num):
    sens_vars = str(excel.get())
    global sp_row_num, univar_row_num
    single_pt_vars = []
    univariate_vars = []
    multivariate_vars = []
    type_of_analysis = analysis_type.get()
    global single_point_var_val
    with open(sens_vars) as f:
        reader = csv.DictReader(f)# Skip the header row
        for row in reader:
            if row['Toggle'].lower().strip() == 'true':          
                if type_of_analysis =='Single Point Analysis':
                    
                    single_pt_vars.append((row["Variable Name"], float(row["Range of Values"].split(',')[0].strip())))
                elif type_of_analysis == 'Multivariate Analysis':
                    multivariate_vars.append(row["Variable Name"])
                else:
                    univariate_vars.append((row["Variable Name"], row["Format of Range"], row['Range of Values'].split(',')))
    #now populate the gui with the appropriate tab and variables stored above
    if type_of_analysis == 'Single Point Analysis':
        sp_row_num = 2
        for name,value in single_pt_vars:
            sp_row_num += 1
            key = str(sp_row_num)
            Label(tab_num, 
            text= name).grid(row=sp_row_num, column= 1, sticky = E,pady = 5,padx = 5)
            key=Entry(tab_num)
            key.grid(row=sp_row_num, column=2,pady = 5,padx = 5)
            key.insert(0,str(value))
            single_point_var_val[name]= key
            

    if type_of_analysis == 'Univariate Sensitivity':
        univar_row_num = 8
        Label(tab_num, 
            text= 'Variable Name').grid(row=univar_row_num, column= 1,pady = 5,padx = 5)
        Label(tab_num, 
            text= 'Sampling Type').grid(row=univar_row_num, column= 2,pady = 5,padx = 5)
        Label(tab_num, 
            text= '# of Trials').grid(row=univar_row_num, column= 3,pady = 5,padx = 5)
        univar_row_num += 1
        for name, format_of_data, vals in univariate_vars:
            Label(tab_num, 
            text= name).grid(row=univar_row_num, column= 1,pady = 5,padx = 5)
            Label(tab_num, 
            text= format_of_data).grid(row=univar_row_num, column= 2,pady = 5,padx = 5)
            
            if not(format_of_data == 'linspace' or format_of_data == 'list'):
                key2=Entry(tab_num)
                key2.grid(row=univar_row_num, column=3,pady = 5,padx = 5)
                #key2.insert(0,univariate_sims)
                univar_var_num_sim[name]= key2
            else:
                Label(tab_num,text= str(len(vals))).grid(row=univar_row_num, column= 3,pady = 5,padx = 5)
            univar_row_num += 1
            
        
        # print to a new univariate analysis tab
        # we want to print the following:
        # Variable Name |  Distribution Type  |  List of values OR NumTrials entry
        
    #if type_of_analysis == 'Multivariate Analysis':
        # print to a new multivariate tab
        # what we want to print is just the variable name and then an 
        # auto-updating graph of its distribution
    
def initialize_multivar_analysis():
    global simulation_vars
    if len(simulation_vars) == 0:
        get_distributions(False)
    run_multivar_sens()
    
def initialize_univar_analysis():
    global simulation_vars
    if len(simulation_vars) == 0:
        get_distributions(True)
    run_univ_sens()
    
def display_time_remaining(time_remaining):
    '''
    THIS NEEDS TO PRINT OUT ESTIMATED TIME REMAINING
    '''
    return None


def check_abort():
    return abort.get()

def check_next_analysis():
    return None
    '''
    #will be for the univariate analysis so that the user can move onto the next analysis
    '''
    '''
    # NOTE, YOU WILL ALSO HAVE TO CHANGE THIS BUTTON BACK TO UNPRESSED ONCE YOU MOVE
    # ONTO THE NEXT VARIABLE
    move_to_next = next_analysis.get()
    if move_to_next:
        # NEED TO UPDATE THE BUTTON TO TURN IT BACK OFF
    return move_to_next
'''
def run_multivar_sens():
    aspenfile= str(aspen.get())
    solverfile= str(solver.get())
    numtrial= int(sim.get())
    outputfile= str(save.get())
    sens_vars = str(excel.get())
    d_f_output = msens.multivariate_sensitivity_analysis(aspenfile,solverfile,sens_vars,numtrial,outputfile, simulation_vars)
        
def run_univ_sens():
    aspenfile= str(aspen.get())
    solverfile= str(solver.get())
    numtrial= int(sim2.get())
    outputfile= str(save2.get())
    sens_vars = str(excel.get())
    global simulations_vars
    for (aspen_variable, aspen_call, fortran_index), values in simulation_vars.items():
        msens.univariate_analysis(aspenfile, solverfile, aspen_call, aspen_variable, values, fortran_index, outputfile)
        
        print('Finished Analysis for Variable: ', aspen_variable)
    print('-----------FINISHED-------------')

def single_point_analysis():
    global sp_row_num
    mfsp = 3.34 #msens.single_point(________)
    Label(tab3, text= 'MFSP = ' + str(mfsp)).grid(row=sp_row_num+1, column = 2)
    
    
    return None

def fill_num_trials():
    global fill_num_sims, univar_var_num_sim
    ntrials = fill_num_sims.get()
    for name, slot in univar_var_num_sim.items():
        slot.insert(0, ntrials)
    

sim = None
fill_nums_sims = None
sim2 = None
save2 = None
save= None
otherbool = None
show_plot = None
boolvar = None
cb = None
tab2 = None
tab1 = None
tab3 = None

def make_new_tab():
    global sim, sim2, save2, save, otherbool, show_plot, boolvar, cb, tab1, tab2, tab3, fill_num_sims
    
    #note.forget(tab5)
    if analysis_type.get() == 'Choose Analysis Type':
        print("ERROR: Select an Analysis")
    elif  analysis_type.get() == 'Univariate Sensitivity':
        tab2 = ttk.Frame(note)
        note.add(tab2,text = "Univariate Analysis")
        ##############Tab 2 LABELS##################
        
        Label(tab2, 
              text="Number of Simulations :").grid(row=3, column= 1, sticky = E,pady = 5,padx = 5)
        sim2 = Entry(tab2)
        sim2.grid(row=3, column=2,pady = 5,padx = 5)
        
        Label(tab2, 
              text="Save As :").grid(row=4, column= 1, sticky = E,pady = 5,padx = 5)
        save2 = Entry(tab2)
        save2.grid(row=4, column=2,pady = 5,padx = 5)
        
        Label(tab2,text = ".csv").grid(row = 4, column = 3, sticky = W)
        
        ##############Tab 2 Buttons###############
        Button(tab2,
               text='Univariate Sensitivity Analysis',
               command=initialize_univar_analysis).grid(row=5,
               column=3, columnspan=2,
               pady=4)
        Button(tab2,
               text='Display Variable Distrbutions',
               command=display_distributions(True)).grid(row=5,
               column=1, columnspan=2,
               pady=4)
        Button(tab2,
               text='Fill Simulations',
               command=fill_num_trials).grid(row=7,
               column=2, sticky = E,
               pady=4)
        fill_num_sims = Entry(tab2)
        fill_num_sims.grid(row=7,column = 3, pady =2, padx = 2)
        
        boolvar = IntVar()
        boolvar.set(False)
        cb = Checkbutton(tab2, text = "Next Variable", variable = boolvar).grid(row=6,columnspan = 1, column = 2, sticky=W)
        
        otherbool = IntVar()
        otherbool.set(False)
        
        cb = Checkbutton(tab2, text = "Abort", variable = otherbool).grid(row= 6,columnspan = 1, column = 3, sticky=W)
        tab_made = tab2
    elif  analysis_type.get() == 'Single Point Analysis':
        tab3 = ttk.Frame(note)
        note.add(tab3, text = 'Single Point')
        
        Button(tab3,
        text='Calculate MFSP',
        command=single_point_analysis).grid(row=7,
        column=2, columnspan=3,
        sticky=W, pady=4)
        tab_made  = tab3
        
    elif  analysis_type.get() == 'Multivariate Sensitivity':
        tab1 = ttk.Frame(note)
        note.add(tab1,text = "Multivariate Analysis")
        ###############TAB 1 LABELS#################


        Label(tab1, 
              text="Number of Simulations :").grid(row=3, column= 1, sticky = E,pady = 5,padx = 5)
        global sim
        sim = Entry(tab1)
        sim.grid(row=3, column=2,pady = 5,padx = 5)
        
        Label(tab1, 
              text="Save As :").grid(row=4, column= 1, sticky = E,pady = 5,padx = 5)
        save = Entry(tab1)
        save.grid(row=4, column=2,pady = 5,padx = 5)
        
        Label(tab1,text = ".csv").grid(row = 4, column = 3, sticky = W)
        
        otherbool = IntVar()
        otherbool.set(False)
        
        cb = Checkbutton(tab1, text = "Abort", variable = otherbool).grid(row=7,columnspan = 1, column = 3, sticky=W)
        
        
        ###############TAB 1 BUTTONS#################
               
        Button(tab1,
               text='Run Multivariate Analysis',
               command=initialize_multivar_analysis).grid(row=6,
               column=3, columnspan=2,
               sticky=W, 
               pady=4)
        Button(tab1,
               text='Load Variable Distrbutions',
               command=display_distributions(False))).grid(row=6,
               column=1, columnspan=2,
               sticky=W, 
               pady=4)
        
        tab_made = tab1
    load_variables_into_GUI(tab_made)
##############INITIALIZE MAIN ROOT AND TAB###############
root = Tk()

note = ttk.Notebook(root)
note.grid()

tab0 = ttk.Frame(note)
note.add(tab0, text = "File Upload")



###############TAB 0 Buttons##################
Button(tab0, 
        text='Upload Excel Data',
        command=open_excel_file).grid(row=0,
        column=1,
        sticky = E,  
        pady = 5,padx = 5)

excel = Entry(tab0)
excel.grid(row=0, column=2)

Button(tab0, 
      text="Upload Aspen Model",
      command=open_aspen_file).grid(row=1, column = 1,sticky = E,
      pady = 5,padx = 5)
aspen = Entry(tab0)
aspen.grid(row=1, column=2,pady = 5,padx = 5)

Button(tab0, 
      text="Upload Excel Model",
      command=open_solver_file).grid(row=2,column = 1,sticky = E,
      pady = 5,padx = 5)
solver = Entry(tab0)
solver.grid(row=2, column=2,pady = 5,padx = 5)

Button(tab0, 
      text="Load Data",
      command=make_new_tab).grid(row=5,column = 4,sticky = E,
      pady = 5,padx = 5)
solver = Entry(tab0)
solver.grid(row=2, column=2,pady = 5,padx = 5)

master = tab0
analysis_type = StringVar(master)
analysis_type.set("Choose Analysis Type") # default value

analysis_type_options = OptionMenu(tab0, analysis_type, "Univariate Sensitivity", "Single Point Analysis", "Multivariate Sensitivity").grid(row = 5,sticky = E,column = 2,padx =5, pady = 5)

        





mainloop()
  
