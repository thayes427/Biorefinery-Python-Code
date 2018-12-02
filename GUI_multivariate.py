from tkinter import *
from tkinter import messagebox
import numpy as np
import sensitivity_analysis as msens
from tkinter import ttk
from tkinter.filedialog import askopenfilename
import matplotlib
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2TkAgg
from matplotlib import pyplot as pplt
from matplotlib.backend_bases import key_press_handler
from matplotlib.figure import Figure
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
    root.filename = askopenfilename(title = "Select file", filetypes = (("csv files","*.csv"),("all files","*.*")))
    excel.delete(0, END)
    excel.insert(0,root.filename)
    
def open_aspen_file():
    root.filename = askopenfilename(title = "Select file", filetypes = (("Aspen Models",["*.bkp", "*.apw"]),("all files","*.*")))
    aspen.delete(0, END)
    aspen.insert(0,root.filename)

def open_solver_file():
    root.filename = askopenfilename(title = "Select file", filetypes = (("Excel Files","*.xlsm"),("all files","*.*")))
    solver.delete(0, END)
    solver.insert(0,root.filename)
    
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
    
    columns = 2
    num_rows= ((len(vars_to_change) + 1) // columns) + 1
    counter = 1
    fig = Figure(figsize = (5,5))
    a = fig.add_subplot(num_rows,columns,counter)
    counter += 1
    total_MFSP = d_f_output["MFSP"]
    num_bins = 100
    try:
        n, bins, patches = a.hist(total_MFSP, num_bins, facecolor='blue', alpha=0.5)
    except Exception:
        pass
    a.set_title ("MFSP Distribution")
    a.set_xlabel("MFSP ($)")
    if len(vars_to_change) != 0:
        for var in vars_to_change:
            a = fig.add_subplot(num_rows,columns,counter)
            counter += 1
            total_data = d_f_output[var]
            num_bins = 100
            try:
                n, bins, patches = a.hist(total_data, num_bins, facecolor='blue', alpha=0.5)
            except Exception:
                pass
            a.set_title(var)

    a = fig.tight_layout()
    canvas = FigureCanvasTkAgg(fig)
    canvas.get_tk_widget().grid(row=8, column = 0,columnspan = 10, rowspan = 10, sticky= W+E+N+S, pady = 5,padx = 5,)
        
    root.update_idletasks()


counter = 1
old_var_name = ''
def plot_univ_on_GUI(dfstreams, var, value, fortran_check):
    '''
    Allows Unviariate to be plotted on the GUI
    
    It will plot the graphs as follows: the number of rows is the number of variables that have
    to be plotted, and there will be a width of two columns, with the first column graphing 
    the variable distribution, and the second one plotting the MFSP distribution.   
    
    '''
    
    global simulation_dist, counter, old_var_name
    columns = 2
    num_rows= ((len(simulation_dist) + 1) // columns) + 1
    fig = Figure(figsize = (5,5))
    a = fig.add_subplot(num_rows,columns,counter)
    if var == old_var_name and old_var_name != '':
        counter += 2
        old_var_name = var
    counter += 1
    num_bins = 100
    try:
        n, bins, patches = a.hist(value, num_bins, facecolor='blue', alpha=0.5)
    except Exception:
        pass
    a.set_title(var)
    if fortran_check:
        a.set_xticklabels([])
        a.set_xlabel(var)
    a = fig.add_subplot(num_rows,columns,counter)
    counter -= 1
    num_bins = 100
    try:
        n, bins, patches = a.hist(dfstreams['MFSP'], num_bins, facecolor='blue', alpha=0.5)
    except Exception:
        pass
    a.set_title('MFSP - ' + var)
    a = fig.tight_layout()
    canvas = FigureCanvasTkAgg(fig)
    canvas.get_tk_widget().grid(row=8, column = 0,columnspan = 10, rowspan = 10, sticky= W+E+N+S, pady = 5,padx = 5,)
    
    #vsb = ttk.Scrollbar(fig, orient="vertical", command=canvas.yview)
    #vsb.grid(row=8, column=1,sticky = 'ns')
    #canvas.configure(yscrollcommand=vsb.set)
    
    #canvas.config(scrollregion=canvas.bbox("all"))
    root.update_idletasks()

def get_distributions(is_univar):
    global simulation_vars, simulation_dist, univar_var_num_sim
    if is_univar:
        max_num_sim = max(int(slot.get()) for slot in univar_var_num_sim.values())
        simulation_vars, simulation_dist = msens.get_distributions(str(excel.get()), max_num_sim)
        for (aspen_variable, aspen_call, fortran_index), dist in simulation_vars.items():
            if aspen_variable in univar_var_num_sim:
                num_trials_per_var = int(univar_var_num_sim[aspen_variable].get())
                simulation_vars[(aspen_variable, aspen_call, fortran_index)] = dist[:num_trials_per_var]
                simulation_dist[aspen_variable] = dist[:num_trials_per_var]
                
    else:
        simulation_vars, simulation_dist = msens.get_distributions(str(excel.get()), ntrials= int(sim.get())) 
    
    
def plot_init_dist():
    '''
    This function will plot the distribution of variable calls prior to running
    the simulation. This will enable users to see whether the distributions are as they expected.
    
    Inputs:
        simulation_dist: dictionary where the key is the variable name, and values are
            lists of the values that will be used in the function.    
    
    '''
    
    global simulation_dist
    columns = 2
    num_rows= ((len(simulation_dist) + 1) // columns) + 1
    counter = 1
    fig = pplt.figure(figsize = (5,5))
    for var, values in simulation_dist.items():
        a = fig.add_subplot(num_rows,columns,counter)
        counter += 1
        num_bins = 100
        try:
            n, bins, patches = pplt.hist(values, num_bins, facecolor='blue', alpha=0.5)
        except Exception:
            pass
        a.set_title(var)
    a = fig.tight_layout()
    canvas = FigureCanvasTkAgg(fig)
    canvas.get_tk_widget().grid(row=8, column = 0,columnspan = 10, rowspan = 10, sticky= W+E+N+S, pady = 5,padx = 5)
    
    print('HIIIIIIIIII')
    if tab3:
        tab_num = tab3
    elif tab2:
        tab_num  = tab2
    else:
        tab_num = tab1
    frame_canvas = ttk.Frame(canvas.get_tk_widget())
    frame_canvas.grid(row=8, column = 0,columnspan = 10, rowspan = 10, sticky= W+E+N+S, pady = 5,padx = 5)
    frame_canvas.grid_rowconfigure(0, weight=1)
    frame_canvas.grid_columnconfigure(0, weight=1)
    frame_canvas.config(height = '5c')
    #canvas = Canvas(frame_canvas)
    #canvas.grid(row=0, column=0, sticky="news")
    #canvas.config(height = '5c')
    vsb = ttk.Scrollbar(frame_canvas, orient="vertical", command=canvas.get_tk_widget().yview)
    vsb.grid(row=8, column=1,sticky = 'ns')
    canvas.get_tk_widget().configure(yscrollcommand=vsb.set)
    
    canvas.get_tk_widget().config(scrollregion=canvas.bbox("all"))
    root.update_idletasks()
    
def display_distributions(is_univar):
    
    get_distributions(is_univar)
    plot_init_dist()   
    

sp_row_num = None
univar_row_num = None
    
def load_variables_into_GUI(tab_num):
    sens_vars = str(excel.get())
    global sp_row_num, univar_row_num, univar_var_num_sim, tab3, tab1, tab2, note
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
                    univariate_vars.append((row["Variable Name"], row["Format of Range"].strip().lower(), row['Range of Values'].split(',')))
    #now populate the gui with the appropriate tab and variables stored above
    if type_of_analysis == 'Single Point Analysis':
        note.select(tab3)
        tab3.config(width = '5c', height = '5c')
        sp_row_num = 2
        
        # Create a frame for the canvas with non-zero row&column weights
        frame_canvas = ttk.Frame(tab_num)
        frame_canvas.grid(row=sp_row_num, column=1, pady=(5, 0))
        frame_canvas.grid_rowconfigure(0, weight=1)
        frame_canvas.grid_columnconfigure(0, weight=1)
        frame_canvas.config(height = '5c')
        # Set grid_propagate to False resizing later
        #frame_canvas.grid_propagate(False)
        
        # Add a canvas in the canvas frame
        canvas = Canvas(frame_canvas)
        canvas.grid(row=0, column=0, sticky="news")
        canvas.config(height = '5c')
        # Link a scrollbar to the canvas
        vsb = ttk.Scrollbar(frame_canvas, orient="vertical", command=canvas.yview)
        vsb.grid(row=0, column=1,sticky = 'ns')
        canvas.configure(yscrollcommand=vsb.set)
        
        # Create a frame to contain the variables
        frame_vars = ttk.Frame(canvas)
        canvas.create_window((0, 0), window=frame_vars, anchor='nw')
        frame_vars.config(height = '5c')
        
        sp_row_num = 0
        for name,value in single_pt_vars:
            sp_row_num += 1
            key = str(sp_row_num)
            Label(frame_vars, 
            text= name).grid(row=sp_row_num, column= 1, sticky = E,pady = 5,padx = 5)
            key=Entry(frame_vars)
            key.grid(row=sp_row_num, column=2,pady = 5,padx = 5)
            key.delete(first=0,last=END)
            key.insert(0,str(value))
            single_point_var_val[name]= key
            
        # Update vars frames idle tasks to let tkinter calculate variable sizes
        frame_vars.update_idletasks()
        # Determine the size of the Canvas
        
        frame_canvas.config(width='5c', height='10c')
        
        # Set the canvas scrolling region
        canvas.config(scrollregion=canvas.bbox("all"))
            
            

    if type_of_analysis == 'Univariate Sensitivity':
        note.select(tab2)
        univar_row_num = 8
        Label(tab_num, 
            text= 'Variable Name').grid(row=univar_row_num, column= 1,pady = 5,padx = 5, sticky= E)
        Label(tab_num, 
            text= 'Sampling Type').grid(row=univar_row_num, column= 2,pady = 5,padx = 5)
        Label(tab_num, 
            text= '# of Trials').grid(row=univar_row_num, column= 3,pady = 5,padx = 5, sticky = W)
        univar_row_num += 1
        # Create a frame for the canvas with non-zero row&column weights
        frame_canvas1 = ttk.Frame(tab_num)
        frame_canvas1.grid(row=univar_row_num, column=1, columnspan =3, pady=(5, 0))
        frame_canvas1.grid_rowconfigure(0, weight=1)
        frame_canvas1.grid_columnconfigure(0, weight=1)
        frame_canvas1.config(height = '3c')
        # Set grid_propagate to False resizing later
        #frame_canvas.grid_propogate(False)
        
        # Add a canvas in the canvas frame
        canvas1 = Canvas(frame_canvas1)
        canvas1.grid(row=0, column=0, sticky="news")
        canvas1.config(height = '3c')
        
        # Link a scrollbar to the canvas
        vsb = ttk.Scrollbar(frame_canvas1, orient="vertical", command=canvas1.yview)
        vsb.grid(row=0, column=1,sticky = 'ns')
        canvas1.configure(yscrollcommand=vsb.set)
        
        # Create a frame to contain the variables
        frame_vars1 = ttk.Frame(canvas1)
        frame_vars1.config(height = '3c')
        canvas1.create_window((0, 0), window=frame_vars1, anchor='nw')
        univar_row_num =0
        for name, format_of_data, vals in univariate_vars:
            Label(frame_vars1, 
            text= name).grid(row=univar_row_num, column= 1,pady = 5,padx = 5)
            Label(frame_vars1, 
            text= format_of_data).grid(row=univar_row_num, column= 2,pady = 5,padx = 5)
            
            if not(format_of_data == 'linspace' or format_of_data == 'list'):
                key2=Entry(frame_vars1)
                key2.grid(row=univar_row_num, column=3,pady = 5,padx = 5)
                #key2.insert(0,univariate_sims)
                univar_var_num_sim[name]= key2
            else:
                Label(frame_vars1,text= str(len(vals))).grid(row=univar_row_num, column= 3,pady = 5,padx = 5)
            univar_row_num += 1
            
        # Update vars frames idle tasks to let tkinter calculate variable sizes
        frame_vars1.update_idletasks()
        # Determine the size of the Canvas
        
        frame_canvas1.config(width='5c', height='5c')
        
        # Set the canvas scrolling region
        canvas1.config(scrollregion=canvas1.bbox("all"))
            
            
        
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
    
def display_time_remaining(time_remaining = 10):
    '''
    THIS NEEDS TO PRINT OUT ESTIMATED TIME REMAINING
    '''
    if tab1:
        Label(tab1, text = 'Time Remaining :' + str(time_remaining)).grid(row = 15, column= 1)
    if tab2:
        Label(tab2, text = 'Time Remaining :' + str(time_remaining)).grid(row = 15, column= 1)
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
    global simulation_vars
    d_f_output = msens.multivariate_sensitivity_analysis(aspenfile,solverfile,sens_vars,numtrial,outputfile, simulation_vars)
        
def run_univ_sens():
    aspenfile= str(aspen.get())
    solverfile= str(solver.get())
    outputfile= str(save2.get())
    sens_vars = str(excel.get())
    global simulation_vars
    for (aspen_variable, aspen_call, fortran_index), values in simulation_vars.items():
        msens.univariate_analysis(aspenfile, solverfile, aspen_call, aspen_variable, values, fortran_index, outputfile)
        
        print('Finished Analysis for Variable: ', aspen_variable)
    print('-----------FINISHED-------------')

def single_point_analysis():
    aspenfile= str(aspen.get())
    solverfile= str(solver.get())
    outputfile= str(save_sp.get())
    sens_vars = str(excel.get())
    global sp_row_num, single_point_var_val
    sp_vars, throwaway = msens.get_distributions(sens_vars, 1)
    for (aspen_variable, aspen_call, fortran_index), values in sp_vars.items():
        sp_vars[(aspen_variable, aspen_call, fortran_index)] = [float(single_point_var_val[aspen_variable].get())]
    mfsp = msens.multivariate_sensitivity_analysis(aspenfile,solverfile,sens_vars, 1,output_file, sp_vars, disp_graphs=False).at[0, 'MFSP']
    Label(tab3, text= 'MFSP = ' + str(mfsp)).grid(row=sp_row_num+1, column = 1)
    
    
    return None

def fill_num_trials():
    global fill_num_sims, univar_var_num_sim
    ntrials = fill_num_sims.get()
    for name, slot in univar_var_num_sim.items():
        slot.delete(0, END)
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
save_sp = None 

def make_new_tab():
    global save_sp,sim, sim2, save2, save, otherbool, show_plot, boolvar, cb, tab1, tab2, tab3, fill_num_sims, note
    
    if tab1:
        note.forget(tab1)
        tab1 = None
    if tab2:
        note.forget(tab2)
        tab2 = None
    if tab3:
        note.forget(tab3)
        tab3 = None
    if analysis_type.get() == 'Choose Analysis Type':
        print("ERROR: Select an Analysis")
    elif  analysis_type.get() == 'Univariate Sensitivity':
        tab2 = ttk.Frame(note)
        note.add(tab2,text = "Univariate Analysis")
        ##############Tab 2 LABELS##################
        
        Label(tab2, 
              text="Save As :").grid(row=4, column= 1, sticky = E,pady = 5,padx = 5)
        save2 = Entry(tab2)
        save2.grid(row=4, column=2,pady = 5,padx = 5)
        
        Label(tab2,text = ".csv").grid(row = 4, column = 3, sticky = W)
        ##############Tab 2 Buttons###############
        Label(tab2, text ='').grid(row= 13, column =1)
        Button(tab2,

               text='Run Univariate Sensitivity Analysis',
               command=initialize_univar_analysis).grid(row=14,
               column=3, columnspan=2,
               pady=4)
        Button(tab2,
               text='Display Variable Distributions',
               command=lambda: display_distributions(True)).grid(row=14,
               column=1, columnspan=2, sticky = W,
               pady=4)
        Button(tab2,
               text='Fill Simulations',
               command=fill_num_trials).grid(row=7, columnspan = 2, sticky =E,
               column=1,
               pady=4)
        fill_num_sims = Entry(tab2)
        fill_num_sims.grid(row=7,column = 3,sticky =W, pady =2, padx = 2)
        fill_num_sims.config(width = 10)
        
        options= ttk.Labelframe(tab2, text='Run Options:')
        options.grid(row = 15,column = 3, pady = 10,padx = 10)
        
        

        boolvar = IntVar()
        boolvar.set(False)
        cb = Checkbutton(options, text = "Next Variable", variable = boolvar).grid(row=6,columnspan = 1, column = 2, sticky=W)
        
        otherbool = IntVar()
        otherbool.set(False)
        
        cb = Checkbutton(options, text = "Abort", variable = otherbool).grid(row= 6,columnspan = 1, column = 3, sticky=W)
        tab_made = tab2
    elif  analysis_type.get() == 'Single Point Analysis':
        tab3 = ttk.Frame(note)
        note.add(tab3, text = 'Single Point')
         
        Label(tab3, 
              text="Save As :").grid(row=0, column= 0, sticky = E,pady = 5,padx = 5)
        save_sp = Entry(tab3)
        save_sp.grid(row=0, column=1,pady = 5,padx = 5)
        
        Button(tab3,
        text='Calculate MFSP',
        command=single_point_analysis).grid(row=7,
        column=1, columnspan=2, pady=4)
        tab_made  = tab3
        
    elif  analysis_type.get() == 'Multivariate Sensitivity':
        tab1 = ttk.Frame(note)
        note.add(tab1,text = "Multivariate Analysis")
        note.select(tab1)
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
               text='Display Variable Distributions',
               command=lambda: display_distributions(False)).grid(row=6,
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
'''
frame= ttk.Frame(note)
note.add(frame,text= 'Scrolling Bar Test')

test2= ttk.Frame(note)
note.add(test2,text= 'Scrolling Test Two')
'''

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
      command=make_new_tab).grid(row=5,column = 3,sticky = E,
      pady = 5,padx = 5)
solver = Entry(tab0)
solver.grid(row=2, column=2,pady = 5,padx = 5)

master = tab0
analysis_type = StringVar(master)
analysis_type.set("Choose Analysis Type") # default value

analysis_type_options = OptionMenu(tab0, analysis_type,"Single Point Analysis","Univariate Sensitivity", "Multivariate Sensitivity").grid(row = 5,sticky = E,column = 2,padx =5, pady = 5)
'''
###################Scroll BAR#####################
label1 = ttk.Label(frame, text="Label 1")
label1.grid(row=0, column=0, pady=(5, 0), sticky='nw')

label2 = ttk.Label(frame, text="Label 2")
label2.grid(row=1, column=0, pady=(5, 0), sticky='nw')

label3 = ttk.Label(frame, text="Label 3")
label3.grid(row=3, column=0, pady=5, sticky='nw')

frame_canvas = ttk.Frame(frame)
frame_canvas.grid(row=2, column=0, pady=(5, 0), sticky='nw')
frame_canvas.grid_rowconfigure(0, weight=1)
frame_canvas.grid_columnconfigure(0, weight=1)
# Set grid_propagate to False to allow 5-by-5 buttons resizing later
frame_canvas.grid_propagate(False)

canvas=Canvas(frame_canvas,width=300,height=300,scrollregion=(0,0,500,500))

hbar=Scrollbar(frame_canvas,orient=HORIZONTAL)
hbar.pack(side=BOTTOM,fill=X)
hbar.config(command=canvas.xview)
vbar=Scrollbar(frame_canvas,orient=VERTICAL)
vbar.pack(side=RIGHT,fill=Y)
vbar.config(command=canvas.yview)
canvas.config(width=300,height=300)
canvas.config(xscrollcommand=hbar.set, yscrollcommand=vbar.set)
canvas.pack(side=LEFT,expand=True,fill=BOTH)


####################




label1 = ttk.Label(test2, text="Label 1")
label1.grid(row=0, column=0, pady=(5, 0), sticky='nw')

label2 = ttk.Label(test2, text="Label 2")
label2.grid(row=1, column=0, pady=(5, 0), sticky='nw')

label3 = ttk.Label(test2, text="Label 3")
label3.grid(row=3, column=0, pady=5, sticky='nw')

# Create a frame for the canvas with non-zero row&column weights
frame_canvas = ttk.Frame(test2)
frame_canvas.grid(row=2, column=0, pady=(5, 0), sticky='nw')
frame_canvas.grid_rowconfigure(0, weight=1)
frame_canvas.grid_columnconfigure(0, weight=1)
# Set grid_propagate to False to allow 5-by-5 buttons resizing later
frame_canvas.grid_propagate(False)

# Add a canvas in that frame
canvas = Canvas(frame_canvas, bg="yellow")
canvas.grid(row=0, column=0, sticky="news")

# Link a scrollbar to the canvas
vsb = ttk.Scrollbar(frame_canvas, orient="vertical", command=canvas.yview)
vsb.grid(row=0, column=1, sticky='ns')
canvas.configure(yscrollcommand=vsb.set)

# Create a frame to contain the buttons
frame_buttons = ttk.Frame(canvas)
canvas.create_window((0, 0), window=frame_buttons, anchor='nw')

# Add 9-by-5 buttons to the frame
rows = 9
columns = 5
buttons = [[ttk.Button() for j in range(columns)] for i in range(rows)]
for i in range(0, rows):
    for j in range(0, columns):
        buttons[i][j] = ttk.Button(frame_buttons, text=("%d,%d" % (i+1, j+1)))
        buttons[i][j].grid(row=i, column=j, sticky='news')

# Update buttons frames idle tasks to let tkinter calculate buttons sizes
frame_buttons.update_idletasks()

# Resize the canvas frame to show exactly 5-by-5 buttons and the scrollbar
first5columns_width = sum([buttons[0][j].winfo_width() for j in range(0, 5)])
first5rows_height = sum([buttons[i][0].winfo_height() for i in range(0, 5)])
frame_canvas.config(width=first5columns_width + vsb.winfo_width(),
                    height=first5rows_height)

# Set the canvas scrolling region
canvas.config(scrollregion=canvas.bbox("all"))
'''
mainloop()
  
