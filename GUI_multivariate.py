from tkinter import *
from tkinter import messagebox
import numpy as np
import obj2_cleaned_up_box as msens
from tkinter import ttk
from tkinter.filedialog import askopenfilename
import matplotlib
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2TkAgg
from matplotlib import pyplot as pplt

#reload(mod)

################Tab 1 Functions###############
 
def quit():
    global root
    root.destroy()


def open_excel_file():
    root.filename = askopenfilename(initialdir = "/",
                                                title = "Select file")
                                    
    excel.insert(0,root.filename)
    
def open_aspen_file():
    root.filename = askopenfilename(initialdir = "/",
                                                title = "Select file")
    aspen.insert(0,root.filename)

def open_solver_file():
    root.filename = askopenfilename(initialdir = "/",
                                                title = "Select file")
    solver.insert(0,root.filename)
    
def plot_on_GUI(d_f_output):
    
    fig = pplt.figure()
    a = fig.add_subplot(111)
    total_MFSP = d_f_output["MFSP"]
    num_bins = 100
    n, bins, patches = pplt.hist(total_MFSP, num_bins, facecolor='blue', alpha=0.5)
    a.set_title ("MFSP Histogram", fontsize=16)
    a.set_ylabel("Count", fontsize=14)
    a.set_xlabel("MFSP ($)", fontsize=14)
    canvas = FigureCanvasTkAgg(fig)
    canvas.get_tk_widget().grid(row=8, column = 0,columnspan = 2, rowspan = 2, sticky= W+E+N+S, pady = 5,padx = 5,)
    root.update_idletasks()
    
def run_multivar_sens():
    aspenfile= str(aspen.get())
    solverfile= str(solver.get())
    numtrial= int(sim.get())
    outputfile= str(save.get())
    sens_vars = str(excel.get())
    graph_plot = int(show_plot.get())
    d_f_output = msens.multivariate_sensitivity_analysis(aspenfile,solverfile,sens_vars,numtrial,outputfile, graph_plot)
        
def run_univ_sens():
    aspenfile= str(aspen.get())
    solverfile= str(solver.get())
    numtrial= int(sim2.get())
    outputfile= str(save2.get())
    sens_vars = str(excel.get())
    simulation_vars = msens.get_distributions(sens_vars, numtrial)
    for (aspen_variable, aspen_call, fortran_index), values in simulation_vars.items():
        msens.univariate_analysis(aspenfile, solverfile, aspen_call, aspen_variable, values, fortran_index, outputfile)
        print('Finished Analysis for Variable: ', aspen_variable)
    print('-----------FINISHED-------------')

        
       
               
       


##############INITIALIZE ROOT AND TABS###############
root = Tk()

note = ttk.Notebook(root)
note.grid()

tab0 = ttk.Frame(note)
note.add(tab0, text = "File Upload")

tab3 = ttk.Frame(note)
note.add(tab3, text = 'Single Point')

tab1 = ttk.Frame(note)
note.add(tab1,text = "Multivariate Analysis")

tab2 = ttk.Frame(note)
note.add(tab2,text = "Univariate Analysis")

###############TAB 0 LABEL##################
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

Label(tab1,text = ".csv").grid(row = 4, column = 3, sticky = W)
###############TAB 1 LABELS#################


Label(tab1, 
      text="Number of Simulations :").grid(row=3, column= 1, sticky = E,pady = 5,padx = 5)
sim = Entry(tab1)
sim.grid(row=3, column=2,pady = 5,padx = 5)

Label(tab1, 
      text="Save As :").grid(row=4, column= 1, sticky = E,pady = 5,padx = 5)
save = Entry(tab1)
save.grid(row=4, column=2,pady = 5,padx = 5)





###############TAB 1 BUTTONS#################
       
Button(tab1,
       text='Run Monte Carlo Simulation',
       command=run_multivar_sens).grid(row=7,
       column=4, columnspan=3,
       sticky=W, 
       pady=4)

show_plot = IntVar()
Checkbutton(tab1, text="Generate MFSP Distribution (Graph)", variable=show_plot).grid(row=5,columnspan = 2, column = 0, sticky=W)

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
       command=run_univ_sens).grid(row=7,
       column=4, columnspan=3,
       sticky=W, 
       pady=4)
show_plot = IntVar()
Checkbutton(tab2, text="Generate MFSP Distribution (Graph)", variable=show_plot).grid(row=5,columnspan = 2, column = 0, sticky=E)



mainloop()
  
