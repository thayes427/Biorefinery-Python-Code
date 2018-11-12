from tkinter import *
from tkinter import messagebox
import numpy as np
import obj2_cleaned_up_box as msens
from tkinter import ttk
from tkinter.filedialog import askopenfilename

#reload(mod)

################Tab 1 Funcitons###############
 
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
                                                title = "Select file",
                                                filetypes = (("csv files","*.csv")))
    solver.insert(0,root.filename)

def run_multivar_sens():
    aspenfile= str(aspen.get())
    solverfile= str(solver.get())
    numtrial= int(sim.get())
    outputfile= str(save.get())
    sens_vars = str(excel.get())
    msens.multivariate_sensitivity_analysis(aspenfile,solverfile,sens_vars,numtrial,outputfile)
##############INITIALIZE ROOT AND TABS###############
root = Tk()

note = ttk.Notebook(root)
note.grid()

tab1 = ttk.Frame(note)
note.add(tab1,text = "Sensitivity Analysis")


###############TAB 1 LABELS#################
Button(tab1, 
        text='Upload Excel Data',
        command=open_excel_file).grid(row=0,
        column=1,
        sticky = E,  
        pady = 5,padx = 5)

excel = Entry(tab1)
excel.grid(row=0, column=2)

Button(tab1, 
      text="Upload Aspen Model",
      command=open_aspen_file).grid(row=1, column = 1,sticky = E,
      pady = 5,padx = 5)
aspen = Entry(tab1)
aspen.grid(row=1, column=2,pady = 5,padx = 5)

Button(tab1, 
      text="Upload Excel Model",
      command=open_solver_file).grid(row=2,column = 1,sticky = E,
      pady = 5,padx = 5)
solver = Entry(tab1)
solver.grid(row=2, column=2,pady = 5,padx = 5)

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



mainloop()
  
