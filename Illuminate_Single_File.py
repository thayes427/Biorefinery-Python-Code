# -*- coding: utf-8 -*-
"""
Created on Sat Dec 15 19:58:47 2018

@author: MENGstudents


Table of Contents:
    Tkinter MainApp Class Lines approx. 40-2000
    Simulation Object Lines approx. 2000-2200
    Process with Error Support Class approx 2200-2230
    Global Functions for 2230-2600
    Compatibility Test Class approx. 2600-END 
"""

from tkinter import Tk, StringVar,E,W,Canvas,END, LEFT,IntVar, Checkbutton, Label, Toplevel
from tkinter.ttk import Entry, Button, Radiobutton, OptionMenu, Labelframe, Scrollbar, Notebook, Frame
from tkinter.filedialog import askopenfilename
from threading import Thread
from pandas import DataFrame, concat, isna, read_excel, ExcelWriter
from multiprocessing import Value, cpu_count, freeze_support, Manager, Process, Pipe
from multiprocessing import Queue as mpQueue
import traceback
from time import time, sleep
from datetime import datetime
from numpy import linspace, random, histogram, subtract, percentile
from psutil import process_iter, virtual_memory
from os import path, makedirs, rmdir, listdir
from shutil import copyfile, rmtree
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
from winreg import EnumKey, CreateKey, EnumValue, HKEY_CLASSES_ROOT
from queue import Queue
from re import search, findall
from random import choices, choice
from math import ceil, isclose
from textwrap import wrap
import pythoncom
import matplotlib.pyplot as plt
import string
from win32com.client import Dispatch, DispatchEx

 

class MainApp(Tk):
    '''
    MainApp object encapsulates all attributes and methods related to the GUI.
    '''

    def __init__(self):
        '''
        Initialize the Tkinter notebook and construct the home tab.
        '''
        Tk.__init__(self)
        #self.iconbitmap('01_128x128.ico')
        self.notebook = Notebook(self)
        self.wm_title("Illuminate")
        self.notebook.grid()
        self.win_lim_x = self.winfo_screenwidth()//2
        self.win_lim_y = int(self.winfo_screenheight()*0.9)
        self.simulations = []
        self.current_simulation = None
        self.sp_error = None
        self.current_tab = None
        self.abort = Value('b', False)
        self.abort_univar_overall = Value('b', False)
        self.simulation_vars = {}
        self.attributes('-topmost', True)
        self.focus_force()
        self.bind('<FocusIn>', OnFocusIn) # used to "focus" on the GUI without forcing it to be in focus
        self.tot_sim_num = 0
        self.sims_completed = Value('i',0)
        self.start_time = None
        self.univar_plot_counter = 1
        self.finished_figures = []
        self.univar_row_num=0
        self.last_results_plotted = None
        self.last_update = None
        # setting the size and location of the Tkinter window
        self.geometry(str(self.win_lim_x) + 'x' + str(self.win_lim_y) + '+0+0')
        self.worker_thread = None
        self.display_tab = None
        self.mapping_pdfs = {}
        self.simulation_dist, self.simulation_vars = {}, {}
        self.graphing_frequency = None
        self.analysis_type_error = None
        self.missing_inputs_error = None
        self.temp_directory = None
        self.warning_keywords = set()
        self.vars_have_been_sampled = False
        self.sp_results_text = []
        self.sp_status = None

        
        self.construct_home_tab()
        
        
    def construct_home_tab(self):
        '''
        Constructs the home tab used to upload input files and test compatibility 
        '''
        self.load_aspen_versions()
        self.home_tab = Frame(self.notebook, style= 'frame.TFrame')
        self.notebook.add(self.home_tab, text = 'File Upload Tab')

        for i in range (5,20):
            Label(self.home_tab, text='                       ').grid(
                    row=100,column=i,columnspan=1)

        for i in range(106,160):
            Label(self.home_tab, text=' ').grid(row=i,column=0,columnspan=1)

        space= Label(self.home_tab, text=" ",font='Helvetica 2')
        space.grid(row=0, column= 1, sticky = E, padx = 5, pady =4)
        space.rowconfigure(0, minsize = 15)
        
        Button(self.home_tab, text='Upload Simulation Inputs',
        command=self.open_input_excel_file).grid(row=1,column=1, sticky = E, pady = 5,
                                    padx = 5)

        self.input_csv_entry = Entry(self.home_tab)
        self.input_csv_entry.grid(row=1, column=2)
        
        Button(self.home_tab, text="Upload Aspen Model",
        command=self.open_aspen_file).grid(row=2, column = 1,sticky = E, pady = 5,padx = 5)
        
        self.aspen_file_entry = Entry(self.home_tab)
        self.aspen_file_entry.grid(row=2, column=2,pady = 5,padx = 5)
        
        Button(self.home_tab, text="Upload Excel Calculator", command=self.open_excel_calculator_file).grid(
                row=3,column = 1,sticky = E,pady = 5,padx = 5)
        
        self.excel_solver_entry = Entry(self.home_tab)
        self.excel_solver_entry.grid(row=3, column=2,pady = 5,padx = 5)
        

        Button(self.home_tab, text="Load Data", command=self.construct_analysis_tab).grid(
                row=9, column = 3, sticky = E, pady = 5,padx = 5)
        
        self.analysis_type = StringVar(self.home_tab)
        self.analysis_type.set("Choose Analysis Type")
        
        OptionMenu(self.home_tab, self.analysis_type,"Choose Analysis Type", 
                   "Single Point Analysis","Univariate Sensitivity", 
                "Multivariate Sensitivity", style = 'raised.TMenubutton').grid(
                        row = 9,sticky = E,column = 2,padx =5, pady = 5)
                        
        select_aspen = Labelframe(self.home_tab, text='Select Aspen Version:')
        select_aspen.grid(row = 5,column = 1, columnspan = 3, sticky = W,
                          pady = 10,padx = 10)

        self.select_version = StringVar()
        row = 6
        column = 0
        aspen_versions = []
        for key,value in self.aspen_versions.items():
            aspen_versions.append(key + '      ')  
        aspen_versions.sort(key=lambda x: -1*float(x[1:-6]))
        for i, version in enumerate(aspen_versions):
            v = Radiobutton(select_aspen, text= version, 
                            variable=self.select_version, 
                            value = self.aspen_versions[version[:-6]])
            v.grid(row=row,column= column, sticky=W)
            if i == 0:
                v.invoke()
            column += 1
            if column == 4:
                column = 0
                row += 1
        
    def print_compatibility_test_status(self):
        '''
        While the compatibility test is running in another thread, this function
        checks a queue of error statements and updates and prints them to the
        GUI to notify the user of the status of the compatibility test.
        '''
        
        char_per_row = 100
        while not self.compatibility_test.status_queue.empty():
            is_error, text = self.compatibility_test.status_queue.get()
            if is_error and not ( "Finished" in text or "Cannot test" in text):
                text = 'ERROR: ' + text
            lines = wrap(text, char_per_row)
            line_num = len(lines)
            text = '\n'.join(lines)
            if is_error:
                Label(self.compat_test_window, text=text, font='Helvetica 10',
                      fg='red',justify=LEFT).place(x= self.compat_x_pos,y=self.compat_y_pos)
                self.compat_y_pos = self.compat_y_pos+20*line_num
            else:
                if 'SUCCESS' in text:
                    Label(self.compat_test_window, text= text, font='Helvetica 10', 
                          justify=LEFT, fg='green').place(x= self.compat_x_pos, y=self.compat_y_pos)
                    self.compat_y_pos = self.compat_y_pos+20*line_num
                    
                elif text == 'Finished with Compatibility Test':
                    Label(self.compat_test_window, text= text, font='Helvetica 10 bold', 
                          justify=LEFT).place(x= self.compat_x_pos, y=self.compat_y_pos)
                    self.compat_y_pos = self.compat_y_pos+20*line_num
                else:
                    Label(self.compat_test_window, text= text, font='Helvetica 10',justify=LEFT).place(x= 
                         self.compat_x_pos, y=self.compat_y_pos)
                    self.compat_y_pos = self.compat_y_pos+20*line_num

        # keep calling this function until the compatibility test is complete
        # and the status queue is empty
        if self.compat_test_thread.isAlive() or not self.compatibility_test.status_queue.empty():
            self.after(100, self.print_compatibility_test_status)
            
        
        
    def test_compatibility(self):
        '''
        Intializes a thread to run compatibility_test which is found in 
        Illuminate_Test_Compatibility
        '''
        
        self.compat_test_window = Toplevel()
        self.compat_test_window.wm_title('Illuminate Compatibility Test')
        self.compat_test_window.wm_geometry(str(self.win_lim_x) + 'x' + str(self.win_lim_y//2) + '+100+100')
        self.compat_test_window.config()
        self.compat_test_window.focus_force()
        self.compat_test_window.bind('<FocusIn>', OnFocusIn) # used to "focus" on the GUI without forcing it to be in focus
        self.compatibility_test = Compatibility_Test()
        self.compat_y_pos= 0#self.win_lim_y *.03 + 35
        self.compat_x_pos= 0#self.win_lim_x *.59 - 150
        self.compat_test_thread = Thread(target=lambda: self.compatibility_test.compatibility_test(
                str(self.input_csv_entry.get()), str(self.excel_solver_entry.get()),
                str(self.aspen_file_entry.get()), str(self.select_version.get())))
        self.compat_test_thread.start()
        self.after(100, self.print_compatibility_test_status)        


    def construct_analysis_tab(self):
        '''
        Depending on the type of analysis chosen, construct a new analysis tab
        and populate all of the entries and buttons on that tab. After
        constructing the analysis tab, this function calls load_input_variables_into_GUI
        to process the user inputs and fill the analysis tabs with the appropriate
        information.
        '''
        

        # if there was previously error messages about analysis type or missing inputs, delete them
        if self.analysis_type_error:
            self.analysis_type_error.destroy()
            self.analysis_type_error = None
        if self.missing_inputs_error:
            self.missing_inputs_error.destroy()
            self.missing_inputs_error = None
            
        if any(len(file.get()) == 0 for file in (
                self.input_csv_entry, self.excel_solver_entry, self.aspen_file_entry)):
            self.missing_inputs_error = Label(
                    self.home_tab, text='ERROR: Please Provide All Input Files', fg='red')
            self.missing_inputs_error.grid(row=10,column=2, columnspan=2)
            return
            
        if self.current_tab:
            # if another analysis tab is open, delete it and forget it
            self.notebook.forget(self.current_tab)
            self.current_tab = None
        if self.analysis_type.get() == 'Choose Analysis Type':
            self.analysis_type_error = Label(self.home_tab,
                                             text='ERROR: Choose An Analysis Type', 
                                             fg='red')
            self.analysis_type_error.grid(row=10,column=2, columnspan=2)
 
        elif self.analysis_type.get() == 'Univariate Sensitivity':
            self.current_tab = Frame(self.notebook)
            self.notebook.add(self.current_tab,text = "Univariate Analysis")
            
            Label(self.current_tab, 
                  text="Save As :").place(x=149,y=6)
            self.save_as_entry= Entry(self.current_tab)
            self.save_as_entry.grid(row=4, column=2, sticky=E, pady=6)
            self.save_as_entry.config(width =18)
            
            Label(self.current_tab,text = ".xlsx").grid(row = 4, column = 3, 
                 sticky = W)
            
            Label(self.current_tab, text='CPU Core Count :').place(x=104,y=39)
            self.num_processes_entry = Entry(self.current_tab)
            self.num_processes_entry.grid(row=5, column=2, sticky=E, pady=6)
            self.num_processes_entry.config(width=18)
            
            rec_core = int(cpu_count()//2)
            Label(self.current_tab, text = 'Recommended Count: ' + str(rec_core)).grid(
                    row = 5, column = 3, sticky = W)
            
            Label(self.current_tab, text = 'Graphing Frequency:').place(x=90, y=72)
            self.graphing_freq_entry = Entry(self.current_tab)
            self.graphing_freq_entry.grid(row=6, column=2, sticky=E, pady=6)
            self.graphing_freq_entry.config(width=18)
            
            Label(self.current_tab, text = '(Input 0 for no Graphs)').grid(
                    row=6,column=3, sticky = W)
                        
            Label(self.current_tab, text ='').grid(row= 13, column =1)
            Button(self.current_tab,
                   text='Run Univariate Sensitivity Analysis',
                   command=self.initialize_univar_analysis).grid(row=14,
                   column=3, columnspan=2,
                   pady=4)
            self.save_bkp = IntVar()
            save_bkp= Checkbutton(self.current_tab, text = "Save .bkp Files", 
                                  variable=self.save_bkp)
            save_bkp.grid(row = 13, column = 3, columnspan =2, pady=4)
            
            Button(self.current_tab,
                   text='Sample and Display Variable Distributions',
                   command=self.plot_variable_distributions).grid(row=14,
                   column=1, columnspan=2, sticky = W,
                   pady=4, padx=6)
            Button(self.current_tab,
                   text='Fill  # Trials',
                   command=self.fill_num_trials).grid(row=7, columnspan = 2, sticky =E,
                   column=1,
                   pady=4)
            self.fill_num_sims = Entry(self.current_tab)
            self.fill_num_sims.grid(row=7,column = 3,sticky =W, pady =2, padx = 2, 
                                    columnspan=2)
            self.fill_num_sims.config(width = 10)
            
        elif  self.analysis_type.get() == 'Single Point Analysis':
            self.current_tab = Frame(self.notebook)
            self.notebook.add(self.current_tab, text = 'Single Point')
             
            Label(self.current_tab, 
                  text="Save As :").grid(row=0, column= 0, sticky = E, pady = 5, padx = 5)
            self.save_as_entry = Entry(self.current_tab)
            self.save_as_entry.grid(row=0, column=1, pady = 5)
            Label(self.current_tab,text = ".xlsx").place(x = 295, y= 6)
            
            self.save_bkp = IntVar()
            save_bkp= Checkbutton(self.current_tab, text = "Save .bkp Files", 
                                  variable=self.save_bkp)
            save_bkp.grid(row = 3, column = 2, columnspan =2, pady=4)
            
            Button(self.current_tab, text='Run Analysis',
            command=self.initialize_single_point).grid(row=3,
            column=1, columnspan=2, pady=4)
            
        elif  self.analysis_type.get() == 'Multivariate Sensitivity':
            self.current_tab = Frame(self.notebook)
            self.notebook.add(self.current_tab,text = "Multivariate Analysis")
            
            Label(self.current_tab, 
                  text="Save As :").grid(row=3, column= 1, sticky = E, pady = 5, padx = 5)
            self.save_as_entry = Entry(self.current_tab)
            self.save_as_entry.grid(row=3, column=2,pady = 5,padx = 5)
            Label(self.current_tab,text = ".xlsx").grid(row = 3, column = 3, sticky = W)
            Label(self.current_tab, 
                  text="Number of Simulations :").grid(row=4, column= 1, sticky = E, 
                                                pady = 5, padx = 5)
            self.num_sim_entry = Entry(self.current_tab)
            self.num_sim_entry.grid(row=4, column=2,pady = 5,padx = 5)
            
            rec_core = int(cpu_count()//2)
            Label(self.current_tab, text='CPU Core Count (Recommend '+ str(rec_core)+ '):').grid(
                    row=5, column=1, sticky=E)
            self.num_processes_entry = Entry(self.current_tab)
            self.num_processes_entry.grid(row=5, column=2, pady=5, padx=5)
                               
            Button(self.current_tab,
                   text='Run Multivariate Analysis',
                   command=self.initialize_multivar_analysis).grid(row=7,
                   column=3, columnspan=2, sticky=W, pady=4)
            
            self.save_bkp = IntVar()
            save_bkp= Checkbutton(self.current_tab, text = "Save .bkp Files", 
                                  variable=self.save_bkp)
            save_bkp.grid(row = 6, column = 3, columnspan =2, sticky=W, pady=4)            

            Label(self.current_tab, text='Plotting Frequency (0 for No Plots):').grid(
                    row=6, column=1, sticky=E, pady=5, padx=5)
            self.graphing_freq_entry = Entry(self.current_tab)
            self.graphing_freq_entry.grid(row=6, column=2)
            
            Button(self.current_tab,
                   text='Sample and Display Variable Distributions',
                   command=self.plot_variable_distributions).grid(row=7,
                   column=1, columnspan=2, sticky=W, pady=4, padx=6)
            
        self.load_input_variables_into_GUI()
        self.notebook.select(self.current_tab)
        
    def standardize_graph_title(self, s, pad=False):
        '''
        It will pad the string for a graph title to make it 37 characters long, 
        or shorten the string if it is more than 37 characters
        
        inputs: 
        s: a string to be used as graph title
        pad: an indicator for whether string needs padding
        '''
        
        if len(s) > 37:
            return s[:34] + '...'
        elif pad:
            return s.ljust(37)    
        return s

    def load_aspen_versions(self):
        '''
        Searches through the user's windows registry to find the versions of
        Aspen installed on the computer. It stores these versions in a dictionary
        called self.aspen_versions where the key is the version name and the value
        is the CLSID. The CLSID is used when dispatching the Aspen COMS object.
        '''
        
        key = CreateKey(HKEY_CLASSES_ROOT, '')
        stop = False
        i=0
        versions = dict()
        while not stop:
            try: 
                if search(r"Apwn.Document", (EnumKey(key,i))):
                    subkey = CreateKey(key, EnumKey(key, i))
                    try:
                        subbkey = CreateKey(subkey, 'DefaultIcon')
                    except:
                        i += 1
                        continue
                    default_icon = EnumValue(subbkey, 0)
                    version = search(r"V(\d)+.\d+", default_icon[1])
                    clsid_key = CreateKey(subkey, 'CLSID')
                    CLSID = EnumValue(clsid_key, 0)[1]
                    if version:
                        versions[version.group()] = CLSID
            except: stop = True
            i += 1
        self.aspen_versions = versions
        
    def load_input_variables_into_GUI(self):
        '''
        This function imports the input variables of interest and checks that there
        are not any duplicates.If Single Point analysis or Univariate Analysis are selected,
        the function will create a canvas on the Analysis Tab and upload all the input
        variables into the GUI in accordance with the the variable distibutions selected.
        '''
        ################### LOAD INPUT VARIABLES #############################
        
        single_pt_vars = []
        univariate_vars = []
        variable_names = set()
        type_of_analysis = self.analysis_type.get()
        gui_excel_input = str(self.input_csv_entry.get())
        col_types = {'Variable Name': str, 'Variable Aspen Call': str, 'Distribution Parameters': str,
                     'Bounds': str, 'Fortran Call':str, 'Fortran Value to Change': str,
                     'Distribution Type': str, 'Toggle': bool}
        
        df = read_excel(gui_excel_input, sheet_name='Inputs', dtype=col_types)
        for index, row in df.iterrows():
            if row['Toggle']:    
                if row['Variable Name'] in variable_names:
                    if type_of_analysis == 'Single Point Analysis':
                        x,y=50,310
                    elif type_of_analysis == 'Univariate Sensitivity':
                        x,y=73,315
                    else:
                        x,y=60,145
                    Label(self.current_tab, text='Note: multiple instances of same variable name ' +\
                          'in input file. \nOnly first instance is received as input.',fg='red').place(
                          x=x, y=y)
                    continue
                variable_names.add(row['Variable Name'])
                if type_of_analysis =='Single Point Analysis':
                    single_pt_vars.append((row["Variable Name"], float(row["Distribution Parameters"
                           ].split(',')[0].strip())))
                else:
                    univariate_vars.append((
                            row["Variable Name"], row["Distribution Type"].strip().lower(
                                    ), row['Distribution Parameters'].split(',')))
                        
        #################### PLACE VARIABLES INTO GUI #########################
        
        #now populate the gui with the appropriate tab and variables stored above
        if type_of_analysis == 'Single Point Analysis':
            self.current_tab.config(width = '10c', height = '5c')
            self.sp_value_entries = {}
            
            # Create a frame for the canvas with non-zero row&column weights
            frame_canvas = Labelframe(self.current_tab, text= 'Input Variables:')
            frame_canvas.grid(row=2, column=0, pady=(5, 0), columnspan =3)
            frame_canvas.grid_rowconfigure(0, weight=1)
            frame_canvas.grid_columnconfigure(0, weight=1)
            frame_canvas.config(height = '5c', width='10c')
            
            # Add a canvas in the canvas frame
            canvas = Canvas(frame_canvas)
            canvas.grid(row=0, column=0, sticky="news")
            canvas.config(height = '5c', width='10c')
            
            # Link a scrollbar to the canvas
            vsb = Scrollbar(frame_canvas, orient="vertical", command=canvas.yview,
                            style='scroll.Vertical.TScrollbar')
            vsb.grid(row=0, column=1,sticky = 'ns')
            canvas.configure(yscrollcommand=vsb.set)
            
            # Create a frame to contain the variables
            frame_vars = Frame(canvas)
            canvas.create_window((0, 0), window=frame_vars, anchor='nw')
            frame_vars.config(height = '5c', width='10c')
            
            self.sp_row_num = 0
            for name,value in single_pt_vars:
                self.sp_row_num += 1
                Label(frame_vars, text= self.standardize_graph_title(name,pad=True)
                ).grid(row=self.sp_row_num, column= 1, sticky = E,pady = 5,padx = 5)
                sp_val=Entry(frame_vars)
                sp_val.grid(row=self.sp_row_num, column=2,pady = 5,padx = 5)
                sp_val.delete(first=0,last=END)
                sp_val.insert(0,str(value))
                self.sp_value_entries[name]= sp_val
                
            # Determine the size of the Canvas
            frame_vars.update_idletasks()
            frame_canvas.config(width='10c', height='5c')
            # Set the canvas scrolling region
            canvas.config(scrollregion=canvas.bbox("all"))
    
        if type_of_analysis == 'Univariate Sensitivity':
            self.univar_ntrials_entries = {}
            Label(self.current_tab, 
                text= 'Variable Name').grid(row=8, column= 1,pady = 5,padx = 5, sticky= E)
            Label(self.current_tab, 
                text= 'Sampling Type').grid(row=8, column= 2,pady = 5,padx = 5, sticky=E)
            Label(self.current_tab, 
                text= '# of Trials').grid(row=8, column= 3,pady = 5,padx = 5)

            frame_canvas1 = Labelframe(self.current_tab)
            frame_canvas1.grid(row=9, column=1, columnspan =3, pady=(5, 0))
            frame_canvas1.grid_rowconfigure(0, weight=1)
            frame_canvas1.grid_columnconfigure(0, weight=1)
            frame_canvas1.config(height = '3c', width='13c')
            
            # Add a canvas in the canvas frame
            canvas1 = Canvas(frame_canvas1)
            canvas1.grid(row=0, column=0, sticky="news")
            canvas1.config(height = '3c', width='13c')
            
            # Link a scrollbar to the canvas
            vsb = Scrollbar(frame_canvas1, orient="vertical", command=canvas1.yview, 
                            style='scroll.Vertical.TScrollbar')
            vsb.grid(row=0, column=1,sticky = 'ns')
            canvas1.configure(yscrollcommand=vsb.set)
            
            # Create a frame to contain the variables
            frame_vars1 = Frame(canvas1)
            frame_vars1.config(height = '3c', width='13c')
            canvas1.create_window((0, 0), window=frame_vars1, anchor='nw')
            for name, format_of_data, vals in univariate_vars:
                Label(frame_vars1, 
                text= self.standardize_graph_title(name, True)).grid(row=self.univar_row_num,
                                                  column= 1,pady = 5,padx = 18)
                Label(frame_vars1, 
                text= self.standardize_graph_title(format_of_data, True)).grid(row=self.univar_row_num,
                                                  column= 2,pady = 5,padx = 18)
                
                if not(format_of_data == 'linspace' or 
                       format_of_data == 'list' or 'mapping' in format_of_data):
                    univar_val=Entry(frame_vars1)
                    univar_val.grid(row=self.univar_row_num, column=3,pady = 5,padx = 5)
                    self.univar_ntrials_entries[name]= univar_val
                    self.univar_ntrials_entries[name].config(width = 8)
                else:
                    if "mapping" in format_of_data:
                        Label(frame_vars1,text= self.standardize_graph_title(vals[-1].strip())).grid(
                                row=self.univar_row_num, column= 3,pady = 5,padx = 5, sticky= W)
                    elif format_of_data == 'linspace':
                        
                        Label(frame_vars1,text= self.standardize_graph_title(str(vals[2]).strip())).grid(
                                row=self.univar_row_num, column= 3,pady = 5,padx = 5, sticky = W)
                    else:
                        Label(frame_vars1,text= self.standardize_graph_title(str(len(vals)))).grid(
                                row=self.univar_row_num, column= 3,pady = 5,padx = 5, sticky= W)
                self.univar_row_num += 1
                
            # Update vars frames idle tasks to let tkinter calculate variable sizes
            frame_vars1.update_idletasks()
            
            # Determine the size of the Canvas
            frame_canvas1.config(width='13c', height='3c')
            
            # Set the canvas scrolling region
            canvas1.config(scrollregion=canvas1.bbox("all"))
            
    def get_distributions(self):
        '''
        Determines the number of trials needed for each input variable and calls
        construct_distributions to fill the simulation_vars and simulation_dist
        dictionaries. For univariate analysis, first the maximum number of trials
        for all of the variables is determined. After constructing the distribution
        for each variable, the extra values are truncated off depending on the number of 
        trials required for each variable.
        '''
        
        if self.analysis_type.get() == 'Univariate Sensitivity':
            # determine the maximum number of trials required
            if self.univar_ntrials_entries:
                max_num_sim = 0
                for slot in self.univar_ntrials_entries.values():
                    try:
                        cur_num_sim = int(slot.get())
                    except:
                        cur_num_sim = 1
                    max_num_sim = max(max_num_sim, cur_num_sim)
            else:
                max_num_sim = 1
            self.simulation_vars, self.simulation_dist = self.construct_distributions(
                    ntrials=max_num_sim)
            for (aspen_variable, aspen_call, fortran_index), dist in self.simulation_vars.items():
                if aspen_variable in self.univar_ntrials_entries:
                    try:
                        num_trials_per_var = int(self.univar_ntrials_entries[aspen_variable].get())
                    except:
                        num_trials_per_var = 1
                        # truncate off extra sampled values
                    self.simulation_vars[(aspen_variable, 
                                          aspen_call, fortran_index)] = dist[:num_trials_per_var]
                    self.simulation_dist[aspen_variable] = \
                    self.simulation_dist[aspen_variable][:num_trials_per_var]                
        else:
            try: 
                ntrials = int(self.num_sim_entry.get())
            except:
                ntrials=1
            self.simulation_vars, self.simulation_dist = self.construct_distributions(ntrials=ntrials) 
            
            
    def construct_distributions(self, ntrials=1):
        '''
        Given the excel input from the user, construct the distributions for 
        each input variable. If input variables have already been sampled from, 
        then a subset can be resampled. The function also checks for duplicate 
        input variables and skips over any duplicates. If the variable is a fortran
        variable, then the the distribution in simulation_vars will be
        converted to a fortran string.
        '''
        
        col_types = {'Variable Name': str, 'Variable Aspen Call': str, 
                     'Distribution Parameters': str, 'Bounds': str, 
                     'Fortran Call':str, 'Fortran Value to Change': str, 
                     'Distribution Type': str, 'Toggle': bool}
        df = read_excel(str(self.input_csv_entry.get()), sheet_name='Inputs', dtype=col_types)
        if not self.simulation_dist:
            simulation_vars = {}
            simulation_dist = {}
        else:
            # if the variables have already been sampled from, we don't want to 
            # wipe those values when we resample for a subset of variables.
            simulation_vars = self.simulation_vars
            simulation_dist = self.simulation_dist
        self.var_bounds = {}
        check_for_duplicate_vars = set()
        for index, row in df.iterrows():
            if row['Toggle']:
                dist_type = row['Distribution Type'].lower()
                aspen_variable = row['Variable Name']
                aspen_call = row['Variable Aspen Call']
                bounds = row['Bounds'].split(',')
                lb = float(bounds[0].strip())
                ub = float(bounds[1].strip())
                if aspen_variable in check_for_duplicate_vars:
                    continue
                if self.vars_have_been_sampled == True and not self.vars_to_resample[aspen_variable].get():
                    # only resample the variables that the user wants to resample
                    continue
                
                self.var_bounds[aspen_variable] = (lb, ub)
                check_for_duplicate_vars.add(aspen_variable)
                if 'mapping' in dist_type:
                    dist_vars = row['Distribution Parameters'].split(',')
                    lb_dist, ub_dist = float(dist_vars[-3].strip()), float(dist_vars[-2].strip())
                    num_trials = int(dist_vars[-1].strip())
                    distribution = linspace(lb_dist, ub_dist, num_trials)
                    if 'normal' in dist_type or 'gaussian' in dist_type:
                        mean, std_dev = float(dist_vars[0].strip()), float(dist_vars[1].strip())
                        if self.analysis_type.get() != "Univariate Sensitivity":
                            distribution = self.sample_gauss(mean, std_dev, lb, ub, ntrials)
                        else:
                            bin_width = (ub_dist - lb_dist)/num_trials
                            pdf_approx = self.sample_gauss(mean, std_dev, 
                                                           lb_dist-0.5*bin_width, 
                                                           ub_dist+0.5*bin_width, 100000)

                    if 'pareto' in dist_type:
                        shape, scale = float(dist_vars[0].strip()), float(dist_vars[1].strip())
                        if self.analysis_type.get() != "Univariate Sensitivity":
                            distribution = self.sample_pareto(shape, scale, lb, ub, ntrials)
                        else:
                            bin_width = (ub_dist - lb_dist)/num_trials
                            pdf_approx = self.sample_pareto(shape, scale, 
                                                            lb_dist-0.5*bin_width, 
                                                            ub_dist+0.5*bin_width, 100000)
                    if 'poisson' in dist_type:
                        lambda_p = float(dist_vars[0].strip())
                        if self.analysis_type.get() != "Univariate Sensitivity":
                            distribution = self.sample_poisson(lambda_p, lb, ub, ntrials)
                        else:
                            bin_width = (ub_dist - lb_dist)/num_trials
                            pdf_approx =self.sample_poisson(
                                    lambda_p, lb_dist-0.5*bin_width, 
                                    ub_dist+0.5*bin_width, 100000)
                        
                    if self.analysis_type.get() == "Univariate Sensitivity":
                        bin_width = (ub_dist - lb_dist)/num_trials
                        lb_pdf = lb_dist - 0.5*bin_width
                        ub_pdf = ub_dist + 0.5*bin_width
                        pdf, bin_edges = histogram(pdf_approx, bins=linspace(
                                lb_pdf, ub_pdf, num_trials+1), density=True)
                        tot_dens = sum(pdf)
                        self.mapping_pdfs[aspen_variable] = [p/tot_dens for p in pdf]

                elif 'normal' in dist_type or 'gaussian' in dist_type:
                    dist_variables = row['Distribution Parameters'].split(',')
                    distribution = self.sample_gauss(float(dist_variables[0].strip()),
                              float(dist_variables[1].strip()), lb, ub, ntrials)  
            
                elif 'linspace' in dist_type: 
                    linspace_vars = row['Distribution Parameters'].split(',')
                    distribution = linspace(float(linspace_vars[0].strip()), 
                                               float(linspace_vars[1].strip()),
                                               int(linspace_vars[2].strip()))
                    if self.analysis_type.get() == 'Multivariate Sensitivity':
                        distribution2 = choices(distribution, k=ntrials)
                        distribution = distribution2
                elif 'poisson' in dist_type:
                    lambda_p = float(row['Distribution Parameters'].strip())
                    distribution = self.sample_poisson(lambda_p, lb, ub, ntrials)
                elif 'pareto' in dist_type:
                    pareto_vals = row['Distribution Parameters'].split(',')
                    shape = float(pareto_vals[0].strip())
                    scale = float(pareto_vals[1].strip())
                    distribution = self.sample_pareto(shape, scale, lb, ub, ntrials)
                elif 'list' in dist_type:
                    lst = row['Distribution Parameters'].split(',')
                    distribution = []
                    for l in lst:
                        distribution.append(float(l.strip()))  
                    if self.analysis_type.get() == 'Multivariate Sensitivity':
                        distribution2 = choices(distribution, k=ntrials)
                        distribution = distribution2
                elif 'uniform' in dist_type:
                    lb_ub = row['Distribution Parameters'].split(',')
                    lb_uniform, ub_uniform = float(
                            lb_ub[0].strip()), float(lb_ub[1].strip())
                    distribution = self.sample_uniform(lb_uniform, ub_uniform, 
                                                       lb, ub, ntrials)
  
                if distribution is None:
                    Label(self.current_tab, text= 'ERROR: Distribution Parameters for ' + \
                          aspen_variable + ' are NOT valid', fg='red').grid(
                                  row=10, column=1, columnspan=3)
                    Label(self.current_tab, 
                          text='Please Adjust Distribution Parameters in'+\
                          'Input File and Restart Illuminate', fg='red').grid(
                                  row=11,column=1,columnspan=3)
                    return {}, {}
                
                simulation_dist[aspen_variable] = distribution[:]
                
                ############## CONVERT FORTRAN VARIABLES TO STRINGS ############
                fortran_index = (0,0)
                if row['Fortran Call'] != 'nan':
                    
                    fortran_call = row['Fortran Call']
                    value_to_change = row['Fortran Value to Change'].strip()
                    len_val = len(value_to_change)

                    for i in range(len(fortran_call)):
                        if fortran_call[i:i+len_val] == value_to_change:
                            fortran_index = (i, i+len_val) #NOT INCLUSIVE
                    distribution2 = list()
                    for i, v in enumerate(distribution):
                        distribution2.append(self.make_fortran(fortran_call, 
                                                               fortran_index, float(v)))
                    distribution = distribution2
                    
                simulation_vars[(aspen_variable, aspen_call, fortran_index)] = distribution
        return simulation_vars, simulation_dist
    
    
    def sample_gauss(self,mean, std, lb, ub, ntrials):
        '''
        Samples all values that are to be sampled from a gaussian distribution.
        Any values sampled outside of the user specified bounds will be discarded 
        and resampled to meet user specified num trials. In the event that the user 
        selects criteria outside the bounds the program will time out and return an
        error statment
        '''
        
        d = []
        for i in range(ntrials):
            rand_sample = random.normal(mean,std)
            st = time()
            stop = False
            while(rand_sample < lb or rand_sample > ub):
                if time() - st > 3:
                    stop = True
                    break
                rand_sample = random.normal(mean,std)
            if stop:
                return None
            d.append(rand_sample)
        return d
    
    
    def sample_uniform(self,lb_uniform, ub_uniform, lb, ub, ntrials):
        '''
        Samples all values that are to be sampled from a uniform distribution.
        Any values sampled outside of the user specified bounds will be discarded 
        and resampled to meet user specified num trials. In the event that the user 
        selects criteria outside the bounds the program will time out and return an
        error statment.
        '''
        
        d = []
        for i in range(ntrials):
            rand_sample = random.uniform(lb_uniform, ub_uniform)
            st = time()
            stop = False
            while(rand_sample < lb or rand_sample > ub):
                if time() - st > 3:
                    stop = True
                    break
                rand_sample = random.uniform(lb_uniform, ub_uniform)
            if stop:
                return None
            d.append(rand_sample)
        return d
    
    
    def sample_poisson(self,lambda_p, lb, ub, ntrials):
        '''
        Samples all values that are to be sampled from a poisson distribution.
        Any values sampled outside of the user specified bounds will be discarded 
        and resampled to meet user specified num trials. In the event that the user 
        selects criteria outside the bounds the program will time out and return an
        error statment.
        '''
        
        d = []
        for i in range(ntrials):
            rand_sample = random.poisson(10000*lambda_p)/10000
            st = time()
            stop = False
            while(rand_sample < lb or rand_sample > ub):
                if time() - st > 3:
                    stop = True
                    break
                
                rand_sample = random.poisson(10000*lambda_p)/10000
            if stop:
                return None
            d.append(rand_sample)
        return d
    
    def sample_pareto(self, shape, scale, lb, ub, ntrials):
        '''
        Samples all values that are to be sampled from a pareto distribution.
        Any values sampled outside of the user specified bounds will be discarded 
        and resampled to meet user specified num trials. In the event that the user 
        selects criteria outside the bounds the program will time out and return an
        error statment.
        '''
        
        d = []
        for i in range(ntrials):
            st = time()
            stop = False
            
            rand_sample = (random.pareto(shape) + 1) * scale
            while(rand_sample < lb or rand_sample > ub):
                if time() - st > 3:
                    stop = True
                    break
                rand_sample = (random.pareto(shape) + 1) * scale
            if stop:
                return None
            d.append(rand_sample)
        return d
    
    
    def make_fortran(self, fortran_call, fortran_index, val):
        '''
        Takes the fortran call string and the indices indicating the values to 
        replace and returns a new fortran string with the new distribution value.
        '''
        
        return fortran_call[:fortran_index[0]] + str(val) + fortran_call[fortran_index[1]:]
    
    
    def disp_single_point_results(self):
        '''
        Periodically check for a result from single point analysis until a result 
        is received. One a result is received, print out the outputs line by 
        line and indicate that the analysis is complete.
        '''
        
        if self.current_simulation and self.current_simulation.results:
            row = 4
            for output_var, toggled in self.graph_toggles.items():
                if toggled.get():
                    output_val = self.current_simulation.results[0].at[1, output_var]
                    if isna(output_val):
                        Label(self.current_tab, text= 'Aspen Failed to Converge', 
                              font='Helvetica 10 bold',fg='red').grid(row=row, column = 1)
                        break
                    output_val = "{:,}".format(float("%.2f" % output_val))
                    output_text_label = Label(self.current_tab, text=str(output_var) + '= ' + output_val)
                    self.sp_results_text.append(output_text_label)
                    output_text_label.grid(row=row, column = 1)
                    row += 1
                    self.sp_status.config(text='Status: Simulation Complete')
        else:
            self.after(5000, self.disp_single_point_results)

    
    def run_single_point_analysis(self):
        '''
        Run a single point analysis using the values specified in the GUI. A 
        single simulation object is created and run.
        '''
        
        if self.sp_error:
            self.sp_error.destroy()
        if self.sp_status:
            self.sp_status.destroy()
        for l in self.sp_results_text:
            l.destroy()
        self.sp_results_text = []
        self.store_user_inputs()
        self.get_distributions()
        # update simulation variable values based on user input in GUI
        for (aspen_variable, aspen_call, fortran_index), values in self.simulation_vars.items():
            value = float(self.sp_value_entries[aspen_variable].get())
            if value < self.var_bounds[aspen_variable][0] or value > self.var_bounds[aspen_variable][1]:
                self.sp_error = Label(self.current_tab, 
                                      text='Error: Value Specified for ' + \
                                      aspen_variable + ' is Outside Bounds', fg='red')
                self.sp_error.grid(row=6, column=1, columnspan=2)
                return
            self.simulation_vars[(aspen_variable, aspen_call, fortran_index)] = [value]
            
        self.create_simulation_object(self.simulation_vars, self.vars_to_change, 
                                      self.output_file, self.num_trial)
        self.sp_status = Label(self.current_tab,text='Status: Simulation Running')
        self.sp_status.grid(row=3,column=0, columnspan=2, sticky=W)
        self.run_simulations()
        
    
    def run_multivar_sens(self):
        '''
        Run a multivariate analysis. A single simulation object is created and
        run.
        '''
        self.store_user_inputs()
        if len(self.simulation_vars) == 0:
            # if the user has pressed the display variable distributions 
            # button, then we don't want to resample
            self.get_distributions()
        self.create_simulation_object(self.simulation_vars, self.vars_to_change, 
                                      self.output_file, self.num_trial)
        self.run_simulations()
        
        
    def run_univ_sens(self):
        '''
        Run a univariate analysis. A distinct simulation object is created for each
        variable in the univariate analysis.
        '''
        
        self.store_user_inputs()
        if len(self.simulation_vars) == 0:
            self.get_distributions()
        for (aspen_variable, aspen_call, fortran_index), values in self.simulation_vars.items():
            # if a variable has a "mapping" distribution, then we need to store
            # the weights of that distribution.
            weights = self.mapping_pdfs.get(aspen_variable, [])
            self.create_simulation_object(
                    {(aspen_variable, aspen_call,fortran_index): values}, 
                    [aspen_variable],self.output_file+'_'+aspen_variable, 
                    len(values), weights)
        self.run_simulations()
    
    def run_simulations(self):
        '''
        Run simulations for all of the simulation objects in self.simulations.
        After all simulations are created, many Main_App attributes are
        reset to their initial values in the event that the user wants to run
        more simulations.
        '''

        for sim in self.simulations:
            self.start_time = time()
            self.sims_completed.value = 0
            self.tot_sim_num = sim.tot_sim
            self.current_simulation = sim
            self.current_simulation.run_simulation() # run the simulation
            
            # the following code only runs once the simulation is complete
            # or is aborted
            if self.abort_univar_overall.value:
                self.abort.value = True
            self.univar_plot_counter += 1
            self.last_update = None
        
        # reset Main_App attributes to original values
        self.abort_univar_overall.value = False
        self.abort.value=False
        self.current_simulation = None
        self.tot_sim_num = 0
        self.sims_completed.value=0
        self.start_time = None
        self.univar_plot_counter = 1
        self.finished_figures = []
        self.last_results_plotted = None
        self.last_update = None
        self.mapping_pdfs = {}
    
    def copy_aspen_to_temp_dir(self):
        '''
        Copies the aspen .apw or .bkp file provided by the user to a temporary
        directory within the 'Output' folder. This is done in order to encapsulate
        all of the extra files that Aspen outputs so that they can be easily removed
        if Aspen crashes or is aborted. It first checks to see if this directory exists,
        and if it does exist, then it deletes the temporary directory and all
        of its contents.
        '''
        
        self.temp_directory = path.join(path.dirname(str(self.input_csv_entry.get())),'Output','Temp')
        
        # delete the directory if it exists
        try:
            rmdir(self.temp_directory)
        except: 
            pass
        try:
            rmtree(self.temp_directory)
        except: 
            pass
        if not path.exists(self.temp_directory):
            makedirs(self.temp_directory)
        aspen_file_names = []
        for i in range(self.num_processes):
            process_specific_dir = self.temp_directory + '\\' + str(i)
            makedirs(process_specific_dir)
            aspen_file_name = path.join(process_specific_dir,path.basename(str(self.aspen_file_entry.get())))
            aspen_file_names.append(aspen_file_name)
            copyfile(str(self.aspen_file_entry.get()), aspen_file_name)
        return aspen_file_names
        
    
    def store_user_inputs(self):
        '''
        Stores inputs provided by the user, including the number of processors,
        number of trials, output file name, aspen and calculator files, 
        and variables and warning keywords specified in the input file.
        '''
        
        self.aspen_file = str(self.aspen_file_entry.get())
        try:
            self.num_processes = int(self.num_processes_entry.get())
        except:
            self.num_processes = 1
        try:
            self.graphing_frequency = int(self.graphing_freq_entry.get())
        except:
            # if no graphing frequency was given, by default it is 1
            self.graphing_frequency = 1
        self.excel_solver_file= str(self.excel_solver_entry.get())
        try:
            self.num_trial = int(self.num_sim_entry.get())
        except: 
            self.num_trial = 1
        self.output_file = str(self.save_as_entry.get())
        if len(self.output_file) == 0:
            self.output_file = str('Simulation_Results')
        self.input_csv = str(self.input_csv_entry.get())
        
        self.vars_to_change = []
        gui_excel_input = str(self.input_csv_entry.get())
        col_types = {'Variable Name': str, 'Variable Aspen Call': str, 
                     'Distribution Parameters': str, 'Bounds': str, 
                     'Fortran Call':str, 'Fortran Value to Change': str, 
                     'Distribution Type': str, 'Toggle': bool}
        df = read_excel(gui_excel_input, sheet_name='Inputs', dtype=col_types)
        variable_names = set()
        for index, row in df.iterrows():
            if row['Toggle'] and row['Variable Name'] not in variable_names:
                variable_names.add(row['Variable Name'])
                self.vars_to_change.append(row["Variable Name"])
        col_types = {'Warning Keywords':str}
        warnings_df = read_excel(gui_excel_input, sheet_name='Warning Messages', dtype=col_types)
        for index, row in warnings_df.iterrows():
            self.warning_keywords.update(
                    [word.strip().lower() for word in row['Warning Keywords'].split()])
            
    
    def parse_output_vars(self):
        '''
        Opens the excel calculator file and stores the output variables of interest
        specified in the "Output" tab.
        '''
        
        self.excel_solver_file= str(self.excel_solver_entry.get())
        excels_to_ignore = set()
        variable_names = set()
        # storing a set of already open excel processes so we don't terminate those
        for p in process_iter():
            if 'excel' in p.name().lower():
                excels_to_ignore.add(p.pid)
        # opening the excel COMS for the calculator file
        excel, book = open_excelCOMS(self.excel_solver_file)
        self.bkp_ref = self.find_bkp_ref(book)
        self.output_vars = []
        self.output_value_cells = []
        row_counter = 3
        while True:
            var_name = book.Sheets('Output').Evaluate("B" + str(row_counter)).Value
            if var_name:
                units = book.Sheets('Output').Evaluate("D" + str(row_counter)).Value
                column_name = var_name + ' (' + units + ')' if units else var_name
                if column_name not in variable_names:
                    variable_names.add(column_name)
                    self.output_vars.append(column_name)
                    self.output_value_cells.append('C' + str(row_counter))
            else:
                break
            row_counter += 1
        self.output_value_cells = ",".join(self.output_value_cells)
        self.output_vars += ['Aspen Error Count', 'Aspen Warning Count', 'Aspen Run Summary'] # adding 2 more columns to output file
        # closing the excel COMS we just opened
        for p in process_iter():
            if 'excel' in p.name().lower() and p.pid not in excels_to_ignore:
                p.terminate()
            
        
    def create_simulation_object(self, simulation_vars, vars_to_change,
                                 output_file, num_trial, weights=[]):
        '''
        Constructs a Simulation object.
        It also calls copy_aspen_to_temp_dir and creates the Output folder and
        temporary directory. 
        '''
        
        self.output_columns = vars_to_change + self.output_vars
        output_directory = path.join(path.dirname(str(self.input_csv_entry.get())),
                                     'Output/',datetime.now().strftime('%Y-%m-%d_%H-%M-%S'))
        makedirs(output_directory)
        if not path.exists(path.join(output_directory,'..','Temp')):
            makedirs(path.join(output_directory,'..','Temp'))
        copyfile(path.abspath(str(self.input_csv_entry.get())), path.join(
                output_directory,'Input_Variables.xlsx'))
        
        aspen_file_names = self.copy_aspen_to_temp_dir()
        new_sim = Simulation(
                self.sims_completed, num_trial, simulation_vars,output_file, 
                output_directory, aspen_file_names, self.excel_solver_file, 
                self.abort, vars_to_change, self.output_value_cells, self.output_columns, 
                self.select_version.get(),weights, self.save_bkp.get(), self.warning_keywords, self.bkp_ref,
                save_freq=2, num_processes=self.num_processes)
        self.simulations.append(new_sim)

        
        
    def initialize_single_point(self):
        '''
        Initializes the single point analysis in a new worker thread. This analysis
        is run in a new thread so that the GUI can remain active in the main thread
        while the simulation is running.
        '''
        self.simulations = []
        if self.worker_thread and self.worker_thread.isAlive():
            #don't start another simulation if one is already running
            print('Simulation Already Running')
            return
        self.worker_thread = Thread(
                target=lambda: self.run_single_point_analysis())
        self.worker_thread.start()
        self.after(5000, self.disp_single_point_results)
        
    def initialize_univar_analysis(self):
        '''
        Initializes a univariate analysis in a new worker thread. This analysis
        is run in a new thread so that the GUI can remain active in the main thread
        while the simulation is running.
        '''
        self.simulations = []
        if self.worker_thread and self.worker_thread.isAlive():
            #don't start another simulation if one is already running
            print('Simulation Already Running')
            return
        self.worker_thread = Thread(
            target=lambda: self.run_univ_sens())
        self.worker_thread.start()
        self.status_label = None
        self.time_rem_label = None
        self.univar_gui_update()

    
    def initialize_multivar_analysis(self):
        '''
        Initializes a multivariate analysis in a new worker thread. This analysis
        is run in a new thread so that the GUI can remain active in the main thread
        while the simulation is running.
        '''
        self.simulations = []
        if self.worker_thread and self.worker_thread.isAlive():
            #don't start another simulation if one is already running
            print('Simulation Already Running')
            return
        self.worker_thread = Thread(
            target=lambda: self.run_multivar_sens())
        self.worker_thread.start()
        self.status_label = None
        self.time_rem_label = None
        self.multivar_gui_update()
        
        
    def find_status_update(self):
        '''
        Constructs the status update string based on the status of the current
        simulation.
        '''
        if self.current_simulation and not self.abort.value:
            if len(self.current_simulation.results) == self.current_simulation.tot_sim:
                status_update = 'Status: Simulation Complete'
            else:
                status_update = 'Status: Simulation Running | {} Results Collected'.format(
                        len(self.current_simulation.results))
            return status_update
        return None
        
    def disp_time_remaining(self, status_update):
        '''
        Calculates the estimated time remaining and displays the full
        status update on the GUI. It only updates the status if new results
        have been collected.
        
        Input: status_update (str) from the function find_status_update
        '''
        if self.start_time and self.sims_completed.value != self.last_update:
            if not status_update:
                status_update = ''
            self.last_update = self.sims_completed.value
            elapsed_time = time() - self.start_time
            if self.sims_completed.value > 0:
                remaining_time = ((elapsed_time / self.sims_completed.value) * (
                        self.tot_sim_num - self.sims_completed.value))//60
                hours, minutes = divmod(remaining_time, 60)
                tmp = Label(self.display_tab, text=status_update + ' | ' + \
                            'Time Remaining: {} Hours, {} Minutes    '.format(int(hours), int(minutes)))
            else:
                tmp = Label(self.display_tab, text=status_update + ' | ' +'Time Remaining: N/A')
            tmp.place(x=6, y=4)
            if self.time_rem_label:
                self.time_rem_label.destroy()
            self.time_rem_label = tmp
    
    def num_bins(self, data):
        '''
        Determines the optimal bin number for histograms based on the number of 
        data points and the distribution of those points. For data with more 
        than 20 values, it implements the Freedman-Diaconis rule for bin numbers.
        '''
        
        if len(data) == 0:
            return 1
        if len(data) < 20:
            return len(data)
        iqr = subtract(*percentile(data, [75, 25]))
        bin_width = (2*iqr)/(len(data)**(1/3))
        if isclose(bin_width, 0.0):
            return 1
        num_bins = ceil((max(data) - min(data))/bin_width)
        return num_bins
            
    def generate_or_update_multivar_plots(self):
        '''
        Generates histograms for the display tab of a multivariate analysis or
        updates those histograms if they already exist. 
        '''
        
        if not self.simulation_dist:
            return
        status_label = None
        if not self.display_tab:
            # create the display tab and put the status label "setting up simulation"
            self.display_tab = Frame(self.notebook)
            self.notebook.add(self.display_tab,text = "Simulation Status")
            
            status_label = Label(self.display_tab, text='Setting Up Simulation...')
            status_label.place(x=6, y=4)
            self.init_plots_constructed = False
            self.plots_dictionary = {}
            Button(self.display_tab, text = "Abort", command=self.abort_sim).place(
                    x=(4*self.win_lim_x)//5, y = 5)
            self.notebook.select(self.display_tab)
        if not self.simulations:
            self.notebook.select(self.display_tab)
            
        # don't want to make histograms if current_simulation hasn't been started,
        # if the graphing_frequency=0, or if there are no new results to plot
        if self.graphing_frequency == 0:
            return
        if not self.current_simulation:
            return
        if len(self.current_simulation.results) == self.last_results_plotted:
            return
        self.last_results_plotted = len(self.current_simulation.results)

        if len(self.current_simulation.results) % self.graphing_frequency != 0:
            return
        if self.current_simulation.results:
            # need to filter out any empty results from aspen failing to converge
            # before attempting to plot them
            results_to_plot = list(filter(lambda x: not isna(x[self.current_simulation.output_columns[len(
                    self.current_simulation.vars_to_change)]].values[0]), self.current_simulation.results))
            if len(results_to_plot) == 0:
                results_filtered = DataFrame(columns=self.output_columns)
                results_unfiltered = results_filtered
            else:
                results_filtered = concat(results_to_plot).sort_index()
                results_unfiltered = concat(self.current_simulation.results).sort_index()
            
        else:
            # if no results, then just generate empty plots
            results_filtered = DataFrame(columns=self.output_columns)
            results_unfiltered = results_filtered
            
        if not self.init_plots_constructed:
            
            ###########    Generate the initial histograms    ###############
            results_fig_list =[] # for results histograms
            for var, toggled in self.graph_toggles.items():
                if toggled.get():
                    fig = Figure(figsize = (3,3), facecolor=[240/255,240/255,237/255], tight_layout=False)
                    ax = fig.add_subplot(111)
                    num_bins = self.num_bins(results_filtered[var])
                    ax.hist(results_filtered[var], num_bins, facecolor='blue', edgecolor='black', alpha=1.0)
                    ax.set_title(self.standardize_graph_title(var))
                    ax.ticklabel_format(axis= 'x', style = 'sci', scilimits= (-3,3))
                    self.plots_dictionary[var] = ax
                    results_fig_list.append(fig)
            
            inputs_fig_list = [] # for input distribution histograms
            for var, values in self.simulation_dist.items():
                fig = Figure(figsize = (3,3), facecolor=[240/255,240/255,237/255], tight_layout=False)
                a = fig.add_subplot(111)
                num_bins = self.num_bins(self.simulation_dist[var])
                _, bins, _ = a.hist(self.simulation_dist[var], num_bins, facecolor='white', 
                                    edgecolor='black',alpha=1.0)
                a.hist(results_unfiltered[var], bins=bins, facecolor='blue',edgecolor='black', alpha=1.0)
                a.set_title(self.standardize_graph_title(var))
                a.ticklabel_format(axis= 'x', style = 'sci', scilimits= (-3,3))
                self.plots_dictionary[var] = a
                inputs_fig_list.append(fig)
            
            ############### Making Frames and Scroll Bars for plotting region of GUI ################
            row_num = 0
            frame_width = self.win_lim_x - 30
            num_graphs_per_row = frame_width//250
            frame_height = 60+(230*((len(inputs_fig_list) + len(
                    results_fig_list)+1)//num_graphs_per_row + 1))  
            window_height = self.win_lim_y - 30
            
            frame_canvas = Frame(self.display_tab)
            frame_canvas.grid(row=row_num, column=1, columnspan = 3,pady=(5, 0))
            frame_canvas.grid_rowconfigure(0, weight=1)
            frame_canvas.grid_columnconfigure(0, weight=1)
            frame_canvas.config(height = window_height, width=frame_width)
            
            main_canvas = Canvas(frame_canvas)
            main_canvas.grid(row=0, column=0, sticky="news")
            main_canvas.config(height = window_height, width=frame_width)
            
            vsb = Scrollbar(frame_canvas, orient="vertical", command=main_canvas.yview, 
                            style='scroll.Vertical.TScrollbar')
            vsb.grid(row=0, column=1,sticky = 'ns')
            main_canvas.configure(yscrollcommand=vsb.set)
            
            figure_frame = Frame(main_canvas)
            main_canvas.create_window((0, 0), window=figure_frame, anchor='nw')
            figure_frame.config(height = frame_height, width=frame_width)
            if status_label:
                status_label.destroy()
        
            ############# Placing plots on the GUI   ###################
            count = 0
            x, y = 10, 30
            output_dis = Label(figure_frame, text = 'Outputs:', font='Helvetica 10 bold')
            output_dis.place(x = x, y = y)
            y += 20
            for figs in results_fig_list:
                figure_canvas = FigureCanvasTkAgg(figs, master=figure_frame)
                x = 10 + 250*(count % num_graphs_per_row)
                figure_canvas.get_tk_widget().place(x = x, y= y, width = 240, height =220)
                if (count+1) % num_graphs_per_row==0 and count != len(results_fig_list) - 1:
                    y += 230
                count += 1
            y += 240
            line= Label(figure_frame, text = '-------------------------------------------'+\
                        '-----------------------------------------------------------------'+\
                        '------------------------------------------------------------------'+\
                        '------------------------------------------------')
            line.place(x = 0, y = y-12)
            input_dis = Label(figure_frame, text = 'Inputs:', font='Helvetica 10 bold')
            input_dis.place(x = 10, y = y)
            y += 20
            x=10
            count = 0
            for figs in inputs_fig_list:
                figure_canvas = FigureCanvasTkAgg(figs, master=figure_frame)
                x = 10 + 250*(count % num_graphs_per_row)
                figure_canvas.get_tk_widget().place(x = x, y= y, width = 240, height =220)
                if (count+1) % num_graphs_per_row==0:
                    y += 230
                count += 1
    
            figure_frame.update_idletasks()
            frame_canvas.config(width=frame_width, height=window_height)
            main_canvas.config(scrollregion=(0,0,x,frame_height))
            Button(self.display_tab, text = "Abort", command=self.abort_sim).place(
                    x=(4*self.win_lim_x)//5, y = 5)


        ######### If init plots have been constructed, just update their axes and data ##########
        else:
            for f in self.plots_dictionary.values():
                # clear the axes first
                try:
                    f.cla()
                except:
                    pass
                try:
                    f.clear()
                except:
                    pass
            for output_var, toggled in self.graph_toggles.items():
                if toggled.get():
                    num_bins = self.num_bins(results_filtered[output_var])
                    self.plots_dictionary[output_var].hist(
                            results_filtered[output_var], num_bins, facecolor='blue', 
                            edgecolor='black', alpha=1.0)
                    self.plots_dictionary[output_var].set_title(
                            self.standardize_graph_title(output_var))
                    self.plots_dictionary[output_var].ticklabel_format(axis= 'x', 
                                         style = 'sci', scilimits= (-3,3))
            for var, values in self.simulation_dist.items():
                num_bins = self.num_bins(self.simulation_dist[var])
                _, bins, _ = self.plots_dictionary[var].hist(
                        self.simulation_dist[var], num_bins, facecolor='white', 
                        edgecolor='black',alpha=1.0)
                self.plots_dictionary[var].hist(results_unfiltered[var], bins=bins, 
                                     facecolor='blue', edgecolor='black', alpha=1.0)
                self.plots_dictionary[var].set_title(self.standardize_graph_title(var))
                self.plots_dictionary[var].ticklabel_format(axis= 'x', style = 'sci', 
                                     scilimits= (-3,3))


            for fig in self.graphs_displayed:
                # updates the plots on the GUI
                fig.draw()
                
    
                
            
    def generate_or_update_univar_plots(self):
        '''
        Generates histograms for the display tab of a univariate analysis or
        updates those histograms if they already exist. 
        '''
        
        status_label = None
        if not self.simulation_dist:
            return
        if not self.display_tab:
            # create display tab and put the status label "setting up simulation"
            self.display_tab = Frame(self.notebook)
            self.notebook.add(self.display_tab,text = "Simulation Status")
            
            status_label = Label(self.display_tab, text='Setting Up Simulation...')
            status_label.place(x=6, y=4)
            self.init_plots_constructed = False
            self.plots_dictionary = {}
            Button(self.display_tab, text = "Next Variable", command=self.abort_sim).place(
                    x=self.win_lim_x - 110, y=3)
            
            Button(self.display_tab, text = "Abort", command=self.abort_univar_overall_fun).place(
                    x=self.win_lim_x-190, y=3)
            self.notebook.select(self.display_tab)
        if not self.simulations:
            self.notebook.select(self.display_tab)
            
        # don't want to make histograms if current_simulation hasn't been started,
        # if the graphing_frequency=0, or if there are no new results to plot
        if self.graphing_frequency == 0:
            return
        if not self.current_simulation:
            return
        if len(self.current_simulation.results) == self.last_results_plotted:
            return
        
        self.last_results_plotted = len(self.current_simulation.results)
        
        
        if len(self.current_simulation.results) % self.graphing_frequency != 0:
            return
        
        current_var = self.current_simulation.vars_to_change[0]
        if self.current_simulation.results:
            # need to filter out any empty results from aspen failing to converge
            # before attempting to plot them
            results_to_plot = list(filter(lambda x: not isna(x[self.current_simulation.output_columns[len(
                    self.current_simulation.vars_to_change)]].values[0]), self.current_simulation.results))
            if len(results_to_plot) == 0:
                results_filtered = DataFrame(columns=self.current_simulation.output_columns)
                results_unfiltered = results_filtered
            else:
                results_filtered = concat(results_to_plot).sort_index()
                results_unfiltered = concat(self.current_simulation.results).sort_index()
        else:
            # if no results, then just generate empty plots
            results_filtered = DataFrame(columns=self.current_simulation.output_columns)
            results_unfiltered = results_filtered
            
        ########### Generate the Initial Histograms  ############
        if not self.init_plots_constructed:
            fig_list = []                
            for var, values in self.simulation_dist.items():
                self.plots_dictionary[var] = {}
                fig = Figure(figsize = (3,3), facecolor=[240/255,240/255,237/255])
                a = fig.add_subplot(111)
                num_bins = self.num_bins(self.simulation_dist[var])
                _, bins, _ = a.hist(self.simulation_dist[var], num_bins, facecolor='white', 
                                    edgecolor='black',alpha=1.0)
                a.set_title(self.standardize_graph_title(var))
                a.ticklabel_format(axis= 'x', style = 'sci', scilimits= (-3,3))
                fig_list.append(fig)
                self.plots_dictionary[var][var] = a
                self.num_toggled = 0
                for output_var, toggled in self.graph_toggles.items():
                    if toggled.get():
                        self.num_toggled += 1
                        fig = Figure(figsize = (3,3), facecolor=[240/255,240/255,237/255])
                        ax = fig.add_subplot(111)
                        num_bins = self.num_bins(results_filtered[output_var])
                        ax.hist(results_filtered[output_var], num_bins, facecolor='blue', 
                                edgecolor='black', alpha=1.0)
                        ax.set_title(self.standardize_graph_title(output_var))
                        ax.ticklabel_format(axis= 'x', style = 'sci', scilimits= (-3,3))
                        fig_list.append(fig)
                        self.plots_dictionary[var][output_var] = ax
                
            ########  Creating the Frame, Canvas, and Scroll bars for plotting region of GUI #####
            row_num = 0
            frame_width = self.win_lim_x - 30
            num_graphs_per_row = self.num_toggled + 1
            graphs_frame_width = 30 + 250*(num_graphs_per_row)
            frame_height = 30+(230*((len(fig_list)+1)//num_graphs_per_row + 1))
            window_height = self.win_lim_y - 60
            
            frame_canvas = Frame(self.display_tab)
            frame_canvas.grid(row=row_num, column=1, columnspan = 3,pady=(5, 0))
            frame_canvas.grid_rowconfigure(0, weight=1)
            frame_canvas.grid_columnconfigure(0, weight=1)
            frame_canvas.config(height = window_height, width=frame_width)
            
            main_canvas = Canvas(frame_canvas)
            main_canvas.grid(row=0, column=0, sticky="news")
            main_canvas.config(height = window_height, width=frame_width)
            
            hsb = Scrollbar(frame_canvas, orient="horizontal", command=main_canvas.xview, 
                            style='scroll.Horizontal.TScrollbar')
            hsb.grid(row=1, column=0,sticky = 'we')
            main_canvas.configure(xscrollcommand=hsb.set)
            
            vsb = Scrollbar(frame_canvas, orient="vertical", command=main_canvas.yview, 
                            style='scroll.Vertical.TScrollbar')
            vsb.grid(row=0, column=1,sticky = 'ns')
            main_canvas.configure(yscrollcommand=vsb.set)
            
            figure_frame = Frame(main_canvas)
            main_canvas.create_window((0, 0), window=figure_frame, anchor='nw')
            figure_frame.config(height = frame_height, width=graphs_frame_width)
        
            ############ Placing plots on the GUI ##############
            count = 0
            x, y = 10, 30
            self.graphs_displayed = []
            for figs in fig_list:
                figure_canvas = FigureCanvasTkAgg(figs, master=figure_frame)
                self.graphs_displayed.append(figure_canvas)
                x = 10 + 250*(count % num_graphs_per_row)
                figure_canvas.get_tk_widget().place(x = x, y= y, width = 240, height =220)
                if (count+1) % num_graphs_per_row==0:
                    y += 230
                count += 1
                
            frame_canvas.config(width=frame_width, height=window_height)
            self.init_plots_constructed = True

            figure_frame.update_idletasks()
            # Set the canvas scrolling region
            main_canvas.config(scrollregion=(0,0,graphs_frame_width,frame_height))
            Button(self.display_tab, text = "Next Variable", command=self.abort_sim).place(
                    x=self.win_lim_x - 110, y=3)
            
            Button(self.display_tab, text = "Abort", command=self.abort_univar_overall_fun).place(
                    x=self.win_lim_x-190, y=3)
        
    
        ######### If init plots already constructed, then just update their axes and data #######
        else:
            for f in self.plots_dictionary[current_var].values():
                f.cla()
                f.clear()
            for output_var, toggled in self.graph_toggles.items():
                if toggled.get():
                    if len(self.current_simulation.weights) > 0:
                        weights = self.current_simulation.weights[0:len(results_filtered)]
                        num_bins = self.num_bins(results_filtered[output_var])
                        self.plots_dictionary[current_var][output_var].hist(
                            results_filtered[output_var], num_bins, weights=weights, 
                            facecolor='blue', edgecolor='black', alpha=1.0)
                    else:
                        num_bins = self.num_bins(results_filtered[output_var])
                        self.plots_dictionary[current_var][output_var].hist(
                                results_filtered[output_var], num_bins, facecolor='blue', 
                                edgecolor='black', alpha=1.0)
                    self.plots_dictionary[current_var][output_var].set_title(
                            self.standardize_graph_title(output_var))
                    self.plots_dictionary[current_var][output_var].ticklabel_format(
                            axis= 'x', style = 'sci', scilimits= (-3,3))
            num_bins = self.num_bins(self.simulation_dist[current_var])
            _, bins, _ = self.plots_dictionary[current_var][current_var].hist(
                    self.simulation_dist[current_var], num_bins, facecolor='white', 
                    edgecolor='black',alpha=1.0)
            self.plots_dictionary[current_var][current_var].hist(
                    results_unfiltered[current_var], bins=bins, facecolor='blue', 
                    edgecolor='black', alpha=1.0)
            self.plots_dictionary[current_var][current_var].set_title(
                    self.standardize_graph_title(current_var))
            self.plots_dictionary[current_var][current_var].ticklabel_format(
                    axis= 'x', style = 'sci', scilimits= (-3,3))

            for fig in self.graphs_displayed:
                fig.draw()
    
        
            
    def plot_variable_distributions(self):
        '''
        This function will get distributions for each input variable and plot them
        prior to running the simulation. It also creates the toggle boxes for
        resample of specific variables. This is called from both univariate
        and multivariate analyses.
        '''


        self.get_distributions()      
        
        # generate the histograms from distributions
        fig_list =[]
        for var, values in self.simulation_dist.items():
            fig = Figure(figsize = (3,3), facecolor=[240/255,240/255,237/255])
            a = fig.add_subplot(111)
            num_bins = self.num_bins(values)
            if var in self.mapping_pdfs:
                a.hist(values, num_bins, weights=self.mapping_pdfs[var], facecolor='blue', 
                       edgecolor='black', alpha=1.0)
            else:
                a.hist(values, num_bins, facecolor='blue', edgecolor='black', alpha=1.0)
            a.set_title(self.standardize_graph_title(var))
            a.ticklabel_format(axis= 'x', style = 'sci', scilimits= (-3,3))
            fig_list.append(fig)
            
        # the location of placement of these plots on the GUI is dependent
        # on the type of analysis
        if self.univar_row_num != 0:
            # univariate
            row_num = 17
        else:
            # multivariate
            row_num = 10

        ######## Create the toggle boxes for individual variable resampling ######
        self.resample_vars= Labelframe(self.current_tab, text='Select Variables to Resample:')
        self.resample_vars.grid(row = row_num, column = 1, columnspan = 5)
        count = 0
        row_track, col_track = 0,0
        if not self.vars_have_been_sampled:
            self.vars_to_resample = {}
            # This variable vars_have_been_sampled is just an indicator that these
            # toggles have been created
            self.vars_have_been_sampled = True
            
            for v in self.simulation_dist.keys():
                self.vars_to_resample[v] = IntVar()
                cb = Checkbutton(self.resample_vars, text = v, variable = self.vars_to_resample[v])
                cb.grid(row= row_track, column = col_track)
                cb.select()
                col_track += 1
                if col_track%5 == 0:
                        row_track +=1
                        col_track = 0
        
            
        ########## Create the frame, canvas, and scroll bars for plotting region of GUI ######
        frame_width = self.win_lim_x - 30
        num_graphs_per_row = frame_width//250
        frame_height = 30+(230*((len(fig_list)-1)//num_graphs_per_row + 1)) 
        if self.univar_row_num != 0:
            
            window_height = self.win_lim_y - (400 + row_track*45)
        else:
            window_height = self.win_lim_y - (160 + row_track*45)
        row_num += 1
        
        frame_canvas = Frame(self.current_tab)
        frame_canvas.grid(row=row_num, column=1, columnspan = 3,pady=(5, 0))
        frame_canvas.grid_rowconfigure(0, weight=1)
        frame_canvas.grid_columnconfigure(0, weight=1)
        frame_canvas.config(height = window_height, width=frame_width)
        
        main_canvas = Canvas(frame_canvas)
        main_canvas.grid(row=0, column=0, sticky="news")
        main_canvas.config(height = window_height, width=frame_width)
        
        vsb = Scrollbar(frame_canvas, orient="vertical", command=main_canvas.yview, 
                        style='scroll.Vertical.TScrollbar')
        vsb.grid(row=0, column=1,sticky = 'ns')
        main_canvas.configure(yscrollcommand=vsb.set)
        
        figure_frame = Frame(main_canvas)
        main_canvas.create_window((0, 0), window=figure_frame, anchor='nw')
        figure_frame.config(height = frame_height, width=frame_width)
        
        count = 0
        x, y = 10, 10
        output_dis = Label(figure_frame, text = 'Inputs:', font='Helvetica 10 bold')
        output_dis.place(x = x, y = y)
        
        ########  Placing plots on the GUI ############
        count =0
        for figs in fig_list:
            figure_canvas = FigureCanvasTkAgg(figs, master=figure_frame)
            x = 10 + 250*(count % num_graphs_per_row)
            figure_canvas.get_tk_widget().place(x = x, y= y, width = 240, height =220)
            if (count+1) % num_graphs_per_row==0:
                y += 230
            count += 1
        figure_frame.update_idletasks()
        frame_canvas.config(width=frame_width, height=window_height)
        main_canvas.config(scrollregion=(0,0,x,frame_height))
        
        
    def univar_gui_update(self):
        '''
        Responsible for calling all of the auto-updating functionality for a 
        univariate analysis. This includes updating the plots and updating the
        status and time remaining.
        '''
        
        self.generate_or_update_univar_plots()
        self.disp_time_remaining(self.find_status_update())
        
        # only keep updating if there is a current simulation or if the current simulation
        # has not been aborted
        if not self.current_simulation or (
                self.current_simulation and not self.abort_univar_overall.value):
            # call itself after 5 seconds
            self.after(5000, self.univar_gui_update)
        
        
    def multivar_gui_update(self):
        '''
        Responsible for calling all of the auto-updating functionality for a 
        multivariate analysis. This includes updating the plots and updating the
        status and time remaining.
        '''
        
        self.generate_or_update_multivar_plots()
        self.disp_time_remaining(self.find_status_update())
        
        # only keep updating if there is a current simulation or if the current simulation
        # has not been aborted
        if not self.current_simulation or (
                self.current_simulation and not self.current_simulation.abort.value):
            # call itself after 5 seconds
            self.after(5000, self.multivar_gui_update)
        
    
    def fill_num_trials(self):
        '''
        Responsible for filling all univariate ntrials entries when the user
        presses the fill #trials button.
        '''
        
        ntrials = self.fill_num_sims.get()
        for name, slot in self.univar_ntrials_entries.items():
            slot.delete(0, END)
            slot.insert(0, ntrials)
        

    def open_input_excel_file(self):
        '''
        Prompts the user to select an input excel file and fills the entry with the
        full path name.
        '''
        filename = askopenfilename(title = "Select file", filetypes = (
                (".xlsx Files","*.xlsx"),))
        self.input_csv_entry.delete(0, END)
        self.input_csv_entry.insert(0, filename)
                
    def find_bkp_ref(self, book):
        '''
        Seaches through the user's VBA code in the .xlsm Calculator file 
        to find the location of the .bkp reference in the "Set-up" tab. If
        we cannot access the VBA code then we assume the .bkp reference is in cell
        B1
        '''
        try:
            vba_code = book.VBProject.VBComponents("GelAllData").CodeModule.Lines(1,500000)
        except:
            try:
                vba_code = book.VBProject.VBComponents("GetAllData").CodeModule.Lines(1,500000)
            except:
                vba_code = ''
        
        i=0
        get_data_VBA = ""
        while i < len(vba_code):
            if vba_code[i:i+31] == 'Public Sub sub_GetSumData_ASPEN':
                s_ind = i
                while vba_code[i-7:i] != 'End Sub':
                    i += 1
                get_data_VBA = vba_code[s_ind:i]
                break
            i +=1
        
        if get_data_VBA:
            try:
                bkp_reference_cell = findall(r"RTrim\(Worksheets\(\"Set-up\"\)\.Range\(\"([A-Z]+[0-9]+)\"\)\.VALUE", 
                                                  get_data_VBA)[0]
            except:
                bkp_reference_cell = 'B1'
            
        else:
            bkp_reference_cell = 'B1'
            
        return bkp_reference_cell
        
    def open_aspen_file(self):
        '''
        Prompts the user to select an Aspen .bkp or .apw file and fills the entry with the
        full path name.
        '''
        filename = askopenfilename(title = "Select file", filetypes = (
                ("Aspen Models",["*.bkp", "*.apw"]),))
        self.aspen_file_entry.delete(0, END)
        self.aspen_file_entry.insert(0, filename)
    
    
    def open_excel_calculator_file(self):
        '''
        Prompts the user to select an Excel calculator .xlsm file and fills the entry with the
        full path name. It also calls self.graph_toggle() to load the output variables
        for the user to select which to graph.
        '''
        filename = askopenfilename(title = "Select file", filetypes = (
                (".xlsm Files","*.xlsm"),))
        self.excel_solver_entry.delete(0, END)
        self.excel_solver_entry.insert(0, filename)
        if filename:
            plot_output_disp_thread = Thread(target=self.output_vars_to_graph_toggle)
            plot_output_disp_thread.start()
            self.wait= Label(self.home_tab, text="Wait While Output Variables Are Loading ...")
            self.wait.grid(row=6, column= 1, columnspan = 2, sticky = E,pady = 5,padx = 5)
              
       
    def output_vars_to_graph_toggle(self):
        '''
        Calls parse_output_vars to open the excel calculator .xlsm file and load
        output variables of interest. It then creates the check boxes for the user
        to toggle which they want to graph. After this is completed, it displays the button 
        for testing compatibility of input files.
        '''
        
        self.parse_output_vars()
        self.graph_toggles = {}
        if len(self.output_vars) < 10:
            # if there are fewer than 10, then we don't need a scroll bar
            self.disp_output_vars= Labelframe(self.home_tab, text='Output Variables to Graph:')
            self.disp_output_vars.grid(row = 6,column = 1, columnspan = 2, pady = 10, 
                                       padx = 10, sticky = E )
            count = 1

            for i,v in enumerate(self.output_vars[:-3]):
                self.graph_toggles[v] = IntVar()
                cb = Checkbutton(self.disp_output_vars, text = v, variable = self.graph_toggles[v])
                cb.grid(row=count,columnspan = 1, column = 2, sticky=W)
                cb.select()
                count += 1
            self.wait.destroy()
        else:
            row_num= 6
            frame_width = self.win_lim_x/3
            frame_height = len(self.output_vars)*25 + 10
            window_height = 300
            
            frame_canvas = Labelframe(self.home_tab,text='Output Variables to Graph:')
            frame_canvas.grid(row=row_num, column=1, columnspan = 3,pady=(5, 0))
            frame_canvas.grid_rowconfigure(0, weight=1)
            frame_canvas.grid_columnconfigure(0, weight=1)
            frame_canvas.config(height = window_height, width=frame_width)
            
            main_canvas = Canvas(frame_canvas)
            main_canvas.grid(row=0, column=0, sticky="news")
            main_canvas.config(height = window_height, width=frame_width)
            
            vsb = Scrollbar(frame_canvas, orient="vertical", command=main_canvas.yview, 
                            style='scroll.Vertical.TScrollbar')
            vsb.grid(row=0, column=1 ,sticky = 'ns')
            main_canvas.configure(yscrollcommand=vsb.set)
            
            figure_frame = Frame(main_canvas)
            main_canvas.create_window((0, 0), window=figure_frame, anchor='nw')
            figure_frame.config(height = frame_height, width=frame_width)

            x , y = 10, 10
            for i,v in enumerate(self.output_vars[:-3]):
                self.graph_toggles[v] = IntVar()
                cb = Checkbutton(figure_frame, text = v, variable = self.graph_toggles[v])
                cb.place(x = x, y = y)
                cb.select()
                y+=25
            figure_frame.update_idletasks()
            frame_canvas.config(width=frame_width, height=window_height)
            main_canvas.config(scrollregion=(0,0,x,frame_height))
            
            
        ##### After variables have been loaded, we display the compatibility test button #######
        compat_button = Button(self.home_tab, text = 'Test Compatibility of Input Files', 
                               command = self.test_compatibility)
        compat_button.place(x = self.win_lim_x *.59, y = self.win_lim_y*.025)
        
       
        
    def abort_sim(self):
        '''
        The function linked to the abort button in a multivariate analysis. It is linked to
        the next variable button in a univariate analysis. It changes the value of abort, which
        the multiprocessing simulations have access to so they know to abort. It also
        starts a cleanup thread which terminates all lingering processes and COMS. While the simulation
        is aborting, it displays a status update notifying the user that it is aborting.
        '''
        
        self.abort.value = True
        self.cleanup_thread = Thread(target=self.cleanup_processes_and_COMS)
        self.cleanup_thread.start()
        try:
            if self.analysis_type.get() == 'Multivariate Sensitivity' or self.abort_univar_overall.value:
                self.time_rem_label.config(text='Status: Aborting Simulation, '+\
                                           'Please Wait Before Starting New Simulation')
            else:
                self.time_rem_label.config(text='Status: Transitioning to Next Variable')

        except: pass
        Thread(target=self.abort_helper).start()

        
    def abort_helper(self):
        '''
        Waits for the cleanup thread to finish before notifying the user that
        it is ready to start a new simulation.
        '''
        self.cleanup_thread.join()
        try:
            self.time_rem_label.config(text='Status: Ready for New Simulation')
        except:
            pass
        
        
    def abort_univar_overall_fun(self):
        '''
        The function linked to the abort button for a univariate analysis.
        It changes the value of abort_univ_overall, which ensures that no further
        simulations are started after the current univariate analysis is aborted.
        '''
        self.abort_univar_overall.value = True
        self.abort_sim()
        
        
    def cleanup_processes_and_COMS(self):
        '''
        Ensures that all COMS have been closed and that all multiprocessing
        processes are terminated. It also saves the final results to the output folder.
        '''
        try:
            self.current_simulation.close_all_COMS()
            self.current_simulation.terminate_processes()
            try:
                self.current_simulation.lock_to_signal_finish.release()
            except:
                pass
            save_data(self.current_simulation.output_file, 
                                  self.current_simulation.results, 
                                  self.current_simulation.directory, 
                                  self.current_simulation.weights)
            save_graphs(self.current_simulation.output_file, 
                                    self.current_simulation.results, 
                                    self.current_simulation.directory, 
                                    self.current_simulation.weights)
        except:
            self.after(400, self.cleanup_processes_and_COMS)
            



#############################################################################################################
class Simulation(object):
    '''
    A class to encapsulate all attributes and methods associated with a simulation.
    Simulation objects are initialized in the Main_App object function
    "create_simulation_object." There are a number of multiprocessing specific
    data structures here that enable pickling/piping of data across different
    processes. Basically, any variable that needs to be modified by different 
    multiprocessing processes must be one of these special multiprocessing
    data structures. 
    
    Each process pulls a trial number from the task queue and runs an analysis
    with the values associated with that trial number. As soon as a trial is 
    completed, the results are appended to the "results" list, which is a list of 
    pandas dataframes, with each dataframe containing the results of just one trial.
    This is why the results are always concatenated together before plotting or
    saving. By using this multiprocessing results list, the results from each 
    process are shuttled out of the process to be accessed by the Main_App threads.
    
    The lock_to_signal_finish is a lock that is released by the processes only 
    once the simulation is completed or properly aborted at the level of the processes.
    Release of this lock notifies the parent worker_thread that the processes are 
    done.
    '''
    
    def __init__(self, sims_completed, tot_sim, simulation_vars, output_file, directory, 
                 aspen_files, excel_solver_file,abort, vars_to_change, output_value_cells,
                 output_columns, dispatch, weights, save_bkps, warning_keywords, bkp_ref,
                 save_freq=2, num_processes=1):
        self.manager = Manager()
        self.num_processes = min(num_processes, tot_sim) #don't need more processors than trials
        self.tot_sim = tot_sim
        self.sims_completed = sims_completed
        self.save_freq = self.manager.Value('i', save_freq)
        self.abort = abort
        self.simulation_vars = self.manager.dict(simulation_vars) 
        self.output_file = self.manager.Value('s', output_file)
        self.directory = self.manager.Value('s', directory)
        self.aspen_files = []
        for f in range(len(aspen_files)):
            self.aspen_files.append(self.manager.Value('s', aspen_files[f]))
        self.excel_solver_file = self.manager.Value('s', excel_solver_file)
        self.output_value_cells = self.manager.Value('s',output_value_cells)
        self.dispatch = dispatch
        self.results = self.manager.list()
        self.output_columns = output_columns
        self.trial_counter = Value('i',0)
        self.results_lock = self.manager.Lock()
        self.processes = []
        self.current_COMS_pids = self.manager.dict()
        self.pids_to_ignore = self.manager.dict()
        self.find_pids_to_ignore()
        self.output_columns = self.manager.list(output_columns)
        self.vars_to_change = self.manager.list(vars_to_change)
        self.aspenlock = self.manager.Lock()
        self.excellock = self.manager.Lock()
        self.lock_to_signal_finish = self.manager.Lock()
        self.weights = self.manager.list(weights)
        self.save_bkps = self.manager.Value('b',save_bkps)
        self.warning_keywords = self.manager.list(list(warning_keywords))
        self.bkp_ref = self.manager.Value('s',bkp_ref)
          
    
    def run_simulation(self):
        '''
        Saves a copy of input variables in the output folder. It then constructs the 
        task queue and starts the simulations. It waits for the
        lock_to_signal_release before closing all COMS and terminating all processes.
        Finally, it saves the data and graphs to the output folder.
        '''
        
        self.save_copy_of_input_variables()
        
        TASKS = [trial for trial in range(0, self.tot_sim)]
        self.lock_to_signal_finish.acquire()
        if not self.abort.value:
            self.start_simulation(TASKS)
            Thread(target=self.check_mp_errors).start()
        else:
            try:
                self.lock_to_signal_finish.release()
            except:
                pass
        print('Simulation Running\nWaiting for Completion Signal')
        # it cannot acquire the lock until it is released from a thread or 
        # is released above if the simulation is aborted.
        self.lock_to_signal_finish.acquire()
        print('Completion Signal Received')
        self.wait()
        self.close_all_COMS()
        self.terminate_processes()
        self.wait()
            
        save_data(self.output_file, self.results, self.directory, self.weights)
        save_graphs(self.output_file, self.results, self.directory, self.weights)        
        self.abort.value = False    
        
        
    def check_mp_errors(self):
        '''
        Periodically checks the multiprocessing Processes to see if there are errors.
        If errors are found, it prints the error and traceback.
        '''
        while not self.abort.value and any(p.is_alive() for p in self.processes):
            for p in self.processes:
                if p.exception:
                    error, traceback = p.exception
                    print(error)
                    print(traceback)
            sleep(4)
        
        
    def save_copy_of_input_variables(self):
        '''
        Saves a copy of the excel input file to the output folder.
        '''
        df = DataFrame(columns=['trial'] + list(self.vars_to_change))
        for key, value in self.simulation_vars.items():
            df[key[0]] = value
            ntrials = len(value)
        df['trial'] = range(1, ntrials+1)
        df.to_csv(path.join(self.directory.value, 'Sampled_Variable_Distributions.csv'),index=False)
        
        
    def terminate_processes(self):
        '''
        Terminates all multiprocessing processes
        '''
        for p in self.processes:
            p.terminate()
            p.join()
         
    def wait(self, t=0.1):
        '''
        Waits until all processes are dead. It repeatedly calls itself until
        no processes are alive
        '''
        if not any(p.is_alive() for p in self.processes):
            return
        else:
            sleep(t)
            self.wait()
            
            
    def start_simulation(self, tasks):
        '''
        Constructs and starts the multiprocessing processes to run the simulation
        '''
        task_queue = mpQueue()
        for task in tasks:
            task_queue.put(task)

        for i in range(self.num_processes):
            process_args = (self.current_COMS_pids, self.pids_to_ignore, self.aspenlock, 
                            self.excellock, self.aspen_files[i], self.save_bkps,self.excel_solver_file,
                            task_queue, self.abort,self.results_lock, self.results, 
                            self.directory, self.output_columns, self.output_value_cells,
                            self.trial_counter, self.save_freq,self.output_file, self.vars_to_change, 
                            self.output_columns, self.simulation_vars, self.sims_completed, 
                            self.lock_to_signal_finish, self.tot_sim, self.dispatch, self.weights,
                            self.warning_keywords, self.bkp_ref)
            self.processes.append(Process_with_Error_Support(target=multiprocessing_worker, args=process_args))
            
        for p in self.processes:
            p.start()
            
        # need to add STOP at the end of the task queue so the processes know when
        # to stop
        for i in range(self.num_processes):
            task_queue.put('STOP')
        
            
    def close_all_COMS(self):
        '''
        If any COMS objects are in the process of being opened, it waits for 
        them to finish opening so that we make sure to have a handle on them 
        so we can delete them. It then terminates all current COMS processes.
        '''
        
        if self.results:
            self.aspenlock.acquire() # wait for aspen COMS to finish opening
            self.excellock.acquire() # wait for excel COMS to finish opening
            for p in process_iter():
                if p.pid in self.current_COMS_pids:
                    p.terminate()
                    del self.current_COMS_pids[p.pid]
            self.aspenlock.release()
            self.excellock.release()
        else:
            for p in process_iter(): # get a handle on the Excel COM we just made
                if (('aspen' in p.name().lower() or 'apwn' in p.name().lower()) or 'excel' in p.name().lower()) and (
                        p.pid not in self.pids_to_ignore):
                    p.terminate()

    
    def find_pids_to_ignore(self):
        '''
        Searches through all processes running to find any Excel or Aspen
        processes that the user already has running so we make sure we don't
        close those when we try to close all of Illuminate's COMS
        '''
        for p in process_iter():
            if ('aspen' in p.name().lower() or 'apwn' in p.name().lower()) or 'excel' in p.name().lower():
                self.pids_to_ignore[p.pid] = 1
        
        
class Process_with_Error_Support(Process):
    '''
    Adding to the multiprocessing Process object so that we can read error messages 
    from within a running process.
    '''
    
    def __init__(self, *args, **kwargs):
        Process.__init__(self, *args, **kwargs)
        self._pconn, self._cconn = Pipe()
        self._exception = None

    def run(self):
        try:
            Process.run(self)
            self._cconn.send(None)
        except Exception as e:
            tb = traceback.format_exc()
            self._cconn.send((e, tb))
            # raise e  # You can still rise this exception if you need to

    @property
    def exception(self):
        if self._pconn.poll():
            self._exception = self._pconn.recv()
        return self._exception

        
############ GLOBAL FUNCTIONS THAT PROCESSES CAN ACCESS ################
                
def open_aspenCOMS(aspenfilename,dispatch):
    '''
    Opens the Aspen COM object and returns a reference to that object
    as well as the Aspen Tree.
    '''
    aspencom = Dispatch(str(dispatch))
    aspencom.InitFromArchive2(path.abspath(aspenfilename), host_type=0, node='', username='', 
                              password='', working_directory='', failmode=0)
    aspen_tree = aspencom.Tree
    #aspencom.SuppressDialogs = False     
    return aspencom,aspen_tree


def open_excelCOMS(excelfilename):
    '''
    Opens the Excel COM object and returns a reference to that object as well
    as the Excel book
    '''
    pythoncom.CoInitialize()
    excel = DispatchEx('Excel.Application')
    book = excel.Workbooks.Open(path.abspath(excelfilename))
    return excel,book  
   
    
def save_data(outputfilename, results, directory, weights):
    '''
    Saves simulation results to the output directory. The weights are only
    for input variables whose distributions have "linspace mapping" type. 
    It also saves summary statistics for each of the output variables in a second tab.
    '''
    if results: 
        collected_data = concat(results).sort_index() # concatenate and sort by trial number
        if len(weights) > 0:
            collected_data['Probability'] = weights[:len(collected_data)]
        writer = ExcelWriter(directory.value + '/' + outputfilename.value + '.xlsx')
        collected_data.to_excel(writer, sheet_name ='Sheet1')
        stats = collected_data.describe()
        stats.to_excel(writer, sheet_name = 'Summary Stats')
        writer.save()
    
def save_graphs(outputfilename, results, directory, weights):
    '''
    Saves histograms for each input variable and output variable to the output
    directory.
    '''
    if results:
        collected_data = list(filter(lambda x: not isna(x[x.columns[-2]].values[0]), results))
        if len(collected_data) > 1:
            collected_data = concat(collected_data).sort_index()
            for index, var in enumerate(collected_data.columns[:-3]):
                fig = plt.figure()
                fig.set_size_inches(6,6)
                ax = fig.add_axes([0.12, 0.12, 0.85, 0.85])
                if len(weights) > 0:
                     plotweight = weights[:len(collected_data)]
                     num_bins = len(collected_data)
                     plt.hist(collected_data[var], num_bins, weights=plotweight, 
                              facecolor='blue', edgecolor='black', alpha=1.0)
                else:
                    data = collected_data[var]
                    if len(data) == 0:
                        num_bins = 1
                    elif len(data) < 20:
                        num_bins = len(data)
                    else:
                        iqr = subtract(*percentile(data, [75, 25]))
                        bin_width = (2*iqr)/(len(data)**(1/3))
                        if isclose(bin_width, 0.0):
                            num_bins = 1
                        else:
                            num_bins = ceil((max(data) - min(data))/bin_width)
                    plt.hist(collected_data[var], num_bins, facecolor='blue', 
                             edgecolor='black', alpha=1.0)
                ax.set_xlabel(var, Fontsize=14)
                ax.set_ylabel('Count', Fontsize=14)
                ax.ticklabel_format(axis= 'x', style = 'sci', scilimits= (-3,3), Fontsize=12)
                ax.ticklabel_format(axis= 'y', Fontsize=12)
                
                # remove any invalid characters from the variable name so we
                # can properly save the file
                var = var.replace('\\','').replace('/','').replace(
                        ':','').replace('?','').replace('*','').replace(
                                '"','').replace('<','').replace('>','').replace('|','')
                plt.savefig(directory.value + '/' + outputfilename.value + \
                            '_' + var + '.png', format='png')
                try:
                    fig.clf()
                except:
                    pass
                try:
                    plt.close('all')
                except:
                    pass

def multiprocessing_worker(current_COMS_pids, pids_to_ignore, aspenlock, excellock, aspenfilename, save_bkps,
           excelfilename, task_queue, abort, results_lock, results, directory, output_columns, 
           output_value_cells, sim_counter, save_freq, outputfilename, vars_to_change, columns, 
           simulation_vars, sims_completed,lock_to_signal_finish, tot_sim, dispatch, weights, 
           warning_keywords, bkp_ref):
    '''
    The function that hosts the multiprocessing running of simulations. Each
    multiprocessing process is running this function when simulations are being
    actively run. 
    
    Here are the steps of this function:
        1. Opens one Aspen and Excel COM object and registers those process 
        IDs as active COMS processes.
        2. If bkp files are to be saved, then it specifies the bkp directory to
        be in the output folder. Otherwise, the bkp directory will be in a temporary folder
        3. Draws trials from the task queue and runs simulations as follows:
            a. Run an Aspen simulation via aspen_run
            b. Save the simulation results to a bkp file
            c. Run the Excel calculator file to calculate all outputs
            d. Append the resulting pandas dataframe of outputs to the results list
            e. Given the save frequency and the number of sims completed, maybe save data collected
            f. Aspen COM objects accumulate memory over many trials. Therefore, if the 
            memory consumption is above 94%, then the Aspen COM is terminated and
            reinitialized to conserve memory. 
            e. After the task queue is emptied, release the lock to signal finish
    '''

    local_pids_to_ignore = {} #local as in not accessible by other processes
    local_pids = {}
    aspenlock.acquire()
    ######## Register any process IDs to ignore before initializing COMS #####
    for p in process_iter():
        if ('aspen' in p.name().lower() or 'apwn' in p.name().lower()):
            local_pids_to_ignore[p.pid] = 1
            
    if not abort.value:
        aspencom,aspen_tree = open_aspenCOMS(aspenfilename.value, dispatch)
    for p in process_iter(): # get a handle on the Aspen COM we just made
        if ('aspen' in p.name().lower() or 'apwn' in p.name().lower()) and (
                p.pid not in local_pids_to_ignore):
            local_pids[p.pid] = 1
    aspenlock.release()
    
    excellock.acquire()
    if not abort.value:
        excel,book = open_excelCOMS(excelfilename.value)
    excellock.release() 
    for p in process_iter(): # get a handle on the Excel COM we just made
        if (('aspen' in p.name().lower() or 'apwn' in p.name().lower()) or 'excel' in p.name().lower()) and (
                p.pid not in pids_to_ignore):
            current_COMS_pids[p.pid] = 1
    
    ######## Determine the directory of the bkp files ##############
    if save_bkps.value:
        aspen_temp_dir = directory.value
    else:
        aspen_temp_dir = path.dirname(aspenfilename.value)
        bkp_name = ''.join([choice(string.ascii_letters + string.digits) for n in range(10)]) + '.bkp'
        full_bkp_name = path.join(aspen_temp_dir,bkp_name)
    
            
    #####################  Run the simulation ########################
    for trial_num in iter(task_queue.get, 'STOP'):
        if abort.value:
            try:
                lock_to_signal_finish.release()
            except:
                continue
        
        # run Aspen simulation
        aspencom, case_values, run_summary, aspen_tree = aspen_run(
                aspencom, aspen_tree, simulation_vars, trial_num, vars_to_change, 
                directory, warning_keywords) 
        
        # save bkp file and tell excel calculator to point to correct bkp file
        if save_bkps.value:
            full_bkp_name = path.join(aspen_temp_dir, 'Trial_' + str(trial_num + 1) + '.bkp')
        aspencom.SaveAs(full_bkp_name)
        book.Sheets('Set-up').Evaluate(bkp_ref.value).Value = full_bkp_name
        
        # run the Excel Calculator
        result = excel_run(excel, book, aspencom, aspen_tree, case_values, columns, run_summary, 
                             trial_num, output_value_cells, directory, aspenfilename.value)
        
        # dump results into results list and maybe save
        results_lock.acquire()
        results.append(result) 
        sim_counter.value = len(results)
        if sim_counter.value % save_freq.value == 0:
            save_data(outputfilename, results, directory, weights)
            save_graphs(outputfilename, results, directory, weights)
        sims_completed.value += 1
        results_lock.release()
        

        # Restart Aspen COM if memory consumption is an issue
        if virtual_memory().percent > 94:
            aspenlock.acquire()
            for p in process_iter():
                if p.pid in local_pids:
                    p.terminate()
                    del current_COMS_pids[p.pid]
                    del local_pids[p.pid]
                    
            for p in process_iter():
                if ('aspen' in p.name().lower() or 'apwn' in p.name().lower()):
                    local_pids_to_ignore[p.pid] = 1
            if not abort.value:
                aspencom,aspen_tree = open_aspenCOMS(aspenfilename.value,dispatch)
            
            for p in process_iter(): #register the pids of COMS objects
                if ('aspen' in p.name().lower() or 'apwn' in p.name().lower()) and (
                        p.pid not in local_pids_to_ignore):
                    current_COMS_pids[p.pid] = 1
                    local_pids[p.pid] = 1
            aspenlock.release()
        
        # restart the Aspen Engine (very important for conservation of memory)            
        aspencom.Engine.Host(host_type=0, node='', username='', password='', working_directory='')    

    try:
        lock_to_signal_finish.release()
    except:
        pass
            


def aspen_run(aspencom, aspen_tree, simulation_vars, trial, vars_to_change, directory, warning_keywords):
    '''
    Runs an Aspen simulation given a trial number and the given variable distributions.
    After running Aspen, it calls find_errors_warnings to identify any errors
    or warnings of interest Aspen outputed.
    '''
    
    variable_values = {}
    for (aspen_variable, aspen_call, fortran_index), dist in simulation_vars.items():
        # change value in Aspen tree
        aspen_tree.FindNode(aspen_call).Value = dist[trial]
        if type(dist[trial]) == str:
            # if fortran variable
            variable_values[aspen_variable] = float(dist[trial][fortran_index[0]:fortran_index[1]])
        else:
            variable_values[aspen_variable] = dist[trial]
    
    ########## STORE THE VARIABLE VALUES FOR THIS TRIAL  ##########
    case_values = []
    for v in vars_to_change:
        case_values.append(variable_values[v])    
    
    aspencom.Reinit()
    aspencom.Engine.Run2()
    run_summary = format_run_summary(retrieve_run_summary(aspencom))
    #errors, warnings = find_errors_warnings(aspencom, warning_keywords)
    
    return aspencom, case_values, run_summary, aspen_tree



def excel_run(excel, book, aspencom, aspen_tree, case_values, columns, run_summary, 
                trial_num, output_value_cells, directory, aspenfilename):
    '''
    Pulls stream results from Aspen simulation .bkp file and run macros to 
    solve for output values. Saves simulation results in a pandas dataframe and
    returns that dataframe.
    '''
   

    if aspen_tree.FindNode(r"\Data\Results Summary\Run-Status\Output\SPDATE").Value == None:
        for file in listdir(path.dirname(aspenfilename)):
            if file.endswith(".his"):
                copyfile(path.join(path.dirname(aspenfilename), file), path.join(
                        directory.value,'history_for_trial_' + str(trial_num) + '.his'))
        
        dfstreams = DataFrame(columns=columns)
        dfstreams.loc[trial_num+1] = case_values + [None]*(len(columns)-1-len(case_values))+["Aspen Failed to Converge"] 
        return dfstreams
    excel.Run('sub_ClearSumData_ASPEN')
    excel.Run('sub_GetSumData_ASPEN')
    excel.Calculate()
    excel.Run('solvedcfror')
      
    results_df = DataFrame(columns=columns)
    results_df.loc[trial_num+1] = case_values + [x.Value for x in book.Sheets('Output').Evaluate(
            output_value_cells.value)] + [run_summary[0]] + [run_summary[1]]  + [";  ".join(run_summary[2])]
    return results_df


    
def retrieve_run_summary(aspencom):
    '''
    Retrieves the run summary from the Aspen Tree. If Aspen failed to converge, 
    this will be empty
    '''
    obj = aspencom.Tree
    node = r'\Data\Results Summary\Run-Status\Output\PER_ERROR'
    not_done = True
    counter = 1
    summary_status = []
    while not_done:
        try:
            line = obj.FindNode(node + '\\' +  str(counter)).Value
            summary_status.append(line)
            counter += 1
        except:
            not_done = False
    return summary_status
        

def format_run_summary(summary):
    '''
    Counts the number of warnings and errors in the summary and also formats the run summary so it
    can be saved in the output excel file.
    '''
    
    if not summary:
        return [0, 0, []]
    error_count = 0
    warning_count = 0 
        
    summary_formatted = []
    summary_formatted.append('')
        
    num_lines = len(summary)
    i = 0 
    while i < num_lines:
        if not len(summary[i]) > 0:
            if i > 0:
                if ":" in summary[(i-1)]:
                    i += 1
                    continue
            summary_formatted.append(summary[i])
        elif "error" in summary[i].lower():
            summary_formatted[len(summary_formatted)-1] = summary_formatted[len(
                    summary_formatted)-1] + summary[i]
            scan_errors = True
            while scan_errors:
                i += 1
                if len(summary[i]) > 0:
                    summary_formatted[len(summary_formatted)-1] = summary_formatted[len(
                            summary_formatted)-1] + ' ' + summary[i]
                    error_count += (summary[i].count('  ') + 1)
                else:
                    i -= 1
                    scan_errors = False                    
        elif "warning" in summary[i].lower():
            summary_formatted[len(summary_formatted)-1] = summary_formatted[len(
                    summary_formatted)-1] + ' ' + summary[i]
            scan_warnings = True
            while scan_warnings:
                i += 1
                if len(summary[i]) > 0:
                    summary_formatted[len(summary_formatted)-1] = summary_formatted[len(
                            summary_formatted)-1] + summary[i]
                    warning_count += (summary[i].count('  ') + 1)
                else:
                    i -= 1
                    scan_warnings = False   
        else:
            summary_formatted[len(summary_formatted)-1] = summary_formatted[len(
                    summary_formatted)-1] + summary[i]
        i += 1

    return [error_count, warning_count, summary_formatted]


def OnFocusIn(event):
    '''
    Brings the GUI into focus but ensures that it is not always in focus (i.e.
    the user can click off of it and it does not remain in front of other windows)
    '''
    if type(event.widget).__name__ == 'MainApp':
        event.widget.attributes('-topmost', False)        

        
        
#############################################################################################################
class Compatibility_Test(object):
    '''
    A class to encapsulate the compatibility test. It has a status queue that is read
    from the GUI to print status updates to the GUI.
    '''
    def __init__(self):
        self.status_queue = Queue()
        
    def compatibility_test(self, excel_input_file, calculator_file, aspen_file, dispatch):
        '''
        Tests for compatibility issues in the user's input files. First tests
        the Excel Calculator file then the Aspen model. The status updates from this
        function are live updated on the GUI in the Main_App.
        '''
        
        aspen_dir = self.copy_aspen_to_temp_dir(aspen_file)
        
        self.status_queue.put((False, 'Testing Compatibility of Excel Input File...'))
        excel_input_errors_found = self.test_excel_input_file(excel_input_file)
        if excel_input_errors_found:
            self.status_queue.put((True, 'Finished Testing Excel Input File, Please Fix Errors'))
        else:
            self.status_queue.put((False, 'SUCCESS: Excel Input File is Compatible with Illuminate'))
        self.status_queue.put((False, 'Testing Compatibility of Excel Calculator File...'))
        errors_found = self.test_calculator_file(calculator_file, aspen_file)
        if errors_found:
            self.status_queue.put((True, 'Finished Testing Excel Calculator File, Please Fix Errors'))
        else:
            self.status_queue.put((False, 'SUCCESS: Excel Calculator File is Compatible with Illuminate'))
            
        if excel_input_errors_found:
            self.status_queue.put((True, 'Cannot test compatibility of Aspen model until Excel input file is compatible. Please fix errors with Excel input file and test compatibility again'))
        else:
            self.status_queue.put((False, 'Testing Compatibility of Aspen Model...'))
            errors_found = self.test_aspen_file(aspen_dir, excel_input_file, dispatch)
            if errors_found:
                self.status_queue.put((True, 'Finished Testing Aspen Model, Please Fix Errors'))
            else:
                self.status_queue.put((False, 'SUCCESS: Aspen Model is Compatible with Illuminate'))
        self.status_queue.put((False, 'Finished with Compatibility Test'))
        
        
    def copy_aspen_to_temp_dir(self, aspen_file):
        '''
        Copies the aspen .apw or .bkp file provided by the user to a temporary
        directory within the 'Output' folder. This is done in order to encapsulate
        all of the extra files that Aspen outputs so that they can be easily removed
        if Aspen crashes or is aborted. It first checks to see if this directory exists,
        and if it does exist, then it deletes the temporary directory and all
        of its contents.
        '''
        
        output_directory = path.join(path.dirname(aspen_file),'Output')
        if not path.exists(output_directory):
            makedirs(output_directory)
        if not path.exists(path.join(output_directory,'Temp')):
            makedirs(path.join(output_directory,'Temp'))
        
        temp_directory = path.join(path.dirname(aspen_file),'Output','Temp')
        
        # delete the directory if it exists
        try:
            rmdir(temp_directory)
        except: 
            pass
        try:
            rmtree(temp_directory)
        except: 
            pass
        makedirs(temp_directory)
        aspen_file_temp = path.join(temp_directory,path.basename(aspen_file))
        copyfile(aspen_file, aspen_file_temp)
        
        return aspen_file_temp
        
    def test_excel_input_file(self, excel_input_file):
        
        errors_found = False
        col_types = {'Variable Name': str, 'Variable Aspen Call': str, 'Distribution Parameters': str, 
                     'Bounds': str, 'Fortran Call':str, 'Fortran Value to Change': str, 
                     'Distribution Type': str, 'Toggle': bool}
        try:
            df = read_excel(excel_input_file, sheet_name='Inputs', dtype=col_types)
        except:
            self.status_queue.put((True, 'There must be a sheet titled "Inputs" in the Excel input file and another sheet titled "Warning Messages"'))
            errors_found = True
            return
        try:
            df = read_excel(excel_input_file, sheet_name='Warning Messages')
        except:
            self.status_queue.put((True, 'There must be a sheet titled "Inputs" in the Excel input file and another sheet titled "Warning Messages"'))
            errors_found = True
        
        
        required_columns = ['Variable Name', 'Variable Aspen Call', 'Distribution Parameters', 
                     'Bounds', 'Fortran Call', 'Fortran Value to Change', 
                     'Distribution Type', 'Toggle']
        df = read_excel(excel_input_file, sheet_name='Inputs', dtype=col_types)
        user_columns = set(df.columns)
        
        for col in required_columns:
            if col not in user_columns:
                self.status_queue.put((True, 'Column "' + col + '" must be in the "Inputs" tab of the Excel input file'))
                errors_found = True
        
        return errors_found
        
        
        
    def test_aspen_file(self, aspen_file,excel_input_file, dispatch):
        '''
        Makes sure the Aspen model can be opened, tests to make sure that all
        Aspen nodes specified in the Excel input file exist in the Aspen model and 
        are not None. Finally, it makes sure that for any Fortran variables, the value
        to change can be found within the Fortran variable string.
        '''
        
        errors_found = False
        ######### Open Aspen COM and get a handle on that COM so we can terminate it #######
        aspens_to_ignore = set()
        for p in process_iter():
            if 'aspen' in p.name().lower() or 'apwn' in p.name().lower():
                aspens_to_ignore.add(p.pid)       
        self.status_queue.put((False, 'Opening Aspen Model...'))
        try:
            aspencom, obj = open_aspenCOMS(aspen_file, dispatch)
        except:
            self.status_queue.put((True, 'Aspen model cannot be opened'))
            errors_found = True
            return
        for p in process_iter():
            if ('aspen' in p.name().lower() or 'apwn' in p.name().lower()) and p.pid not in aspens_to_ignore:
                aspen_to_delete = p
            
                
        # Make sure that all nodes in the tree exist
        self.status_queue.put((False, 'Testing Aspen Paths Specified in Input File...'))
        col_types = {'Variable Name': str, 'Variable Aspen Call': str, 'Distribution Parameters': str, 
                     'Bounds': str, 'Fortran Call':str, 'Fortran Value to Change': str, 
                     'Distribution Type': str, 'Toggle': bool}
        df = read_excel(open(excel_input_file,'rb'), dtype=col_types)
        for index, row in df.iterrows():
            if row['Toggle']: 
                try:
                    if obj.FindNode(row['Variable Aspen Call']).Value is None:     
                        self.status_queue.put((True, 'The value at the node "'+ row['Variable Aspen Call'] + \
                                          '" for variable "' + row['Variable Name'] + \
                                          '" is None. Are you sure this is the right path?'))
                        errors_found = True
                except:
                    self.status_queue.put((True, 'Aspen call "'+ row['Variable Aspen Call'] + \
                                      '" for variable "' + row['Variable Name'] + \
                                      '" does not exist in the Aspen model'))
                    errors_found = True
        for index, row in df.iterrows():
            if row['Toggle'] and row['Fortran Call']:
                if row['Fortran Value to Change'] not in row['Fortran Call']:
                    self.status_queue.put((True, 'The fortran value to change "' + \
                                      row['Fortran Value to Change'] + '" for variable "' + \
                                      row['Variable Name'] + '" is not in the Fortran call "' +
                                      row['Fortran Call'] + '"'))
                    errors_found = True
         
        aspen_to_delete.terminate()
        return errors_found
    
    
    def test_calculator_file(self, calculator_file, aspen_file):
        '''
        Tests for compatibility issues in the Excel Calculator file. First makes sure
        the Output tab exists and is configured properly. It then makes sure the .bkp
        reference in the setup tab is configured as expected. Finally, it tests
        to make sure macros are named properly and are functional.
        '''
        
        errors_found = False
        ########### Open Excel COM and get a handle on it to terminate it later ########
        excels_to_ignore = set()
        for p in process_iter():
            if 'excel' in p.name().lower():
                excels_to_ignore.add(p.pid)
        excel, book = open_excelCOMS(calculator_file)
        for p in process_iter():
            if 'excel' in p.name().lower() and p.pid not in excels_to_ignore:
                excel_to_delete = p
        
        ########### Make sure that the output tab exists  ###################
        output_tab_exists = False
        try:
            book.Sheets('Output')
            output_tab_exists = True
        except:
            self.status_queue.put((True,'"Output" tab missing from Excel calculator .xlsm file. Please add this tab'))
            errors_found = True
            
            
        ########## Make sure output tab is set up as it is supposed to be  ######## 
        if output_tab_exists:
            if any(str(v) != "Variable Name" for v in book.Sheets('Output').Evaluate('B2')):
                self.status_queue.put((True,'Output tab is not configured properly. The column header for '+\
                                  '"Variable Name" should be in B2 so that the first variable name is in B3'))
                errors_found = True
            elif any(str(v) != "Value" for v in book.Sheets('Output').Evaluate('C2')):
                self.status_queue.put((True,'Output tab is not configured properly. The column header '+\
                                  'for "Value" should be in C2 so that the first variable value is in C3'))
                errors_found = True
                
                
        ######### Make sure the bkp file reference is where it should be #########
        try:
            vba_code = book.VBProject.VBComponents("GelAllData").CodeModule.Lines(1,500000)
        except:
            try:
                vba_code = book.VBProject.VBComponents("GetAllData").CodeModule.Lines(1,500000)
            except:
                vba_code = ''
        
        i=0
        get_data_VBA = ""
        while i < len(vba_code):
            if vba_code[i:i+31] == 'Public Sub sub_GetSumData_ASPEN':
                s_ind = i
                while vba_code[i-7:i] != 'End Sub':
                    i += 1
                get_data_VBA = vba_code[s_ind:i]
                break
            i +=1
        
        if get_data_VBA:
            bkp_reference_cell = findall(r"RTrim\(Worksheets\(\"Set-up\"\)\.Range\(\"([A-Z]+[0-9]+)\"\)\.VALUE", 
                                                  get_data_VBA)[0]
        else:
            bkp_reference_cell = 'B1'
        
        
        
        
        setup_tab_functional = True
        try:
            book.Sheets('Set-up')
        except:
            setup_tab_functional = False
            self.status_queue.put((True,'"Set-up" tab missing from Excel calculator .xlsm file. '+\
                              'Please rename this tab.'))
            errors_found = True
        try:
            filename, file_extension = path.splitext(book.Sheets('Set-up').Evaluate(bkp_reference_cell).Value)
            if not (file_extension=='.bkp' or file_extension == '.apw'):
                self.status_queue.put((True,'In the "Set-up" tab, the name of the .apw or .bkp '+\
                                  'should be in cell B1. If, however, you have made VBA '+\
                                  'accessible to Illuminate, then you can have this bkp '+\
                                  'reference in a different location. If it is not in B1, '+\
                                  'then the reference in "sub_GetSumData_ASPEN" must be updated. '+\
                                  'If the location of this reference needs '+\
                                  'to be changed, make sure that you also change it in the "sub_GetSumData" macro'))
                errors_found = True
        except:
            setup_tab_functional = False
            self.status_queue.put((True,'In the "Set-up" tab, the name of the .apw or .bkp '+\
                                  'should be in cell B1. If, however, you have made VBA accessible to '+\
                                  'Illuminate, then you can have this bkp reference in a different '+\
                                  'location. If it is not in B1, then the reference in "sub_GetSumData_ASPEN" '+\
                                  'must be updated. If the location of this reference needs '+\
                                  'to be changed, make sure that you also change it in the "sub_GetSumData" macro'))
            errors_found = True
        
            
        ####################  Test all important macros ########################
        
        try:
            clear_load_cell = findall(r"Range\(\"([A-Z]+[0-9]+)\"\)\.End\(xlDown\)\.Row", 
                                      book.VBProject.VBComponents("GelAllData").CodeModule.Lines(1,500000))[0]
        except:
            try:
                clear_load_cell = findall(r"Range\(\"([A-Z]+[0-9]+)\"\)\.End\(xlDown\)\.Row", 
                                          book.VBProject.VBComponents("GetAllData").CodeModule.Lines(1,500000))[0]
            except:
                clear_load_cell = 'C7'
        try:
            excel.Run('sub_ClearSumData_ASPEN')
            try:
                if book.Sheets('aspen').Evaluate(clear_load_cell).Value is not None:
                    self.status_queue.put((True,'Excel macro sub_ClearSumData_ASPEN does not appear to '+\
                                      'be working. Values in column C of sheet "aspen" are not being cleared.'))
                    errors_found = True
            except:
                pass
        except:
            self.status_queue.put((True, 'Macro with name "sub_ClearSumData_ASPEN" does '+\
                              'not exist or is broken'))
            errors_found = True
        
    
        
        
        if setup_tab_functional:
            try:
                book.Sheets('Set-up').Evaluate(bkp_reference_cell).Value = aspen_file
                excel.Run('sub_GetSumData_ASPEN')
                if book.Sheets('aspen').Evaluate(clear_load_cell).Value is None:
                    self.status_queue.put((True,'"sub_GetSumData_ASPEN" does not appear to be '+\
                                      'working. Values should be populated in column C of sheet "aspen"'))
                    errors_found = True
                    
            except:
                self.status_queue.put((True,'Macro with name "sub_GetSumData_ASPEN" does not '+\
                                  'exist or is broken'))
                errors_found = True
            
        
        try:
            excel.Run('solvedcfror')
        except:
            self.status_queue.put((True, 'Macro with name "solvedcfror" does not exist or is broken.'))
            errors_found = True
            return
    
        
        try:
            module1_VBA = book.VBProject.VBComponents("Module1").CodeModule.Lines(1,50000000)
            vba_code_access = True
        except:
            self.status_queue.put((True, 'Unable to access "solvedcfror" VBA code and therefore cannot test'+\
                              '"solvedcfror" functionality. ' +\
                              'If you would like Illuminate to be able to test this, you must enable access'+\
                              'by opening the .xlsm file and going to' +\
                              'File -> Options -> Trust Center -> Trust Center Settings -> '+\
                              'Macro Settings -> Trust Access to VBA project object model'))
            errors_found = True
            vba_code_access = False
        if vba_code_access:
            i=0
            DCFROR_VBA = ""
            while i < len(module1_VBA):
                if module1_VBA[i:i+15] == 'Sub solvedcfror':
                    s_ind = i
                    while module1_VBA[i-7:i] != 'End Sub':
                        i += 1
                    DCFROR_VBA = module1_VBA[s_ind:i]
                    break
                i +=1
                
            DCFROR_cells = findall(
                    r"Range\(\"([A-Z]+[0-9]+)\"\)\.GoalSeek Goal\:\=0\, ChangingCell\:\=Range\(\"([A-Z]+[0-9]+)\"\)", 
                    DCFROR_VBA)
            DCFROR_sheetname = findall(r"Sheets\(\"(.*)\"\).Select",DCFROR_VBA)[0]
            
            if not DCFROR_cells:
                self.status_queue.put((True, 'Cannot find VBA code for "solvedcfror" in Module1. This is not critical, '+\
                                  ', but it means that Illuminate cannot test the functionality of this macro. If you are confident' +\
                                  'that it is working, then disregard this message.'))
                errors_found = True
            else:
                goal_seek, change_cell = DCFROR_cells[0]
                seek_val = float(book.Sheets(DCFROR_sheetname).Evaluate(goal_seek).Value)
                book.Sheets(DCFROR_sheetname).Evaluate(change_cell).Value = 5.0
                
                
                
                if isclose(book.Sheets(DCFROR_sheetname).Evaluate(goal_seek).Value, seek_val):
                    self.status_queue.put((True, 'The "goal seek" and "change cell" cells indicated '+\
                                      'in the "solvedcfror" code do not appear to be linked.'+\
                                      'Make sure these are the correct cells referenced in the macro' ))
                    errors_found = True
                
                excel.Run('solvedcfror')
                if not isclose(float(book.Sheets(DCFROR_sheetname).Evaluate(goal_seek).Value), 0.0):
                    self.status_queue.put((True, 'The "solvedcfror" function is not minimizing the '+\
                                      '"goal seek" cell to 0 as it should be.' ))
                    errors_found = True
    
            
            
        excel_to_delete.terminate()
        return errors_found
    
    

        
if __name__ == "__main__":
    freeze_support()
    main_app = MainApp()
    main_app.mainloop()
    
    ####### If there is a current simulation, make sure it is aborted before closing
    if main_app.current_simulation:
        main_app.abort_univar_overall.value = True
        main_app.abort_sim()
        print('Waiting for Clearance to Exit Program...')
        main_app.current_simulation.wait()
        main_app.worker_thread.join()
        print('Cleaning Up Processes/Threads...')
        main_app.cleanup_thread.join()
    exit()
        
        
    

    

