# -*- coding: utf-8 -*-
"""
Created on Sat Dec 15 19:58:47 2018

@author: MENGstudents
"""

from tkinter import Tk, StringVar,E,W,Canvas,END, IntVar, Checkbutton, Label
from tkinter.ttk import Entry, Button, Radiobutton, OptionMenu, Labelframe, Scrollbar, Notebook, Frame
from tkinter.filedialog import askopenfilename
from threading import Thread
from pandas import ExcelWriter, DataFrame, concat, isna, read_excel
from multiprocessing import Value, Manager, Lock, Queue, Process, cpu_count
from time import time, sleep
from datetime import datetime
from numpy import linspace, random, histogram
from psutil import process_iter, virtual_memory
from win32com.client import Dispatch, DispatchEx
import pythoncom
from os import path, makedirs
from shutil import copyfile
from multiprocessing import freeze_support
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
from winreg import EnumKey, CreateKey, EnumValue, HKEY_CLASSES_ROOT
from re import search, findall
from random import choices
 





class MainApp(Tk):

    def __init__(self):
        Tk.__init__(self)
        #self.iconbitmap('01_128x128.ico')
        self.notebook = Notebook(self)
        self.wm_title("Illuminate")
        self.notebook.grid()
        self.construct_home_tab()
        
        self.simulations = []
        self.current_simulation = None
        self.current_tab = None
        self.abort = Value('b', False)
        self.abort_univar_overall = Value('b', False)
        self.simulation_vars = {}
        self.attributes('-topmost', True)
        self.focus_force()
        self.bind('<FocusIn>', OnFocusIn)
        #self.attributes("-topmost", True)
        self.tot_sim_num = 0
        self.sims_completed = Value('i',0)
        self.start_time = None
        self.univar_plot_counter = 1
        self.finished_figures = []
        self.univar_row_num=0
        self.last_results_plotted = None
        self.last_update = None
        self.win_lim_x = self.winfo_screenwidth()//2
        self.win_lim_y = int(self.winfo_screenheight()*0.9)
        self.geometry(str(self.win_lim_x) + 'x' + str(self.win_lim_y) + '+0+0')
        self.worker_thread = None
        self.display_tab = None
        self.mapping_pdfs = {}
        self.simulation_dist, self.simulation_vars = {}, {}
        
#        style = ttk.Style()
#        style.configure('Kim.TButton', foreground='blue', bg='blue', activebackground='red', relief='raised')
#        style.configure('label.TLabel', background='red',foreground='blue')
#        style.configure('TabStyle.TNotebook.Tab', background='green')
#        style.configure('frame.TFrame', background='blue')

#        style.configure('Wild.TButton', background='black', foreground='white', font=('Helvetica', 12, 'bold'))
#        style.map('Wild.TButton',
#              foreground=[('disabled', 'yellow'),
#                    ('pressed', 'red'),
#                    ('active', 'blue')],
#                          background=[('disabled', 'magenta'),
#                    ('pressed', '!focus', 'cyan'),
#                    ('active', 'green')],
#                    highlightcolor=[('focus', 'green'),
#                        ('!focus', 'red')],
#                                    relief=[('pressed', 'groove'),
#                ('!pressed', 'ridge')])
#
#        style.theme_create("st_app", parent='alt', settings={
#        "TButton":     {"configure": {'foreground':'maroon', 'relief': 'raised'}}})
        #style.theme_use("st_app")

#              "TNotebook.Tab": {
#            "configure": {"padding": [5, 1], "background": mygreen },
#            "map":       {"background": [("selected", myred)],
#                          "expand": [("selected", [1, 1, 1, 0])] } } 



    def construct_home_tab(self):
        self.load_aspen_versions()
        self.home_tab = Frame(self.notebook, style= 'frame.TFrame')
        self.notebook.add(self.home_tab, text = 'File Upload Tab')

        for i in range (5,20):
            Label(self.home_tab, text='                       ').grid(row=100,column=i,columnspan=1)
        for i in range(106,160):
            Label(self.home_tab, text=' ').grid(row=i,column=0,columnspan=1)

        
        
        space= Label(self.home_tab, text=" ",font='Helvetica 2')
        space.grid(row=0, column= 1, sticky = E, padx = 5, pady =4)
        space.rowconfigure(0, minsize = 15)
        

        Button(self.home_tab, text='Upload Simulation Inputs',
        command=self.open_excel_file).grid(row=1,column=1, sticky = E, pady = 5,padx = 5)
        self.input_csv_entry = Entry(self.home_tab)
        self.input_csv_entry.grid(row=1, column=2)
        
        Button(self.home_tab, 
              text="Upload Aspen Model",
              command=self.open_aspen_file).grid(row=2, column = 1,sticky = E,
              pady = 5,padx = 5)
        self.aspen_file_entry = Entry(self.home_tab)
        self.aspen_file_entry.grid(row=2, column=2,pady = 5,padx = 5)
        
        Button(self.home_tab,
              text="Upload Excel Model",
              command=self.open_solver_file).grid(row=3,column = 1,sticky = E,
              pady = 5,padx = 5)
        self.excel_solver_entry = Entry(self.home_tab)
        self.excel_solver_entry.grid(row=3, column=2,pady = 5,padx = 5)
        
        Button(self.home_tab,
              text="Load Data",
              command=self.make_new_tab).grid(row=9,column = 3,sticky = E,
              pady = 5,padx = 5)
        
        test= Label(self.home_tab, 
                  text=" ",font='Helvetica 2')
        test.grid(row=4, column= 1, sticky = E, padx = 5)
        test.rowconfigure(4, minsize = 4)
        
        self.analysis_type = StringVar(self.home_tab)
        self.analysis_type.set("Choose Analysis Type")
        
        OptionMenu(self.home_tab, self.analysis_type,"Choose Analysis Type", "Single Point Analysis","Univariate Sensitivity", 
                "Multivariate Sensitivity", style = 'raised.TMenubutton').grid(row = 9,sticky = E,column = 2,padx =5, pady = 5)
                        
        select_aspen = Labelframe(self.home_tab, text='Select Aspen Version:')
        select_aspen.grid(row = 5,column = 1, columnspan = 3, sticky = W,pady = 10,padx = 10)

        self.select_version = StringVar()
        row = 6
        column = 0
        aspen_versions = []
        for key,value in self.aspen_versions.items():
            aspen_versions.append(key + '      ')

        aspen_versions.sort(key=lambda x: -1*float(x[1:-6]))

        for i, version in enumerate(aspen_versions):
            v = Radiobutton(select_aspen, text= version, variable=self.select_version, value = self.aspen_versions[version[:-6]])
            v.grid(row=row,column= column, sticky=W)
            if i == 0:
                v.invoke()
            
            column += 1
            if column == 4:
                column = 0
                row += 1

            
        #Label(self.home_tab, text='CPU Core Count :').grid(row=3, column=1, sticky=E)
        #self.num_processes_entry = Entry(self.home_tab)
        #self.num_processes_entry.grid(row=3, column=2, pady=5, padx=5)

    def make_new_tab(self):
        if self.current_tab:
            self.notebook.forget(self.current_tab)
            self.current_tab = None
        if self.analysis_type.get() == 'Choose Analysis Type':
            print("ERROR: Select an Analysis Type")
 
        elif self.analysis_type.get() == 'Univariate Sensitivity':
            self.current_tab = Frame(self.notebook)
            self.notebook.add(self.current_tab,text = "Univariate Analysis")
            ##############Tab 2 LABELS##################
            
            Label(self.current_tab, 
                  text="Save As :").place(x=149,y=6)
            self.save_as_entry= Entry(self.current_tab)
            self.save_as_entry.grid(row=4, column=2, sticky=E, pady=6)
            self.save_as_entry.config(width =18)
            
            Label(self.current_tab,text = ".xlsx").grid(row = 4, column = 3, sticky = W)
            
            Label(self.current_tab, text='CPU Core Count :').place(x=104,y=39)
            self.num_processes_entry = Entry(self.current_tab)
            self.num_processes_entry.grid(row=5, column=2, sticky=E, pady=6)
            self.num_processes_entry.config(width=18)
            
            rec_core = int(cpu_count()//2)
            Label(self.current_tab, text = 'Recommended Count: ' + str(rec_core)).grid(row = 5, column = 3, sticky = W)
            
            Label(self.current_tab, text ='').grid(row= 13, column =1)
            Button(self.current_tab,
                   text='Run Univariate Sensitivity Analysis',
                   command=self.initialize_univar_analysis).grid(row=14,
                   column=3, columnspan=2,
                   pady=4)
            Button(self.current_tab,
                   text='Display Variable Distributions',
                   command=self.plot_init_dist).grid(row=14,
                   column=1, columnspan=2, sticky = W,
                   pady=4)
            Button(self.current_tab,
                   text='Fill  # Trials',
                   command=self.fill_num_trials).grid(row=7, columnspan = 2, sticky =E,
                   column=1,
                   pady=4)
            self.fill_num_sims = Entry(self.current_tab)
            self.fill_num_sims.grid(row=7,column = 3,sticky =W, pady =2, padx = 2)
            self.fill_num_sims.config(width = 10)
            
        elif  self.analysis_type.get() == 'Single Point Analysis':
            self.current_tab = Frame(self.notebook)
            self.notebook.add(self.current_tab, text = 'Single Point')
             
            Label(self.current_tab, 
                  text="Save As :").grid(row=0, column= 0, sticky = E, pady = 5, padx = 5)
            self.save_as_entry = Entry(self.current_tab)
            self.save_as_entry.grid(row=0, column=1, pady = 5)
            Label(self.current_tab,text = ".xlsx").place(x = 295, y= 6)
            
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
                  text="Number of Simulations :").grid(row=4, column= 1, sticky = E, pady = 5, padx = 5)
            self.num_sim_entry = Entry(self.current_tab)
            self.num_sim_entry.grid(row=4, column=2,pady = 5,padx = 5)
            
            rec_core = int(cpu_count()//2)
            Label(self.current_tab, text='CPU Core Count (Recommend '+ str(rec_core)+ '):').grid(row=5, column=1, sticky=E)
            self.num_processes_entry = Entry(self.current_tab)
            self.num_processes_entry.grid(row=5, column=2, pady=5, padx=5)
                               
            Button(self.current_tab,
                   text='Run Multivariate Analysis',
                   command=self.initialize_multivar_analysis).grid(row=6,
                   column=3, columnspan=2, sticky=W, pady=4)
            Button(self.current_tab,
                   text='Display Variable Distributions',
                   command=self.plot_init_dist).grid(row=6,
                   column=1, columnspan=2, sticky=W, pady=4)
            
        self.load_variables_into_GUI()
        self.notebook.select(self.current_tab)
        
    def conv_title(self, s, pad=False):
        if len(s) > 37:
            return s[:34] + '...'
        elif pad:
            return s.ljust(37)    
        return s

    def load_aspen_versions(self):
        
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
        
    def load_variables_into_GUI(self):
        single_pt_vars = []
        univariate_vars = []
        multivariate_vars = []
        type_of_analysis = self.analysis_type.get()
        gui_excel_input = str(self.input_csv_entry.get())
        col_types = {'Variable Name': str, 'Variable Aspen Call': str, 'Distribution Parameters': str, 'Bounds': str, 'Fortran Call':str, 'Fortran Value to Change': str, 'Distribution Type': str, 'Toggle': bool}
        df = read_excel(open(gui_excel_input,'rb'), dtype=col_types)
        for index, row in df.iterrows():
            if row['Toggle']:          
                if type_of_analysis =='Single Point Analysis':
                    single_pt_vars.append((row["Variable Name"], float(row["Distribution Parameters"].split(',')[0].strip())))
                elif type_of_analysis == 'Multivariate Analysis':
                    multivariate_vars.append(row["Variable Name"])
                else:
                    univariate_vars.append((
                            row["Variable Name"], row["Distribution Type"].strip().lower(
                                    ), row['Distribution Parameters'].split(',')))
                        
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
            vsb = Scrollbar(frame_canvas, orient="vertical", command=canvas.yview, style='scroll.Vertical.TScrollbar')
            vsb.grid(row=0, column=1,sticky = 'ns')
            canvas.configure(yscrollcommand=vsb.set)
            
            # Create a frame to contain the variables
            frame_vars = Frame(canvas)
            canvas.create_window((0, 0), window=frame_vars, anchor='nw')
            frame_vars.config(height = '5c', width='10c')
            
            self.sp_row_num = 0
            for name,value in single_pt_vars:
                self.sp_row_num += 1
                key = str(self.sp_row_num)
                Label(frame_vars, 
                text= self.conv_title(name,pad=True)).grid(row=self.sp_row_num, column= 1, sticky = E,pady = 5,padx = 5)
                key=Entry(frame_vars)
                key.grid(row=self.sp_row_num, column=2,pady = 5,padx = 5)
                key.delete(first=0,last=END)
                key.insert(0,str(value))
                self.sp_value_entries[name]= key
                
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
            # Create a frame for the canvas with non-zero row&column weights
            #label_frame = Labelframe(self.current_tab)
            #label_frame.grid(row=9, column=1, columnspan=3)
            frame_canvas1 = Frame(self.current_tab)
            frame_canvas1.grid(row=9, column=1, columnspan =3, pady=(5, 0))
            frame_canvas1.grid_rowconfigure(0, weight=1)
            frame_canvas1.grid_columnconfigure(0, weight=1)
            frame_canvas1.config(height = '3c', width='13c')
            
            # Add a canvas in the canvas frame
            canvas1 = Canvas(frame_canvas1)
            canvas1.grid(row=0, column=0, sticky="news")
            canvas1.config(height = '3c', width='13c')
            
            # Link a scrollbar to the canvas
            vsb = Scrollbar(frame_canvas1, orient="vertical", command=canvas1.yview, style='scroll.Vertical.TScrollbar')
            vsb.grid(row=0, column=1,sticky = 'ns')
            canvas1.configure(yscrollcommand=vsb.set)
            
            # Create a frame to contain the variables
            frame_vars1 = Frame(canvas1)
            frame_vars1.config(height = '3c', width='13c')
            canvas1.create_window((0, 0), window=frame_vars1, anchor='nw')
            for name, format_of_data, vals in univariate_vars:
                Label(frame_vars1, 
                text= self.conv_title(name, True)).grid(row=self.univar_row_num, column= 1,pady = 5,padx = 5)
                Label(frame_vars1, 
                text= self.conv_title(format_of_data)).grid(row=self.univar_row_num, column= 2,pady = 5,padx = 5)
                
                if not(format_of_data == 'linspace' or format_of_data == 'list' or 'mapping' in format_of_data):
                    key2=Entry(frame_vars1)
                    key2.grid(row=self.univar_row_num, column=3,pady = 5,padx = 5)
                    #key2.insert(0,univariate_sims)
                    self.univar_ntrials_entries[name]= key2
                else:
                    if "mapping" in format_of_data:
                        Label(frame_vars1,text= self.conv_title(vals[-1].strip())).grid(row=self.univar_row_num, column= 3,pady = 5,padx = 5)
                    elif format_of_data == 'linspace':
                        
                        Label(frame_vars1,text= self.conv_title(str(vals[2]).strip())).grid(row=self.univar_row_num, column= 3,pady = 5,padx = 5)
                    else:
                        Label(frame_vars1,text= self.conv_title(str(len(vals)))).grid(row=self.univar_row_num, column= 3,pady = 5,padx = 5)
                self.univar_row_num += 1
                
            # Update vars frames idle tasks to let tkinter calculate variable sizes
            frame_vars1.update_idletasks()
            # Determine the size of the Canvas
            
            frame_canvas1.config(width='13c', height='3c')
            
            # Set the canvas scrolling region
            canvas1.config(scrollregion=canvas1.bbox("all"))
            
    def get_distributions(self):
        if self.analysis_type.get() == 'Univariate Sensitivity':
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
            self.simulation_vars, self.simulation_dist = self.construct_distributions(ntrials=max_num_sim)
            for (aspen_variable, aspen_call, fortran_index), dist in self.simulation_vars.items():
                if aspen_variable in self.univar_ntrials_entries:
                    try:
                        num_trials_per_var = int(self.univar_ntrials_entries[aspen_variable].get())
                    except:
                        num_trials_per_var = 1
                    self.simulation_vars[(aspen_variable, aspen_call, fortran_index)] = dist[:num_trials_per_var]
                    self.simulation_dist[aspen_variable] = self.simulation_dist[aspen_variable][:num_trials_per_var]                
        else:
            try: 
                ntrials = int(self.num_sim_entry.get())
            except:
                ntrials=1
            self.simulation_vars, self.simulation_dist = self.construct_distributions(ntrials=ntrials) 
            
            
    def construct_distributions(self, ntrials=1):
        '''
        Given the excel input from the user in the GUI, produce a list_of_variables
        the user wants to change as well as their distributions that should be 
        randomly sampled from. 
        '''
        
        gui_excel_input = str(self.input_csv_entry.get())
        col_types = {'Variable Name': str, 'Variable Aspen Call': str, 'Distribution Parameters': str, 'Bounds': str, 'Fortran Call':str, 'Fortran Value to Change': str, 'Distribution Type': str, 'Toggle': bool}
        df = read_excel(open(gui_excel_input,'rb'), dtype=col_types)
        simulation_vars = {}
        simulation_dist = {}
        for index, row in df.iterrows():
            if row['Toggle']:
                dist_type = row['Distribution Type'].lower()
                aspen_variable = row['Variable Name']
                aspen_call = row['Variable Aspen Call']
                bounds = row['Bounds'].split(',')
                lb = float(bounds[0].strip())
                ub = float(bounds[1].strip())
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
                            pdf_approx = self.sample_gauss(mean, std_dev, lb_dist, ub_dist, 10000)

                    if 'pareto' in dist_type:
                        shape, scale = float(dist_vars[0].strip()), float(dist_vars[1].strip())
                        if self.analysis_type.get() != "Univariate Sensitivity":
                            distribution = self.sample_pareto(shape, scale, lb, ub, ntrials)
                        else:
                            pdf_approx = self.sample_pareto(shape, scale, lb_dist, ub_dist, num_trials)
                    if 'poisson' in dist_type:
                        lambda_p = float(dist_vars[0].strip())
                        if self.analysis_type.get() != "Univariate Sensitivity":
                            distribution = self.sample_poisson(lambda_p, lb, ub, ntrials)
                        else:
                            pdf_approx =self.sample_poisson(lambda_p, lb_dist, ub_dist, num_trials)
                        
                    if self.analysis_type.get() == "Univariate Sensitivity":
                        bin_width = (ub_dist - lb_dist)/num_trials
                        lb_pdf = lb_dist - 0.5*bin_width
                        ub_pdf = ub_dist + 0.5*bin_width
                        pdf, bin_edges = histogram(pdf_approx, bins=linspace(lb_pdf, ub_pdf, num_trials+1), density=True)
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
#                        print(distribution)
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
                    lb_uniform, ub_uniform = float(lb_ub[0].strip()), float(lb_ub[1].strip())
                    distribution = self.sample_uniform(lb_uniform, ub_uniform, lb, ub, ntrials)
  
                if distribution is None:
                    Label(self.current_tab, text= 'ERROR: Distribution Parameters for ' + aspen_variable + ' are NOT valid', fg='red').grid(row=10, column=1, columnspan=3)
                    Label(self.current_tab, text='Please Adjust Distribution Parameters in Input File and Restart Illuminate', fg='red').grid(row=11,column=1,columnspan=3)
                    return {}, {}
#                print(distribution)
                simulation_dist[aspen_variable] = distribution[:]
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
                        distribution2.append(self.make_fortran(fortran_call, fortran_index, float(v)))
                    distribution = distribution2
                simulation_vars[(aspen_variable, aspen_call, fortran_index)] = distribution
        return simulation_vars, simulation_dist
    
    
    def sample_gauss(self,mean, std, lb, ub, ntrials):
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
        return fortran_call[:fortran_index[0]] + str(val) + fortran_call[fortran_index[1]:]
    
    def disp_sp_mfsp(self):
        if self.current_simulation and self.current_simulation.results:
            row = 4
            for output_var, toggled in self.graph_toggles.items():
                if toggled.get():
                    output_val = self.current_simulation.results[0].at[1, output_var]
                    if isna(output_val):
                        Label(self.current_tab, text= 'Aspen Failed to Converge', font='Helvetica 10 bold',fg='red').grid(row=row, column = 1)
                        break
                    output_val = "{:,}".format(float("%.2f" % output_val))
                    Label(self.current_tab, text=str(output_var) + '= ' + output_val).grid(
                    row=row, column = 1)
                    row += 1
        else:
            self.after(5000, self.disp_sp_mfsp)

    
    def single_point_analysis(self):
        self.store_user_inputs()
        self.get_distributions()
        if not self.simulation_vars:
            return
        # update simulation variable values based on user input in GUI
        for (aspen_variable, aspen_call, fortran_index), values in self.simulation_vars.items():
            self.simulation_vars[(aspen_variable, aspen_call, fortran_index)] = [float(
                    self.sp_value_entries[aspen_variable].get())]
        self.create_simulation_object(self.simulation_vars, self.vars_to_change, self.output_file, self.num_trial)
        self.run_simulations()
        
    
    def run_multivar_sens(self):
        self.store_user_inputs()
        if len(self.simulation_vars) == 0:
            self.get_distributions()
        if not self.simulation_vars:
            return
        self.create_simulation_object(self.simulation_vars, self.vars_to_change, self.output_file, self.num_trial)
        self.run_simulations()
        
        
    def run_univ_sens(self):
        self.store_user_inputs()
        if len(self.simulation_vars) == 0:
            self.get_distributions()
        if not self.simulation_vars:
            return
        for (aspen_variable, aspen_call, fortran_index), values in self.simulation_vars.items():
            weights = self.mapping_pdfs.get(aspen_variable, [])
            self.create_simulation_object({(aspen_variable, aspen_call, 
                                            fortran_index): values}, [aspen_variable], 
        self.output_file+'_'+aspen_variable, len(values), weights)
        self.run_simulations()
    
    
    def store_user_inputs(self):
        self.aspen_file = str(self.aspen_file_entry.get())
        try:
            self.num_processes = int(self.num_processes_entry.get())
        except:
            self.num_processes = 1
        self.excel_solver_file= str(self.excel_solver_entry.get())
        try:
            self.num_trial = int(self.num_sim_entry.get())
        except: 
            self.num_trial = 1
        self.output_file = str(self.save_as_entry.get())
        self.input_csv = str(self.input_csv_entry.get())
        
        self.vars_to_change = []
        
        gui_excel_input = str(self.input_csv_entry.get())
        col_types = {'Variable Name': str, 'Variable Aspen Call': str, 'Distribution Parameters': str, 'Bounds': str, 'Fortran Call':str, 'Fortran Value to Change': str, 'Distribution Type': str, 'Toggle': bool}
        df = read_excel(open(gui_excel_input,'rb'), dtype=col_types)
        for index, row in df.iterrows():
            if row['Toggle']:
                self.vars_to_change.append(row["Variable Name"])
        
        
    def run_simulations(self):
        
        for sim in self.simulations:
            self.start_time = time()
            self.current_simulation = sim
            self.current_simulation.init_sims()
            
            if self.abort_univar_overall.value:
                self.abort.value = True
            self.univar_plot_counter += 1
            self.last_update = None
        self.simulations = []
    
    def parse_output_vars(self):
        self.excel_solver_file= str(self.excel_solver_entry.get())
        excels_to_ignore = {}
        for p in process_iter():
            if 'excel' in p.name().lower():
                excels_to_ignore[p.pid] = 1
        excel, book = open_excelCOMS(self.excel_solver_file)
        self.output_vars = []
        row_counter = 3
        while True:
            var_name = book.Sheets('Output').Evaluate("B" + str(row_counter)).Value
            if var_name:
                units = book.Sheets('Output').Evaluate("D" + str(row_counter)).Value
                column_name = var_name + ' (' + units + ')' if units else var_name
                self.output_vars.append(column_name)
            else:
                break
            row_counter += 1
        self.output_value_cells = "C3:C" + str(row_counter - 1)
        self.output_vars += ['Aspen Errors']
        for p in process_iter():
            if 'excel' in p.name().lower() and p.pid not in excels_to_ignore:
                p.terminate()
            
        
    def create_simulation_object(self, simulation_vars, vars_to_change, output_file, num_trial, weights=[]):
#        print(simulation_vars)
#        print(vars_to_change)
        self.output_columns = vars_to_change + self.output_vars
        print(self.output_columns)
        output_directory = path.join(path.dirname(str(self.input_csv_entry.get())),'Output/',datetime.now().strftime('%Y-%m-%d_%H-%M-%S'))
        makedirs(output_directory)
        copyfile(path.abspath(str(self.input_csv_entry.get())), path.join(output_directory,'Input_variables.xlsx'))
        new_sim = Simulation(self.sims_completed, num_trial, simulation_vars, output_file, output_directory,
                             self.aspen_file, self.excel_solver_file, self.abort, vars_to_change, self.output_value_cells,
                             self.output_columns, self.select_version.get(), weights, save_freq=2, num_processes=self.num_processes)
        self.simulations.append(new_sim)
        self.tot_sim_num += num_trial
        
        
    def initialize_single_point(self):
        if self.worker_thread and self.worker_thread.isAlive():
            print('simulations already running')
            return
        self.worker_thread = Thread(
                target=lambda: self.single_point_analysis())
        self.worker_thread.start()
        self.after(5000, self.disp_sp_mfsp)
        
    def initialize_univar_analysis(self):
        if self.worker_thread and self.worker_thread.isAlive():
            print('simulations already running')
            return
        self.worker_thread = Thread(
            target=lambda: self.run_univ_sens())
        self.worker_thread.start()
        self.status_label = None
        self.time_rem_label = None
        self.after(5000, self.univar_gui_update)

    
    def initialize_multivar_analysis(self):
        if self.worker_thread and self.worker_thread.isAlive():
            print('simulations already running')
            return
        self.worker_thread = Thread(
            target=lambda: self.run_multivar_sens())
        self.worker_thread.start()
        print('started new worker thread')
        self.status_label = None
        self.time_rem_label = None
        self.multivar_gui_update()
        
        
    def disp_status_update(self):
        if self.current_simulation and not self.abort.value:
            if len(self.current_simulation.results) == self.current_simulation.tot_sim:
                status_update = 'Status: Simulation Complete'
            else:
                status_update = 'Status: Simulation Running | {} Results Collected'.format(
                        len(self.current_simulation.results))
            return status_update
        return None
        
    def disp_time_remaining(self, status_update):
        if self.start_time and self.sims_completed.value != self.last_update:
            if not status_update:
                status_update = ''
            self.last_update = self.sims_completed.value
            elapsed_time = time() - self.start_time
            if self.sims_completed.value > 0:
                remaining_time = ((elapsed_time / self.sims_completed.value) * (self.tot_sim_num - self.sims_completed.value))//60
                hours, minutes = divmod(remaining_time, 60)
                tmp = Label(self.display_tab, text=status_update + ' | ' + 'Time Remaining: {} Hours, {} Minutes    '.format(int(hours), int(minutes)))
            else:
                tmp = Label(self.display_tab, text=status_update + ' | ' +'Time Remaining: N/A')
            tmp.place(x=6, y=4)
            if self.time_rem_label:
                self.time_rem_label.destroy()
            self.time_rem_label = tmp
            
            
    def plot_on_GUI(self):
        status_label = None
        if not self.simulation_dist:
            return
        if not self.display_tab:
            self.display_tab = Frame(self.notebook)
            self.notebook.add(self.display_tab,text = "Simulation Status")
            self.notebook.select(self.display_tab)
            status_label = Label(self.display_tab, text='Setting Up Simulation...')
            status_label.place(x=6, y=4)
            self.init_plots_constructed = False
            self.plots_dictionary = {}
            
        if not self.current_simulation:
            return
        if len(self.current_simulation.results) == self.last_results_plotted:
            return
        self.last_results_plotted = len(self.current_simulation.results)
        
        if self.current_simulation.results:
            results_to_plot = list(filter(lambda x: not isna(x[self.current_simulation.output_columns[len(
                    self.current_simulation.vars_to_change)]].values[0]), self.current_simulation.results))
            if len(results_to_plot) == 0:
                results_filtered = DataFrame(columns=self.output_columns)
                results_unfiltered = results_filtered
            else:
                results_filtered = concat(results_to_plot).sort_index()
                results_unfiltered = concat(self.current_simulation.results).sort_index()
            
        else:
            results_filtered = DataFrame(columns=self.output_columns)
            results_unfiltered = results_filtered
            
        if not self.init_plots_constructed:
            results_fig_list =[]
            num_bins = 15
            for var, toggled in self.graph_toggles.items():
                if toggled.get():
                    fig = Figure(figsize = (3,3), facecolor=[240/255,240/255,237/255], tight_layout=False)
                    ax = fig.add_subplot(111)
                    ax.hist(results_filtered[var], num_bins, facecolor='blue', edgecolor='black', alpha=1.0)
                    ax.set_title(self.conv_title(var))
                    ax.ticklabel_format(axis= 'x', style = 'sci', scilimits= (-3,3))
                    self.plots_dictionary[var] = ax
                    results_fig_list.append(fig)
            
            inputs_fig_list = []
            for var, values in self.simulation_dist.items():
                fig = Figure(figsize = (3,3), facecolor=[240/255,240/255,237/255], tight_layout=False)
                a = fig.add_subplot(111)
                _, bins, _ = a.hist(self.simulation_dist[var], num_bins, facecolor='white', edgecolor='black',alpha=1.0)
                a.hist(results_unfiltered[var], bins=bins, facecolor='blue',edgecolor='black', alpha=1.0)
                a.set_title(self.conv_title(var))
                a.ticklabel_format(axis= 'x', style = 'sci', scilimits= (-3,3))
               # a.get_xaxis().set_major_formatter(matplotlib.ticker.ScalarFormatter(set_powerlimits((n,m))
                self.plots_dictionary[var] = a
                inputs_fig_list.append(fig)
            
            row_num = 0
            frame_width = self.win_lim_x - 30
            num_graphs_per_row = frame_width//250
            frame_height = 60+(230*((len(inputs_fig_list) + len(results_fig_list)+1)//num_graphs_per_row + 1))  
            window_height = self.win_lim_y - 30
            
            frame_canvas = Frame(self.display_tab)
            frame_canvas.grid(row=row_num, column=1, columnspan = 3,pady=(5, 0))
            frame_canvas.grid_rowconfigure(0, weight=1)
            frame_canvas.grid_columnconfigure(0, weight=1)
            frame_canvas.config(height = window_height, width=frame_width)
            
            main_canvas = Canvas(frame_canvas)
            main_canvas.grid(row=0, column=0, sticky="news")
            main_canvas.config(height = window_height, width=frame_width)
            
            vsb = Scrollbar(frame_canvas, orient="vertical", command=main_canvas.yview, style='scroll.Vertical.TScrollbar')
            vsb.grid(row=0, column=1,sticky = 'ns')
            main_canvas.configure(yscrollcommand=vsb.set)
            
            figure_frame = Frame(main_canvas)
            main_canvas.create_window((0, 0), window=figure_frame, anchor='nw')
            figure_frame.config(height = frame_height, width=frame_width)
            if status_label:
                status_label.destroy()
        
    #        row_num = 0
    #        column = False
    #        count = 1
    #        for figs in results_fig_list + inputs_fig_list:
    #            figure_canvas = FigureCanvasTkAgg(figs, master=figure_frame)
    #            if column:
    #                col = 4
    #            else:
    #                col = 1
    #            #figure_canvas.draw()
    #            figure_canvas.get_tk_widget().grid(
    #                    row=row_num, column=col,columnspan=2, rowspan = 5, pady = 5,padx = 8, sticky=E)
    #            #figure_canvas._tkcanvas.grid(row=row_num, column = 0,columnspan = 10, rowspan = 10, sticky= W+E+N+S, pady = 5,padx = 5)
    #            if column:
    #                row_num += 5
    #            column = not column
    #            count += 1
            
            
            
            count = 0
            x, y = 10, 30
            output_dis = Label(figure_frame, text = 'Outputs:', font='Helvetica 10 bold')
            output_dis.place(x = x, y = y)
            y += 20
            for figs in results_fig_list:
                figure_canvas = FigureCanvasTkAgg(figs, master=figure_frame)
                x = 10 + 250*(count % num_graphs_per_row)
                figure_canvas.get_tk_widget().place(x = x, y= y, width = 240, height =220)
                if (count+1) % num_graphs_per_row==0:
                    y += 230
                count += 1
            y += 230
            line= Label(figure_frame, text = '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------')
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
            Button(self.display_tab, text = "Abort", command=self.abort_sim).place(x=(4*self.win_lim_x)//5, y = 5)
        else:
            for f in self.plots_dictionary.values():
                f.cla()
                f.clear()
            num_bins = 15
            for output_var, toggled in self.graph_toggles.items():
                if toggled.get():
                    self.plots_dictionary[output_var].hist(
                            results_filtered[output_var], num_bins, facecolor='blue', edgecolor='black', alpha=1.0)
                    self.plots_dictionary[output_var].set_title(self.conv_title(output_var))
                    self.plots_dictionary[output_var].ticklabel_format(axis= 'x', style = 'sci', scilimits= (-3,3))
            for var, values in self.simulation_dist.items():
                _, bins, _ = self.plots_dictionary[var].hist(self.simulation_dist[var], num_bins, facecolor='white', edgecolor='black',alpha=1.0)
                self.plots_dictionary[var].hist(results_unfiltered[var], bins=bins, facecolor='blue', edgecolor='black', alpha=1.0)
                self.plots_dictionary[var].set_title(self.conv_title(var))
                self.plots_dictionary[var].ticklabel_format(axis= 'x', style = 'sci', scilimits= (-3,3))


            for fig in self.graphs_displayed:
                fig.draw()
    
                
            
    def plot_univ_on_GUI(self):
        status_label = None
        if not self.simulation_dist:
            return
        
        if not self.display_tab:
            self.display_tab = Frame(self.notebook)
            self.notebook.add(self.display_tab,text = "Simulation Status")
            self.notebook.select(self.display_tab)
            status_label = Label(self.display_tab, text='Setting Up Simulation...')
            status_label.place(x=6, y=4)
            self.init_plots_constructed = False
            self.plots_dictionary = {}

            
        if not self.current_simulation:
            return
        if len(self.current_simulation.results) == self.last_results_plotted:
            return
        
        self.last_results_plotted = len(self.current_simulation.results)
        
        
        current_var = self.current_simulation.vars_to_change[0]
        if self.current_simulation.results:
            results_to_plot = list(filter(lambda x: not isna(x[self.current_simulation.output_columns[len(
                    self.current_simulation.vars_to_change)]].values[0]), self.current_simulation.results))
            if len(results_to_plot) == 0:
                results_filtered = DataFrame(columns=self.current_simulation.output_columns)
                results_unfiltered = results_filtered
            else:
                results_filtered = concat(results_to_plot).sort_index()
                results_unfiltered = concat(self.current_simulation.results).sort_index()
        else:
            results_filtered = DataFrame(columns=self.current_simulation.output_columns)
            results_unfiltered = results_filtered
            
        if not self.init_plots_constructed:
            num_bins = 15
            fig_list = []                
            for var, values in self.simulation_dist.items():
                self.plots_dictionary[var] = {}
                fig = Figure(figsize = (3,3), facecolor=[240/255,240/255,237/255])
                a = fig.add_subplot(111)
                _, bins, _ = a.hist(self.simulation_dist[var], num_bins, facecolor='white', edgecolor='black',alpha=1.0)
                #a.hist(results_unfiltered[var], bins=bins, facecolor='blue',edgecolor='black', alpha=1.0)
                a.set_title(self.conv_title(var))
                a.ticklabel_format(axis= 'x', style = 'sci', scilimits= (-3,3))
                fig_list.append(fig)
                self.plots_dictionary[var][var] = a
                self.num_toggled = 0
                for output_var, toggled in self.graph_toggles.items():
                    if toggled.get():
                        self.num_toggled += 1
                        fig = Figure(figsize = (3,3), facecolor=[240/255,240/255,237/255])
                        ax = fig.add_subplot(111)
                        ax.hist(results_filtered[output_var], num_bins, facecolor='blue', edgecolor='black', alpha=1.0)
                        ax.set_title(self.conv_title(output_var))
                        ax.ticklabel_format(axis= 'x', style = 'sci', scilimits= (-3,3))
                        fig_list.append(fig)
                        self.plots_dictionary[var][output_var] = ax
                
        
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
            
            hsb = Scrollbar(frame_canvas, orient="horizontal", command=main_canvas.xview, style='scroll.Horizontal.TScrollbar')
            hsb.grid(row=1, column=0,sticky = 'we')
            main_canvas.configure(xscrollcommand=hsb.set)
            
            vsb = Scrollbar(frame_canvas, orient="vertical", command=main_canvas.yview, style='scroll.Vertical.TScrollbar')
            vsb.grid(row=0, column=1,sticky = 'ns')
            main_canvas.configure(yscrollcommand=vsb.set)
            
            figure_frame = Frame(main_canvas)
            main_canvas.create_window((0, 0), window=figure_frame, anchor='nw')
            figure_frame.config(height = frame_height, width=graphs_frame_width)
        
    
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
        
    
            Button(self.display_tab, text = "Next Variable", command=self.abort_sim).place(x=self.win_lim_x - 110, y=3)
            
            Button(self.display_tab, text = "Abort", command=self.abort_univar_overall_fun).place(x=self.win_lim_x-190, y=3)
        else:
            for f in self.plots_dictionary[current_var].values():
                f.cla()
                f.clear()
            num_bins = 15
            for output_var, toggled in self.graph_toggles.items():
                if toggled.get():
                    if len(self.current_simulation.weights) > 0:
                        weights = self.current_simulation.weights[0:len(results_filtered)]
                        self.plots_dictionary[current_var][output_var].hist(
                            results_filtered[output_var], num_bins, weights=weights, facecolor='blue', edgecolor='black', alpha=1.0)
                    else:
                        self.plots_dictionary[current_var][output_var].hist(
                                results_filtered[output_var], num_bins, facecolor='blue', edgecolor='black', alpha=1.0)
                    self.plots_dictionary[current_var][output_var].set_title(self.conv_title(output_var))
                    self.plots_dictionary[current_var][output_var].ticklabel_format(axis= 'x', style = 'sci', scilimits= (-3,3))
            _, bins, _ = self.plots_dictionary[current_var][current_var].hist(self.simulation_dist[current_var], num_bins, facecolor='white', edgecolor='black',alpha=1.0)
            self.plots_dictionary[current_var][current_var].hist(results_unfiltered[current_var], bins=bins, facecolor='blue', edgecolor='black', alpha=1.0)
            self.plots_dictionary[current_var][current_var].set_title(self.conv_title(current_var))
            self.plots_dictionary[current_var][current_var].ticklabel_format(axis= 'x', style = 'sci', scilimits= (-3,3))


            for fig in self.graphs_displayed:
                fig.draw()
    
        
            
    def plot_init_dist(self):
        '''
        This function will plot the distribution of variable calls prior to running
        the simulation. This will enable users to see whether the distributions are as they expected.
        
        '''

#        if self.display_tab:
#                self.notebook.forget(self.display_tab)
        self.get_distributions()  
        if not self.simulation_dist:
            return
        
#        self.display_tab = Frame(self.notebook)
#        self.notebook.add(self.display_tab,text = "Results (Graphed)")
        fig_list =[]
        for var, values in self.simulation_dist.items():
            fig = Figure(figsize = (3,3), facecolor=[240/255,240/255,237/255])
            a = fig.add_subplot(111)
            num_bins = 15
            a.hist(values, num_bins, facecolor='blue', edgecolor='black', alpha=1.0)
            a.set_title(self.conv_title(var))
            a.ticklabel_format(axis= 'x', style = 'sci', scilimits= (-3,3))
            fig_list.append(fig)
            
        if self.univar_row_num != 0:
            row_num = 17
        else:
            row_num = 10
        
        frame_width = self.win_lim_x - 30
        num_graphs_per_row = frame_width//250
        frame_height = 30+(230*((len(fig_list)-1)//num_graphs_per_row + 1)) 
        if self.univar_row_num != 0:
            
            window_height = self.win_lim_y - 330
        else:
            window_height = self.win_lim_y - 160
        
        frame_canvas = Frame(self.current_tab)
        frame_canvas.grid(row=row_num, column=1, columnspan = 3,pady=(5, 0))
        frame_canvas.grid_rowconfigure(0, weight=1)
        frame_canvas.grid_columnconfigure(0, weight=1)
        frame_canvas.config(height = window_height, width=frame_width)
        
        main_canvas = Canvas(frame_canvas)
        main_canvas.grid(row=0, column=0, sticky="news")
        main_canvas.config(height = window_height, width=frame_width)
        
        
        vsb = Scrollbar(frame_canvas, orient="vertical", command=main_canvas.yview, style='scroll.Vertical.TScrollbar')
        vsb.grid(row=0, column=1,sticky = 'ns')
        main_canvas.configure(yscrollcommand=vsb.set)
        
        figure_frame = Frame(main_canvas)
        main_canvas.create_window((0, 0), window=figure_frame, anchor='nw')
        figure_frame.config(height = frame_height, width=frame_width)
        
        count = 0
        x, y = 10, 10
        output_dis = Label(figure_frame, text = 'Inputs:', font='Helvetica 10 bold')
        output_dis.place(x = x, y = y)
        y = 30
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
        self.plot_univ_on_GUI()
        self.disp_time_remaining(self.disp_status_update())
        
        self.after(10000, self.univar_gui_update)
        
        
    def multivar_gui_update(self):
        self.plot_on_GUI()
        self.disp_time_remaining(self.disp_status_update())
        self.after(10000, self.multivar_gui_update)
        
        
    
    def fill_num_trials(self):
        ntrials = self.fill_num_sims.get()
        for name, slot in self.univar_ntrials_entries.items():
            slot.delete(0, END)
            slot.insert(0, ntrials)
        

    def open_excel_file(self):
        filename = askopenfilename(title = "Select file", filetypes = ((".xlsx Files","*.xlsx"),))
        self.input_csv_entry.delete(0, END)
        self.input_csv_entry.insert(0, filename)
                
        
        
    def open_aspen_file(self):
        filename = askopenfilename(title = "Select file", filetypes = (("Aspen Models",["*.bkp", "*.apw"]),))
        self.aspen_file_entry.delete(0, END)
        self.aspen_file_entry.insert(0, filename)
    
    
    def open_solver_file(self):
        filename = askopenfilename(title = "Select file", filetypes = ((".xlsm Files","*.xlsm"),))
        self.excel_solver_entry.delete(0, END)
        self.excel_solver_entry.insert(0, filename)
        if filename:
            plot_output_disp_thread = Thread(target=self.graph_toggle)
            plot_output_disp_thread.start()
            self.wait= Label(self.home_tab, text="Wait While Output Variables Are Loading ...")
            self.wait.grid(row=6, column= 1, columnspan = 2, sticky = E,pady = 5,padx = 5)
       
    def graph_toggle(self):
        self.parse_output_vars()
        self.graph_toggles = {}
        if len(self.output_vars) < 10:
            self.disp_output_vars= Labelframe(self.home_tab, text='Output Variables to Graph:')
            self.disp_output_vars.grid(row = 3,column = 1, columnspan = 2, pady = 10, padx = 10, sticky = E )
            count = 1

            for i,v in enumerate(self.output_vars[:-1]):
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
            
            vsb = Scrollbar(frame_canvas, orient="vertical", command=main_canvas.yview, style='scroll.Vertical.TScrollbar')
            vsb.grid(row=0, column=1 ,sticky = 'ns')
            main_canvas.configure(yscrollcommand=vsb.set)
            
            figure_frame = Frame(main_canvas)
            main_canvas.create_window((0, 0), window=figure_frame, anchor='nw')
            figure_frame.config(height = frame_height, width=frame_width)

        
            x , y = 10, 10
            self.graphs_displayed = []
            for i,v in enumerate(self.output_vars[:-1]):
                self.graph_toggles[v] = IntVar()
                cb = Checkbutton(figure_frame, text = v, variable = self.graph_toggles[v])
                cb.place(x = x, y = y)
                cb.select()
                y+=25
            figure_frame.update_idletasks()
            frame_canvas.config(width=frame_width, height=window_height)
            main_canvas.config(scrollregion=(0,0,x,frame_height))
        
       
        
    def abort_sim(self):
        self.abort.value = True
        self.cleanup_thread = Thread(target=self.cleanup_processes_and_COMS)
        self.cleanup_thread.start()
        
    def abort_univar_overall_fun(self):
        self.abort_univar_overall.value = True
        self.abort_sim()
        
    def cleanup_processes_and_COMS(self):
        try:
            self.current_simulation.close_all_COMS()
            self.current_simulation.terminate_processes()
            try:
                self.current_simulation.lock_to_signal_finish.release()
            except:
                pass
            save_data(self.current_simulation.output_file, self.current_simulation.results, self.current_simulation.directory, self.current_simulation.weights)
        except:
            self.after(1000, self.cleanup_processes_and_COMS)
            


        
        
       
        
        
        
        
################################################################################    
        
        
        
        
class Simulation(object):
    def __init__(self, sims_completed, tot_sim, simulation_vars, output_file, directory, 
                 aspen_file, excel_solver_file,abort, vars_to_change, output_value_cells,
                 output_columns, dispatch, weights, save_freq=10, num_processes=1):
        self.manager = Manager()
        self.num_processes = min(num_processes, tot_sim)
        self.tot_sim = tot_sim
        self.sims_completed = sims_completed
        self.save_freq = self.manager.Value('i', save_freq)
        self.abort = abort
        self.simulation_vars = self.manager.dict(simulation_vars) 
        self.output_file = self.manager.Value('s', output_file)
        self.directory = self.manager.Value('s', directory)
#        self.output_directory = path.join(directory,'/Output/',datetime.now().strftime('%Y-%m-%d_%H-%M-%S'))
        self.aspen_file = self.manager.Value('s', aspen_file)
        self.excel_solver_file = self.manager.Value('s', excel_solver_file)
        self.output_value_cells = self.manager.Value('s',output_value_cells)
        self.dispatch = dispatch
        self.results = self.manager.list()
        self.trial_counter = Value('i',0)
        self.results_lock = Lock()
        self.processes = []
        self.current_COMS_pids = self.manager.dict()
        self.pids_to_ignore = self.manager.dict()
        self.find_pids_to_ignore()
        self.output_columns = self.manager.list(output_columns)
        self.vars_to_change = self.manager.list(vars_to_change)
        self.aspenlock = Lock()
        self.excellock = Lock()
        self.lock_to_signal_finish = Lock()
        self.weights = self.manager.list(weights)
          
    
    def init_sims(self):
        
        df = DataFrame(columns=['trial'] + list(self.vars_to_change))
#        print(self.vars_to_change)
        for key, value in self.simulation_vars.items():
            df[key[0]] = value
            ntrials = len(value)
        df['trial'] = range(1, ntrials+1)
        df.to_csv('trials_for_'+self.output_file.value + '.csv',index=False)
        
        
        
        TASKS = [trial for trial in range(0, self.tot_sim)]
        self.lock_to_signal_finish.acquire()
        if not self.abort.value:
            self.run_sim(TASKS)
        else:
            try:
                self.lock_to_signal_finish.release()
            except:
                pass
        print('waiting for acquire')
        self.lock_to_signal_finish.acquire()
        print('acquired')
        self.wait()
        self.close_all_COMS()
        self.terminate_processes()
        self.wait()
            
        save_data(self.output_file, self.results, self.directory, self.weights)
        self.abort.value = False    
        
        
    def terminate_processes(self):
        for p in self.processes:
            p.terminate()
            p.join()
         
    def wait(self, t=0.25):
        if not any(p.is_alive() for p in self.processes):
            return
        else:
            sleep(t)
            self.wait()
            
            
    def run_sim(self, tasks):
        task_queue = Queue()
        for task in tasks:
            task_queue.put(task)

        for i in range(self.num_processes):
            self.processes.append(Process(target=worker, args=(self.current_COMS_pids, self.pids_to_ignore, 
                                                                self.aspenlock, self.excellock, self.aspen_file, 
                                                                self.excel_solver_file, task_queue, self.abort, 
                                                                self.results_lock, self.results, self.directory, self.output_columns, self.output_value_cells,
                                                                self.trial_counter, self.save_freq, 
                                                                self.output_file, self.vars_to_change, 
                                                                self.output_columns, self.simulation_vars, self.sims_completed, 
                                                                self.lock_to_signal_finish, self.tot_sim, self.dispatch, self.weights)))
        for p in self.processes:
            p.start()
        for i in range(self.num_processes):
            task_queue.put('STOP')
        
            
    def close_all_COMS(self):
        self.aspenlock.acquire()
        self.excellock.acquire()
        sleep(3)
        for p in process_iter():
            if p.pid in self.current_COMS_pids:
                p.terminate()
                del self.current_COMS_pids[p.pid]
        self.aspenlock.release()
        self.excellock.release()
                
                
    def find_pids_to_ignore(self):
        for p in process_iter():
            if 'aspen' in p.name().lower() or 'excel' in p.name().lower():
                self.pids_to_ignore[p.pid] = 1
        
        
        
############ GLOBAL FUNCTIONS ################
                
def open_aspenCOMS(aspenfilename,dispatch):
    aspencom = Dispatch(str(dispatch))
    aspencom.InitFromArchive2(path.abspath(aspenfilename), host_type=0, node='', username='', password='', working_directory='', failmode=0)
    obj = aspencom.Tree     
    return aspencom,obj


def open_excelCOMS(excelfilename):
    pythoncom.CoInitialize()
    excel = DispatchEx('Excel.Application')
    book = excel.Workbooks.Open(path.abspath(excelfilename))
    return excel,book  
   
    
def save_data(outputfilename, results, directory, weights):
    if results: 
        collected_data = concat(results).sort_index()
        if len(weights) > 0:
            collected_data['Probability'] = weights[:len(collected_data)]
        writer = ExcelWriter(directory.value + '/' + outputfilename.value + '.xlsx')
        collected_data.to_excel(writer, sheet_name ='Sheet1')
        stats = collected_data.describe()
        stats.to_excel(writer, sheet_name = 'Summary Stats')
        writer.save()
    
    
def save_graphs(outputfilename, results, directory, weights):
    if results:
        collected_data = concat(results).sort_index()
        for index, var in enumerate(collected_data.columns[:-1]):
            fig = plt.figure()
            fig.set_size_inches(6,6)
            ax = fig.add_axes([0.12, 0.12, 0.85, 0.85])
            if len(weights) > 0:
                 plotweight = weights[:len(collected_data)]
                 num_bins = len(collected_data)
                 plt.hist(collected_data[var], num_bins, weights=plotweight, facecolor='blue', edgecolor='black', alpha=1.0)
            else:
                num_bins = 20
                plt.hist(collected_data[var], num_bins, facecolor='blue', edgecolor='black', alpha=1.0)
            ax.set_xlabel(var, Fontsize=14)
            ax.set_ylabel('Count', Fontsize=14)
            ax.ticklabel_format(axis= 'x', style = 'sci', scilimits= (-3,3), Fontsize=12)
            ax.ticklabel_format(axis= 'y', Fontsize=12)
            char_omit = findall(r"([^\\\/\:\?\*\"\<\>\|]*)[\\\/\:\?\*\"\<\>\|]([^\\\/\:\?\*\"\<\>\|]*)",var)
            if char_omit:
                var = "".join([s[0]+s[1] for s in char_omit])
            plt.savefig(directory.value + '/' + outputfilename.value + '_' + var + '.png', format='png')
    

def worker(current_COMS_pids, pids_to_ignore, aspenlock, excellock, aspenfilename, 
           excelfilename, task_queue, abort, results_lock, results, directory, output_columns, output_value_cells,
           sim_counter, save_freq, outputfilename, vars_to_change, columns, simulation_vars, sims_completed, lock_to_signal_finish, tot_sim, dispatch, weights):
    
    local_pids_to_ignore = {}
    local_pids = {}
    aspenlock.acquire()
    for p in process_iter():
        if 'aspen' in p.name().lower():
            local_pids_to_ignore[p.pid] = 1
    if not abort.value:
        aspencom,obj = open_aspenCOMS(aspenfilename.value, dispatch)
    for p in process_iter():
        if 'aspen' in p.name().lower() and p.pid not in local_pids_to_ignore:
            local_pids[p.pid] = 1
    aspenlock.release()
    excellock.acquire()
    if not abort.value:
        excel,book = open_excelCOMS(excelfilename.value)
    excellock.release() 
    
    for p in process_iter(): #register the pids of COMS objects
        if ('aspen' in p.name().lower() or 'excel' in p.name().lower()) and p.pid not in pids_to_ignore:
            current_COMS_pids[p.pid] = 1
            
            
    for trial_num in iter(task_queue.get, 'STOP'):
        if abort.value:
            try:
                lock_to_signal_finish.release()
            except:
                continue
        
        aspencom, case_values, errors, obj = aspen_run(aspencom, obj, simulation_vars, trial_num, vars_to_change) 
        result = mp_excelrun(excel, book, aspencom, obj, case_values, columns, errors, trial_num, output_value_cells)
        
        results_lock.acquire()
        results.append(result) 
        sim_counter.value = len(results)
        if sim_counter.value % save_freq.value == 0:
            save_data(outputfilename, results, directory, weights)
            save_graphs(outputfilename, results, directory, weights)
        sims_completed.value += 1
        results_lock.release()
        

        if virtual_memory().percent > 94:
            aspenlock.acquire()
            for p in process_iter():
                if p.pid in local_pids:
                    p.terminate()
                    del current_COMS_pids[p.pid]
                    del local_pids[p.pid]
                    
            
            for p in process_iter():
                if 'aspen' in p.name().lower():
                    local_pids_to_ignore[p.pid] = 1
            if not abort.value:
                aspencom,obj = open_aspenCOMS(aspenfilename.value,dispatch)
            
            for p in process_iter(): #register the pids of COMS objects
                if 'aspen' in p.name().lower() and p.pid not in local_pids_to_ignore:
                    current_COMS_pids[p.pid] = 1
                    local_pids[p.pid] = 1
            aspenlock.release()
        
        aspencom.Engine.Host(host_type=0, node='', username='', password='', working_directory='')    

    try:
        lock_to_signal_finish.release()
    except:
        pass
      
        
def aspen_run(aspencom, obj, simulation_vars, trial, vars_to_change):
    
    #SUC_LOC = r"\Data\Blocks\A300\Data\Blocks\B1\Input\FRAC\TOC5"
    #suobj.FindNode(SUC_LOC).Value = 0.4
    
    variable_values = {}
    for (aspen_variable, aspen_call, fortran_index), dist in simulation_vars.items():
        obj.FindNode(aspen_call).Value = dist[trial]
        if type(dist[trial]) == str:
            variable_values[aspen_variable] = float(dist[trial][fortran_index[0]:fortran_index[1]])
        else:
            variable_values[aspen_variable] = dist[trial]
    
    ########## STORE THE RANDOMLY SAMPLED VARIABLE VALUES  ##########
    case_values = []
    for v in vars_to_change:
        case_values.append(variable_values[v])    
    
    aspencom.Reinit()
    aspencom.Engine.Run2()
    errors = FindErrors(aspencom)
    
    return aspencom, case_values, errors, obj


def mp_excelrun(excel, book, aspencom, obj, case_values, columns, errors, trial_num, output_value_cells):

#    column = [x for x in book.Sheets('Aspen_Streams').Evaluate("D1:D100") if x.Value != None] 
#    
#    if obj.FindNode(column[0]) == None: # basically, if the massflow out of the system is None, then it failed to converge
#        dfstreams = DataFrame(columns=columns)
#        dfstreams.loc[trial_num+1] = case_values + [None]*(len(columns) - 1 - len(case_values)) + ["Aspen Failed to Converge"]
#        return dfstreams
#    stream_values = []
#    for index,stream in enumerate(column):
#        stream_value = obj.FindNode(stream).Value   
#        stream_values.append((stream_value,))
#    cell_string = "C1:C" + str(len(column))
#    book.Sheets('ASPEN_Streams').Evaluate(cell_string).Value = stream_values
#    
#    excel.Calculate()
#    excel.Run('SOLVE_DCFROR')
#
#     
#    dfstreams = DataFrame(columns=columns)
#    dfstreams.loc[trial_num+1] = case_values + [x.Value for x in book.Sheets('Output').Evaluate(output_value_cells.value)] + ["; ".join(errors)]
#    return dfstreams
    
    
    excel.Run('sub_ClearSumData_ASPEN')
    excel.Run('sub_GetSumData_ASPEN')
    excel.Calculate()
#    excel.Run('SolveProductCost')
    excel.Run('solvedcfror')
      
    dfstreams = DataFrame(columns=columns)
    dfstreams.loc[trial_num+1] = case_values + [x.Value for x in book.Sheets('Output').Evaluate(output_value_cells.value)] + ["; ".join(errors)]
    return dfstreams


def FindErrors(aspencom):
    obj = aspencom.Tree
    error = r'\Data\Results Summary\Run-Status\Output\PER_ERROR'
    not_done = True
    counter = 1
    error_number = 0
    error_statements = []
    while not_done:
        try:
            check_for_errors = obj.FindNode(error + '\\' +  str(counter)).Value
            if "error" in check_for_errors.lower():
                error_statements.append(check_for_errors)
                scan_errors = True
                counter += 1
                while scan_errors:
                    if len(obj.FindNode(error + '\\' + str(counter)).Value.lower()) > 0:
                        error_statements[error_number] = error_statements[error_number] + obj.FindNode(error + '\\' + str(counter)).Value
                        counter += 1
                    else:
                        scan_errors = False
                        error_number += 1
                        counter += 1
            else:
                counter += 1
        except:
            not_done = False
    return error_statements

def OnFocusIn(event):
    if type(event.widget).__name__ == 'MainApp':
        event.widget.attributes('-topmost', False)

        
if __name__ == "__main__":
    freeze_support()
    main_app = MainApp()
    main_app.mainloop()
    if main_app.current_simulation:
        main_app.abort_univar_overall.value = True
        main_app.abort_sim()
        print('Waiting for Clearance to Exit...')
        main_app.current_simulation.wait()
        print('Waiting for Worker Thread to Terminate...')
        main_app.worker_thread.join()
        print('Waiting for Cleanup Thread to Terminate...')
        main_app.cleanup_thread.join()
    exit()
        
        
    

    

