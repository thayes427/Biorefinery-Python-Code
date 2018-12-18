# -*- coding: utf-8 -*-
"""
Created on Sat Dec 15 19:58:47 2018

@author: MENGstudents
"""
import tkinter as tk
from tkinter import *
from tkinter import messagebox
from tkinter import ttk
from tkinter.filedialog import askopenfilename
import threading
import pandas as pd
import multiprocessing as mp
import time
import numpy as np
import psutil
import random
import win32com.client as win32
import pythoncom
import os
from math import ceil
import csv
from multiprocessing import freeze_support
import matplotlib
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2TkAgg
from matplotlib.figure import Figure
 

class MainApp(tk.Tk):

    def __init__(self):
        ####### Do something ######
        tk.Tk.__init__(self)
        self.notebook = ttk.Notebook(self)
        self.wm_title("Sensitivity Analysis Tool")
        self.notebook.grid()
        self.construct_home_tab()
        
        self.simulations = []
        self.current_simulation = None
        self.current_tab = None
        self.abort = mp.Value('b', False)
        self.abort_univar_overall = mp.Value('b', False)
        self.simulation_vars = {}
        self.attributes("-topmost", True)
        self.tot_sim_num = 0
        self.sims_completed = mp.Value('i',0)
        self.start_time = None
        self.univar_plot_counter = 1
        self.univar_old_name = ''
        self.univar_row_num = 0


    def construct_home_tab(self):
        self.home_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.home_tab, text = 'File Upload')
        Button(self.home_tab, 
        text='Upload Excel Data',
        command=self.open_excel_file).grid(row=0,column=1, sticky = E, pady = 5,padx = 5)
        self.input_csv_entry = Entry(self.home_tab)
        self.input_csv_entry.grid(row=0, column=2)
        
        Button(self.home_tab, 
              text="Upload Aspen Model",
              command=self.open_aspen_file).grid(row=1, column = 1,sticky = E,
              pady = 5,padx = 5)
        self.aspen_file_entry = Entry(self.home_tab)
        self.aspen_file_entry.grid(row=1, column=2,pady = 5,padx = 5)
        
        Button(self.home_tab, 
              text="Upload Excel Model",
              command=self.open_solver_file).grid(row=2,column = 1,sticky = E,
              pady = 5,padx = 5)
        self.excel_solver_entry = Entry(self.home_tab)
        self.excel_solver_entry.grid(row=2, column=2,pady = 5,padx = 5)
        
        Button(self.home_tab, 
              text="Load Data",
              command=self.make_new_tab).grid(row=5,column = 3,sticky = E,
              pady = 5,padx = 5)
        
        self.analysis_type = StringVar(self.home_tab)
        self.analysis_type.set("Choose Analysis Type")
        
        analysis_type_options = OptionMenu(
                self.home_tab, self.analysis_type,"Single Point Analysis","Univariate Sensitivity", 
                "Multivariate Sensitivity").grid(row = 5,sticky = E,column = 2,padx =5, pady = 5)
        
        Label(self.home_tab, text='Number of Processes :').grid(row=3, column=1, sticky=E)
        self.num_processes_entry = Entry(self.home_tab)
        self.num_processes_entry.grid(row=3, column=2, pady=5, padx=5)

    def make_new_tab(self):
        if self.current_tab:
            self.notebook.forget(self.current_tab)
            self.current_tab = None
        if self.analysis_type.get() == 'Choose Analysis Type':
            print("ERROR: Select an Analysis Type")
        elif self.analysis_type.get() == 'Univariate Sensitivity':
            self.current_tab = ttk.Frame(self.notebook)
            self.notebook.add(self.current_tab,text = "Univariate Analysis")
            ##############Tab 2 LABELS##################
            
            Label(self.current_tab, 
                  text="Save As :").grid(row=4, column= 1, sticky = E,pady = 5,padx = 5)
            self.save_as_entry= Entry(self.current_tab)
            self.save_as_entry.grid(row=4, column=2,pady = 5,padx = 5)
            
            Label(self.current_tab,text = ".csv").grid(row = 4, column = 3, sticky = W)
            
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
            
            self.options_box = ttk.Labelframe(self.current_tab, text='Run Options:')
            self.options_box.grid(row = 15,column = 3, pady = 10,padx = 10)
    
            Button(self.options_box, text = "Next Variable", command=self.abort_sim).grid(
                    row=6,columnspan = 1, column = 2, sticky=W)
            
            Button(self.options_box, text = "Abort", command=self.abort_univar_overall_fun).grid(
                    row= 6,columnspan = 1, column = 3, sticky=W)
        elif  self.analysis_type.get() == 'Single Point Analysis':
            self.current_tab = ttk.Frame(self.notebook)
            self.notebook.add(self.current_tab, text = 'Single Point')
             
            Label(self.current_tab, 
                  text="Save As :").grid(row=0, column= 0, sticky = E,pady = 5,padx = 5)
            self.save_as_entry = Entry(self.current_tab)
            self.save_as_entry.grid(row=0, column=1,pady = 5,padx = 5)
            
            Button(self.current_tab, text='Calculate MFSP',
            command=self.initialize_single_point).grid(row=7,
            column=1, columnspan=2, pady=4)
            
        elif  self.analysis_type.get() == 'Multivariate Sensitivity':
            self.current_tab = ttk.Frame(self.notebook)
            self.notebook.add(self.current_tab,text = "Multivariate Analysis")
    
            Label(self.current_tab, 
                  text="Number of Simulations :").grid(row=3, column= 1, sticky = E,pady = 5,padx = 5)
            self.num_sim_entry = Entry(self.current_tab)
            self.num_sim_entry.grid(row=3, column=2,pady = 5,padx = 5)
            
            Label(self.current_tab, 
                  text="Save As :").grid(row=4, column= 1, sticky = E,pady = 5,padx = 5)
            self.save_as_entry = Entry(self.current_tab)
            self.save_as_entry.grid(row=4, column=2,pady = 5,padx = 5)
            
            Label(self.current_tab,text = ".csv").grid(row = 4, column = 3, sticky = W)
                   
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
        
    def load_variables_into_GUI(self):
        sens_vars = str(self.input_csv_entry.get())
        single_pt_vars = []
        univariate_vars = []
        multivariate_vars = []
        type_of_analysis = self.analysis_type.get()
        with open(sens_vars) as f:
            reader = csv.DictReader(f)# Skip the header row
            for row in reader:
                if row['Toggle'].lower().strip() == 'true':          
                    if type_of_analysis =='Single Point Analysis':
                        single_pt_vars.append((row["Variable Name"], float(row["Range of Values"].split(',')[0].strip())))
                    elif type_of_analysis == 'Multivariate Analysis':
                        multivariate_vars.append(row["Variable Name"])
                    else:
                        univariate_vars.append((
                                row["Variable Name"], row["Format of Range"].strip().lower(
                                        ), row['Range of Values'].split(',')))
                        
        #now populate the gui with the appropriate tab and variables stored above
        if type_of_analysis == 'Single Point Analysis':
            self.current_tab.config(width = '5c', height = '5c')
            self.sp_value_entries = {}
            
            # Create a frame for the canvas with non-zero row&column weights
            frame_canvas = ttk.Frame(self.current_tab)
            frame_canvas.grid(row=2, column=1, pady=(5, 0))
            frame_canvas.grid_rowconfigure(0, weight=1)
            frame_canvas.grid_columnconfigure(0, weight=1)
            frame_canvas.config(height = '5c')
            
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
            
            self.sp_row_num = 0
            for name,value in single_pt_vars:
                self.sp_row_num += 1
                key = str(self.sp_row_num)
                Label(frame_vars, 
                text= name).grid(row=self.sp_row_num, column= 1, sticky = E,pady = 5,padx = 5)
                key=Entry(frame_vars)
                key.grid(row=self.sp_row_num, column=2,pady = 5,padx = 5)
                key.delete(first=0,last=END)
                key.insert(0,str(value))
                self.sp_value_entries[name]= key
                
            # Determine the size of the Canvas
            frame_canvas.config(width='5c', height='10c')
            # Set the canvas scrolling region
            canvas.config(scrollregion=canvas.bbox("all"))
    
        if type_of_analysis == 'Univariate Sensitivity':
            self.univar_ntrials_entries = {}
            Label(self.current_tab, 
                text= 'Variable Name').grid(row=8, column= 1,pady = 5,padx = 5, sticky= E)
            Label(self.current_tab, 
                text= 'Sampling Type').grid(row=8, column= 2,pady = 5,padx = 5)
            Label(self.current_tab, 
                text= '# of Trials').grid(row=8, column= 3,pady = 5,padx = 5, sticky = W)
            # Create a frame for the canvas with non-zero row&column weights
            frame_canvas1 = ttk.Frame(self.current_tab)
            frame_canvas1.grid(row=9, column=1, columnspan =3, pady=(5, 0))
            frame_canvas1.grid_rowconfigure(0, weight=1)
            frame_canvas1.grid_columnconfigure(0, weight=1)
            frame_canvas1.config(height = '3c')
            
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
            for name, format_of_data, vals in univariate_vars:
                Label(frame_vars1, 
                text= name).grid(row=self.univar_row_num, column= 1,pady = 5,padx = 5)
                Label(frame_vars1, 
                text= format_of_data).grid(row=self.univar_row_num, column= 2,pady = 5,padx = 5)
                
                if not(format_of_data == 'linspace' or format_of_data == 'list'):
                    key2=Entry(frame_vars1)
                    key2.grid(row=self.univar_row_num, column=3,pady = 5,padx = 5)
                    #key2.insert(0,univariate_sims)
                    self.univar_ntrials_entries[name]= key2
                else:
                    if format_of_data == 'linspace':
                        
                        Label(frame_vars1,text= str(vals[2])).grid(row=univar_row_num, column= 3,pady = 5,padx = 5)
                    else:
                        Label(frame_vars1,text= str(len(vals))).grid(row=self.univar_row_num, column= 3,pady = 5,padx = 5)
                self.univar_row_num += 1
                
            # Update vars frames idle tasks to let tkinter calculate variable sizes
            frame_vars1.update_idletasks()
            # Determine the size of the Canvas
            
            frame_canvas1.config(width='5c', height='5c')
            
            # Set the canvas scrolling region
            canvas1.config(scrollregion=canvas1.bbox("all"))
            
    def get_distributions(self):
        if self.analysis_type.get() == 'Univariate Sensitivity':
            if self.univar_ntrials_entries:
                max_num_sim = max(int(slot.get()) for slot in self.univar_ntrials_entries.values())
            else:
                max_num_sim = 1
            self.simulation_vars, self.simulation_dist = self.construct_distributions(ntrials=max_num_sim)
            for (aspen_variable, aspen_call, fortran_index), dist in self.simulation_vars.items():
                if aspen_variable in self.univar_ntrials_entries:
                    num_trials_per_var = int(self.univar_ntrials_entries[aspen_variable].get())
                    self.simulation_vars[(aspen_variable, aspen_call, fortran_index)] = dist[:num_trials_per_var]
                    self.simulation_dist[aspen_variable] = dist[:num_trials_per_var]                
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
        with open(gui_excel_input) as f:
            reader = csv.DictReader(f)# Skip the header row
            simulation_vars = {}
            simulation_dist = {}
            for row in reader:
                if row['Toggle'].lower().strip() == 'true':
                    dist_type = row['Format of Range'].lower()
                    aspen_variable = row['Variable Name']
                    aspen_call = row['Variable Aspen Call']
                    bounds = row['Bounds'].split(',')
                    lb = float(bounds[0].strip())
                    ub = float(bounds[1].strip())
                    if 'normal' in dist_type or 'gaussian' in dist_type:
                        dist_variables = row['Range of Values'].split(',')
                        distribution = self.sample_gauss(float(dist_variables[0].strip()),
                                  float(dist_variables[1].strip()), lb, ub, ntrials)
                    if 'linspace' in dist_type:
                        linspace_vars = row['Range of Values'].split(',')
                        distribution = np.linspace(float(linspace_vars[0].strip()), 
                                                   float(linspace_vars[1].strip()),
                                                   float(linspace_vars[2].strip()))
                    if 'poisson' in dist_type:
                        lambda_p = float(row['Range of Values'].strip())
                        distribution = self.sample_poisson(lambda_p, lb, ub, ntrials)
                    if 'pareto' in dist_type:
                        pareto_vals = row['Range of Values'].split(',')
                        shape = float(pareto_vals[0].strip())
                        scale = float(pareto_vals[1].strip())
                        distribution = self.sample_pareto(shape, scale, lb, ub, ntrials)
                    if 'list' in dist_type:
                        lst = row['Range of Values'].split(',')
                        distribution = []
                        for l in lst:
                            distribution.append(float(l.strip()))                
                    if 'uniform' in dist_type:
                        lb_ub = row['Range of Values'].split(',')
                        lb_uniform, ub_uniform = float(lb_ub[0].strip()), float(lb_ub[1].strip())
                        distribution = self.sample_uniform(lb_uniform, ub_uniform, lb, ub, ntrials)
                    simulation_dist[aspen_variable] = distribution
                    fortran_index = (0,0)
                    if row['Fortran Call'].strip() != "":
                        fortran_call = row['Fortran Call']
                        value_to_change = row['Fortran Value to Change'].strip()
                        len_val = len(value_to_change)
                        for i in range(len(fortran_call)):
                            if fortran_call[i:i+len_val] == value_to_change:
                                fortran_index = (i, i+len_val) #NOT INCLUSIVE
                        for i, v in enumerate(distribution):
                            distribution[i] = self.make_fortran(fortran_call, fortran_index, v)
                    simulation_vars[(aspen_variable, aspen_call, fortran_index)] = distribution
        return simulation_vars, simulation_dist
    
    def sample_gauss(self,mean, std, lb, ub, ntrials):
        d = []
        for i in range(ntrials):
            rand_sample = np.random.normal(mean,std)
            while(rand_sample < lb or rand_sample > ub):
                rand_sample = np.random.normal(mean,std)
            d.append(rand_sample)
        return d
    
    def sample_uniform(self,lb_uniform, ub_uniform, lb, ub, ntrials):
        d = []
        for i in range(ntrials):
            rand_sample = np.random.uniform(lb_uniform, ub_uniform)
            while(rand_sample < lb or rand_sample > ub):
                rand_sample = np.random.uniform(lb_uniform, ub_uniform)
            d.append(rand_sample)
        return d
    
    
    def sample_poisson(self,lambda_p, lb, ub, ntrials):
        d = []
        for i in range(ntrials):
            rand_sample = np.random.poisson(10000*lambda_p)/10000
            while(rand_sample < lb or rand_sample > ub):
                rand_sample = np.random.poisson(10000*lambda_p)/10000
            d.append(rand_sample)
        return d
    
    def sample_pareto(self, shape, scale, lb, ub, ntrials):
        d = []
        for i in range(ntrials):
            rand_sample = (np.random.pareto(shape) + 1) * scale
            while(rand_sample < lb or rand_sample > ub):
                rand_sample = (np.random.pareto(shape) + 1) * scale
            d.append(rand_sample)
        return d
    
    def make_fortran(self, fortran_call, fortran_index, val):
        return fortran_call[:fortran_index[0]] + str(val) + fortran_call[fortran_index[1]:]
    
    def disp_sp_mfsp(self):
        print('checking for result')
        try:
            if self.current_simulation.results:
                mfsp = self.current_simulation.results[0].at[0, 'MFSP']
                Label(self.current_tab, text= 'MFSP = ${:.2f}'.format(mfsp)).grid(
                        row=self.sp_row_num+1, column = 1)
            else:
                self.after(5000, self.disp_sp_mfsp)
        except:
            self.after(5000, self.disp_sp_mfsp)
    
    def single_point_analysis(self):
        self.store_user_inputs()
        self.get_distributions()
        # update simulation variable values based on user input in GUI
        for (aspen_variable, aspen_call, fortran_index), values in self.simulation_vars.items():
            self.simulation_vars[(aspen_variable, aspen_call, fortran_index)] = [float(
                    self.sp_value_entries[aspen_variable].get())]
        self.create_simulation_object(self.simulation_vars, self.vars_to_change, self.output_file, self.num_trial)
        self.run_simulations()
        
    
    def run_multivar_sens(self):
        Button(self.current_tab, text = "Abort", command=self.abort_sim).grid(
                    row=7,columnspan = 1, column = 3, sticky=W)
        self.store_user_inputs()
        if len(self.simulation_vars) == 0:
            self.get_distributions()
        self.create_simulation_object(self.simulation_vars, self.vars_to_change, self.output_file, self.num_trial)
        self.run_simulations()
        
        
    def run_univ_sens(self):
        self.store_user_inputs()
        if len(self.simulation_vars) == 0:
            self.get_distributions()
        for (aspen_variable, aspen_call, fortran_index), values in self.simulation_vars.items():
            self.create_simulation_object({(aspen_variable, aspen_call, 
                                            fortran_index): values}, [aspen_variable], 
    self.output_file+'_'+aspen_variable, len(values))
        self.run_simulations()
    
    
    def store_user_inputs(self):
        self.aspen_file = str(self.aspen_file_entry.get())
        self.num_processes = int(self.num_processes_entry.get())
        self.excel_solver_file= str(self.excel_solver_entry.get())
        try:
            self.num_trial = int(self.num_sim_entry.get())
        except: 
            self.num_trial = 1
        self.output_file = str(self.save_as_entry.get())
        self.input_csv = str(self.input_csv_entry.get())
        
        self.vars_to_change = []
        with open(self.input_csv) as f:
            reader = csv.DictReader(f)# Skip the header row
            for row in reader:
                if row['Toggle'].lower().strip() == 'true':
                    self.vars_to_change.append(row["Variable Name"])
        
        
    def run_simulations(self):
        self.start_time = time.time()
        
        for sim in self.simulations: 
            self.current_simulation = sim
            self.current_simulation.init_sims()
            if self.abort_univar_overall.value:
                self.abort.value = True
            
        
    def create_simulation_object(self, simulation_vars, vars_to_change, output_file, num_trial):
        self.output_columns = vars_to_change + ['Biofuel Output', 'Succinic Acid Output', 'Fixed Op Costs',\
              'Var OpCosts', ' Capital Costs', 'MFSP','Fixed Capital Investment',\
              'Capital Investment with Interest','Loan Payment per Year','Depreciation','Cash on Hand',\
              'Steam Plant Value','Bag Cost', 'Aspen Errors']
        
        new_sim = Simulation(self.sims_completed, num_trial, simulation_vars, output_file, 
                             self.aspen_file, self.excel_solver_file, self.abort, vars_to_change, 
                             self.output_columns, save_freq=5, num_processes=self.num_processes, reinit_coms_freq=25)
        self.simulations.append(new_sim)
        self.tot_sim_num += num_trial
        
        
    def initialize_single_point(self):
        self.worker_thread = threading.Thread(
                target=lambda: self.single_point_analysis())
        self.worker_thread.start()
        self.after(5000, self.disp_sp_mfsp)
        
    def initialize_univar_analysis(self):
        self.worker_thread = threading.Thread(
            target=lambda: self.run_univ_sens())
        self.worker_thread.start()
        self.status_label = None
        self.time_rem_label = None
        self.after(5000, self.univar_gui_update)

    
    def initialize_multivar_analysis(self):
        self.worker_thread = threading.Thread(
            target=lambda: self.run_multivar_sens())
        self.worker_thread.start()
        self.status_label = None
        self.time_rem_label = None
        self.multivar_gui_update()
        
        
    def disp_status_update(self):
        if self.current_simulation and not self.abort.value:
            if len(self.current_simulation.results) == self.current_simulation.tot_sim:
                tmp = Label(self.current_tab, text= 'Status: Simulation Complete                   ')
            else:
                tmp = Label(self.current_tab, text= 'Status: Simulation Running | {} Results Collected'.format(
                        len(self.current_simulation.results)))
            tmp.grid(row=15, column = 1, sticky=W, columnspan=2)
            if self.status_label:
                self.status_label.destroy()
            self.status_label = tmp
        
        
    def disp_time_remaining(self):
        if self.start_time:
            elapsed_time = time.time() - self.start_time
            if self.sims_completed.value > 0:
                remaining_time = ((elapsed_time / self.sims_completed.value) * (self.tot_sim_num - self.sims_completed.value))//60
                hours, minutes = divmod(remaining_time, 60)
                tmp = Label(self.current_tab, text='Time Remaining: {} Hours, {} Minutes    '.format(int(hours), int(minutes)))
            else:
                tmp = Label(self.current_tab, text='Time Remaining: N/A')
            tmp.grid(row=16, column=1, columnspan=2,sticky=W)
            if self.time_rem_label:
                self.time_rem_label.destroy()
            self.time_rem_label = tmp
            
            
    def plot_on_GUI(self):
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
        if self.current_simulation and self.current_simulation.results:
            d_f_output = pd.concat(self.current_simulation.results).sort_index()
            vars_to_change = self.current_simulation.vars_to_change
            columns = 2
            num_rows= ((len(vars_to_change) + 1) // columns) + 1
            counter = 1
            fig = Figure(figsize = (5,5))
            a = fig.add_subplot(num_rows,columns,counter)
            counter += 1
            total_MFSP = d_f_output["MFSP"]
            num_bins = 15
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
                    num_bins = 15
                    try:
                        n, bins, patches = a.hist(total_data, num_bins, facecolor='blue', alpha=0.5)
                    except Exception:
                        pass
                    a.set_title(var)
        
            #a = fig.tight_layout()
            canvas = FigureCanvasTkAgg(fig)
            canvas.get_tk_widget().grid(row=8, column = 0,columnspan = 10, rowspan = 10, sticky= W+E+N+S, pady = 5,padx = 5,)
                
    def plot_univ_on_GUI(self):
        '''
        Allows Unviariate to be plotted on the GUI
        
        It will plot the graphs as follows: the number of rows is the number of variables that have
        to be plotted, and there will be a width of two columns, with the first column graphing 
        the variable distribution, and the second one plotting the MFSP distribution.   
        
        '''
        
        if self.current_simulation and self.current_simulation.results:
            dfstreams = pd.concat(self.current_simulation.results).sort_index()
            var = self.current_simulation.vars_to_change[0]
            c = self.current_simulation.trial_counter.value + 1
            columns = 2
            num_rows= ((len(self.simulation_dist) + 1) // columns) + 1
            fig = Figure(figsize = (5,5))
            a = fig.add_subplot(num_rows,columns,self.univar_plot_counter)
            if var != self.univar_old_name and self.univar_old_name != '':
                self.univar_plot_counter += 2
                self.univar_old_name = var
            if var != self.univar_old_name:
                self.univar_old_name = var
            self.univar_plot_counter += 1
            num_bins = 15
            try:
                n, bins, patches = a.hist(self.simulation_dist[var][:c], num_bins, facecolor='blue', alpha=0.5)
            except Exception:
                pass
            a.set_title(var)
            a = fig.add_subplot(num_rows,columns,self.univar_plot_counter)
            self.univar_plot_counter -= 1
            num_bins = 15
            try:
                n, bins, patches = a.hist(dfstreams['MFSP'], num_bins, facecolor='blue', alpha=0.5)
            except Exception:
                pass
            a.set_title('MFSP - ' + var)
            #a = fig.tight_layout()
            canvas = FigureCanvasTkAgg(fig)
            canvas.get_tk_widget().grid(row=8, column = 0,columnspan = 10, rowspan = 10, sticky= W+E+N+S, pady = 5,padx = 5,)
            
            #vsb = ttk.Scrollbar(fig, orient="vertical", command=canvas.yview)
            #vsb.grid(row=8, column=1,sticky = 'ns')
            #canvas.configure(yscrollcommand=vsb.set)
            
            #canvas.config(scrollregion=canvas.bbox("all"))
            
    def plot_init_dist(self):
        '''
        This function will plot the distribution of variable calls prior to running
        the simulation. This will enable users to see whether the distributions are as they expected.
        
        '''
        
        self.get_distributions()
        
        columns = 2
        num_rows= ((len(self.simulation_dist) + 1) // columns) + 1
        counter = 1
        
        fig_list =[]

        for var, values in self.simulation_dist.items():
            fig = Figure(figsize = (3,3), facecolor=[240/255,240/255,237/255])
            a = fig.add_subplot(num_rows,columns,counter)
            #counter += 1
            num_bins = 15
            try:
                n, bins, patches = a.hist(values, num_bins, facecolor='blue', alpha=0.5)
            except Exception:
                pass
            a.set_title(var)
            fig_list.append(fig)
        #a = fig.tight_layout()
        if self.univar_row_num != 0:
            row_num = 16
        else:
            row_num = 8
        frame_canvas = ttk.Frame(self.current_tab)
        frame_canvas.grid(row=row_num, column=columns, pady=(5, 0))
        frame_canvas.grid_rowconfigure(0, weight=1)
        frame_canvas.grid_columnconfigure(0, weight=1)
        frame_canvas.config(height = '5c')
        
        main_canvas = Canvas(frame_canvas)
        main_canvas.grid(row=0, column=0, sticky="news")
        main_canvas.config(height = '5c')
        
        vsb = ttk.Scrollbar(frame_canvas, orient="vertical", command=main_canvas.yview)
        vsb.grid(row=0, column=1,sticky = 'ns')
        main_canvas.configure(yscrollcommand=vsb.set)
        
        figure_frame = ttk.Frame(main_canvas)
        main_canvas.create_window((0, 0), window=figure_frame, anchor='nw')
        figure_frame.config(height = '5c')
    
        row_num = 0
        
        for figs in fig_list:
            figure_canvas = FigureCanvasTkAgg(figs, master=figure_frame)
            #figure_canvas.draw()
            figure_canvas.get_tk_widget().grid(row=row_num, column = 0,columnspan = 10, rowspan = 10, sticky= W+E+N+S, pady = 5,padx = 5)
            #figure_canvas._tkcanvas.grid(row=row_num, column = 0,columnspan = 10, rowspan = 10, sticky= W+E+N+S, pady = 5,padx = 5)
            row_num += 5
        
        figure_canvas.update_idletasks()
        
        frame_canvas.config(width='5c', height='5c')
        
        # Set the canvas scrolling region
        main_canvas.config(scrollregion=main_canvas.bbox("all"))
        
    def univar_gui_update(self):
        self.disp_status_update()
        self.disp_time_remaining()
        self.plot_univ_on_GUI()
        ####################### ADD AUTO UPDATE GRAPHS #############
        self.after(10000, self.univar_gui_update)
        
        
    def multivar_gui_update(self):
        self.disp_status_update()
        self.disp_time_remaining()
        self.plot_on_GUI()
        ####################### ADD AUTO UPDATE GRAPHS #############
        self.after(10000, self.multivar_gui_update)
        
    
    def fill_num_trials(self):
        ntrials = self.fill_num_sims.get()
        for name, slot in self.univar_ntrials_entries.items():
            slot.delete(0, END)
            slot.insert(0, ntrials)
        

    def open_excel_file(self):
        filename = askopenfilename(title = "Select file", filetypes = (("csv files","*.csv"),("all files","*.*")))
        self.input_csv_entry.delete(0, END)
        self.input_csv_entry.insert(0, filename)
        
        
    def open_aspen_file(self):
        filename = askopenfilename(title = "Select file", filetypes = (("Aspen Models",["*.bkp", "*.apw"]),("all files","*.*")))
        self.aspen_file_entry.delete(0, END)
        self.aspen_file_entry.insert(0, filename)
    
    
    def open_solver_file(self):
        filename = askopenfilename(title = "Select file", filetypes = (("Excel Files","*.xlsm"),("all files","*.*")))
        self.excel_solver_entry.delete(0, END)
        self.excel_solver_entry.insert(0, filename)
        
        
    def abort_sim(self):
        self.abort.value = True
        self.cleanup_thread = threading.Thread(target=self.cleanup_processes_and_COMS)
        self.cleanup_thread.start()
        
    def abort_univar_overall_fun(self):
        self.abort_univar_overall.value = True
        self.abort_sim()
        
    def cleanup_processes_and_COMS(self):
        try:
            self.current_simulation.close_all_COMS()
            self.current_simulation.terminate_processes()
            save_data(self.current_simulation.output_file, self.current_simulation.results)
        except:
            self.after(1000, self.cleanup_processes_and_COMS)
            


        
        
       
        
        
        
        
################################################################################    
        
        
        
        
class Simulation(object):
    def __init__(self, sims_completed, tot_sim, simulation_vars, output_file, aspen_file, excel_solver_file,
                 abort, vars_to_change, output_columns, 
                 save_freq=10, num_processes=1, reinit_coms_freq=15):
        self.manager = mp.Manager()
        self.num_processes = min(num_processes, tot_sim, reinit_coms_freq)
        self.tot_sim = tot_sim
        self.sims_completed = sims_completed
        self.reinit_coms_freq = reinit_coms_freq
        self.save_freq = self.manager.Value('i', save_freq)
        self.abort = abort
        self.simulation_vars = self.manager.dict(simulation_vars) 
        self.output_file = self.manager.Value('s', output_file)
        self.aspen_file = self.manager.Value('s', aspen_file)
        self.excel_solver_file = self.manager.Value('s', excel_solver_file)
        
        self.results = self.manager.list()
        self.trial_counter = mp.Value('i',0)
        self.results_lock = mp.Lock()
        self.processes = []
        self.current_COMS_pids = self.manager.dict()
        self.pids_to_ignore = self.manager.dict()
        self.find_pids_to_ignore()
        self.output_columns = self.manager.list(output_columns)
        self.vars_to_change = self.manager.list(vars_to_change)
        self.aspenlock = mp.Lock()
        self.excellock = mp.Lock()
          
    
    def init_sims(self):
        for i in range(0, self.tot_sim, self.reinit_coms_freq):
            upper_bound = min(i + self.reinit_coms_freq, self.tot_sim)
            TASKS = [trial for trial in range(i, upper_bound)]
            if not self.abort.value:
                self.run_sim(TASKS)
            self.wait()
            self.terminate_processes()
            self.processes = []
            self.close_all_COMS()
            
        save_data(self.output_file, self.results)
        self.abort.value = False    
        
        
    def terminate_processes(self):
        for p in self.processes:
            p.terminate()
            p.join()
         
    def wait(self):
        if not any(p.is_alive() for p in self.processes):
            return
        else:
            time.sleep(5)
            self.wait()
            
            
    def run_sim(self, tasks):
        task_queue = mp.Queue()
        for task in tasks:
            task_queue.put(task)

        for i in range(self.num_processes):
            self.processes.append(mp.Process(target=worker, args=(self.current_COMS_pids, self.pids_to_ignore, 
                                                                self.aspenlock, self.excellock, self.aspen_file, 
                                                                self.excel_solver_file, task_queue, self.abort, 
                                                                self.results_lock, self.results,
                                                                self.trial_counter, self.save_freq, 
                                                                self.output_file, self.vars_to_change, 
                                                                self.output_columns, self.simulation_vars, self.sims_completed)))
        for p in self.processes:
            p.start()
        for i in range(self.num_processes):
            task_queue.put('STOP')
        
            
    def close_all_COMS(self):
        self.aspenlock.acquire()
        self.excellock.acquire()
        time.sleep(3)
        for p in psutil.process_iter():
            if p.pid in self.current_COMS_pids:
                p.terminate()
                del self.current_COMS_pids[p.pid]
        self.aspenlock.release()
        self.excellock.release()
                
                
    def find_pids_to_ignore(self):
        for p in psutil.process_iter():
            if 'aspen' in p.name().lower() or 'excel' in p.name().lower():
                self.pids_to_ignore[p.pid] = 1
        
        
        
############ GLOBAL FUNCTIONS ################
                
def open_aspenCOMS(aspenfilename):
    aspencom = win32.Dispatch('Apwn.Document')
    aspencom.InitFromArchive(os.path.abspath(aspenfilename))
    obj = aspencom.Tree     
    return aspencom,obj


def open_excelCOMS(excelfilename):
    excel = win32.DispatchEx('Excel.Application')
    book = excel.Workbooks.Open(os.path.abspath(excelfilename))
    return excel,book  
   
    
def save_data(outputfilename, results):
    if results:
        collected_data = pd.concat(results).sort_index()
        writer = pd.ExcelWriter(outputfilename.value + '.xlsx')
        collected_data.to_excel(writer, sheet_name ='Sheet1')
        stats = collected_data['MFSP'].describe()
        stats.to_excel(writer, sheet_name = 'Summary Stats')
        writer.save()
    

def worker(current_COMS_pids, pids_to_ignore, aspenlock, excellock, aspenfilename, 
           excelfilename, task_queue, abort, results_lock, results,
           sim_counter, save_freq, outputfilename, vars_to_change, columns, simulation_vars, sims_completed):
    
    aspenlock.acquire()
    if not abort.value:
        aspencom,obj = open_aspenCOMS(aspenfilename.value)
    aspenlock.release()
    excellock.acquire()
    if not abort.value:
        excel,book = open_excelCOMS(excelfilename.value)
    excellock.release() 
    
    for p in psutil.process_iter(): #register the pids of COMS objects
        if ('aspen' in p.name().lower() or 'excel' in p.name().lower()) and p.pid not in pids_to_ignore:
            current_COMS_pids[p.pid] = 1
            
    for trial_num in iter(task_queue.get, 'STOP'):
        if abort.value:
            continue
        
        aspencom, case_values, errors, obj = aspen_run(aspencom, obj, simulation_vars, trial_num, vars_to_change) 
        result = mp_excelrun(excel, book, aspencom, obj, case_values, columns, errors, trial_num)
        
        results_lock.acquire()
        results.append(result) 
        sim_counter.value = len(results)
        if sim_counter.value % save_freq.value == 0:
            save_data(outputfilename, results)
        sims_completed.value += 1
        results_lock.release()
        aspencom.Engine.ConnectionDialog()


def aspen_run(aspencom, obj, simulation_vars, trial, vars_to_change):
    
    SUC_LOC = r"\Data\Blocks\A300\Data\Blocks\B1\Input\FRAC\TOC5"
    obj.FindNode(SUC_LOC).Value = 0.4
    
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
    stop = CheckConverge(aspencom)
    errors = FindErrors(aspencom)
    
    return aspencom, case_values, errors, obj

def mp_excelrun(excel, book, aspencom, obj, case_values, columns, errors, trial_num):

    column = [x for x in book.Sheets('Aspen_Streams').Evaluate("D1:D100") if x.Value != None] 
    
    if obj.FindNode(column[0]) == None:
        print('ERROR in Aspen for fraction '+ str(case_values))
        return pd.DataFrame(columns=columns)
    stream_values = []
    for index,stream in enumerate(column):
        stream_value = obj.FindNode(stream).Value   
        stream_values.append((stream_value,))
    cell_string = "C1:C" + str(len(column))
    book.Sheets('ASPEN_Streams').Evaluate(cell_string).Value = stream_values
    
    excel.Calculate()
    excel.Run('SOLVE_DCFROR')

    
    dfstreams = pd.DataFrame(columns=columns)
    dfstreams.loc[trial_num] = case_values + [x.Value for x in book.Sheets('Output').Evaluate("C3:C15")] + [" ; ".join(errors)]
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



def CheckConverge(aspencom):
    obj = aspencom.Tree
    error = r'\Data\Blocks\REFINE\Data\Blocks\FRAC\Output\PER_ERROR\1'
    stage = r'\Data\Blocks\REFINE\Data\Blocks\FRAC\Input\NSTAGE'
    fracstm = r'\Data\Blocks\REFINE\Data\Blocks\FRAC\Input\FEED_STAGE\FRACSTM'
    fracfd = r'\Data\Blocks\REFINE\Data\Blocks\FRAC\Input\FEED_STAGE\FRACFD' 
    stm_stage = r'\Data\Blocks\REFINE\Data\Blocks\FRAC\Input\FEED_CONVEN\FRACSTM'
    #fd_stage = r'\Data\Blocks\REFINE\Data\Blocks\FRAC\Input\FEED_CONVEN\FRACFD'
    nstage = obj.FindNode(stage)
    
    while obj.FindNode(error) != None:
        
        nstage = obj.FindNode(stage)
        
        obj.FindNode(stm_stage).Value = "ABOVE-STAGE"
        nstage.Value -= 1
        obj.FindNode(fracstm).Value -= 1
        obj.FindNode(stm_stage).Value = "ON-STAGE"
        obj.FindNode(fracfd).Value = ceil(nstage.Value/2)
        
        print('Failed to Converge, Adjusting stages and Feed Stage #')
        print('Number of Stages: ', nstage.Value)
        print('Feed Stage: ', obj.FindNode(fracfd).Value)
        
        if nstage.Value < 2:
            return True
        
        aspencom.Reinit()
        aspencom.Engine.Run2()
        
    print("Converged with " + str(nstage.Value) + ' stages')
    print('Feed Stage: ', obj.FindNode(fracfd).Value)
    return False

        
if __name__ == "__main__":
    freeze_support()
    main_app = MainApp()
    main_app.mainloop()
    if main_app.current_simulation:
        main_app.abort_sim()
        print('now waiting for clearance')
        main_app.current_simulation.wait()
    exit()
        
        

