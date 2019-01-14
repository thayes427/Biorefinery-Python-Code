# -*- coding: utf-8 -*-
"""
Created on Sat Dec 15 19:58:47 2018

@author: MENGstudents
"""

from tkinter import Tk,Button,Label,Entry,StringVar,E,W,OptionMenu,Canvas,END
from tkinter.ttk import Frame, Labelframe, Scrollbar, Notebook
from tkinter.filedialog import askopenfilename
from threading import Thread
from pandas import ExcelWriter, DataFrame, concat
from multiprocessing import Value, Manager, Lock, Queue, Process
from time import time, sleep
from numpy import linspace, random
from psutil import process_iter, virtual_memory
from win32com.client import Dispatch, DispatchEx
import pythoncom ### I DONT THINK YOU NEED THIS
from os import path
from csv import DictReader
from multiprocessing import freeze_support
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
 

class MainApp(Tk):

    def __init__(self):
        Tk.__init__(self)
        self.notebook = Notebook(self)
        self.wm_title("Sensitivity Analysis Tool")
        self.notebook.grid()
        self.construct_home_tab()
        
        self.simulations = []
        self.current_simulation = None
        self.current_tab = None
        self.abort = Value('b', False)
        self.abort_univar_overall = Value('b', False)
        self.simulation_vars = {}
        self.attributes("-topmost", True)
        self.tot_sim_num = 0
        self.sims_completed = Value('i',0)
        self.start_time = None
        self.univar_plot_counter = 1
        self.finished_figures = []
        self.univar_row_num=0
        self.last_results_plotted = None
        self.last_update = None


    def construct_home_tab(self):
        self.home_tab = Frame(self.notebook)
        self.notebook.add(self.home_tab, text = 'Home')
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
        
        OptionMenu(self.home_tab, self.analysis_type,"Single Point Analysis","Univariate Sensitivity", 
                "Multivariate Sensitivity").grid(row = 5,sticky = E,column = 2,padx =5, pady = 5)
        
        Label(self.home_tab, text='CPU Core Count :').grid(row=3, column=1, sticky=E)
        self.num_processes_entry = Entry(self.home_tab)
        self.num_processes_entry.grid(row=3, column=2, pady=5, padx=5)

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
            
            self.options_box = Labelframe(self.current_tab, text='Run Options:')
            self.options_box.grid(row = 15,column = 3, pady = 10,padx = 10)
    
            Button(self.options_box, text = "Next Variable", command=self.abort_sim).grid(
                    row=6,columnspan = 1, column = 2, sticky=W)
            
            Button(self.options_box, text = "Abort", command=self.abort_univar_overall_fun).grid(
                    row= 6,columnspan = 1, column = 3, sticky=W)
        elif  self.analysis_type.get() == 'Single Point Analysis':
            self.current_tab = Frame(self.notebook)
            self.notebook.add(self.current_tab, text = 'Single Point')
             
            Label(self.current_tab, 
                  text="Save As :").grid(row=0, column= 0, sticky = E,pady = 5,padx = 5)
            self.save_as_entry = Entry(self.current_tab)
            self.save_as_entry.grid(row=0, column=1,pady = 5,padx = 5)
            
            Button(self.current_tab, text='Calculate MFSP',
            command=self.initialize_single_point).grid(row=7,
            column=1, columnspan=2, pady=4)
            
        elif  self.analysis_type.get() == 'Multivariate Sensitivity':
            self.current_tab = Frame(self.notebook)
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
            reader = DictReader(f)# Skip the header row
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
            frame_canvas = Frame(self.current_tab)
            frame_canvas.grid(row=2, column=1, pady=(5, 0))
            frame_canvas.grid_rowconfigure(0, weight=1)
            frame_canvas.grid_columnconfigure(0, weight=1)
            frame_canvas.config(height = '5c')
            
            # Add a canvas in the canvas frame
            canvas = Canvas(frame_canvas)
            canvas.grid(row=0, column=0, sticky="news")
            canvas.config(height = '5c')
            # Link a scrollbar to the canvas
            vsb = Scrollbar(frame_canvas, orient="vertical", command=canvas.yview)
            vsb.grid(row=0, column=1,sticky = 'ns')
            canvas.configure(yscrollcommand=vsb.set)
            
            # Create a frame to contain the variables
            frame_vars = Frame(canvas)
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
            frame_vars.update_idletasks()
            frame_canvas.config(width='5c', height='10c')
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
            frame_canvas1 = Frame(self.current_tab)
            frame_canvas1.grid(row=9, column=1, columnspan =3, pady=(5, 0))
            frame_canvas1.grid_rowconfigure(0, weight=1)
            frame_canvas1.grid_columnconfigure(0, weight=1)
            frame_canvas1.config(height = '3c')
            
            # Add a canvas in the canvas frame
            canvas1 = Canvas(frame_canvas1)
            canvas1.grid(row=0, column=0, sticky="news")
            canvas1.config(height = '3c')
            
            # Link a scrollbar to the canvas
            vsb = Scrollbar(frame_canvas1, orient="vertical", command=canvas1.yview)
            vsb.grid(row=0, column=1,sticky = 'ns')
            canvas1.configure(yscrollcommand=vsb.set)
            
            # Create a frame to contain the variables
            frame_vars1 = Frame(canvas1)
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
                        
                        Label(frame_vars1,text= str(vals[2])).grid(row=self.univar_row_num, column= 3,pady = 5,padx = 5)
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
            reader = DictReader(f)# Skip the header row
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
                        distribution = linspace(float(linspace_vars[0].strip()), 
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
                    simulation_dist[aspen_variable] = distribution[:]
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
            rand_sample = random.normal(mean,std)
            while(rand_sample < lb or rand_sample > ub):
                rand_sample = random.normal(mean,std)
            d.append(rand_sample)
        return d
    
    def sample_uniform(self,lb_uniform, ub_uniform, lb, ub, ntrials):
        d = []
        for i in range(ntrials):
            rand_sample = random.uniform(lb_uniform, ub_uniform)
            while(rand_sample < lb or rand_sample > ub):
                rand_sample = random.uniform(lb_uniform, ub_uniform)
            d.append(rand_sample)
        return d
    
    
    def sample_poisson(self,lambda_p, lb, ub, ntrials):
        d = []
        for i in range(ntrials):
            rand_sample = random.poisson(10000*lambda_p)/10000
            while(rand_sample < lb or rand_sample > ub):
                rand_sample = random.poisson(10000*lambda_p)/10000
            d.append(rand_sample)
        return d
    
    def sample_pareto(self, shape, scale, lb, ub, ntrials):
        d = []
        for i in range(ntrials):
            rand_sample = (random.pareto(shape) + 1) * scale
            while(rand_sample < lb or rand_sample > ub):
                rand_sample = (random.pareto(shape) + 1) * scale
            d.append(rand_sample)
        return d
    
    def make_fortran(self, fortran_call, fortran_index, val):
        return fortran_call[:fortran_index[0]] + str(val) + fortran_call[fortran_index[1]:]
    
    def disp_sp_mfsp(self):
        try:
            if self.current_simulation.results:
                mfsp = self.current_simulation.results[0].at[0, 'MFSP']
                if mfsp:
                    Label(self.current_tab, text= 'MFSP = ${:.2f}'.format(mfsp)).grid(
                        row=self.sp_row_num+1, column = 1)
                else:
                    Label(self.current_tab, text= 'Aspen Failed to Converge').grid(
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
        with open(self.input_csv) as f:
            reader = DictReader(f)# Skip the header row
            for row in reader:
                if row['Toggle'].lower().strip() == 'true':
                    self.vars_to_change.append(row["Variable Name"])
        
        
    def run_simulations(self):
        self.start_time = time()
        
        for sim in self.simulations: 
            self.current_simulation = sim
            self.current_simulation.init_sims()
            if self.abort_univar_overall.value:
                self.abort.value = True
            self.univar_plot_counter += 1
    
    def parse_output_vars(self):
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
            
        
    def create_simulation_object(self, simulation_vars, vars_to_change, output_file, num_trial):
        self.parse_output_vars()
        self.output_columns = vars_to_change + self.output_vars
        
        new_sim = Simulation(self.sims_completed, num_trial, simulation_vars, output_file, path.dirname(str(self.input_csv_entry.get())),
                             self.aspen_file, self.excel_solver_file, self.abort, vars_to_change, self.output_value_cells,
                             self.output_columns, save_freq=5, num_processes=self.num_processes)
        self.simulations.append(new_sim)
        self.tot_sim_num += num_trial
        
        
    def initialize_single_point(self):
        self.worker_thread = Thread(
                target=lambda: self.single_point_analysis())
        self.worker_thread.start()
        self.after(5000, self.disp_sp_mfsp)
        
    def initialize_univar_analysis(self):
        self.worker_thread = Thread(
            target=lambda: self.run_univ_sens())
        self.worker_thread.start()
        self.status_label = None
        self.time_rem_label = None
        self.after(5000, self.univar_gui_update)

    
    def initialize_multivar_analysis(self):
        self.worker_thread = Thread(
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
            if self.univar_row_num != 0:
                row = 15
            else:
                row = 8
            tmp.grid(row=row, column = 1, sticky=W, columnspan=2)
            if self.status_label:
                self.status_label.destroy()
            self.status_label = tmp
        
        
    def disp_time_remaining(self):
        if self.start_time and self.sims_completed.value != self.last_update:
            self.last_update = self.sims_completed.value
            elapsed_time = time() - self.start_time
            if self.sims_completed.value > 0:
                remaining_time = ((elapsed_time / self.sims_completed.value) * (self.tot_sim_num - self.sims_completed.value))//60
                hours, minutes = divmod(remaining_time, 60)
                tmp = Label(self.current_tab, text='Time Remaining: {} Hours, {} Minutes    '.format(int(hours), int(minutes)))
            else:
                tmp = Label(self.current_tab, text='Time Remaining: N/A')
            if self.univar_row_num != 0:
                row = 16
            else:
                row = 9
            tmp.grid(row=row, column=1, columnspan=2,sticky=W)
            if self.time_rem_label:
                self.time_rem_label.destroy()
            self.time_rem_label = tmp
            
            
    def plot_on_GUI(self):
        
        if not self.current_simulation:
            return
        if len(self.current_simulation.results) == self.last_results_plotted:
            return
        self.last_results_plotted = len(self.current_simulation.results)
        
        if self.current_simulation.results:
            results = concat(self.current_simulation.results).sort_index()
            results = results[[d is not None for d in results['MFSP']]] # filter to make sure you aren't plotting None results
        else:
            results = DataFrame(columns=self.output_columns)

        fig_list =[]
        num_bins = 15
        mfsp_fig = Figure(figsize = (3,3), facecolor=[240/255,240/255,237/255], tight_layout=True)
        b = mfsp_fig.add_subplot(111)
        b.hist(results['MFSP'], num_bins, facecolor='blue', edgecolor='black', alpha=1.0)
        b.set_title('MFSP')
        fig_list.append(mfsp_fig)
        
        for var, values in self.simulation_dist.items():
            fig = Figure(figsize = (3,3), facecolor=[240/255,240/255,237/255], tight_layout=True)
            a = fig.add_subplot(111)
            _, bins, _ = a.hist(self.simulation_dist[var], num_bins, facecolor='white', edgecolor='black',alpha=1.0)
            a.hist(results[var], bins=bins, facecolor='blue',edgecolor='black', alpha=1.0)
            a.set_title(var)
            fig_list.append(fig)
        
        if self.univar_row_num != 0:
            row_num = 17
        else:
            row_num = 10
        
        frame_canvas = Frame(self.current_tab)
        frame_canvas.grid(row=row_num, column=1, columnspan = 3,pady=(5, 0))
        frame_canvas.grid_rowconfigure(0, weight=1)
        frame_canvas.grid_columnconfigure(0, weight=1)
        frame_canvas.config(height = '10c', width='16c')
        
        main_canvas = Canvas(frame_canvas)
        main_canvas.grid(row=0, column=0, sticky="news")
        main_canvas.config(height = '10c', width='16c')
        
        vsb = Scrollbar(frame_canvas, orient="vertical", command=main_canvas.yview)
        vsb.grid(row=0, column=1,sticky = 'ns')
        main_canvas.configure(yscrollcommand=vsb.set)
        
        figure_frame = Frame(main_canvas)
        main_canvas.create_window((0, 0), window=figure_frame, anchor='nw')
        figure_frame.config(height = '10c', width='16c')
    
        row_num = 0
        column = False
        for figs in fig_list:
            figure_canvas = FigureCanvasTkAgg(figs, master=figure_frame)
            if column:
                col = 4
            else:
                col = 1
            #figure_canvas.draw()
            figure_canvas.get_tk_widget().grid(
                    row=row_num, column=col,columnspan=2, rowspan = 5, pady = 5,padx = 8, sticky=E)
            #figure_canvas._tkcanvas.grid(row=row_num, column = 0,columnspan = 10, rowspan = 10, sticky= W+E+N+S, pady = 5,padx = 5)
            if column:
                row_num += 5
            column = not column
        

        figure_frame.update_idletasks()
        frame_canvas.config(width='16c', height='10c')
        
        # Set the canvas scrolling region
        main_canvas.config(scrollregion=figure_frame.bbox("all"))
                
            
    def plot_univ_on_GUI(self):
        
        if not self.current_simulation:
            return
        if len(self.current_simulation.results) == self.last_results_plotted:
            return
        self.last_results_plotted = len(self.current_simulation.results)
        
        current_var = self.current_simulation.vars_to_change[0]
        if self.current_simulation.results:
            results = concat(self.current_simulation.results).sort_index()
            results = results[[d is not None for d in results['MFSP']]] # filter to make sure you aren't plotting None results
        else:
            results = DataFrame(columns=self.output_columns)

        fig_list =[]
        var_fig = Figure(figsize = (3,3), facecolor=[240/255,240/255,237/255], tight_layout=True)
        a = var_fig.add_subplot(111)
        num_bins = 15
        _, bins, _ = a.hist(self.simulation_dist[current_var], num_bins, facecolor='white',edgecolor='black', alpha=1.0)
        a.hist(results[current_var], bins=bins, facecolor='blue',edgecolor='black', alpha=1.0)
        a.set_title(current_var)
        fig_list.append(var_fig)
        
        mfsp_fig = Figure(figsize = (3,3), facecolor=[240/255,240/255,237/255], tight_layout=True)
        b = mfsp_fig.add_subplot(111)
        b.hist(results['MFSP'], num_bins, facecolor='blue', edgecolor='black', alpha=1.0)
        b.set_title('MFSP - ' + current_var)
        fig_list.append(mfsp_fig)
        
        figs_to_plot = self.finished_figures[:] + fig_list
        if len(self.current_simulation.results) == self.current_simulation.tot_sim:
            self.finished_figures += fig_list
        
        if self.univar_row_num != 0:
            row_num = 17
        else:
            row_num = 10
        
        frame_canvas = Frame(self.current_tab)
        frame_canvas.grid(row=row_num, column=1, columnspan = 3,pady=(5, 0))
        frame_canvas.grid_rowconfigure(0, weight=1)
        frame_canvas.grid_columnconfigure(0, weight=1)
        frame_canvas.config(height = '10c', width='16c')
        
        main_canvas = Canvas(frame_canvas)
        main_canvas.grid(row=0, column=0, sticky="news")
        main_canvas.config(height = '10c', width='16c')
        
        vsb = Scrollbar(frame_canvas, orient="vertical", command=main_canvas.yview)
        vsb.grid(row=0, column=1,sticky = 'ns')
        main_canvas.configure(yscrollcommand=vsb.set)
        
        figure_frame = Frame(main_canvas)
        main_canvas.create_window((0, 0), window=figure_frame, anchor='nw')
        figure_frame.config(height = '10c', width='16c')
    
        row_num = 0
        column = False
        for figs in figs_to_plot:
            figure_canvas = FigureCanvasTkAgg(figs, master=figure_frame)
            if column:
                col = 4
            else:
                col = 1
            #figure_canvas.draw()
            figure_canvas.get_tk_widget().grid(
                    row=row_num, column=col,columnspan=2, rowspan = 5, pady = 5,padx = 8, sticky=E)
            #figure_canvas._tkcanvas.grid(row=row_num, column = 0,columnspan = 10, rowspan = 10, sticky= W+E+N+S, pady = 5,padx = 5)
            if column:
                row_num += 5
            column = not column
        

        figure_frame.update_idletasks()
        frame_canvas.config(width='16c', height='10c')
        
        # Set the canvas scrolling region
        main_canvas.config(scrollregion=figure_frame.bbox("all"))
        
            
    def plot_init_dist(self):
        '''
        This function will plot the distribution of variable calls prior to running
        the simulation. This will enable users to see whether the distributions are as they expected.
        
        '''
        
        self.get_distributions()        
        fig_list =[]
        for var, values in self.simulation_dist.items():
            fig = Figure(figsize = (3,3), facecolor=[240/255,240/255,237/255], tight_layout=True)
            a = fig.add_subplot(111)
            num_bins = 15
            a.hist(values, num_bins, facecolor='blue', edgecolor='black', alpha=1.0)
            a.set_title(var)
            fig_list.append(fig)
            
        if self.univar_row_num != 0:
            row_num = 17
        else:
            row_num = 10
        frame_canvas = Frame(self.current_tab)
        frame_canvas.grid(row=row_num, column=1, columnspan = 3,pady=(5, 0))
        frame_canvas.grid_rowconfigure(0, weight=1)
        frame_canvas.grid_columnconfigure(0, weight=1)
        frame_canvas.config(height = '10c', width='16c')
        
        main_canvas = Canvas(frame_canvas)
        main_canvas.grid(row=0, column=0, sticky="news")
        main_canvas.config(height = '10c', width='16c')
        
        vsb = Scrollbar(frame_canvas, orient="vertical", command=main_canvas.yview)
        vsb.grid(row=0, column=2,sticky = 'ns')
        main_canvas.configure(yscrollcommand=vsb.set)
        
        figure_frame = Frame(main_canvas)
        main_canvas.create_window((0, 0), window=figure_frame, anchor='nw')
        figure_frame.config(height = '10c', width='16c')
    
        row_num = 0
        column = False
        for figs in fig_list:
            figure_canvas = FigureCanvasTkAgg(figs, master=figure_frame)
            if column:
                col = 4
            else:
                col = 1
            figure_canvas.get_tk_widget().grid(
                    row=row_num, column=col,columnspan=2, rowspan = 5, pady = 5,padx = 8, sticky=E)

            if column:
                row_num += 5
            column = not column

        figure_frame.update_idletasks()
        frame_canvas.config(width='16c', height='10c')
        main_canvas.config(scrollregion=figure_frame.bbox("all"))
        
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
        self.cleanup_thread = Thread(target=self.cleanup_processes_and_COMS)
        self.cleanup_thread.start()
        
    def abort_univar_overall_fun(self):
        self.abort_univar_overall.value = True
        self.abort_sim()
        
    def cleanup_processes_and_COMS(self):
        try:
            self.current_simulation.close_all_COMS()
            self.current_simulation.terminate_processes()
            save_data(self.current_simulation.output_file, self.current_simulation.results, self.current_simulation.directory)
        except:
            self.after(1000, self.cleanup_processes_and_COMS)
            


        
        
       
        
        
        
        
################################################################################    
        
        
        
        
class Simulation(object):
    def __init__(self, sims_completed, tot_sim, simulation_vars, output_file, directory, 
                 aspen_file, excel_solver_file,abort, vars_to_change, output_value_cells,
                 output_columns, save_freq=10, num_processes=1):
        self.manager = Manager()
        self.num_processes = min(num_processes, tot_sim)
        self.tot_sim = tot_sim
        self.sims_completed = sims_completed
        self.save_freq = self.manager.Value('i', save_freq)
        self.abort = abort
        self.simulation_vars = self.manager.dict(simulation_vars) 
        self.output_file = self.manager.Value('s', output_file)
        self.directory = self.manager.Value('s', directory)
        self.aspen_file = self.manager.Value('s', aspen_file)
        self.excel_solver_file = self.manager.Value('s', excel_solver_file)
        self.output_value_cells = self.manager.Value('s',output_value_cells)
        
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
          
    
    def init_sims(self):
        TASKS = [trial for trial in range(0, self.tot_sim)]
        self.lock_to_signal_finish.acquire()
        if not self.abort.value:
            self.run_sim(TASKS)
        self.lock_to_signal_finish.acquire()
        self.wait(t=5)
        self.close_all_COMS()
        self.terminate_processes()
        self.wait()
            
        save_data(self.output_file, self.results, self.directory)
        self.abort.value = False    
        
        
    def terminate_processes(self):
        for p in self.processes:
            p.terminate()
            p.join()
         
    def wait(self, t=2):
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
                                                                self.results_lock, self.results, self.directory, self.output_value_cells,
                                                                self.trial_counter, self.save_freq, 
                                                                self.output_file, self.vars_to_change, 
                                                                self.output_columns, self.simulation_vars, self.sims_completed, 
                                                                self.lock_to_signal_finish, self.tot_sim)))
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
                
def open_aspenCOMS(aspenfilename):
    aspencom = Dispatch('Apwn.Document')
    aspencom.InitFromArchive(path.abspath(aspenfilename))
    obj = aspencom.Tree     
    return aspencom,obj


def open_excelCOMS(excelfilename):
    pythoncom.CoInitialize()
    excel = DispatchEx('Excel.Application')
    book = excel.Workbooks.Open(path.abspath(excelfilename))
    return excel,book  
   
    
def save_data(outputfilename, results, directory):
    if results:
        collected_data = concat(results).sort_index()
        writer = ExcelWriter(directory.value + '/' + outputfilename.value + '.xlsx')
        collected_data.to_excel(writer, sheet_name ='Sheet1')
        stats = collected_data.describe()
        stats.to_excel(writer, sheet_name = 'Summary Stats')
        writer.save()
    

def worker(current_COMS_pids, pids_to_ignore, aspenlock, excellock, aspenfilename, 
           excelfilename, task_queue, abort, results_lock, results, directory, output_value_cells,
           sim_counter, save_freq, outputfilename, vars_to_change, columns, simulation_vars, sims_completed, lock_to_signal_finish, tot_sim):
    
    local_pids = {}
    aspenlock.acquire()
    if not abort.value:
        aspencom,obj = open_aspenCOMS(aspenfilename.value)
    aspenlock.release()
    excellock.acquire()
    if not abort.value:
        excel,book = open_excelCOMS(excelfilename.value)
    excellock.release() 
    
    for p in process_iter(): #register the pids of COMS objects
        if ('aspen' in p.name().lower() or 'excel' in p.name().lower()) and p.pid not in pids_to_ignore:
            current_COMS_pids[p.pid] = 1
            local_pids[p.pid] = 1
            
            
    for trial_num in iter(task_queue.get, 'STOP'):
        if abort.value:
            continue
        
        aspencom, case_values, errors, obj = aspen_run(aspencom, obj, simulation_vars, trial_num, vars_to_change) 
        result = mp_excelrun(excel, book, aspencom, obj, case_values, columns, errors, trial_num, output_value_cells)
        
        results_lock.acquire()
        results.append(result) 
        sim_counter.value = len(results)
        if sim_counter.value % save_freq.value == 0:
            save_data(outputfilename, results, directory)
        sims_completed.value += 1
        results_lock.release()
        
        if virtual_memory().percent > 95:
            for p in process_iter():
                if p.pid in local_pids:
                    p.terminate()
                    del current_COMS_pids[p.pid]
                    del local_pids[p.pid]
            aspenlock.acquire()
            if not abort.value:
                aspencom,obj = open_aspenCOMS(aspenfilename.value)
            aspenlock.release()
            excellock.acquire()
            if not abort.value:
                excel,book = open_excelCOMS(excelfilename.value)
            excellock.release() 
            
            for p in process_iter(): #register the pids of COMS objects
                if ('aspen' in p.name().lower() or 'excel' in p.name().lower()) and p.pid not in pids_to_ignore:
                    current_COMS_pids[p.pid] = 1
                    local_pids[p.pid] = 1
                    
        aspencom.Engine.ConnectionDialog()
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

    column = [x for x in book.Sheets('Aspen_Streams').Evaluate("D1:D100") if x.Value != None] 
    
    if obj.FindNode(column[0]) == None: # basically, if the massflow out of the system is None, then it failed to converge
        print('ERROR in Aspen for '+ str(case_values))
        dfstreams = DataFrame(columns=columns)
        dfstreams.loc[trial_num] = case_values + [None]*13 + ["Aspen Failed to Converge"]
        return dfstreams
    stream_values = []
    for index,stream in enumerate(column):
        stream_value = obj.FindNode(stream).Value   
        stream_values.append((stream_value,))
    cell_string = "C1:C" + str(len(column))
    book.Sheets('ASPEN_Streams').Evaluate(cell_string).Value = stream_values
    
    excel.Calculate()
    excel.Run('SOLVE_DCFROR')

    
    dfstreams = DataFrame(columns=columns)
    dfstreams.loc[trial_num] = case_values + [x.Value for x in book.Sheets('Output').Evaluate(output_value_cells.value)] + ["; ".join(errors)]
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


        
if __name__ == "__main__":
    freeze_support()
    main_app = MainApp()
    main_app.mainloop()
    if main_app.current_simulation:
        main_app.abort_sim()
        print('Waiting for Clearance to Exit...')
        main_app.current_simulation.wait()
        print('Waiting for Worker Thread to Terminate...')
        main_app.worker_thread.join()
        print('Waiting for Cleanup Thread to Terminate...')
        main_app.cleanup_thread.join()
    exit()
        
        

