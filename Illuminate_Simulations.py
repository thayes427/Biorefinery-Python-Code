from pandas import ExcelWriter, DataFrame, concat, isna
from multiprocessing import Value, Manager, Lock, Queue
from time import sleep
from psutil import process_iter, virtual_memory
from win32com.client import Dispatch, DispatchEx
import pythoncom
from os import path, listdir
import matplotlib.pyplot as plt
from numpy import subtract, percentile
from random import choice
import string
from math import ceil
from shutil import copyfile
from threading import Thread

import multiprocessing as mp
import traceback

class Process_with_Error_Support(mp.Process):
    def __init__(self, *args, **kwargs):
        mp.Process.__init__(self, *args, **kwargs)
        self._pconn, self._cconn = mp.Pipe()
        self._exception = None

    def run(self):
        try:
            mp.Process.run(self)
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
                 output_columns, dispatch, weights, save_bkps, warning_keywords, save_freq=2, 
                 num_processes=1):
        print('hhhllllo')
        self.manager = Manager()
        print('hlsls')
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
        self.save_bkps = self.manager.Value('b',save_bkps)
        self.warning_keywords = self.manager.list(list(warning_keywords))
          
    
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
        for p in self.processes:
            if p.exception:
                error, traceback = p.exception
                print(traceback)
        sleep(4)
        if not self.abort.value and any(p.is_alive() for p in self.processes):
            self.check_mp_errors()
        
        
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
        task_queue = Queue()
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
                            self.warning_keywords)
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
        if len(collected_data) > 0:
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
                        if bin_width == 0:
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
           warning_keywords):
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
        book.Sheets('Set-up').Evaluate('B1').Value = full_bkp_name
        
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
                            summary_formatted)-1] + summary[i]
                    error_count += (summary[i].count(' ') + 1)
                else:
                    i -= 1
                    scan_errors = False                    
        elif "warning" in summary[i].lower():
            summary_formatted[len(summary_formatted)-1] = summary_formatted[len(
                    summary_formatted)-1] + summary[i]
            scan_warnings = True
            while scan_warnings:
                i += 1
                if len(summary[i]) > 0:
                    summary_formatted[len(summary_formatted)-1] = summary_formatted[len(
                            summary_formatted)-1] + summary[i]
                    warning_count += (summary[i].count(' ') + 1)
                else:
                    i -= 1
                    scan_warnings = False   
        else:
            summary_formatted[len(summary_formatted)-1] = summary_formatted[len(
                    summary_formatted)-1] + summary[i]
        i += 1

    return [error_count, warning_count, summary_formatted]


#def find_errors_warnings(aspencom, warning_keywords):
#    '''
#    Searches through the Aspen Run Status node in the Aspen Tree to find any 
#    errors in convergence or warning messages that contain the user-defined
#    keywords of interest. Returns a list of strings for errors and warnings 
#    separately.
#    '''
#    aspen_tree = aspencom.Tree
#    error = r'\Data\Results Summary\Run-Status\Output\PER_ERROR'
#    not_done = True
#    counter = 1
#    error_number = 0
#    warning_number = 0
#    error_statements = []
#    warning_statements = []
#    while not_done:
#        try:
#            check_for_errors = aspen_tree.FindNode(error + '\\' +  str(counter)).Value
#            if "error" in check_for_errors.lower():
#                ############   FOUND AN ERROR ###################
#                error_statements.append(check_for_errors)
#                scan_errors = True
#                counter += 1
#                while scan_errors:
#                    if len(aspen_tree.FindNode(error + '\\' + str(counter)).Value.lower()) > 0:
#                        error_statements[error_number] = error_statements[error_number] + aspen_tree.FindNode(
#                                error + '\\' + str(counter)).Value
#                        counter += 1
#                    else:
#                        scan_errors = False
#                        error_number += 1
#                        counter += 1
#            elif any(keyword in check_for_errors.lower() for keyword in warning_keywords):
#                ############# FOUND A WARNING WITH WARNING KEYWORD ##############
#                warning_statements.append(check_for_errors)
#                scan_errors = True
#                counter += 1
#                while scan_errors:
#                    if len(aspen_tree.FindNode(error + '\\' + str(counter)).Value.lower()) > 0:
#                        warning_statements[warning_number] = warning_statements[warning_number] + aspen_tree.FindNode(
#                                error + '\\' + str(counter)).Value
#                        counter += 1
#                    else:
#                        scan_errors = False
#                        warning_number += 1
#                        counter += 1
#                
#            else:
#                counter += 1
#        except:
#            not_done = False
#    return error_statements, warning_statements
#if __name__ == "__main__":
#    freeze_support()