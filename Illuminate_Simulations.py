from pandas import ExcelWriter, DataFrame, concat, isna
from multiprocessing import Value, Manager, Lock, Queue, Process
from time import sleep
from psutil import process_iter, virtual_memory
from win32com.client import Dispatch, DispatchEx
import pythoncom
from os import path
import matplotlib.pyplot as plt
from re import findall
from random import choice
import string




class Simulation(object):
    def __init__(self, sims_completed, tot_sim, simulation_vars, output_file, directory, 
                 aspen_file, excel_solver_file,abort, vars_to_change, output_value_cells,
                 output_columns, dispatch, weights, save_bkps, save_freq=2, num_processes=1):
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
          
    
    def init_sims(self):
        
        df = DataFrame(columns=['trial'] + list(self.vars_to_change))
        for key, value in self.simulation_vars.items():
            df[key[0]] = value
            ntrials = len(value)
        df['trial'] = range(1, ntrials+1)
        df.to_csv(path.join(self.directory.value, 'Variable_Distributions.csv'),index=False)
        
        
        
        TASKS = [trial for trial in range(0, self.tot_sim)]
        self.lock_to_signal_finish.acquire()
        if not self.abort.value:
            self.run_sim(TASKS)
        else:
            try:
                self.lock_to_signal_finish.release()
            except:
                pass
        print('Simulation Running\nWaiting for Completion Signal')
        self.lock_to_signal_finish.acquire()
        print('Completion Signal Received')
        self.wait()
        self.close_all_COMS()
        self.terminate_processes()
        self.wait()
            
        save_data(self.output_file, self.results, self.directory, self.weights)
        save_graphs(self.output_file, self.results, self.directory, self.weights)        
        self.abort.value = False    
        
        
    def terminate_processes(self):
        for p in self.processes:
            p.terminate()
            p.join()
         
    def wait(self, t=0.1):
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
                                                                self.aspenlock, self.excellock, self.aspen_file, self.save_bkps,
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
        for p in process_iter():
            if p.pid in self.current_COMS_pids:
                p.terminate()
                del self.current_COMS_pids[p.pid]
        self.aspenlock.release()
        self.excellock.release()

    
                
    def find_pids_to_ignore(self):
        for p in process_iter():
            if ('aspen' in p.name().lower() or 'apwn' in p.name().lower()) or 'excel' in p.name().lower():
                self.pids_to_ignore[p.pid] = 1
        
        
        
############ GLOBAL FUNCTIONS ################
                
def open_aspenCOMS(aspenfilename,dispatch):
    aspencom = Dispatch(str(dispatch))
    aspencom.InitFromArchive2(path.abspath(aspenfilename), host_type=0, node='', username='', password='', working_directory='', failmode=0)
    obj = aspencom.Tree
    #aspencom.SuppressDialogs = False     
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
        collected_data = list(filter(lambda x: not isna(x[x.columns[-2]].values[0]), results))
        if len(collected_data) > 0:
            collected_data = concat(collected_data).sort_index()
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
                try:
                    fig.clf()
                except:
                    pass
                try:
                    plt.close('all')
                except:
                    pass

def worker(current_COMS_pids, pids_to_ignore, aspenlock, excellock, aspenfilename, save_bkps,
           excelfilename, task_queue, abort, results_lock, results, directory, output_columns, output_value_cells,
           sim_counter, save_freq, outputfilename, vars_to_change, columns, simulation_vars, sims_completed, lock_to_signal_finish, tot_sim, dispatch, weights):
    
    local_pids_to_ignore = {}
    local_pids = {}
    
    aspenlock.acquire()
    for p in process_iter():
        if ('aspen' in p.name().lower() or 'apwn' in p.name().lower()):
            local_pids_to_ignore[p.pid] = 1
    if not abort.value:
        aspencom,obj = open_aspenCOMS(aspenfilename.value, dispatch)
    for p in process_iter():
        if ('aspen' in p.name().lower() or 'apwn' in p.name().lower()) and p.pid not in local_pids_to_ignore:
            local_pids[p.pid] = 1
    aspenlock.release()
    excellock.acquire()
    if not abort.value:
        
        if save_bkps.value:
            aspen_temp_dir = directory.value
        else:
            aspen_temp_dir = path.dirname(aspenfilename.value)
            bkp_name = ''.join([choice(string.ascii_letters + string.digits) for n in range(10)]) + '.bkp'
            full_bkp_name = path.join(aspen_temp_dir,bkp_name)
        #aspencom.SaveAs(full_bkp_name)
        excel,book = open_excelCOMS(excelfilename.value)
        #book.Sheets('Set-up').Evaluate('B1').Value = full_bkp_name
    excellock.release() 
    
    for p in process_iter(): #register the pids of COMS objects
        if (('aspen' in p.name().lower() or 'apwn' in p.name().lower()) or 'excel' in p.name().lower()) and p.pid not in pids_to_ignore:
            current_COMS_pids[p.pid] = 1
            
            
    for trial_num in iter(task_queue.get, 'STOP'):
        if abort.value:
            try:
                lock_to_signal_finish.release()
            except:
                continue
        
        aspencom, case_values, errors, obj = aspen_run(aspencom, obj, simulation_vars, trial_num, vars_to_change, directory) 
        
        # save bkp file and tell excel calculator to point to correct bkp file
        if save_bkps.value:
            full_bkp_name = path.join(aspen_temp_dir, 'trial_' + str(trial_num + 1) + '.bkp')
        aspencom.SaveAs(full_bkp_name)
        book.Sheets('Set-up').Evaluate('B1').Value = full_bkp_name
        
        result = mp_excelrun(excel, book, aspencom, obj, case_values, columns, errors, trial_num, output_value_cells, directory)
        
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
                if ('aspen' in p.name().lower() or 'apwn' in p.name().lower()):
                    local_pids_to_ignore[p.pid] = 1
            if not abort.value:
                aspencom,obj = open_aspenCOMS(aspenfilename.value,dispatch)
            
            for p in process_iter(): #register the pids of COMS objects
                if ('aspen' in p.name().lower() or 'apwn' in p.name().lower()) and p.pid not in local_pids_to_ignore:
                    current_COMS_pids[p.pid] = 1
                    local_pids[p.pid] = 1
            aspenlock.release()
        
                    
        aspencom.Engine.Host(host_type=0, node='', username='', password='', working_directory='')    

    try:
        lock_to_signal_finish.release()
    except:
        pass
            
            
            


def aspen_run(aspencom, obj, simulation_vars, trial, vars_to_change, directory):
    
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



def mp_excelrun(excel, book, aspencom, obj, case_values, columns, errors, trial_num, output_value_cells, directory):

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
    
    
    
#    excel.Run('sub_ClearSumData_ASPEN')
#    excel.Run('sub_GetSumData_ASPEN')
#    excel.Calculate()
##    excel.Run('SolveProductCost')
#    excel.Run('solvedcfror')
#      
#    dfstreams = DataFrame(columns=columns)
#    dfstreams.loc[trial_num+1] = case_values + [x.Value for x in book.Sheets('Output').Evaluate(output_value_cells.value)] + ["; ".join(errors)]
#    return dfstreams


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