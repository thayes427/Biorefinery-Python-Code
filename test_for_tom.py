# -*- coding: utf-8 -*-
"""
Created on Sun Dec 16 14:57:12 2018

@author: MENGstudents
"""

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Fri Dec 14 14:23:28 2018

@author: Tom1
"""
import tkinter as tk
import threading
import pandas as pd
import multiprocessing as mp
import time
import numpy as np
from multiprocessing import freeze_support
 
    
class MainApp(tk.Tk):

    def __init__(self):
        ####### Do something ######
        tk.Tk.__init__(self)
        self.myframe = tk.Frame(self)
        self.myframe.grid(row=0, column=0, sticky='nswe')
        self.mylabel = tk.Label(self.myframe) # Element to be updated 
        self.mylabel.config(text='No message')
        self.mylabel.grid(row=0, column=0)
        self.mybutton = tk.Button(
            self.myframe, 
            text='Run Simulation',
            command=lambda: self.update_text())
        self.mybutton.grid(row=1, column=0)
        self.sim = None
        self.abort_button = tk.Button(self.myframe, text='Abort', 
                                      command=lambda: self.abort_sim())
        self.abort_button.grid(row=2, column=0)
        self.abort = mp.Value('b', False)
        self.attributes("-topmost", True)
        self.continue_checking = True
        
        
    def abort_sim(self):
        self.abort.value = True

    def start_sim(self):
        self.sim = Sim(400, self)
        self.sim.init_sims()
        
        
    def update_text(self):
        '''
        Spawn a new thread for running long loops in background
        '''
        
        self.mylabel.config(text='Running loop')
        self.new_thread = threading.Thread(
            target=lambda: self.start_sim())
        self.new_thread.start()
        self.after(2000, self.listen_for_result)

    def listen_for_result(self):
        '''
        Check if there is something in the queue
        '''

        try:
            r = pd.concat(self.sim.done).sort_index()
            #r.sort()
            if len(r) == 10000:
                self.continue_checking = False
            print(r)
        except:
            pass
        if self.continue_checking:
            if self.abort.value:
                self.continue_checking = False
            self.after(2000, self.listen_for_result)
        
        
        
class Sim(object):
    def __init__(self, tot_sim, GUI):
        self.numprocess = 4
        self.done = mp.Manager().list()
        self.sim_counter = mp.Value('i',0)
        self.results_lock = mp.Lock()
        self.delete_coms_freq = 100
        self.tot_sim = tot_sim
        self.save_freq = mp.Value('i',50)
        self.processes = []
        self.GUI = GUI
        
        
    def init_sims(self):
        start = time.time()
        #pool = mp.Pool(initializer=initializer)
        for i in range(0, self.tot_sim, self.delete_coms_freq):
            upper_bound = min(i + self.delete_coms_freq, self.tot_sim)
            TASKS = [trial for trial in range(i, upper_bound)]
            if not self.GUI.abort.value:
                #pool.map(pool_fun ,iter((it, self.done, self.GUI.abort, self.sim_counter, 
                                     #self.save_freq) for it in range(i, upper_bound)))
                self.run_sim(TASKS)
            self.wait_til_done()
            self.terminate_processes()
            self.processes = []
        #pool.close()
        #pool.join()
            
        self.save_data()
        print('Total simulation time = ',time.time() - start)
        
    def terminate_processes(self):
        for p in self.processes:
            p.terminate()
            p.join()
        
        
    def wait_til_done(self):
        if not any(p.is_alive() for p in self.processes):
            return
        else:
            time.sleep(1)
            self.wait_til_done()
            
            
    def run_sim(self, tasks):
        task_queue = mp.Queue()
        for task in tasks:
            task_queue.put(task)
        for i in range(self.numprocess):
            new_process = mp.Process(target=worker, args=(
                    task_queue,self.GUI.abort,self.results_lock, self.done, self.sim_counter, self.save_freq))
            self.processes.append(new_process)
            new_process.start()
        for i in range(self.numprocess):
            task_queue.put('STOP')
        
        
    def funs_function(self, it):
        result = []
        for i in range(5):
            time.sleep(0.03)
            v = np.random.normal(it,0.1)
            result.append(v)
        cols = [1,2,3,4,5]
        df = pd.DataFrame(columns=cols)
        df.loc[it] = result
        return df
    
    
    def save_data(self):
        try:
            collected_data = pd.concat(self.done)
        except:
            pass
        #print(collected_data)
        # make writer object and save the data to save location
    
    
def worker(task_queue, abort, results_lock, done, sim_counter, save_freq):
    for it in iter(task_queue.get, 'STOP'):
        if abort.value:
            continue

        result = funnn(it)
        results_lock.acquire()
        done.append(result) 
        sim_counter.value = len(done) - 1
        if not sim_counter.value % save_freq.value:
            print('save')
            #sim_obj.save_data()
        results_lock.release()


def funnn(it):
    result = []
    for i in range(5):
        time.sleep(0.001)
        v = np.random.normal(it,0.1)
        result.append(v)
    cols = [1,2,3,4,5]
    df = pd.DataFrame(columns=cols)
    df.loc[it] = result
    return df

def initializer():
    print('new thread initialized')
            
if __name__ == "__main__":
    freeze_support()
    main_app = MainApp()
    main_app.mainloop()
    main_app.abort_sim()
    time.sleep(1)
    try:
        main_app.sim.terminate_processes()
    except: pass
    time.sleep(0.5)
    #exit()
    
    
        
        

