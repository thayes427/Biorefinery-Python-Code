# -*- coding: utf-8 -*-
"""
Created on Fri Dec 14 11:02:26 2018

@author: MENGstudents
"""

def run_univar_sens(aspenfile, solverfile, outputfile, simulation_variables):
    for (aspen_variable, aspen_call, fortran_index), values in simulation_variables.items():
        start=time.time()
        print('\nStarting univariate simulation of Variable: ', aspen_variable)
        univariate_analysis(aspenfile, solverfile, aspen_call, aspen_variable, values, fortran_index, outputfile)
        print('Finished univariate simulation of Variable: ', aspen_variable)
        print("Simulation time = %.3f min" % ((time.time()-start)/60))
    print('-----------FINISHED-------------')
   
    
def univariate_analysis(aspenfilename, excelfilename, aspen_call, aspen_var_name, values, fortran_index, outputfilename):
    
    columns= ['Biofuel Output', 'Succinic Acid Output', 'Fixed Op Costs',\
              'Var OpCosts', ' Capital Costs', 'MFSP','Fixed Capital Investment',\
              'Capital Investment with Interest','Loan Payment per Year','Depreciation','Cash on Hand',\
              'Steam Plant Value','Bag Cost']
    dfstreams = pd.DataFrame(columns=columns)
      
    aspenlock = mp.Lock()
    excellock = mp.Lock()
    NUMBER_OF_PROCESSES = 4
    TASKS = [(mp_aspenrun, (aspenfilename, excelfilename, aspen_call, aspen_var_name, case, fortran_index)) for case in values]
    task_queue = mp.Queue()
    done_queue = mp.Queue()
    for task in TASKS:
        task_queue.put(task)
    
    # Start worker processes
    for i in range(NUMBER_OF_PROCESSES):
        mp.Process(target=mp_worker, args=(aspenlock, excellock, task_queue, done_queue)).start()
   # Get and store results
    
#    print('Finished univariate simulation of Variable: '+str(aspen_var_name))
    for case in values:
        dfstreams.loc[case] = done_queue.get()
        print('\tFinished value = '+str(case))
    # Tell child processes to stop
    for i in range(NUMBER_OF_PROCESSES):
        task_queue.put('STOP')

    writer = pd.ExcelWriter(outputfilename + '_' + aspen_var_name + '.xlsx')
    dfstreams.to_excel(writer,'Sheet1')
    writer.save()


def mp_worker(aspenlock, excellock, input, output):
    aspenlock.acquire()
    aspencom,obj = open_aspenCOMS(aspenfilename)
    aspenlock.release()
    excellock.acquire()
    excel,book = open_excelCOMS(excelfilename)
    excellock.release()
    for func, args in iter(input.get, 'STOP'):
        aspencom,value = func(aspencom, obj, *args)
#        excellock.acquire()
        result = mp_excelrun(excel,book,aspencom,value)
#        excellock.release()
        output.put(result)
#    excel.Workbooks.Close(False)
    map(lambda book: book.Close(False), excel.Workbooks)
    excel.Quit()
    excel = None
    del excel

def mp_excelrun(excel, book, aspencom, value):
    
    obj = aspencom.Tree
    column = [x for x in book.Sheets('Aspen_Streams').Evaluate("D1:D100") if x.Value != None] 
    
    if obj.FindNode(column[0]) == None:
        print('ERROR in Aspen for fraction '+ str(value))
        return()
    stream_values = []
    for index,stream in enumerate(column):
        stream_value = obj.FindNode(stream).Value   
        stream_values.append((stream_value,))
    cell_string = "C1:C" + str(len(column))
    book.Sheets('ASPEN_Streams').Evaluate(cell_string).Value = stream_values

    excel.Calculate()
    excel.Run('SOLVE_DCFROR')

#  NEEDS TO BE IMPLEMENTED ABOVE???    
#    if type(case) == str:
#        case = float(case[fortran_index[0]:fortran_index[1]])
    
    dfstream = [x.Value for x in book.Sheets('Output').Evaluate("C3:C15")]
    return dfstream


def mp_aspenrun(aspencom, obj, aspenfilename, excelfilename, aspen_call, aspen_var_name, value, fortran_index):
    '''
    THIS FUNCTION ONLY NEEDS TO BE RUN ONCE
    
    Function fills a dataframe with information needed to perform
    a monte carlo simulation on profitability.
    This function interfaces with an ASPEN file for an
    integrated biorefinery and the NREL TEA file. 
    
    Inputs:
        aspenfilename: string
        excelfilename: string
    Outputs:
        dfstreams
            index is the SA fractionalization
            columns hold info from the TEA calcs
        ***function also outputs an excel file with the same info 
        in the dataframe
    '''
    

    #  ELIMINATE OR MOVE TO EXCEL INPUT FILE???    
    SUC_LOC = r"\Data\Blocks\A300\Data\Blocks\B1\Input\FRAC\TOC5"
    obj.FindNode(SUC_LOC).Value = 0.4
    
    print("Simulating "+str(aspen_var_name)+" value = "+str(value))
    obj.FindNode(aspen_call).Value = value
    
    aspencom.Reinit()
    aspencom.Engine.Run2()
    stop = CheckConverge(aspencom)
    errors = FindErrors(aspencom)
    for e in errors:
        print(e)
    
    return aspencom,value




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
    __spec__ = "ModuleSpec(name='builtins', loader=<class '_frozen_importlib.BuiltinImporter'>)"
    freeze_support()
    pass