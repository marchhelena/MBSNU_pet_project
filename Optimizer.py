#!/usr/bin/env python
# coding: utf-8

# Скрипт для определения оптимального положения МБСНУ
# Имя МБСНУ должно иметь соответствующий ей трубопровод, которое оно заменяет

import scipy.optimize as opt
import time
import pandas as pd
from win32com.client import Dispatch

# Выгрузка имен сепараторов
def all_sep(Server, count):

    sep_list = list()

    for i in range(count):
        sep = Server.GetValue('GAP.MOD[{PROD}].SEP[%i].LABEL' %i)
        sep_list.append(sep)
    return sep_list

# расчетная функция
def MBSNU_optimizer(x0):

    count = [0 for v in range(count_pipe)]
    x0 = [int(round(v)) for v in x0]

    for i in range(len(x0)):
        count[x0[i]] = 1
    #print(count)
    df = pd.read_csv(r'optimize.txt', sep = '\t')
    option = ', '.join([str(e) for e in count])
    if len(df['Option']) != 0 and option in df['Option'].tolist():
        ind = df.index[df['Option'] == option].tolist()[0]
        return float(df['Oil Rate'][ind])
    
    else:
        for i in range(len(count)):
            if count[i] == 1:
                # создаем ситерируемый вариант
                Server.DoCommand('GAP.MOD[{PROD}].SEP[{%s}].UNMASK()' %sep_list[i])
                #pipe = 'P_' + sep_list[i]
                Server.DoCommand('GAP.MOD[{PROD}].PIPE[{%s}].MASK()' %(pipe[i]))

        Server.DoCommand(Server.DoCommand('GAP.SOLVENETWORK(0, MOD[{PROD}], 0)'))

        for i in range(len(count)):
            if count[i] == 1:
                # возвращаем в исходное состояние
                Server.DoCommand('GAP.MOD[{PROD}].SEP[{%s}].MASK()' %sep_list[i])
                Server.DoCommand('GAP.MOD[{PROD}].PIPE[{%s}].UNMASK()' %(pipe[i]))
     
        result = - float(Server.GetValue('GAP.MOD[{PROD}].GROUP[{GR_All_Separators}].SolverResults[0].OilRate')) * 0.8362
        liq_rate = float(Server.GetValue('GAP.MOD[{PROD}].GROUP[{GR_All_Separators}].SolverResults[0].LiqRate'))
        gas_rate = float(Server.GetValue('GAP.MOD[{PROD}].GROUP[{GR_All_Separators}].SolverResults[0].GasRate'))
        
        

        print('Oil rate = ', result, 'On/off step = ', count)
        file = open(r'optimize.txt', 'a')
        for i in range(len(x0)):
            try:
                sep_oil = float(Server.GetValue('GAP.MOD[{PROD}].SEP[{%s}].SolverResults[0].OilRate' %sep_list[x0[i]])) * 0.8362
                sep_liq = float(Server.GetValue('GAP.MOD[{PROD}].SEP[{%s}].SolverResults[0].LiqRate' %sep_list[x0[i]]))
                sep_gas = float(Server.GetValue('GAP.MOD[{PROD}].SEP[{%s}].SolverResults[0].GasRate' %sep_list[x0[i]]))
            except:
                sep_oil, sep_liq, sep_gas = 0, 0, 0
                
            file.write(option + '\t' + str(result) + '\t' + str(liq_rate) + '\t' + str(gas_rate) + '\t' + str(sep_list[x0[i]])
                        + '\t' + str(sep_oil) + '\t' + str(sep_liq) + '\t' + str(sep_gas)+ '\n')
        file.close()    
        return result

if __name__ == '__main__':

    N = 5  #Количество искомых мест
    start_program = time.time()
    Server = Dispatch("PX32.OpenServer.1")

    df_sep = pd.read_csv(r'Sep_and_pipe.txt', sep = '\t')
    sep_list = df_sep['Separators label'].values.tolist()
    pipe = df_sep['Label'].values.tolist()
    count_pipe = len(pipe)
    
    bounds = [(0, count_pipe - 1) for i in range(N)] 

    with open(r'optimize.txt', 'w') as file:
        file.write('Option\tOil Rate\tLiq Rate\tGas Rate\tCluster\tSep Oil Rate\tSep Liq Rate\tSep Gas Rate\n')

    # метод дифференциальной эволюции
    try:
        res = opt.differential_evolution(MBSNU_optimizer, bounds = bounds, strategy = 'rand1exp', updating='deferred')
        result = [round(v) for v in res.x.tolist()]
        print(result, res.fun)
    except:
        print('!!!!!! метод дифференциальной эволюции не сработал !!!!!!')
    print('\n!!!!!!!!!!!!!!!!!!!!!!!!!!\nметод дифференциальной эволюции закончил считать\n!!!!!!!!!!!!!!!!!!!!!!!!!!!!\n')

    # метод имитационного отжига
    try:
        res = opt.dual_annealing(MBSNU_optimizer, bounds)
        result = [round(v) for v in res.x.tolist()]
        print(result, res.fun)
    except:
        print('!!!!!! метод имитационного отжига не сработал !!!!!!')

    '''
    instrum = ng.p.Instrumentation(
        ng.p.Array(shape=(6,)).set_bounds(lower=0, upper=1)
    )
    optimizer = ng.optimizers.NGOpt(parametrization=instrum, budget=150, num_workers=1)
    recommendation = optimizer.minimize(MBSNU_optimizer, verbosity=2)
    print(recommendation.value)
    '''

    Server = None
    end_prog = time.time()
    print('\n=== Time: ' + str(end_prog - start_program) + ' sec ===')
