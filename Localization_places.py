#!/usr/bin/env python
# coding: utf-8

import numpy as np
import pandas as pd
from win32com.client import Dispatch
import re

Server = Dispatch("PX32.OpenServer.1")

#Введите путь к файлу
full_path = r'D:\PROFILES\Martyanova.EI\Desktop\Работа Евгений\Orenburg\Pipe rate.txt'

#Задать условия для поиска Родителя трубопровода (их совместного узла)
joint_choke_B = list()
joint_pipe_A = list()
well_pipes = list()

choke_count = Server.GetValue('GAP.MOD[{PROD}].INLCHK.COUNT')
pipe_count = Server.GetValue('GAP.MOD[{PROD}].PIPE.COUNT')

#Фиксирование конечного узла для Родителя и начального для трубопровода в список
for i in range(int(choke_count)):
    joint = Server.GetValue('GAP.MOD[{PROD}].INLCHK[%i].EndB'%i)
    joint = joint.replace('GAP.MOD[{PROD}].JOINT[{', '').replace('}]', '')
    joint_choke_B.append(joint)
print(joint_choke_B)

i = 0
joint_pipe_A = list()
for i in range(int(pipe_count)):
    joint_p = Server.GetValue('GAP.MOD[{PROD}].PIPE[%i].EndA'%i)
    joint_p = joint_p.replace('GAP.MOD[{PROD}].JOINT[{', '').replace('}]', '')
    joint_pipe_A.append(joint_p)
print(joint_pipe_A)

#Если конечный узел для Родителя и начальный для трубопровода совпадают,
#то этот трубопровод заносится в список для сохранения инфо, но в работе не участвует
k=0
well_pipes_1 = list()
for i in range(len(joint_pipe_A)):
    if joint_pipe_A[i] in joint_choke_B:
        well_pipes_1.append(i)
        k+=1
well_pipes = list(set(well_pipes_1))
print(well_pipes)

#Если конечный узел для Родителя и начальный для трубопровода совпадают,
#то этот трубопровод заносится в список для сохранения инфо для дальнейшего перебора скважин
k=0
not_well_pipes = list()
pipe_index = list()

for i in range(int(pipe_count)):
    pipe_index.append(i)
    
for i in range(len(pipe_index)):
    if pipe_index[i] not in well_pipes:
        not_well_pipes.append(pipe_index[i])
        k+=1

print(not_well_pipes)
print(k)

#Проверка на соответсвие расхода по трубопровода тех.характеристикам МБСНУ. Формирование окончательного списка.
full_path = r'D:\PROFILES\Martyanova.EI\Desktop\Работа Евгений\Orenburg\Pipe rate.txt'
path = r'D:\PROFILES\Martyanova.EI\Desktop\Работа Евгений\Orenburg\Pipe rate.xlsx'
q_oil = list()
q_liq = list()
q_gas = list()
label = list()
      
for i in range(len(not_well_pipes)):
    q_liq_i = Server.GetValue('GAP.MOD[{PROD}].PIPE[%i].SolverResults[0].LiqRate' %(not_well_pipes[i]))
    q_gas_i = Server.GetValue('GAP.MOD[{PROD}].PIPE[%i].SolverResults[0].GasRate' %(not_well_pipes[i]))
    label_i = Server.GetValue('GAP.MOD[{PROD}].PIPE[%i].Label' %(not_well_pipes[i]))
    
    if (abs(float(q_liq_i))<300.0) and (abs(float(q_liq_i))>0) and (abs(float(q_gas_i))<400.0) and (label_i!=''):
        q_oil.append(Server.GetValue('GAP.MOD[{PROD}].PIPE[%i].SolverResults[0].OilRate' %(not_well_pipes[i])))
        q_gas.append(Server.GetValue('GAP.MOD[{PROD}].PIPE[%i].SolverResults[0].GasRate' %(not_well_pipes[i])))
        q_liq.append(Server.GetValue('GAP.MOD[{PROD}].PIPE[%i].SolverResults[0].LiqRate' %(not_well_pipes[i])))
        label.append(Server.GetValue('GAP.MOD[{PROD}].PIPE[%i].Label' %(not_well_pipes[i])))


pipe_rate = pd.DataFrame ({
    'Label':label,
    'Liquid rate':q_liq,
    'Oil rate':q_oil,
    'Gas rate':q_gas}) 

pipe_rate.to_excel(path)

# pipe_rate.to_csv(full_path)

# pipe_rate


Server = None





