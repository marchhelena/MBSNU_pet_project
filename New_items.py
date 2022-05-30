#!/usr/bin/env python
# coding: utf-8

import numpy as np
import pandas as pd
from win32com.client import Dispatch
import re

Server = Dispatch("PX32.OpenServer.1")


#Ввести путь к файлу, где лежит список трубопроводов 
path = r'D:\PROFILES\Martyanova.EI\Desktop\Работа Евгений\Orenburg\Pipe rate_work.xlsx'
df_pipe_original = pd.read_excel(path, sep = '\t')
number = df_pipe_original['Number'].values.tolist()
label = df_pipe_original['Label'].values.tolist()
sep_label = df_pipe_original ['Separators label'].values.tolist()


# Создание элементов + их соединение + изменение их имен на нужные, чтобы в дальнейшем ссылаться на имена

for i in range(len(label)):
    label_i = Server.GetValue('GAP.MOD[{PROD}].PIPE[{%s}].Label' %str(label[i]))
#label_i = Server.GetValue('GAP.MOD[{PROD}].PIPE[{P_1443}].Label')
    Server.DoCommand('GAP.NEWITEM("SEP", "Sep_new", "ABOVE", MOD[{PROD}].EQUIP[{%s}])' %(label_i))
    Server.SetValue('GAP.MOD[{PROD}].SEP[{Sep_new}].Label', 'Sep_'+str(label_i))
    Server.DoCommand('GAP.NEWITEM("SOURCE", "Source_new", "RIGHT", MOD[{PROD}].EQUIP[{%s}])' %(label_i))
    Server.SetValue('GAP.MOD[{PROD}].SOURCE[{Source_new}].Label', 'Source_'+str(label_i))
    Server.DoCommand('GAP.NEWITEM("JOINT", "J_new", "RIGHT", MOD[{PROD}].EQUIP[{%s}])' %(label_i))
    Server.SetValue('GAP.MOD[{PROD}].JOINT[{J_new}].Label', 'J_'+str(label_i))

    joint_parent = Server.GetValue('GAP.MOD[{PROD}].PIPE[{%s}].EndA' %(label_i))
    joint_parent = joint_parent.replace('GAP.MOD[{PROD}].JOINT[{', '').replace('}]', '')
    joint_child = Server.GetValue('GAP.MOD[{PROD}].PIPE[{%s}].EndB' %(label_i))
    joint_child = joint_child.replace('GAP.MOD[{PROD}].JOINT[{', '').replace('}]', '')

    Server.DoCommand('GAP.LINKITEMS(MOD[0].JOINT[{name_j_p}],MOD[0].SEP[{name_sep}])'.format(name_j_p = ('{'+str(joint_parent)+'}'), name_sep = ('{Sep_'+str(label_i)+'}')))
    Server.DoCommand('GAP.LINKITEMS(MOD[0].SEP[{name_sep}],MOD[0].SOURCE[{name_source}])'.format(name_sep = ('{Sep_'+str(label_i)+'}'), name_source = ('{Source_'+str(label_i)+'}')))
    Server.DoCommand('GAP.LINKITEMS(MOD[0].SOURCE[{name_source}],MOD[0].JOINT[{name_j_n}])'.format(name_source = ('{Source_'+str(label_i)+'}'), name_j_n = ('{J_'+str(label_i)+'}')))
    Server.DoCommand('GAP.LINKITEMS(MOD[0].JOINT[{name_j_n}],MOD[0].JOINT[{name_j_c}], "{name_new_pipe}")'.format(name_j_n = ('{J_'+str(label_i)+'}'), name_j_c = ('{'+str(joint_child)+'}'), name_new_pipe = ('COPY_'+str(label_i))))


# Внесение основной информации по трубопроводам в соответсвии с их копией

for i in range(len(label)):
#label_i = Server.GetValue('GAP.MOD[{PROD}].PIPE[{P_1105/2}].Label')
    label_i = Server.GetValue('GAP.MOD[{PROD}].PIPE[{%s}].Label' %str(label[i]))
    desc_count = Server.GetValue('GAP.MOD[{PROD}].PIPE[{%s}].Desc.COUNT' %label_i)
      
    pipe_type = Server.GetValue('GAP.MOD[{PROD}].PIPE[{%s}].PipeModel' %label_i)
    Server.SetValue('GAP.MOD[{PROD}].PIPE[{%s}].PipeModel' %('COPY_'+str(label_i)), str(pipe_type))
    
    pipe_evn_t = Server.GetValue('GAP.MOD[{PROD}].PIPE[{%s}].TMPSUR' %label_i)
    pipe_evn_t = pipe_evn_t.replace(' deg C', '')
    Server.SetValue('GAP.MOD[{PROD}].PIPE[{%s}].TMPSUR' %('COPY_'+str(label_i)), str(pipe_evn_t))
    
    pipe_evn_1 = Server.GetValue('GAP.MOD[{PROD}].PIPE[{%s}].HTCSUR' %label_i)
    pipe_evn_1 = pipe_evn_1.replace(' W/m2/K', '')
    Server.SetValue('GAP.MOD[{PROD}].PIPE[{%s}].HTCSUR' %('COPY_'+str(label_i)), str(pipe_evn_1))
    
    pipe_evn_2 = Server.GetValue('GAP.MOD[{PROD}].PIPE[{%s}].CPO' %label_i)
    pipe_evn_2 = pipe_evn_2.replace(' KJ/Kg/K', '')
    Server.SetValue('GAP.MOD[{PROD}].PIPE[{%s}].CPO' %('COPY_'+str(label_i)), str(pipe_evn_2))
    
    pipe_evn_3 = Server.GetValue('GAP.MOD[{PROD}].PIPE[{%s}].CPG' %label_i)
    pipe_evn_3 = pipe_evn_3.replace(' KJ/Kg/K', '')
    Server.SetValue('GAP.MOD[{PROD}].PIPE[{%s}].CPG' %('COPY_'+str(label_i)), str(pipe_evn_3))
    
    pipe_evn_4 = Server.GetValue('GAP.MOD[{PROD}].PIPE[{%s}].CPW' %label_i)
    pipe_evn_4 = pipe_evn_4.replace(' KJ/Kg/K', '')
    Server.SetValue('GAP.MOD[{PROD}].PIPE[{%s}].CPW' %('COPY_'+str(label_i)), str(pipe_evn_4))
    
    pipe_corr = Server.GetValue('GAP.MOD[{PROD}].PIPE[{%s}].PIPECORR' %label_i)
    Server.SetValue('GAP.MOD[{PROD}].PIPE[{%s}].PIPECORR' %('COPY_'+str(label_i)), str(pipe_corr))
    
    pipe_gravity = Server.GetValue('GAP.MOD[{PROD}].PIPE[{%s}].Matching.AVALS[{PetroleumExperts4}][0]' %label_i)
    Server.SetValue('GAP.MOD[{PROD}].PIPE[{%s}].Matching.AVALS[{PetroleumExperts4}][0]' %('COPY_'+str(label_i)), str(pipe_gravity))
    
    pipe_friction = Server.GetValue('GAP.MOD[{PROD}].PIPE[{%s}].Matching.AVALS[{PetroleumExperts4}][1]' %label_i)
    Server.SetValue('GAP.MOD[{PROD}].PIPE[{%s}].Matching.AVALS[{PetroleumExperts4}][1]' %('COPY_'+str(label_i)), str(pipe_friction))
    
    pipe_tvd_0 = Server.GetValue('GAP.MOD[{PROD}].PIPE[{%s}].Desc[0].TVD' %label_i)
    pipe_tvd_0 = pipe_tvd_0.replace(' m', '')
    Server.SetValue('GAP.MOD[{PROD}].PIPE[{%s}].Desc[0].TVD' %('COPY_'+str(label_i)), str(pipe_tvd_0))
    
    for j in range(1,int(desc_count)):
        pipe_or_not_pipe =  Server.GetValue('GAP.MOD[0].PIPE[{label_i}].Desc[{num_desc}].Type'.format(label_i = ("{"+ str(label_i)+"}"), num_desc = j))
        Server.SetValue('GAP.MOD[0].PIPE[{name_new_pipe}].Desc[{num_desc}].Type'.format(name_new_pipe = ('{COPY_'+str(label_i)+'}'), num_desc = j), str(pipe_or_not_pipe))
    
        pipe_length = Server.GetValue('GAP.MOD[0].PIPE[{label_i}].Desc[{num_desc}].Length'.format(label_i = ("{"+ str(label_i)+"}"), num_desc = j))
        pipe_length = pipe_length.replace(' m', '')
        Server.SetValue('GAP.MOD[0].PIPE[{name_new_pipe}].Desc[{num_desc}].Length'.format(name_new_pipe = ('{COPY_'+str(label_i)+'}'), num_desc = j), str(pipe_length))
        
        pipe_tvd_all = Server.GetValue('GAP.MOD[0].PIPE[{label_i}].Desc[{num_desc}].TVD'.format(label_i = ('{'+ str(label_i)+'}'), num_desc = j))
        pipe_tvd_all = pipe_tvd_all.replace(' m', '')
        Server.SetValue('GAP.MOD[0].PIPE[{name_new_pipe}].Desc[{num_desc}].TVD'.format(name_new_pipe = ('{COPY_'+str(label_i)+'}'), num_desc = j), str(pipe_tvd_all))
        
        pipe_id = Server.GetValue('GAP.MOD[0].PIPE[{label_i}].Desc[{num_desc}].ID'.format(label_i = ('{'+ str(label_i)+'}'), num_desc = j))
        pipe_id = pipe_id.replace(' m', '')
        Server.SetValue('GAP.MOD[0].PIPE[{name_new_pipe}].Desc[{num_desc}].ID'.format(name_new_pipe = ('{COPY_'+str(label_i)+'}'), num_desc = j), str(pipe_id))
        
        pipe_rough = Server.GetValue('GAP.MOD[0].PIPE[{label_i}].Desc[{num_desc}].Roughness'.format(label_i = ('{'+ str(label_i)+'}'), num_desc = j))
        pipe_rough = pipe_rough.replace(' m', '')
        Server.SetValue('GAP.MOD[0].PIPE[{name_new_pipe}].Desc[{num_desc}].Roughness'.format(name_new_pipe = ('{COPY_'+str(label_i)+'}'), num_desc = j), str(pipe_rough))

#Сохранение названий созданных ММБСНУ для дальнейшей работы
path_sep = r'D:\PROFILES\Martyanova.EI\Desktop\Работа Евгений\Orenburg\Separators name.xlsx'
sep_label = list()
sep_count = Server.GetValue('GAP.MOD[{PROD}].SEP.COUNT')
for i in range(int(sep_count)):
    sep_label.append(Server.GetValue('GAP.MOD[{PROD}].SEP[%i].Label' %i))
df_sep_label = pd.DataFrame ({
    'Separators Label':sep_label})   
df_sep_label.to_excel(path_sep)


Server = None







