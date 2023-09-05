# -*- coding: utf-8 -*-
"""
Created on Wed May  4 18:34:45 2022

@author: 86131
"""

import numpy as np
import pickle
import xlwt

Air=open('Airlines','rb')
F,cityname=pickle.load(Air)#A:始发地、目的地、航班里程、每周班次
Air.close()

#生成邻接矩阵
A=np.zeros((185,185))#加权邻接矩阵
for i in range(len(F)):
    A[F[i][0]][F[i][1]]=A[F[i][0]][F[i][1]]+F[i][3]
    
#HITS
[U,S,V]=np.linalg.svd(A)

authority=V[0]
a_result_num=np.argsort(authority)#svd算法得到的奇异向量非正，因此从小到大排
a_result_city=[cityname[a_result_num[i]] for i in range(185)]

hub=U.T[0]
h_result_num=np.argsort(hub)
h_result_city=[cityname[h_result_num[i]] for i in range(185)]

#结果输出
city=xlwt.Workbook()
worksheet1=city.add_sheet('authority')
for i in range(len(a_result_city)):
    worksheet1.write(i%35,2*(i//35),i+1)
    worksheet1.write(i%35,2*(i//35)+1,a_result_city[i])
worksheet2=city.add_sheet('hub')
for i in range(len(h_result_city)):
    worksheet2.write(i%35,2*(i//35),i+1)
    worksheet2.write(i%35,2*(i//35)+1,h_result_city[i])
city.save('hits1.xls')

