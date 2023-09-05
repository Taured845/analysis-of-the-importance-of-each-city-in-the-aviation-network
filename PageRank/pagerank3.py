# -*- coding: utf-8 -*-
"""
Created on Wed May  4 13:15:45 2022

@author: 86131
"""

import pickle
import numpy as np
from pagerank1 import P_S
from pagerank2 import P_D
import xlwt

Air=open('Airlines','rb')
F,cityname=pickle.load(Air)#A:始发地、目的地、航班里程、每周班次
Air.close()


#生成转移概率矩阵
P=0.5*P_S+0.5*P_D


#验证P是本原的
k=1
flag=True
P1=P
while(k<=1000 and flag==True):
    k=k+1
    P1=P1@P
    flag=False
    for i in range(185):
        for j in range(185):
            if P1[i][j]==0:
                flag=True


#PageRank
R0=np.ones((1,185))/185#初分布
a=1
step=0#迭代次数
R=np.zeros((1,185))#迭代结果
while(step<=1000):
    I1=np.ones((1,185))
    R1=a*R0@P+(1-a)*I1/185
    step=step+1
    s_R01=np.linalg.norm((R1-R0),ord=1)
    if(s_R01<=1*10**(-6)*185):
        R=R1
        break
    else:
        R0=R1.copy()
        
result_num=np.argsort(-R,1)#按R的值从大到小排城市编号
result_city=[cityname[result_num[0,i]] for i in range(185)]


#结果输出
City=xlwt.Workbook()
worksheet=City.add_sheet('city')
for i in range(len(result_city)):
    worksheet.write(i%35,2*(i//35),i+1)
    worksheet.write(i%35,2*(i//35)+1,result_city[i])
City.save('pagerank3.xls')