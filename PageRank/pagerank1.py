# -*- coding: utf-8 -*-
"""
Created on Fri Apr 22 15:37:24 2022

@author: 86131
"""

import numpy as np
import pickle
import xlwt

Air=open('Airlines','rb')
F,cityname=pickle.load(Air)#A:始发地、目的地、航班里程、每周班次
Air.close()
    

#生成转移概率矩阵
P_S=np.zeros((185,185))#转移概率矩阵
S_cc=np.zeros((185,185))#两城市间航班总数
for i in range(len(F)):
    S_cc[F[i][0]][F[i][1]]=S_cc[F[i][0]][F[i][1]]+F[i][3]
S_c=np.sum(S_cc,1)#每座城市航班总数
S_c_0=np.where(S_c==0)[0]#没有航班起飞的城市
S_cc1=S_cc.copy()
for i in range(len(S_c_0)):
    S_cc1[S_c_0[i]]=S_cc1.T[S_c_0[i]]  
S_c=np.sum(S_cc1,1).reshape(185,1)#每座城市航班总数
P_S=S_cc1/S_c

    
#验证P是本原的
k=1
flag=True
P1=P_S
while(k<=1000 and flag==True):
    k=k+1
    P1=P1@P_S
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
    R1=a*R0@P_S+(1-a)*I1/185
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
City.save('pagerank1.xls')


    
    



        
    