# -*- coding: utf-8 -*-
"""
Created on Sat May  7 12:21:09 2022

@author: 86131
"""

import numpy as np
import pickle
import xlwt

Air=open('Airlines','rb')
F,cityname=pickle.load(Air)#A:始发地、目的地、航班里程、每周班次
Air.close()

#生成邻接矩阵
#-----------------------------------------#
D=np.zeros((185,185))#两城市间距离
A=np.zeros((185,185))#邻接矩阵
#生成D、A
num1=np.zeros((185,185))
for i in range(len(F)):
    A[F[i][0]][F[i][1]]=1
    if(F[i][2]>0):
        num1[F[i][0]][F[i][1]]=num1[F[i][0]][F[i][1]]+1
        D[F[i][0]][F[i][1]]=D[F[i][0]][F[i][1]]+F[i][2]
D=np.where(num1!=0,D/num1,D)

#预处理
number0=[]#有航班里程数为0的城市       
for i in range(len(D)):
    for h in range(len(D)):
        if(A[i][h]==1 and D[i][h]==0):
            number0.append([i,h])

Dmax=np.max(D)
d_=[]
for i in range(len(number0)):
    d=0
    if(D[number0[i][1]][number0[i][0]]!=0):
        d=D[number0[i][1]][number0[i][0]]
    else:      
        d1=0
        num3=0
        d2=0
        num4=0
        for h in range(185):
            if((0<D[number0[i][0]][h]<500 or 0<D[h][number0[i][0]]<500) and D[h][number0[i][0]]<D[h][number0[i][1]]):
                num3=num3+1
                d1=d1+D[h][number0[i][1]]
            if((0<D[number0[i][1]][h]<500 or 0<D[h][number0[i][1]]<500) and D[h][number0[i][1]]<D[h][number0[i][0]]):
                num4=num4+1
                d2=d2+D[h][number0[i][0]]
        if(num3!=0):        
            d1=d1/num3 
        if(num4!=0):
            d2=d2/num4
        d=(d1+d2)/2
        if(d==0):
            num2=0
            for h in range(185):
                if(D[number0[i][0]][h]!=0 and D[h][number0[i][1]]!=0 and D[number0[i][0]][h]+D[h][number0[i][1]]<Dmax):
                    num2=num2+1
                    d=d+D[number0[i][0]][h]+D[h][number0[i][1]]
    d_.append(d)

d_[11]=897.1 #丽江到安顺
d_[23]=1733.4 #杭州到安顺
d_[26]=1060.9 #海口到安顺
d_[32]=951.2 #西双版纳到安顺
d_[33]=1143.9 #西安到安顺

for i in range(len(d_)):
    D[number0[i][0]][number0[i][1]]=d_[i] 
    
D=np.where(D==0,0,1/D)   
#-----------------------------------------#    
    
#HITS
[U,S,V]=np.linalg.svd(D)

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
city.save('hits2.xls')