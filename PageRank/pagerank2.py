# -*- coding: utf-8 -*-
"""
Created on Tue May  3 20:19:17 2022

@author: 86131
"""

import numpy as np
import pickle
import xlwt

Air=open('Airlines','rb')
F,cityname=pickle.load(Air)#A:始发地、目的地、航班里程、每周班次
Air.close()


#生成转移概率矩阵
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

D_S=np.sum(D,1)
D0=np.where(D_S==0)[0] #全0行
for i in range(len(D0)):
    D[D0[i]]=D.T[D0[i]]
    
#计算转移概率矩阵
D1=np.where(D==0,0,1/D)
S_D1=np.sum(D1,1).reshape(185,1)
P_D=np.where(S_D1==0,0,D1/S_D1)
#-----------------------------------------#           


#验证P是本原的
k=1
flag=True
P1=P_D
while(k<=1000 and flag==True):
    k=k+1
    P1=P1@P_D
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
    R1=a*R0@P_D+(1-a)*I1/185
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
City.save('pagerank2.xls')

            
        