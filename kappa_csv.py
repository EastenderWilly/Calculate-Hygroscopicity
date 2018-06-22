# include all functions from the math and numpy library
#import xlrd
#import xlwt
#from xlutils.copy import copy as xlutils_copy
#from openpyxl import *#Workbook, load_workbook
from math import *
import math
import numpy as np
from scipy import fmin
import csv
import os
from os import listdir
from os.path import isfile, isdir, join
from scipy import optimize
import pandas as pd
from datetime import datetime
from datetime import timedelta
#import sysconfig
#print(sysconfig.__file__)
#print(sysconfig._init_posix.func_code)

DATA_DIR_CPC=input('Which folder are the CPC data at? (use correct case)\n')

if DATA_DIR_CPC=='':
   DATA_DIR_CPC = "./CPC" #default folder
   print('\nUse default folder: ',DATA_DIR_CPC)

DATA_DIR_CCN=input('\nWhich folder are the CCN data at? (please use correct case)\n')

if DATA_DIR_CCN=='':
   DATA_DIR_CCN = "./CCN" #default folder
   print('\nUse default folder: ',DATA_DIR_CCN)

DATA_DIR_SMPS=input('\nWhich folder are the SMPS data at? (please use correct case)\n')

if DATA_DIR_SMPS=='':
   DATA_DIR_SMPS = "./SMPS"#default folder
   print('\nUse default folder: ',DATA_DIR_SMPS)

SSr_calibr=['0.155905052','0.278318','0.49774','0.798991']

SSr_calibr_01=input('\nWhat is the calibrated value of SSr=0.1?    ')

try:
    SSr_calibr_01=float(SSr_calibr_01)
    SSr_calibr[0]=SSr_calibr_01
except:
    print('\nUse default value 0.155905052')

SSr_calibr_02=input('\nWhat is the calibrated value of SSr=0.2?    ')

try:
    SSr_calibr_02=float(SSr_calibr_02)
    SSr_calibr[1]=SSr_calibr_02
except:
    print('\nUse default value 0.278318')

SSr_calibr_03=input('\nWhat is the calibrated value of SSr=0.5?    ')

try:
    SSr_calibr_03=float(SSr_calibr_03)
    SSr_calibr[2]=SSr_calibr_03
except:
    print('\nUse default value 0.49774')

SSr_calibr_04=input('\nWhat is the calibrated value of SSr=0.8?    ')

try:
    SSr_calibr_04=float(SSr_calibr_04)
    SSr_calibr[3]=SSr_calibr_04
except:
    print('\nUse default value 0.798991')


#All CPC data are in a single file
try:
 for filename in os.listdir(DATA_DIR_CPC):
  print ("Loading: %s" % filename)
  file_name=join(DATA_DIR_CPC,filename)
  date=[]
  CPC={'date':[],'time':[],'CN':[]}
  try:
    #open(file_name,'r',encoding='windows-1252')
    with open(file_name,'r',encoding='windows-1252') as csvfile:
      cnreader=csv.reader(csvfile,delimiter=',',quotechar='"')
      line=1
      line_CPC=1
      for row in cnreader:                  
          if line==5:
             date=list(row[1:-1:2])
          elif line>=19:
             CPC['date'].append(date)
             for i in range(0,len(row[0:-2:2])):
                try:
                   row[2*i]=datetime.strptime(row[2*i],"%H:%M:%S") 
                   #transfer to time pattern
                except:
                   row[2*i]='NaN'
             line_CPC=line_CPC+1
             CPC['time'].append(row[0:-2:2])
             CPC['CN'].append(row[1:-1:2]) 
          line=line+1
  except IOError:
    print('\nFailed to open file ',file_name) 
 del date
except IOError:
 print('\nFailed to open folder ',DATA_DIR_CPC)


CCN={'date':[],'time':[],'SSr':[],'stat':[],'CCN':[]}#data with stat=0 to be dropped 
try:
 for filename in os.listdir(DATA_DIR_CCN):
  print ("Loading: %s" % filename)
  file_name=join(DATA_DIR_CCN,filename)
  try:
    with open(file_name,'r') as csvfile:
      ccnreader=csv.reader(csvfile,delimiter=',',quotechar='"')
      line=1
      for row in ccnreader:
          if line==2: 
             date=row[1]
          elif not row:
             continue
          elif line >= 5:
             CCN['date'].append(date)
             CCN['time'].append(datetime.strptime(row[0],"%H:%M:%S"))
             if float(row[1])==0.1:
                CCN['SSr'].append(SSr_calibr[0])
             elif float(row[1])==0.2:
                CCN['SSr'].append(SSr_calibr[1])
             elif float(row[1])==0.5:
                CCN['SSr'].append(SSr_calibr[2])
             elif float(row[1])==0.8:
                CCN['SSr'].append(SSr_calibr[3])
             else:
                CCN['SSr'].append('NaN')

             CCN['stat'].append(row[2])
             if float(row[2]) == 1.0:                
                CCN['CCN'].append(row[45])
             elif float(row[2]) == 0.0:
                CCN['CCN'].append('NaN')
          line=line+1
  except IOError:
    print('\nFailed to open file ',file_name) 
except IOError:
 print('\nFailed to open folder ',DATA_DIR_CCN)

diameter=[]
SMPS={'date':[],'time':[],'diameter':[],'number':[],'total number':[]}
CCN_CN_new={'date':[],'time':[],'CCN':[],'CN':[],'SSr':[],'critical diameter':[],\
'kappa':[],'SMPS':[],'predict_kappa':[]}

try:
 for filename in os.listdir(DATA_DIR_SMPS):
  print ("Loading: %s" % filename)
  file_name=join(DATA_DIR_SMPS,filename)
  try:
    with open(file_name,'r',encoding='windows-1252') as csvfile:
      smpsreader=csv.reader(csvfile,delimiter=',',quotechar='"')
      line=1
      for row in smpsreader:
          if line==19:
             diameter=row[8:120]
          elif line>=20:
             SMPS['diameter'].append(diameter)
             SMPS['date'].append(row[1])
             SMPS['time'].append(row[2])
             SMPS['number'].append(row[8:120])
             SMPS['total number'].append(row[144])
            
             CCN_CN_new['date'].append(row[1])

             #x is the start time of 5 minute interval of SMPS data
             #y is the end time
             x=datetime.strptime(row[2],"%H:%M:%S")
             if x.second!=0:
                x=x-timedelta(seconds=x.second)
             if x.minute%5!=0:
                x=x-timedelta(minutes=x.minute%5)
             y=x+timedelta(minutes=4,seconds=59)
             
             row[2]=datetime.strftime(x,"%H:%M:%S")

             CCN_CN_new['time'].append(row[2])
             CCN_CN_new['SMPS'].append(row[144])


             #find where are the corresponding CCN and CPC data?

             date_CN=[]
             for i in range(0,len(CPC['date'][0])):
                 if CPC['date'][0][i]==row[1]:
                    date_CN.append(i)
              
             time_CN=[]
             for j in range(0,line_CPC):
                 try:
                    if CPC['time'][j][date_CN[0]]>=x and CPC['time'][j][date_CN[0]]<=y: 
                       time_CN.append(j)
                 except:
                    break                   
             date_CCN=[i for i,s in enumerate(CCN['date']) if row[1] in s]
             time_CCN=[]
             for i in range(0,len(CCN['time'])):
                 if CCN['time'][i]>=x and CCN['time'][i]<=y:
                    time_CCN.append(i)
             
             date_time_CCN=set(time_CCN).intersection(set(date_CCN))
              
             if len(date_time_CCN)>=1 and len(time_CN)>=1 and len(date_CN)==1:

                CCN_count = 0.0
                CCN_mean=0.0
                SSr_count = 0.0
                SSr_mean=0.0

                for i in date_time_CCN:
                    if (CCN['CCN'][i]!='NaN'):
                       CCN_count += 1.0
                       CCN_mean = CCN_mean + float(CCN['CCN'][i])
                    else:
                       CCN_count += 0.0
                       CCN_mean = CCN_mean

                    if (CCN['SSr'][i]!='NaN'):
                       SSr_count += 1.0
                       SSr_mean = SSr_mean + float(CCN['SSr'][i])
                    else:
                       SSr_count += 0.0
                       SSr_mean = SSr_mean

                if CCN_count!=0.0:
                   CCN_mean=CCN_mean/CCN_count
                else:
                   CCN_mean='NaN'
                CCN_CN_new['CCN'].append(CCN_mean)

                if SSr_count!=0.0:
                   SSr_mean=SSr_mean/SSr_count
                   if abs(SSr_mean-float(CCN['SSr'][i]))<0.000001:
                      CCN_CN_new['SSr'].append(CCN['SSr'][i])
                   else:
                      CCN_CN_new['SSr'].append('NaN')
                else:
                   SSr_mean='NaN'
                   CCN_CN_new['SSr'].append(SSr_mean)

                CN_count = 0.0
                CN_mean=0.0                
                for i in date_CN:
                    for j in time_CN:
                        if CPC['CN'][j][i]!='NaN' and CPC['CN'][j][i]!='':
                           if CPC['CN'][j][i]!='NaN' and CPC['CN'][j][i]!='':
                              CN_count += 1.0
                              CN_mean = CN_mean + float(CPC['CN'][j][i])
                           else:
                              CN_count += 0.0
                              CN_mean = CN_mean
                if CN_count!=0.0:
                   CN_mean=CN_mean/CN_count
                else:
                   CN_mean='NaN'
                CCN_CN_new['CN'].append(CN_mean)

                if CN_mean!=0.0 and CN_mean!='NaN':
                   CCN_CN_ratio=CCN_mean/CN_mean
                else:
                   CCN_CN_ratio='NaN'


                if CCN_CN_ratio=='NaN':
                   continue
                elif CCN_CN_ratio>1.0 and CCN_CN_ratio<=1.1:
                   CCN_CN_ratio=1.0
                elif CCN_CN_ratio>=1.1 or CCN_CN_ratio<0.0:
                   CCN_CN_ratio='NaN'
 
                if row[144]!='NaN' and CCN_CN_ratio!='NaN': #row[144]:total number
                   critical_number=float(row[144])*64.0*CCN_CN_ratio
                   for i in range(1,len(row[8:120])+1):
                       if i==len(row[8:120]):
                          next_diameter=13.1
                       else:
                          next_diameter=float(diameter[-i-1])
                       critical_number=critical_number-float(row[120-i])
                       if critical_number==0.0:
                          lowerbound=(math.log10(float(diameter[-i]))+math.log10(next_diameter))/2.0
                          lowerbound=math.pow(10.0,lowerbound)
                          critical_diameter=lowerbound
                       elif critical_number<0.0:
                          upperbound=(math.log10(float(diameter[-i]))+math.log10(float(diameter[-i+1])))/2.0
                          upperbound=math.pow(10.0,upperbound)
                          lowerbound=(math.log10(float(diameter[-i]))+math.log10(next_diameter))/2.0
                          lowerbound=math.pow(10.0,lowerbound)
                          critical_diameter=(lowerbound*(float(row[120-i])+critical_number)\
+upperbound*(critical_number)*(-1.0))/float(row[120-i])
                          break
                       elif critical_number>0.0:
                          continue
                   CCN_CN_new['critical diameter'].append(critical_diameter) 
                else:
                   CCN_CN_new['critical diameter'].append('NaN')
                
             else:
                CCN_CN_new['CN'].append('NaN')
                CCN_CN_new['CCN'].append('NaN')
                CCN_CN_new['critical diameter'].append('NaN')
                CCN_CN_new['SSr'].append('NaN')
          line=line+1
  except IOError:
    print('\nFailed to open file ',file_name)
except IOError:
 print('\nFailed to open folder ',DATA_DIR_SMPS)


T=298.15
sigma=0.072
A=4.0*sigma*18.0/(8.314*T*1.0e6)
Dd=0.0


for i in range(0,len(CCN_CN_new['date'])):
    if CCN_CN_new['SSr'][i]!='NaN':
       d=float(CCN_CN_new['critical diameter'][i])*1.0e-9
       try:
          def g(k):
              def f(dwet):
                  s=-(dwet**3.0-d**3.0)/(dwet**3.0-d**3.0*(1.0-k))*exp(A/dwet)
                  return s
              dmin=optimize.fmin(f,d,disp=0)
              smax=(dmin**3.0-d**3.0)/(dmin**3.0-d**3.0*(1.0-k))*exp(A/dmin)
              t=abs(smax[0]-(1.0+float(CCN_CN_new['SSr'][i])*0.01))
              return t
          kmin=optimize.fmin(g,0.1,disp=0)
          kappa=kmin[0]
       
          CCN_CN_new['kappa'].append(kappa)
       except:
          CCN_CN_new['kappa'].append('NaN')
       kappa_temp=4.0*(A**3.0)/(27.0*(d**3.0)*(log(float(CCN_CN_new['SSr'][i])*0.01+1.0))**2.0)
       CCN_CN_new['predict_kappa'].append(kappa_temp)
    else:
       CCN_CN_new['kappa'].append('NaN')
       CCN_CN_new['predict_kappa'].append('NaN')


#write data into excel
with open('kappa_csv.csv','w',newline='') as csvfile:
     data_writer=csv.writer(csvfile)
     header=['date','time','SSr='+SSr_calibr[0],'SSr='+SSr_calibr[1],\
'SSr'+SSr_calibr[2],'SSr'+SSr_calibr[3],'predict kappa']
     data_writer.writerow(header)
     for i in range(0,len(CCN_CN_new['CCN'])):
         if CCN_CN_new['SSr'][i]==SSr_calibr[0]:
            line=[CCN_CN_new['date'][i],CCN_CN_new['time'][i],\
                  CCN_CN_new['kappa'][i],'','','',CCN_CN_new['predict_kappa'][i]]
         elif CCN_CN_new['SSr'][i]==SSr_calibr[1]:
            line=[CCN_CN_new['date'][i],CCN_CN_new['time'][i],'',\
                  CCN_CN_new['kappa'][i],'','',CCN_CN_new['predict_kappa'][i]]
         elif CCN_CN_new['SSr'][i]==SSr_calibr[2]:
            line=[CCN_CN_new['date'][i],CCN_CN_new['time'][i],'','',\
                  CCN_CN_new['kappa'][i],'',CCN_CN_new['predict_kappa'][i]]
         elif CCN_CN_new['SSr'][i]==SSr_calibr[3]:
            line=[CCN_CN_new['date'][i],CCN_CN_new['time'][i],'','','',\
                  CCN_CN_new['kappa'][i],CCN_CN_new['predict_kappa'][i]]
         else:
            line=[CCN_CN_new['date'][i],CCN_CN_new['time'][i],'']
         data_writer.writerow(line)


with open('AR_csv.csv','w',newline='') as csvfile:
     data_writer=csv.writer(csvfile)
     header=['date','time','SSr='+SSr_calibr[0],'SSr='+SSr_calibr[1],\
'SSr'+SSr_calibr[2],'SSr'+SSr_calibr[3]]
     data_writer.writerow(header)
     for i in range(0,len(CCN_CN_new['CCN'])):
         if CCN_CN_new['SSr'][i]==SSr_calibr[0]:
            line=[CCN_CN_new['date'][i],CCN_CN_new['time'][i],\
                  float(CCN_CN_new['CCN'][i])/float(CCN_CN_new['CN'][i]),'','','']
         elif CCN_CN_new['SSr'][i]==SSr_calibr[1]:
            line=[CCN_CN_new['date'][i],CCN_CN_new['time'][i],'',\
                  float(CCN_CN_new['CCN'][i])/float(CCN_CN_new['CN'][i]),'','']
         elif CCN_CN_new['SSr'][i]==SSr_calibr[2]:
            line=[CCN_CN_new['date'][i],CCN_CN_new['time'][i],'','',\
                  float(CCN_CN_new['CCN'][i])/float(CCN_CN_new['CN'][i]),'']
         elif CCN_CN_new['SSr'][i]==SSr_calibr[3]:
            line=[CCN_CN_new['date'][i],CCN_CN_new['time'][i],'','','',\
                  float(CCN_CN_new['CCN'][i])/float(CCN_CN_new['CN'][i])]
         else:
            line=[CCN_CN_new['date'][i],CCN_CN_new['time'][i],'']
         data_writer.writerow(line)

with open('critical_diameter_csv.csv','w',newline='') as csvfile:
     data_writer=csv.writer(csvfile)
     header=['date','time','SSr='+SSr_calibr[0],'SSr='+SSr_calibr[1],\
'SSr'+SSr_calibr[2],'SSr'+SSr_calibr[3]]
     data_writer.writerow(header)
     for i in range(0,len(CCN_CN_new['CCN'])):
         if CCN_CN_new['SSr'][i]==SSr_calibr[0]:
            line=[CCN_CN_new['date'][i],CCN_CN_new['time'][i],\
                  CCN_CN_new['critical diameter'][i],'','','']
         elif CCN_CN_new['SSr'][i]==SSr_calibr[1]:
            line=[CCN_CN_new['date'][i],CCN_CN_new['time'][i],'',\
                  CCN_CN_new['critical diameter'][i],'','']
         elif CCN_CN_new['SSr'][i]==SSr_calibr[2]:
            line=[CCN_CN_new['date'][i],CCN_CN_new['time'][i],'','',\
                  CCN_CN_new['critical diameter'][i],'']
         elif CCN_CN_new['SSr'][i]==SSr_calibr[3]:
            line=[CCN_CN_new['date'][i],CCN_CN_new['time'][i],'','','',\
                  CCN_CN_new['critical diameter'][i]]
         else:
            line=[CCN_CN_new['date'][i],CCN_CN_new['time'][i],'']
         data_writer.writerow(line)

