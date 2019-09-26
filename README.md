# Plain-text-to-Excel-Python-Script
Python Script to automatically import data from plain text file to Excel
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import csv, os, openpyxl,re,string

#This function writes a code that test whether row[1] is number 
def is_number(s):
    try:
        float(s)
        return True
    except ValueError:
        return False
#open the Results excel file
from openpyxl import Workbook
wbr=Workbook()
wsr=wbr.active
col=-2; col2=-2
regT=re.compile(r'Max Static Temperature'); 
regM=re.compile(r'MFR')
regH=re.compile(r'Heat Flux')
regI=re.compile(r'Iteration')
os.makedirs('fluenttest', exist_ok=True)
for csvFilename in os.listdir('.'):
    if not csvFilename.endswith('.out'):
        continue # skip non-csv files
    regX=re.compile(r'(.*).out')    #to find filename without .out
    xcelN=regX.findall(csvFilename)
    print('Removing header from '+csvFilename+'.....')
    
    csvRows=[]
    csvFileObj=open(csvFilename,newline='')
    readerObj=csv.reader(csvFileObj,delimiter=' ')
    for row in readerObj:
        if readerObj.line_num==1:
            continue
        csvRows.append(row)

    rcnt=1  #row count in csvRows
    col=col+2   #column upto 26 files
    col2=col2+2 #column for above 26 files
    
    #add name of the file at the top of excel
    for row in csvRows:
        rcnt=rcnt+1
        if len(row)!=2: break
       
        if col<=25:
        #figures out whether element in row is number and then convert to float if needed
            if rcnt==2:
                wsr[string.ascii_uppercase[col]+str(rcnt-1)]=str(xcelN[0]) #add name of the file at the top of excel
                wsr[string.ascii_uppercase[col]+str(rcnt)]=row[0]
                wsr[string.ascii_uppercase[col+1]+str(rcnt)]=row[1]
            else:
                wsr[string.ascii_uppercase[col]+str(rcnt)]=float(row[0])
                wsr[string.ascii_uppercase[col+1]+str(rcnt)]=float(row[1])

        else:
            if rcnt==2:
                wsr[str('A')+string.ascii_uppercase[col2-26]+str(rcnt-1)]=str(xcelN[0]) #add name of the file at the top of excel
                wsr[str('A')+string.ascii_uppercase[col2-26]+str(rcnt)]=row[0]
                wsr[str('A')+string.ascii_uppercase[col2-25]+str(rcnt)]=row[1]
            else:
                wsr[str('A')+string.ascii_uppercase[col2-26]+str(rcnt)]=float(row[0])
                wsr[str('A')+string.ascii_uppercase[col2-25]+str(rcnt)]=float(row[1])

wbr.save('results.xlsx')
csvFileObj.close()
            
wbr.close()
 
    
        
    
        
    
