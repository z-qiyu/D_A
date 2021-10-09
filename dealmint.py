#!/usr/bin/python
# -*- coding: utf-8 -*-
#coding:utf-8
import datetime
import os
import xlrd
import xlwt
         
def getFiles(sourceDir):
    listfile = []
    for file in os.listdir(sourceDir):
        sourceFile = os.path.join(sourceDir,  file)
        if os.path.isfile(sourceFile):
            listfile.append(sourceFile)
    return listfile

def taskinfo_excel_fun(mypddate,file_name,extension_name):
    wb = xlwt.Workbook()   # 创建excel文件
    sheet = wb.add_sheet('My Sheet')   # 为第一个表命名
    content = mypddate
    for i in range(len(content)):
        for j in range(len(content[i])):
            sheet .write(i, j, content[i][j])
    file_path = os.getcwd()  # 指定要保存的目录
    if not os.path.exists(file_path):  # 如果目录不存在，生成
        os.mkdir(file_path)
    file_path2 = os.path.join(file_path,file_name + '_deal' + extension_name)  # 文件的绝对路径
    print(file_path2)
    wb.save(file_path2)



#vHost
def vhost():
    myfiles = getFiles("d:\\1\\")
    for myfile in myfiles:
        bok = xlrd.open_workbook(myfile)
        myfilename= os.path.basename(myfile)
        file_name, extension_name = os.path.splitext(myfilename)
        sht = bok.sheets()[0]
        mylist = [] #小于4的所有行数据
        for num in range(1,sht.nrows):
            row=sht.row_values(num)
            mind=row[3]  #日最低气温
           
            if(mind != '' and maxd != ''):
                if(float(mind) <= 4): #日最低气温小于4
                    mylist.append(row)

    
        mypddate=[] 
        for i in range(0,len(mylist)):
            getd = datetime.datetime.strptime(mylist[i][0], '%Y-%m-%d').date() #读的excel每行日期值 
            addgetd = getd + datetime.timedelta(days=2)  #24
            addgetdd = getd + datetime.timedelta(days=3) #48
            addgetddd = getd + datetime.timedelta(days=4) #72
            for j in range(0,len(mylist)):  
                 getdj = datetime.datetime.strptime(mylist[j][0], '%Y-%m-%d').date()  #读的excel每行日期值 
                 if(getd == getdj):
                     if(float(mylist[i][3]) - float(mylist[j][3]) >= 8):
                         mypddate.append(mylist[i])
                         break
                 if(addgetd == getdj):
                     if(float(mylist[i][3]) - float(mylist[j][3]) >= 10):
                         min1 = float(mylist[i][3])
                         min2 = float(mylist[j][3])
                         
                         max1 = float(mylist[i][2])
                         max2 = float(mylist[j][2])
                         
                         if(min1>min2):
                             mypddate.append(mylist[i])
                             break
                 if(addgetdd == getdj):
                     if(float(mylist[i][3]) - float(mylist[j][3]) >= 12):
                         min1 = float(mylist[i][3])
                         max1 = float(mylist[i][2])
                         
                         min3 = float(mylist[j][3])
                         max3 = float(mylist[j][2])
                       
                         for w1 in range(0,len(mylist)):
                             if(addgetd ==  datetime.datetime.strptime(mylist[w1][0], '%Y-%m-%d').date()):
                                min2 = float(mylist[w1][3])
                                max2 = float(mylist[w1][2])
                                
                         if(min1>min2>min3):
                             mypddate.append(mylist[i])
                         break    
        mypddate.insert(0,['站号','平均温度','最高温度','最低温度','日较差'])           
        taskinfo_excel_fun(mypddate,file_name,extension_name)
    
if __name__ == "__main__":
    vhost()
