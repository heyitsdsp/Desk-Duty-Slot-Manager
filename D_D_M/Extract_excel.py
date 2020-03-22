import os
import random
try:
    import xlrd as xlr
except ImportError:
    print("This file needs XLRD to run , please execute 'pip install xlrd' and run this file again")
    exit()

loc=os.getcwd()+"\\slots.xlsx"

wb=xlr.open_workbook(loc)
names=[]
day={}

c=0

while(c!=1):
    sheet=wb.sheet_by_index(c)
    for j in range(1,sheet.ncols):
        names.clear()
        for i in range(2,sheet.nrows):
            if(sheet.cell_value(i,j)=='yes' or sheet.cell_value(i,j)=='Yes'):
                names.append(sheet.cell_value(i,0))
            day[sheet.cell_value(1,j)]=names[0:len(names)]
    tt={}
    sjt={}

    for i in day.keys():
        tt[i]=[]
        sjt[i]=[]

    name=[]

    for i in day.keys():
        name.clear()
        while(len(name)!=3):
            value1=random.randrange(0,len(day[i])-1,1)
            names=day[i]
            name.append(names[value1])
            day[i].remove(names[value1])
        tt[i]=name[0:len(name)]

    print(tt)
    print('\n')
    for i in day.keys():
        if(len(day[i])>=3):
            sjt[i]=day[i][:3]
        else:
            sjt[i]=day[i]

    print(sjt)

    c+=1
    

