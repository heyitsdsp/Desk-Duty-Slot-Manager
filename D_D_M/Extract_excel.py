import os
import random

try:
    import xlrd as xlr
except ImportError:
    print("This file needs XLRD to run , please execute 'pip install xlrd' and run this file again")
    exit()

try:
    import openpyxl
except ImportError:
    print("This file needs Openpyxl to run , please execute 'pip install openpyxl' and run this file again")
    exit()

loc=os.getcwd()+"\\slots.xlsx"

wb=xlr.open_workbook(loc)
new_wb=openpyxl.load_workbook('Template.xlsx')
sheet_names=new_wb.get_sheet_names()
names=[]
day={}

c=0

while(c!=5):
    new_sheet=new_wb.get_sheet_by_name(sheet_names[c])
    rval=cval=1

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
        rval=1
        tt[i]=[]
        sjt[i]=[]
        new_sheet.cell(row=rval,column=cval).value=i
        rval+=5
        new_sheet.cell(row=rval,column=cval).value=i
        cval+=1

    name=[]
 
    cval=1
    for i in day.keys():
        rval=2
        x=0
        name.clear()
        if(len(day[i])>3):
            while(len(name)!=3):
                value1=random.randrange(0,len(day[i])-1)
                names=day[i]
                name.append(names[value1])
                day[i].remove(names[value1])
            tt[i]=name[0:len(name)]
        else:
            tt[i]=day[i][1]

        #print(i,':',tt[i])
        while(x!=len(tt[i])):
            new_sheet.cell(row=rval,column=cval).value=tt[i][x]
            x+=1
            rval+=1
        cval+=1
    #print(tt[i])
    #print('\n')

    cval=1
    for i in day.keys():
        rval=7
        x=0
        if(len(day[i])>=3):
            sjt[i]=day[i][:3]
        else:
            sjt[i]=day[i]
        #print(i,':',sjt[i])
        while(x!=len(sjt[i])):
            new_sheet.cell(row=rval,column=cval).value=sjt[i][x]
            x+=1
            rval+=1
        cval+=1
    #print(sjt)
    #print("\n\n")
    #print('\n')
    
    c+=1

new_wb.save('test.xlsx')
