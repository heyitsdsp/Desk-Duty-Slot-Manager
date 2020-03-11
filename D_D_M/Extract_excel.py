import xlrd as xlr
import random

loc=("C:\\Users\\Dev Shankar Paul\\Desktop\\D_D_M")

wb=xlr.open_workbook("Free Slots Timetable - IEEE.xlsx","rb")

names=[]
monday={}
tuesday={}
wednesday={}
c=0;


while(c!=3):
    sheet=wb.sheet_by_index(c)
    for j in range(1,sheet.ncols):
        names.clear()
        for i in range(2,sheet.nrows):
            if(sheet.cell_value(i,j)=='yes' or sheet.cell_value(i,j)=='Yes'):
                names.append(sheet.cell_value(i,0))
        if(c==0):
            monday[sheet.cell_value(1,j)]=names[0:len(names)]
        elif(c==1):
            tuesday[sheet.cell_value(1,j)]=names[0:len(names)]
        else:
            wednesday[sheet.cell_value(1,j)]=names[0:len(names)]
    c+=1

mondaytt={}
mondaysjt={}

for i in monday.keys():
    mondaytt[i]=[]
    mondaysjt[i]=[]

name=[]

for i in mondaytt.keys():
    name.clear()
    while(len(name)!=3):
        value1=random.randrange(0,len(monday[i])-1,1)
        names=monday[i]
        name.append(names[value1])
        monday[i].remove(names[value1])
    mondaytt[i]=name[0:len(name)]

print(mondaytt)
