# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""
import pandas as pd
import numpy as np
#-------------------- GUI file selection---------------
import os
os.system('clear')
from tkinter import *
from tkinter import filedialog
root=Tk()
root.geometry("400x240")
root.title('Input file ')
root.filename = filedialog.askopenfilename(initialdir = "/",\
                                           title ="Select the Excel file")
my_label = Label(root,text= root.filename).pack()

text=root.filename
button = Button(root,text = "Submit", command = root.destroy)
button.pack()
root.mainloop()
#--------------------End of GUI file selection---------------
file=pd.ExcelFile(text)
 
#file=pd.ExcelFile('D:/MRM-Aug.xlsm')
worksheets = file.sheet_names
print('\n  All Worksheets Details  \n\n',worksheets)

#ata_sheet_list =\
#x for x in worksheets if re.search('r',x,re.IGNORECASE)]
#rint('sheet containing sales',data_sheet_list)
 
each_data_worksheet ={}

for data in worksheets:
    each_data_worksheet[data]=pd.read_excel(text,\
    sheet_name = data,header = 1,index_col=0,usecols="B:M")

length =len(worksheets)
#print(length)
# the dataframe to be concatenated should be of same dimension else error msg
df= pd.DataFrame()
if length > 1:
    for data in worksheets:
        df= pd.concat([df,each_data_worksheet[data]],axis =1,sort=False)     
else:
    for data in worksheets:
        df= each_data_worksheet[data]
#print(Sales & Marketing)       
#print(df.head())
print ("\n   All Column List    \n")        
print(df.columns)      
#df.to_excel('C:/Users/1012031/MRM_outt1.xlsx')
# user input for the columns
mylist =[]
Title =(input("  Enter Department Name  :"))
numbers=int(input("  Enter number of columns [Option 3,4,5,7] :"))

if numbers == 3:
      
    #print(numbers)
    for i in range(numbers):
        name =str(input('  Enter the column Name :'))
        if name not in df.columns:
    #    if name not in['Sales Cost ($)','Marketing Cost ($)','Total Cost',\
    #                  'YTD cost as a % of Revenue']:
            print("!!!!! Incorrect inputs !!!!!!!")
            break
        mylist.append(name)         
#    print(mylist)  
    df2 = df[df.columns.intersection(mylist)]
    df1 = df2.copy()
    #print(df1)
    for ind,row in df1.iterrows():
        value =row[0]+row[1]
    #    value1 =value.cumsum()
        df1.loc[ind,"Total Cost"]= value
        df1["YTD Cost"]=df1["Total Cost"].cumsum()
        df1["YTD Revenue"]=df1.iloc[:,numbers-1].cumsum()
        df1["YTD Cost_% of Revenue"]=round(100*df1["YTD Cost"]/\
           df1["YTD Revenue"],2)
    #    df1.loc[ind,"YTD Cost_% of Revenue)"]= value.cumsum(axis=None)
        pd.options.display.float_format = "{:,.0f}".format
    #    df1.round(decimals=2)
        df_plot1=df1.iloc[0:13]
    #print(df_plot1) 
    df_plot1.to_excel(r'Excel_output.xlsx')  
    #______________ graph ploting variable_____________________________
    
    import matplotlib.pyplot as plt
    #% matplotlib inline
    plt.rcParams["figure.figsize"] = [26, 10]
    plt.clf()
    
    fig1, (ax1,ax2) = plt.subplots(2,sharex=True)
    ax22= ax2.twinx()
    #fig2, ax1 = plt.subplots()
    ax22.plot (df_plot1.index,df_plot1.iloc[:,numbers+3],label='YTD cost %',\
               marker='o')
    ax2.bar (df_plot1.index,df_plot1.iloc[:,0],label= mylist[0],\
             edgecolor ='black', color = 'cornflowerblue')
    #cornflowerblue
    ax2.bar (df_plot1.index,df_plot1.iloc[:,1],label=mylist[1],\
            bottom=df_plot1.iloc[:,0],edgecolor ='black',color = 'coral')
    #'coral'
    #ax2.plot(df_new.index,df_new['Total Cost'],label='Total Cost',\
    #edgecolor ='black', color = 'blue')
    ax1.plot(df_plot1.index,df_plot1.iloc[:,numbers],label='Total Cost',\
             marker ='s',markersize=15, color = 'g') 
    #___________________________ LOOP 2 START-----------------------
elif numbers == 4:

    for i in range(numbers):
        name =str(input('  Enter the column Name :'))
        if name not in df.columns:
            print("!!!!! Incorrect inputs !!!!!!!")
            break
        mylist.append(name)         
#    print(mylist)  
    df2 = df[df.columns.intersection(mylist)]
    df1 = df2.copy()
    #print(df1)
    for ind,row in df1.iterrows():
        value =row[0]+row[1]+row[2]
    #    value1 =value.cumsum()
        df1.loc[ind,"Total Cost"]= value
        df1["YTD Cost"]=df1["Total Cost"].cumsum()
        df1["YTD Revenue"]=df1.iloc[:,numbers-1].cumsum()
        df1["YTD Cost_% of Revenue"]=round(100*df1["YTD Cost"]/\
           df1["YTD Revenue"],2)
    #    df1.loc[ind,"YTD Cost_% of Revenue)"]= value.cumsum(axis=None)
        pd.options.display.float_format = "{:,.0f}".format
    #    df1.round(decimals=2)
        df_plot1=df1.iloc[0:13]
    #print(df_plot1) 
    df_plot1.to_excel('Excel_output.xlsx')  
    #______________ graph ploting variable_____________________________
    
    import matplotlib.pyplot as plt
    #% matplotlib inline
    plt.rcParams["figure.figsize"] = [20, 10]
    plt.clf()
    
    fig1, (ax1,ax2) = plt.subplots(2,sharex=True)
    ax22= ax2.twinx()
    #fig2, ax1 = plt.subplots()
    ax22.plot (df_plot1.index,df_plot1.iloc[:,numbers+3],label='YTD cost %',\
               marker='o')
    ax2.bar (df_plot1.index,df_plot1.iloc[:,0],label= mylist[0],\
             edgecolor ='black', color = 'cornflowerblue')
    #cornflowerblue
    
    ax2.bar (df_plot1.index,df_plot1.iloc[:,1],label=mylist[1],\
            bottom=df_plot1.iloc[:,0],edgecolor ='black',color = 'orange')
    sum = df_plot1.iloc[:,0]+df_plot1.iloc[:,1]
    ax2.bar (df_plot1.index,df_plot1.iloc[:,2],label=mylist[2],\
            bottom=sum,edgecolor ='black',color = 'tan')
    #'coral'
    #ax2.plot(df_new.index,df_new['Total Cost'],label='Total Cost',\
    #edgecolor ='black', color = 'blue')
    ax1.plot(df_plot1.index,df_plot1.iloc[:,numbers],label='Total Cost',\
             marker ='s',markersize=12, color = 'g')
#______________________ LOOP 2 END --------------------------- 
elif numbers == 5:
      
    #print(numbers)
    for i in range(numbers):
        name =str(input('  Enter the column Name :'))
        if name not in df.columns:
    #    if name not in['Sales Cost ($)','Marketing Cost ($)','Total Cost',\
    #                  'YTD cost as a % of Revenue']:
            print("!!!!! Incorrect inputs !!!!!!!")
            break
        mylist.append(name)   
#    print(mylist)  
    df2 = df[df.columns.intersection(mylist)]
    df1 = df2.copy()    
    #print(df1)
    for ind,row in df1.iterrows():
        value =row[0]+row[1]+row[2]+row[3]
    #    value1 =value.cumsum()
        df1.loc[ind,"Total Cost"]= value
        df1["YTD Cost"]=df1["Total Cost"].cumsum()
        df1["YTD Revenue"]=df1.iloc[:,numbers-1].cumsum()
        df1["YTD Cost_% of Revenue"]=round(100*df1["YTD Cost"]/\
           df1["YTD Revenue"],2)
    #    df1.loc[ind,"YTD Cost_% of Revenue)"]= value.cumsum(axis=None)
        pd.options.display.float_format = "{:,.0f}".format
    #    df1.round(decimals=2)
        df_plot1=df1.iloc[0:13]
    #print(df_plot1) 
    df_plot1.to_excel('Excel_output.xlsx')  
    #______________ graph ploting variable_____________________________
    
    import matplotlib.pyplot as plt
    #% matplotlib inline
    plt.rcParams["figure.figsize"] = [20, 10]
    plt.clf()
    
    fig1, (ax1,ax2) = plt.subplots(2,sharex=True)
    ax22= ax2.twinx()
    #fig2, ax1 = plt.subplots()
    ax22.plot (df_plot1.index,df_plot1.iloc[:,numbers+3],label='YTD cost %',\
               marker='o',color = 'g')
    ax2.bar (df_plot1.index,df_plot1.iloc[:,0],label= mylist[0],\
             edgecolor ='black', color = 'cornflowerblue')
    #cornflowerblue
    
    ax2.bar (df_plot1.index,df_plot1.iloc[:,1],label=mylist[1],\
            bottom=df_plot1.iloc[:,0],edgecolor ='black',color = 'orange')
    sum = df_plot1.iloc[:,0]+df_plot1.iloc[:,1]
    ax2.bar (df_plot1.index,df_plot1.iloc[:,2],label=mylist[2],\
            bottom=sum,edgecolor ='black',color = 'grey')
    sum1 = sum+df_plot1.iloc[:,2]
    ax2.bar (df_plot1.index,df_plot1.iloc[:,3],label=mylist[3],\
            bottom=sum1,edgecolor ='black',color = 'tan')
    #'coral'
    #ax2.plot(df_new.index,df_new['Total Cost'],label='Total Cost',\
    #edgecolor ='black', color = 'blue')
    ax1.plot(df_plot1.index,df_plot1.iloc[:,numbers],label='Total Cost',\
             marker ='s',markersize=12, color = 'g') 
    #___________________________ LOOP 2 START-----------------------
elif numbers == 7:
      
    #print(numbers)
    for i in range(numbers):
        name =str(input('  Enter the column Name :'))
        if name not in df.columns:
    #    if name not in['Sales Cost ($)','Marketing Cost ($)','Total Cost',\
    #                  'YTD cost as a % of Revenue']:
            print("!!!!! Incorrect inputs !!!!!!!")
            break
        mylist.append(name)   
#    print(mylist)  
    df2 = df[df.columns.intersection(mylist)]
    df1 = df2.copy()    
    #print(df1)
    for ind,row in df1.iterrows():
        value =row[0]+row[1]+row[2]+row[3]+row[4]+row[5]
    #    value1 =value.cumsum()
        df1.loc[ind,"Total Cost"]= value
        df1["YTD Cost"]=df1["Total Cost"].cumsum()
        df1["YTD Revenue"]=df1.iloc[:,numbers-1].cumsum()
        df1["YTD Cost_% of Revenue"]=round(100*df1["YTD Cost"]/\
           df1["YTD Revenue"],2)
    #    df1.loc[ind,"YTD Cost_% of Revenue)"]= value.cumsum(axis=None)
        pd.options.display.float_format = "{:,.0f}".format
    #    df1.round(decimals=2)
        df_plot1=df1.iloc[0:13]
    #print(df_plot1) 
    df_plot1.to_excel('Excel_output.xlsx')  
    #______________ graph ploting variable_____________________________
    
    import matplotlib.pyplot as plt
    #% matplotlib inline
    plt.rcParams["figure.figsize"] = [20, 10]
    plt.clf()
    
    fig1, (ax1,ax2) = plt.subplots(2,sharex=True)
    ax22= ax2.twinx()
    #fig2, ax1 = plt.subplots()
    ax22.plot (df_plot1.index,df_plot1.iloc[:,numbers+3],label='YTD cost %',\
               marker='o',color = 'g')
    ax2.bar (df_plot1.index,df_plot1.iloc[:,0],label= mylist[0],\
             edgecolor ='black', color = 'cornflowerblue')
    #cornflowerblue
    
    ax2.bar (df_plot1.index,df_plot1.iloc[:,1],label=mylist[1],\
            bottom=df_plot1.iloc[:,0],edgecolor ='black',color = 'orange')
    sum = df_plot1.iloc[:,0]+df_plot1.iloc[:,1]
    ax2.bar (df_plot1.index,df_plot1.iloc[:,2],label=mylist[2],\
            bottom=sum,edgecolor ='black',color = 'grey')
    sum1 = sum+df_plot1.iloc[:,2]
    ax2.bar (df_plot1.index,df_plot1.iloc[:,3],label=mylist[3],\
            bottom=sum1,edgecolor ='black',color = 'tan')
    #-------- additional two bars added
    sum2 = sum1+df_plot1.iloc[:,3]
    ax2.bar (df_plot1.index,df_plot1.iloc[:,4],label=mylist[4],\
            bottom=sum2,edgecolor ='black',color = 'sandybrown')
    sum3 = sum2+df_plot1.iloc[:,4]
    ax2.bar (df_plot1.index,df_plot1.iloc[:,5],label=mylist[5],\
            bottom=sum3,edgecolor ='black',color = 'wheat')
    #'coral'lightgreen
    #ax2.plot(df_new.index,df_new['Total Cost'],label='Total Cost',\
    #edgecolor ='black', color = 'blue')
    ax1.plot(df_plot1.index,df_plot1.iloc[:,numbers],label='Total Cost',\
             marker ='s',markersize=12, color = 'g') 
    #___________________________ LOOP 2 START-----------------------    
else:   
    print(" !!! Incorrect option ! Check the Column number and enter ")
    exit()
    ##--------- annotation for columns----------------
##---------------------------------------------

for p in ax2.patches:
    width, height = p.get_width(), p.get_height()
    x, y = p.get_xy()
    ax2.text(x+width/1.8,
            y+height/1.8,
            '{:,.0f}'.format(height),
            horizontalalignment='center',
            verticalalignment='center',
            fontsize = 14,color ='black')

    # fontweight='bold'
#-------------annotation for total cost---------------------   
for x,y in zip(df_plot1.index,df_plot1["Total Cost"]):
    label ="{:,.0f} ".format(y)
  
    ax1.annotate(label, (x,y),textcoords="offset points",                    
                        xytext=(-10,15),ha='center',fontsize = 14,color='k')
#    ax1.grid(axis ='x')
 #---------------------annotation for YTD %------
for x,y in zip(df_plot1.index,df_plot1["YTD Cost_% of Revenue"]):
    label ="{:,.01f} %".format(y)
  
    plt.annotate(label, (x,y),textcoords="offset points",                    
                        xytext=(0,10),ha='center',fontsize = 14,color='brown')
    ax22.margins(y=0.3)
    ax2.margins(y=0.25)
 #--------------------------
ax1.set_title(Title,fontsize=20)

import math
low = min(df_plot1["Total Cost"])
high = max(df_plot1["Total Cost"])
ax1.set_ylim([math.ceil(low-0.5*(high-low)), math.ceil(high+0.5*(high-low))])

#_________________Combining primary and secondary axis legends__________________
ln_1,lb_1 = ax2.get_legend_handles_labels()
ln_2,lb_2 = ax22.get_legend_handles_labels()
lines=ln_1+ln_2
labels=lb_1+lb_2
ax2.legend(lines,labels,loc='best', ncol=3,fontsize=14)
#_________________end of legend block__________________
#ax22.set_ylim([3.5,7.0])
ax1.legend()
#ax2.legend() 
ax2.set_ylabel('Cost in $ X1000')
plt.tight_layout()
plt.show()
fig1.savefig(Title+'.jpg',bbox_inches='tight', dpi=100) 
