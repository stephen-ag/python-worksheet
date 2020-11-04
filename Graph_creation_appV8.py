# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
#-------------------- GUI file selection---------------
import os
os.system('clear')
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from PIL import ImageTk, Image
root=Tk()
#tk.Label(root, text="How should i change border color", width=50, height=4,\
#         bg="White", highlightthickness=4, highlightbackground="#37d3ff").place(x=10, y=10
root.geometry("400x800+10+20")
#canvas =Canvas(root,width=250,height=50,bg="white")
#canvas.pack()
#my_img=ImageTk.PhotoImage(Image.open("Sales---.jpg"))
#canvas.create_image(40,40,anchor=S,image=my_img)
##canvas.create_line(0, 10, 50, 40, fill="#476042")
##canvas.create_image(20, 20, anchor=NW, image=img)

root['bg']='lightblue'
root['bd']= 3
#oldlace'azure'
root.title('Graph Creation Tool ')
Label(root,text = "Graph Creation Tool ", bg="Dodgerblue3",height ="2",\
      width = "400", fg ="white",
      font = ("Calibri",14)).pack()
Label(root,text = "Note: This tool is specific to finance team\n Input to this\
  tool is an excel sheet with all required data,\n binary excel format not supported",
      height ="3",
      width = "400",
      font = ("Calibri",8)).pack()

text1 = Text(root,height = 1, width = 400,bg = "lightgrey")
text1.insert('1.0',
'''\n This graph tool will allow you to create chart for different departments.\
Twin axis is used, left axis for bar graphs & right for percentage.
stacked bar graph is displayed for selected departments and Line graphs for\
 percentage display, top half of the chart shows the line graph for total cost\
 of departments selected. you can create the graph for selected department\
 with corresponding columns chosen. Picture of the graph is saved and \
 iteractive graph displayed. Computed Excel Worksheet is saved for verification.''')
text1.pack(side="bottom", fill=BOTH, expand=1)
#readme=Entry(root, textvariable = "tool readme", width="25")
#readme.pack(side= "bottom")
#root.pack(fill=BOTH, expand=True)
#user_department = StringVar()
#Label(root,text="Department name").pack(side="left",fill=X, expand =1)
#department = Entry(root, textvariable = user_department )
#department.pack(side="right",fill=X,expand =1)
plt.rcParams['axes.spines.top'] = False
plt.rcParams['axes.spines.right'] = False

def openfile():
    global text
    global file
    
    root.filename = filedialog.askopenfilename(initialdir = "/",\
                                            title ="Select the Excel file")
    my_label = Label(root,width = "45",
                     text= "Selected file path-"+root.filename)
    my_label.pack(side = "bottom")
    text=root.filename
    file=pd.ExcelFile(text) 
#file=pd.ExcelFile
    worksheets = file.sheet_names
    count =0
    for sheet in worksheets:
        count= count+1
   
    my_label1 = Label(root,text= count,width = "45").pack(side ="bottom")
    my_label = Label(root, width = "45",
                     text= "\n Total number of sheets are :").pack(side ="bottom") 
    
    return text
def assign():
    global user_name
    global user_column
    user_name = department_name.get()
    user_column = column_num.get()
    print(user_name)
    print(user_column)
    Label(screen1,text= "Data entered",fg ="green").pack()
    entry1.delete(0,END)
    entry2.delete(0,END)
   
    
def register():
    global department_name
    global column_num
    global screen1
    global entry1
    global entry2
    screen1=Toplevel(root)
    screen1.title("Department and Column Details")
    screen1.geometry("400x250")
    
    department_name = StringVar()
    column_num = IntVar()
    
    Label(screen1,text= "").pack()
    Label(screen1,text= "Enter Department name :").pack()
    Label(screen1,text= "").pack()
    entry1=Entry(screen1, textvariable = department_name,width="25")
    entry1.pack()
    Label(screen1,text= "").pack()
    Label(screen1,text= "Enter number of columns :\n Option[3,4,5,7]").pack()
    Label(screen1,text= "").pack()
    entry2=Entry(screen1, textvariable = column_num, width="25")
    entry2.pack()
    Button(screen1,text ="enter", height ="1",width = "15",
           font = ("Calibri",13),
           command = assign).pack(padx=10,pady=10)  
    #print("started")
button1 = Button(root,text = "Open input File", height ="2", width = "25",\
                 font = ("Calibri",13),bg="dodgerblue3",fg ="white",
                 command = openfile)
button1.pack(padx = 10, pady = 25, )

#print(openfile())
#    user_entry = Entry(root,textvariable= department)
#    user_entry.pack()
#    button2 = Button(root,text = "Enter Dept",  height ="2", width = "25",\
#                    font = ("Calibri",13),command = " ")
#    button2.pack(padx = 10, pady = 10, )
button2 = Button(root,text = "Enter Department & Column",  height ="2", width = "25",\
                font = ("Calibri",13),bg="dodgerblue3",fg ="white",
                command = register)
#command = root.destroy
button2.pack(padx = 10, pady = 25 )

   
def close_window():
    root.destroy() 

button3 = Button(root,text = " Close ",height ="2", width = "25",\
                 font = ("Calibri",13),bg="dodgerblue3",fg ="white",
                 command = close_window)
button3.pack(padx = 10,pady = 25)
#Label(root,text="Department name").pack(side="left",fill=X, expand =1)
#department = Entry(root, textvariable = user_department )
#department.pack(side="right",fill=X,expand =1)

root.mainloop()
#--------------------End of GUI file selection---------------

#file=pd.ExcelFile(text) 
#file=pd.ExcelFile
worksheets = file.sheet_names
#count =0
#for sheet in worksheets:
#    count= count+1
#print(count)

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
global counter
df= pd.DataFrame()
counter =0
if length > 1:
    for data in worksheets:
        df= pd.concat([df,each_data_worksheet[data]],axis =1,sort=False)     
else:
    for data in worksheets:
        counter = counter+1
        df= each_data_worksheet[data]
print(counter)
#print(Sales & Marketing)       
#print(df.head())
print ("\n   All Column List    \n")        
print(df.columns) 

#mylist =[]
#Title =(input("  Enter Department Name  :"))
Title = user_name
#numbers=int(input("  Enter number of columns [Option 3,4,5,7] :"))
numbers = user_column
###################################################
######################################################
########################################################
window = Tk()
window.geometry("400x600+10+20") 
window.title('Multiple Column selection') 
window['bg']='lightblue'
window['bd']= 5  
# for scrolling vertically 
yscrollbar = Scrollbar(window) 
yscrollbar.pack(side = RIGHT, fill = Y)   
label = Label(window, 
              text = "Select the Columns below :  ", 
              font = ("Calibri", 13),
              width = "400",bg="dodgerblue3",fg ="white",
              padx = 10, pady = 10)
#Label(root,text = "Graph Creation tool [Finance]", bg="Dodgerblue3",height ="2",\
#      width = "400",font = ("Calibri",13)).pack()
label.pack() 
label1 = Label(window, 
              text = "Selected Columns:  ", 
              font = ("Calibri", 13),
              padx = 10, pady = 10)

list = Listbox(window, selectmode = "multiple",  
               yscrollcommand = yscrollbar.set)   
# Widget expands horizontally and  
# vertically by assigning both to 
# fill option 
list.pack(padx = 10, pady = 10, 
          expand = YES, fill = "x") 
x= df.columns
  
for each_item in range(len(x)):       
    list.insert(END, x[each_item])
    list.itemconfig(each_item, bg = "white")
yscrollbar.config(command = list.yview)     
#-----looping the address of list to get the string values-----
  ##https://www.geeksforgeeks.org/creating-a-multiple-selection-using-tkinter/
#print(list)
def select():
    global result
    global choise
    result = ''
    choise =[]
    for item in list.curselection():
        result = result + str(list.get(item)) + '\n' 
        name =str(list.get(item))
        label1.pack()
        label1.config(text = result)
        choise.append(name)
#---------closing the window ----
def close_window():
    messagebox.showinfo("User Message! ", " Graph will be saved in .JPG file ")
    window.destroy()
# Attach listbox to vertical scrollbar 
yscrollbar.config(command = list.yview) 
my_button1 = Button(window,text = "Select",height ="2", width = "25",\
                 font = ("Calibri",13),bg="steelblue",fg ="white",
                 command = select)
my_button1.pack(pady = 10)
ok_button2 = Button(window,text = " Submit ",height ="2", width = "25",\
                 font = ("Calibri",13),bg="steelblue",fg ="white",
                 command = close_window)
ok_button2.pack(pady = 10)
window.mainloop()
############  
mylist =choise
print(mylist)
###################################################
######################################################
########################################################   
#df.to_excel('C:/Users/1012031/MRM_outt1.xlsx')
# user input for the columns

#select_all()
if numbers == 3:
#    mylist = mylist + choise  
    #print(numbers)
#    for i in range(numbers):
#        name =str(input('  Enter the column Name :'))
#        if name not in df.columns:
#    #    if name not in['Sales Cost ($)','Marketing Cost ($)','Total Cost',\
#    #                  'YTD cost as a % of Revenue']:
#            print("!!!!! Incorrect inputs !!!!!!!")
#            break
#        mylist.append(name)         
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
    df_plot1.to_excel( Title+'_Computed_sheet.xlsx')  
    #______________ graph ploting variable_____________________________
    
    import matplotlib.pyplot as plt
    #% matplotlib inline
    plt.rcParams["figure.figsize"] = [26, 10]
#    plt.rcParams['axes.spines.left'] = False
    plt.clf()
    
#    fig1, (ax1,ax2) = plt.subplots(4,1,sharex=True)
    fig1 = plt.figure()
#    ax1 = fig1.add_subplot(211)
#    ax2 = fig1.add_subplot(212)
    ax2 = plt.subplot(212)
    ax1 = plt.subplot(211)    

    ax2=plt.subplot2grid((4,1),(1,0),rowspan=3,colspan=1)
    ax1=plt.subplot2grid((4,1),(0,0),rowspan=1,colspan=1,sharex = ax2)
    ax22= ax2.twinx()
#    ax22=plt.subplot2grid((4,1),(1,0),rowspan=3,colspan=1)    
#    ax22= ax2.twinx()
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
#
#    for i in range(numbers):
#        name =str(input('  Enter the column Name :'))
#        if name not in df.columns:
#            print("!!!!! Incorrect inputs !!!!!!!")
#            break
#        mylist.append(name)         
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
    df_plot1.to_excel('Computed_sheet.xlsx')  
    #______________ graph ploting variable_____________________________
    
    import matplotlib.pyplot as plt
    #% matplotlib inline
    plt.rcParams["figure.figsize"] = [20, 10]
    plt.clf()
    fig1 = plt.figure()
    ax2 = plt.subplot(212)
    ax1 = plt.subplot(211, sharex = ax2)    
    ax2=plt.subplot2grid((4,1),(1,0),rowspan=3,colspan=1)
    ax1=plt.subplot2grid((4,1),(0,0),rowspan=1,colspan=1,sharex = ax2)

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
#    for i in range(numbers):
#        name =str(input('  Enter the column Name :'))
#        if name not in df.columns:
#    #    if name not in['Sales Cost ($)','Marketing Cost ($)','Total Cost',\
#    #                  'YTD cost as a % of Revenue']:
#            print("!!!!! Incorrect inputs !!!!!!!")
#            break
#        mylist.append(name)   
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
    df_plot1.to_excel('Computed_sheet.xlsx')  
    #______________ graph ploting variable_____________________________
    
    import matplotlib.pyplot as plt
    #% matplotlib inline
    plt.rcParams["figure.figsize"] = [20, 10]
    plt.clf()
    
    fig1 = plt.figure()
    ax2 = plt.subplot(212)
    ax1 = plt.subplot(211, sharex = ax2)    
    ax2=plt.subplot2grid((4,1),(1,0),rowspan=3,colspan=1)
    ax1=plt.subplot2grid((4,1),(0,0),rowspan=1,colspan=1,sharex = ax2)
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
    df_plot1.to_excel('Computed_sheet.xlsx')  
    #______________ graph ploting variable_____________________________
    
    import matplotlib.pyplot as plt
    #% matplotlib inline
    plt.rcParams["figure.figsize"] = [20, 10] 
    plt.clf()
    
#    fig1, (ax1,ax2) = plt.subplots(4,sharex=True)
    fig1 = plt.figure()
    ax2 = plt.subplot(212)
    ax1 = plt.subplot(211, sharex = ax2)    
    ax2=plt.subplot2grid((4,1),(1,0),rowspan=3,colspan=1)
    ax1=plt.subplot2grid((4,1),(0,0),rowspan=1,colspan=1,sharex = ax2)
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
    
#window.mainloop()    
#    exit()
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
#ax1.axes.xaxis.set_visible(False)
ax1.axis('off')

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
ax1.legend(fontsize=14)

#ax2.legend() 
ax2.set_ylabel('Cost in $ X1000', fontsize = 14)
plt.tight_layout()
#plt.tick_params(top='off', bottom='off', left='off', right='off',
#                labelleft='off', labelbottom='on')
fig1.savefig(Title+'.jpg',bbox_inches='tight', dpi=100) 
plt.show()