import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
#%matplotlib inline

#______________ read the input excel file and store it in dataframe
df = pd.read_excel(r'D:\New folder\MRM sample.xlsm',header=2,index_col=0,usecols="B:M")
df_new=df.iloc[0:10] # filtering all rows from 0 to 10

#______________export to excel file for backup_________________
df_new.to_excel('MRM sample_copy.xlsx')

#______________user input for selecting columns______________
mylist =[]
numbers=int(input("Enter number of columns "))
print(numbers)
for i in range(numbers):
    name =str(input('Enter the column Name '))
    if name not in['Sales Cost ($)','Marketing Cost ($)','Total Cost','YTD cost as a % of Revenue']:
        print("!!!!! Incorrect inputs !!!!!!!")
        break
    mylist.append(name)
#______________
df_plot1=df_new.loc[:,(mylist)]

#______________ graph ploting variable_____________________________

fig1, (ax1,ax2) = plt.subplots(2)
ax22= ax2.twinx()
#fig2, ax1 = plt.subplots()
ax22.plot (df_plot1.index,df_plot1['YTD cost as a % of Revenue'],label='YTD cost %',marker='o')
ax2.bar (df_plot1.index,df_plot1['Sales Cost ($)'],label='Sales Cost',edgecolor ='black', color = 'cornflowerblue')
ax2.bar (df_plot1.index,df_plot1['Marketing Cost ($)'],label='Mkt Cost',bottom=df_new['Sales Cost ($)'],edgecolor ='black',color = 'coral')
#ax2.plot(df_new.index,df_new['Total Cost'],label='Total Cost',edgecolor ='black', color = 'blue')
ax1.plot(df_plot1.index,df_plot1['Total Cost'],label='Total Cost',marker ='s',markersize=15, color = 'brown')

#plt.setp(ax1.get_xticklabels(), visible=False)
#ax1.axis('off')

#ax2.scatter(df_new.index,df_new['Total Cost'],label='Total Cost',marker ='s',markersize=16,color = 'blue')
for p in ax2.patches:
    width, height = p.get_width(), p.get_height()
    x, y = p.get_xy()
    ax2.text(x+width/1.8,
            y+height/1.8,
            '{:,.0f}'.format(height),
            horizontalalignment='center',
            verticalalignment='center',
            fontsize = 12,color ='black')
#ax22.yaxis.set_visible(False)
for x,y in zip(df_plot1.index,df_plot1['Total Cost']):
    label ="{:,.0f} ".format(y)
  
    ax1.annotate(label, (x,y),textcoords="offset points",                    
                        xytext=(0,10),ha='center',fontsize = 12,color='red')   
    
#    ax1.annotate('Happy',xy=(5,3500),fontsize = 12)
#------------------------


for x,y in zip(df_plot1.index,df_new['YTD cost as a % of Revenue']):
    label ="{:,.02f} %".format(y)
  
    plt.annotate(label, (x,y),textcoords="offset points",                    
                        xytext=(0,10),ha='center',fontsize = 12,color='red')
##
#for X,Y in enumerate(df_new):



# !!!!!!!!!!!!!!!!below block works but needs to be tweeked
#for i in ax2.patches: 
#    ax2.text(i.get_x()+0.2, i.get_height()+0.5,
#             str(round((i.get_height()), 2)),
#             fontsize = 10, fontweight ='bold', 
#            color ='grey') 
# !!!!!!!!!!!!!!!!
#for p in ax2.patches:
#   ax2.annotate("{:,.0f}".format(p.get_height()),(p.get_x()+p.get_width()/2, p.get_height()),ha='center',
#               va='center',xytext=(0,10 ),textcoords='offset points',fontsize=12)    
ax1.legend()
#ax1.set_ylabel('YTD')
#ax2.set_title('Sales and Marketting',fontsize=18)
ax1.set_title('Sales and Marketting',fontsize=20)
#_________________Combining primary and secondary axis legends__________________
#ax2.legend(loc='best', ncol=2,fontsize=12,frameon=False)
#ax22.legend(loc='best', ncol=1,fontsize=12)
ln_1,lb_1 = ax2.get_legend_handles_labels()
ln_2,lb_2 = ax22.get_legend_handles_labels()

lines=ln_1+ln_2
labels=lb_1+lb_2

ax2.legend(lines,labels,loc='best', ncol=3,fontsize=12)
#_________________end of legend block__________________

ax2.set_ylabel('Cost in $ X1000')
#ax22.set_ylabel('YTD cost %')
#ax22.set_ylim([0.043,0.054])
ax22.set_ylim([0.043,0.054])
#ax2.set_title('second graph')
#ax2.set_xlabel('MONTHS')
#plt.tight_layout()
#plt.xticks(rotation=0, horizontalalignment="center",fontsize=20)
plt.tick_params(top='off', bottom='off', left='off', right='off', labelleft='off', labelbottom='on')
#plt.box(False)
plt.show()
#df_new[["Total Cost","Revenue ($)","Sales Cost ($)","Revenue ($)"]].plot(ax=axes[0], kind='bar')
#df_new[["Sales Cost ($)", "Marketing Cost ($)","Total Cost"]].plot(ax=axes[1], kind='bar');
fig1.savefig('plot3333.jpg',bbox_inches='tight', dpi=100)
#fig2.savefig('plot111.jpg',bbox_inches='tight', dpi=150)