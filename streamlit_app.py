# -*- coding: utf-8 -*-
"""
Created on Thu Sep 23 21:05:34 2021

@author: vutukurimadhurireddy
"""

# import docx2txt
# import docx
# # read in word

# doc = docx.Document('D:/documents/thesis/inv/mn/winhh.docx')
# result = docx2txt.process('D:/documents/thesis/inv/mn/winhh.docx')
# result1 = [p.text for p in doc.paragraphs]
from docx2python import docx2python
 
# extract docx content
doc_result = docx2python('D:/documents/thesis/inv/mn/win.docx')
# get separate components of the document
my_list = doc_result.body
ds = doc_result.text
table_list = []
body_list = []
for i in range(0,len(my_list)):
    if  len(my_list[i]) > 1 or len(my_list[i][0]) > 1:
        table_list.append(my_list[i])
    else:
        body_list.append(my_list[i])
# hg = len(table_list[1])
update_tabl = []
def single_rows(table_list):
    for i in range(0,len(table_list)):
        if len(table_list[i]) == 1:
            update_tabl.append(table_list[i])
    orig_list = [x for x in table_list if (x not in update_tabl)]
    return orig_list
orig_list = single_rows(table_list)
counter,total_count = [[] for i in range(3)],0
for i in range(0,len(orig_list)):
     for j in range(0,len(orig_list[i])):
         counter[i].append(len(orig_list[i][j]))
     counter[i] = list(dict.fromkeys(counter[i]))
     total_count = total_count +len(counter[i])
def remove_blanks(update_table, blank):
    x = len(update_table)
    for i in range(1,x):
        # print(i)
        if len(update_table[i]) == 0:
            blank.append(update_table[i])
    update_table = [x for x in update_table if (x not in blank)]
    return update_table
blank = []
counter = remove_blanks(counter, blank)  
# def partion(orig_list): 
update_table,update_table1,update_table11 = [[] for i in range(total_count+10)],[[] for i in range(total_count+10)],[[] for i in range(total_count+10)] 
k = 0           
for i in range(0,len(orig_list)):
    for j in range(0,len(orig_list[i])):
        # print(i,j)
        if  j == 0 and len(orig_list[i][j]) != len(orig_list[i][j+1]) :
              # print(i)
              update_table[k].append(orig_list[i][0])
        elif j == 0 or len(orig_list[i][j]) == len(orig_list[i][j-1]):
                update_table1[k].append(orig_list[i][j])
                # print(i,j)
        elif j+2 <= len(orig_list[i]):
            if len(update_table1[k]) != 0:
                k = k+1
            if len(orig_list[i][j]) == len(orig_list[i][j+1]):
                update_table1[k].append(orig_list[i][j])
                # print('j')  
            elif len(orig_list[i][j]) != len(orig_list[i][j+1]) and len(orig_list[i][j]) != len(orig_list[i][j+1]):
                update_table11[k].append(orig_list[i][j])                  
        elif j-1 == len(orig_list[i]):
            if len(update_table1[k]) != 0:
                k = k+1
            if len(orig_list[i][j]) == len(orig_list[i][j+1]):
                update_table1[k].append(orig_list[i][j])
                # print('j')  
            elif len(orig_list[i][j]) != len(orig_list[i][j+1]) and len(orig_list[i][j]) != len(orig_list[i][j+1]):
                update_table11[k].append(orig_list[i][j])
        else:
            # print('murder')
            update_table11[i].append(orig_list[i][j])
    k = k+1            
update_table = update_table1 + update_table11 + update_table
update_table = remove_blanks(update_table, blank)
# return update_table
# update_table =update_table + update_tabl
# update_table = partion(orig_list)
if total_count == len(update_table):
    update_table = update_table + update_tabl
    # print('sucess')
var = [[] for i in range(total_count+10)]
import pandas as pd
df4 = pd.DataFrame()
for m in range(0,len(update_table)):
    df = pd.DataFrame(update_table[m])
    # print(len(df.columns))
    test = 0
    for i in range(0,len(df.columns)):
          for j in range(0,len(df)):
              if df[i][j] == ['']:
                  test = test+1
                  # print(test)
              if test == len(df):
                  df.drop(i,axis=1,inplace=True)
          test = 0
    var[m] = df.values.tolist()

modified_table  = [x for x in var if x != []]


df = pd.DataFrame(modified_table)
# new_header = df.iloc[0] #grab the first row for the header
# df = df[1:] #take the data less the header row
# df.columns = new_header #set the header row as the df header
# check = modified_table[1][0] + modified_table[1][3]
df.to_excel(r'D:/documents/thesis/inv/mn/win.xlsx', index = False, header=True)
df.to_csv(r'D:/documents/thesis/inv/mn/win1.csv', index = False, header=True)
import csv

# from itertools import zip_longest as zip
# a = zip(*csv.reader(open('D:/documents/thesis/inv/mn/win1.csv', "rb", encoding='unicode-escape')))
# csv.writer(open('D:/documents/thesis/inv/mn/wino.csv', "wb")).writerows(a)
import pandas as pd

csv = pd.read_csv('D:/documents/thesis/inv/mn/win1.csv', header=[0], skiprows=0)
# use skiprows if you want to skip headers
df_csv = pd.DataFrame(data=csv)
transposed_csv = df_csv.T
transposed_csv.to_csv(r'D:/documents/thesis/inv/mn/win11.csv', index = False, header=True)
df2 = pd.read_csv('D:/documents/thesis/inv/mn/win11.csv',header=[0])
cd = len(df2. columns) 
header1,query,header,headings = [],[],[],[]
for a in range(len(df2.columns)):
    def function(a):
        df12 = df2.iloc[0:,a:a+1]
        df12 = df12.dropna()
        list12 = df12.values.tolist()
        # list112 = remove_blanks(list12, blank)
        rem = [[] for i in range(len(list12))]
        rem2 = [[] for i in range(len(list12))]
        for i in range(len(list12)):
            rem[i] = list12[i][0].replace('"','fdz').replace("fdz","'").split("'],")
        # THIS CODE IS FOR REMOVING LISTS WITH SIZE 1 IF NEEDED MAYBE USED IN LATER PART
        # if len(rem[i]) == 1:
        #     print(rem[i])
        #     rem234.append(rem[i])
        #     rem = [x for x in rem if (x not in rem234)] 
        #     rem = [x for x in rem if (x not in blank)]
            rem1 = [[[]] for i in range(len(rem[i]))]
            rem1 = [[[]] for i in range(len(rem[i]))]
            for j in range(len(rem[i])):
                rem1[j] = (rem[i][j].replace("\t","").replace("[['","").replace("['","").replace(" ","").replace("'","").replace("]]",""))
                rem2[i].append(rem1[j])
        return rem2
    rem2 = function(a)
    #CODE FOR CHECKIMG NUMBERIC ALPHABET AND SPECIAL CHARACTERS
    
    # import re
    # val = re.search('[a-zA-Z1-9!@#$%.]+', rem2[1][0])
    # if val != None:
    #     val[0].isalpha() # returns True if the variable is an alphabet
    #     print(val[0]) # this will print the first instance of the matching value
    # ab = any(c.isalpha() for c in rem2[1][1])
    # if ab == True:
    #     print(ab)
    
    def removetabs(rem2):    
        for i in range(len(rem2)):
            for j in range(len(rem2[i])):
                rem2[i][j] = rem2[i][j].replace('\\t','').replace('\\','')
    removetabs(rem2)

    #CODE FOR FIRST LINE IS HEADING
    if len(rem2[0]) == 2:
        k = 0
        for i in range(len(rem2[0])):
            if rem2[0][i] == '':
                k = k+1
        if k-i == 0:
            headings.append(rem2[0])
    
    
    #checking conditions for horizontal and vertical
    
    # #2 ROWS 
    
    if len(rem2[0]) == 2:
        for i in range(len(rem2)):
            if rem2[i][0] != None:
                header.append(rem2[i][0])
                query.append(rem2[i][1])

        
            
    # # 3 ROWS
    # if len(rem2[i]) == 3:
    #     k = 0
    #     for i in range(len(rem2[0])):
    #         if rem2[0][i] == '':
    #             k = k+1
    #     header = []
    #     for i in range(len(rem2)):
    #         if k == 1:
    #                 rem2[i].remove('')
    #                 header.append(rem2[i][0])
    #         elif len(rem2[i]) == 3:
    #             header.append(rem2[i][0])
    
    #not imp #headings
    
    # for i in range(len(rem2)):
    #     k = 0
    #     for j in range(len(rem2[i])):
    #         if rem2[i][j] == '':
    #             k = k+1
    #     if k-j == 0:
    #             headings.append(rem2[i])
    #             headings.append(rem2)
          
    # MORE THAN 3 ROWS
    if len(rem2[0]) > 2:
        k = 0
        for i in range(len(rem2[0])):
            if rem2[0][i] == '':
                k = k+1
        if k-i == 0:
            headings.append(rem2[0])
            headings.append(rem2[1])
            rem2.remove(rem2[0])  
            k = 0
            for i in range(len(rem2[0])):
                if rem2[0][i] == '':
                    k = k+1
            for i in range(len(rem2)):
                if k == 1 and len(rem2[i]) == 3:
                        rem2[i].remove('')
                        header.append(rem2[i][0])  
                        query.append(rem2[i][1])
            if k != 1:
                query1 = []
                header.append(rem2[0])
                for j in range(1,(len(rem2))):
                    query1.append(rem2[j])
                query.append(query1)
        else:
            query1 = []
            header.append(rem2[0])
            for j in range(1,(len(rem2))):
                    query1.append(rem2[j])
            query.append(query1)
    header1.append(header)
    
    

for i in range(len(header)):
    print(i,":",header[i],end='\n')      
# x = input("please select do you want for ouput:") 
# search_list = x.replace('[','').replace(']','').replace("'","")
# search_list = list(search_list.split(", "))
# print(x)  
from tabulate import tabulate
# for i in range(len(header)):
#     if x == header[i]:
#         print((query[i]),end='\n')
#         for n in range(len(headings)):
#             # for m in range(len(headings[n])):
#             if x == headings[n][0]:
#                 print(headings[n-1])
#     elif set(search_list) == set(header[i]):
#         print(tabulate(query[i]),end='\n')
#         for n in range(len(headings)):
#             # for m in range(len(headings[n])):
#             if set(search_list) == set(headings[n]):
#                     print(headings[n-1])
#     else:
#         for j in range(len(header[i])):
#             if x == header[i][j]:
#                 df = i
#                 gh = j
#                 for k in range(len(query[i])):
#                     print(tabulate(query[df][k][gh]),end='\n')   


x = input("please select do you want for ouput:")
x = int(x)
if len(header[x]) < 2 or type(header[x]) == str:
    print(query[x])
else:
    print(tabulate(query[x]))

    
if len(header[x]) < 2 or type(header[x]) == str:
    
    from tkinter import *
    from tkinter import ttk
    
    ws=Tk()
    
    ws.title('PythonGuides')
    ws.geometry('500x500')
    
    set = ttk.Treeview(ws)
    set.pack()
    
    set['columns']= (header[x])
    set.column("#0", width=0,  stretch=NO)
    # set.column(header[3][0],anchor=CENTER, width=80)
    # set.column(header[0][1],anchor=CENTER, width=80)
    # set.column(header[0][2],anchor=CENTER, width=80)
    set.column(header[x],anchor=CENTER, width=280)
    
    set.heading("#0",text="",anchor=CENTER)
    
    set.heading(header[x],text=header[x],anchor=CENTER)
    # set.heading(header[0][0],text="ID",anchor=CENTER)
    # set.heading(header[0][1],text="Full_Name",anchor=CENTER)
    # set.heading(header[0][2],text="Award",anchor=CENTER)
    data = []
    #data
    data.append(query[x])
    print(data)
    global count
    count=0
    valus = [] 
    valus =tuple(data) 
    for record in data:
            set.insert(parent='',index='end',iid = count,text='',values= valus)      
            
    count += 1
    
    
    Input_frame = Frame(ws)
    Input_frame.pack()
    
    id = Label(Input_frame,text=header[0][0])
    id.grid(row=0,column=0)
    
    full_Name= Label(Input_frame,text=header[0][1])
    full_Name.grid(row=0,column=1)
    
    award = Label(Input_frame,text=header[0][2])
    award.grid(row=0,column=1)
    
    id_entry = Entry(Input_frame)
    id_entry.grid(row=1,column=0)
    
    fullname_entry = Entry(Input_frame)
    fullname_entry.grid(row=1,column=1)
    
    award_entry = Entry(Input_frame)
    award_entry.grid(row=1,column=1)
    
    def input_record():
        
    
        global count
       
        set.insert(parent='',index='end',iid = count,text='',values=(id_entry.get(),fullname_entry.get(),award_entry.get()))
        count += 1
    
       
        id_entry.delete(0,END)
        fullname_entry.delete(0,END)
        award_entry.delete(0,END)
         
    #button
    Input_button = Button(ws,text = "Input Record",command= input_record)
    
    Input_button.pack()
    
    
    
    ws.mainloop()

if len(header[x]) >= 3 and type(header[x]) != str and (query[x] != []):
    from tkinter import *
    from tkinter import ttk
    
    ws=Tk()
    
    ws.title('PythonGuides')
    ws.geometry('500x500')
    
    set = ttk.Treeview(ws)
    set.pack()
    
    set['columns']= (header[x])
    set.column("#0", width=0,  stretch=YES)
    # set.column(header[3][0],anchor=CENTER, width=80)
    # set.column(header[0][1],anchor=CENTER, width=80)
    # set.column(header[0][2],anchor=CENTER, width=80)
    for i in range(len(header[x])):
        set.column(header[x][i],anchor=CENTER, width=180)
    
    set.heading("#0",text="",anchor=CENTER)
    for i in range(len(header[x])):
        set.heading(header[x][i],text=header[x][i],anchor=CENTER)
    # set.heading(header[0][0],text="ID",anchor=CENTER)
    # set.heading(header[0][1],text="Full_Name",anchor=CENTER)
    # set.heading(header[0][2],text="Award",anchor=CENTER)
    data = []
    #data
    for i in range(len(query[x])):
        data.append(query[x][i])
    print(data)
    count=0
    valus = []  
    for i in range(len(query[x])):
        mnc = [tuple(x) for x in data]
        mnc = tuple(mnc)
        #         valus = (' '.join(data[i]))
    for i in range(len(mnc)):
        valus = ()
        valus = mnc[i]
        set.insert(parent='',index='end',iid = count,text='',values= valus)      
        count += 1
    
    
    Input_frame = Frame(ws)
    Input_frame.pack()
    
    id = Label(Input_frame,text=header[0][0])
    id.grid(row=0,column=0)
    
    full_Name= Label(Input_frame,text=header[0][1])
    full_Name.grid(row=0,column=1)
    
    award = Label(Input_frame,text=header[0][2])
    award.grid(row=0,column=3)
    
    id_entry = Entry(Input_frame)
    id_entry.grid(row=1,column=0)
    
    fullname_entry = Entry(Input_frame)
    fullname_entry.grid(row=1,column=1)
    
    award_entry = Entry(Input_frame)
    award_entry.grid(row=1,column=2)
    
    def input_record():
        
    
        global count
       
        set.insert(parent='',index='end',iid = count,text='',values=(id_entry.get(),fullname_entry.get(),award_entry.get()))
        count += 1
    
       
        id_entry.delete(0,END)
        fullname_entry.delete(0,END)
        award_entry.delete(0,END)
         
    #button
    Input_button = Button(ws,text = "Input Record",command= input_record)
    
    Input_button.pack()
    
    
    
    ws.mainloop()
     
if len(header[x]) >= 3 and type(header[x]) != str and query[x] != []:
    from tkinter import *
    from tkinter import ttk
    
    ws=Tk()
    
    ws.title('PythonGuides')
    ws.geometry('500x500')
    
    set = ttk.Treeview(ws)
    set.pack()
    
    set['columns']= (header[x])
    set.column("#0", width=0,  stretch=YES)
    # set.column(header[3][0],anchor=CENTER, width=80)
    # set.column(header[0][1],anchor=CENTER, width=80)
    # set.column(header[0][2],anchor=CENTER, width=80)
    for i in range(len(header[x])):
        set.column(header[x][i],anchor=CENTER, width=180)
    
    set.heading("#0",text="",anchor=CENTER)
    for i in range(len(header[x])):
        set.heading(header[x][i],text=header[x][i],anchor=CENTER)
    # set.heading(header[0][0],text="ID",anchor=CENTER)
    # set.heading(header[0][1],text="Full_Name",anchor=CENTER)
    # set.heading(header[0][2],text="Award",anchor=CENTER)
    data = []
    #data
    for i in range(len(query[x])):
        data.append(query[x][i])
    print(data)
    count=0
    valus = []  
    for i in range(len(query[x])):
        mnc = [tuple(x) for x in data]
        mnc = tuple(mnc)
        #         valus = (' '.join(data[i]))
    for i in range(len(mnc)):
        valus = ()
        valus = mnc[i]
        set.insert(parent='',index='end',iid = count,text='',values= valus)      
        count += 1
    
    
    Input_frame = Frame(ws)
    Input_frame.pack()
    
    id = Label(Input_frame,text=header[0][0])
    id.grid(row=0,column=0)
    
    full_Name= Label(Input_frame,text=header[0][1])
    full_Name.grid(row=0,column=1)
    
    award = Label(Input_frame,text=header[0][2])
    award.grid(row=0,column=3)
    
    id_entry = Entry(Input_frame)
    id_entry.grid(row=1,column=0)
    
    fullname_entry = Entry(Input_frame)
    fullname_entry.grid(row=1,column=1)
    
    award_entry = Entry(Input_frame)
    award_entry.grid(row=1,column=2)
    
    def input_record():
        
    
        global count
       
        set.insert(parent='',index='end',iid = count,text='',values=(id_entry.get(),fullname_entry.get(),award_entry.get()))
        count += 1
    
       
        id_entry.delete(0,END)
        fullname_entry.delete(0,END)
        award_entry.delete(0,END)
         
    #button
    Input_button = Button(ws,text = "Input Record",command= input_record)
    
    Input_button.pack()
    
    
    
    ws.mainloop()    

import re
def reggx(regx,ans):
    for i in range(len(body_list)): 
        for j in range(len(body_list[i])): 
            for k in range(len(body_list[i][j])): 
                for n in range(len(body_list[i][j][k])):
                    if regx.match(body_list[i][j][k][n]):
                        findstring = re.findall(ans, body_list[i][j][k][n])
                        listtostring = ''.join(map(str, findstring))
                        return listtostring
kundenummer = reggx(re.compile(r'Monatsrechnung Nr.\s+'),'\d+')
st.table(df2.style.set_precision(2))
