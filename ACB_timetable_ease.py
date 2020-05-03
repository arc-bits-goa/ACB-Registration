
# coding: utf-8

# In[4]:


import numpy as np
import pandas as pd
import csv
import xlrd
import math
from os import listdir
from os.path import isfile, join


# In[5]:


def load_reg():
    files = [f for f in listdir("./data") if isfile(join("./data", f))]
    files.sort()
    files
    registration_data=pd.DataFrame()
    for i in files:
        data = pd.read_excel("./data/"+i,"sheet1",header=1)
        registration_data=registration_data.append(data)
    return registration_data


# In[6]:


def load_pre():
    pre_req=pd.read_excel('Pre-requisite_27-12-2018.xlsx',header=1)
    return pre_req


# In[7]:


def load_students():
    students_id=[]
    files = [f for f in listdir("./Input") if isfile(join("./Input", f))]
    files.sort()
    for i in files:
        if(i[0]!='.'):
            students_id.append(i)     
    return students_id


# In[8]:


students_id=load_students()
students_id


# In[9]:


#RETURNS LIST OF COURSES THAT HAVE PRE-REQ TO BE SATISFIED ALONG WITH THE REQUIRED COURSES

def pre_req_satisfied(courses,reg_data):
    dict1=dict()
    for course_no in range(len(courses[0])):
        
        pre_line1=pre_req.loc[((pre_req["Subject"])==courses[0][course_no])]
        if pre_line1.empty==True:
            pre_line1=pre_req.loc[((pre_req["Subject"])==" "+courses[0][course_no])]
        
        pre_line=pre_line1.loc[pre_line1["Catalog"]==(courses[1][course_no])]
        if pre_line.empty==True:
            pre_line=pre_line1.loc[(pre_line1["Catalog"]=="    "+courses[1][course_no])]
        if pre_line.empty==True:
            pre_line=pre_line1.loc[(pre_line1["Catalog"]=="   "+courses[1][course_no])]
        if pre_line.empty==True:
            pre_line=pre_line1.loc[(pre_line1["Catalog"]=="  "+courses[1][course_no])]
        if pre_line.empty==True:
            pre_line=pre_line1.loc[(pre_line1["Catalog"]==" "+courses[1][course_no])]
        if(pre_line.empty==True):
            continue
        
        course_name=pre_line.iloc[0]['Title']
        
        p1_1=pre_line.iloc[0]['preq1 subject']
        p1_2=pre_line.iloc[0]['preq1 catalog']
        p2_1=pre_line.iloc[0]['preq2 sub']
        p2_2=pre_line.iloc[0]['preq2 cat']
        p3_1=pre_line.iloc[0]['preq3 no']
        p3_2=pre_line.iloc[0]['preq3 cat']
        p4_1=pre_line.iloc[0]['preq4 no']
        p4_2=pre_line.iloc[0]['preq4 cat']

        p1=True
        p3=True
        p2=True
        p4=True
        
        c1=pre_line.iloc[0]['AND/OR']
        c2=pre_line.iloc[0]['AND/OR.1']
        c3=pre_line.iloc[0]['AND/OR.2']

        course1=pre_line.iloc[0]["pereq1 title "]
        course2=pre_line.iloc[0]["pereq2 title "]
        course3=pre_line.iloc[0]["pereq3 title "]
        course4=pre_line.iloc[0]["pereq4 title "]
              
        var1=True
        var2=True
        var3=True
        var4=True
        if(course1==course1):
            var1=check_done(course1,reg_data)
        if(course2==course2):
            var2=check_done(course2,reg_data)
        if(course3==course3):
            var3=check_done(course3,reg_data)
        if(course4==course4):
            var4=check_done(course4,reg_data)
        
        var=False
        if(course4==course4):
            if(c3=="OR"):
                var=(var4 or var3)
            elif(c3=="AND"):
                var=var4 and var3
            if(c2=="OR"):
                var=var or var2
            elif( c2=="AND"):
                var=var and var2
            if(c1=="OR"):
                var=var or var1
            elif( c1=="AND"):
                var=var and var1
            
        elif(course3==course3):
            if(c2=="OR"):
                var=var3 or var2
            elif(c2=="AND"):
                var=var3 and var2
            if(c1=="OR"):
                var=var or var1
            elif( c1=="AND"):
                var=var and var1
        elif(course2==course2):
            if(c1=="OR"):
                var=var2 or var1
            elif(c1=="AND"):
                var=var2 and var1
        elif(course1==course1):
            var=var1
        else:
            var=True
            
        not_completed=list()
        
        if(var!=True):
            if(not var1):
                not_completed.append(course1)
            if(not var2):
                not_completed.append(course2)
            if(not var3):
                not_completed.append(course3)
            if(not var4):
                not_completed.append(course4)
        
        
        if(len(not_completed)!=0):
            dict1[course_name]=not_completed
    return dict1


# In[10]:


def pre_req_future(courses):
    dict1=dict()
    for course_no in range(len(courses[0])):
        pre_line1=pre_req.loc[((pre_req["Subject"])==courses[0][course_no])]
        if pre_line1.empty==True:
            pre_line1=pre_req.loc[((pre_req["Subject"])==" "+courses[0][course_no])]
        
        pre_line=pre_line1.loc[pre_line1["Catalog"]==(courses[1][course_no])]
        if pre_line.empty==True:
            pre_line=pre_line1.loc[(pre_line1["Catalog"]=="    "+courses[1][course_no])]
        if pre_line.empty==True:
            pre_line=pre_line1.loc[(pre_line1["Catalog"]=="   "+courses[1][course_no])]
        if pre_line.empty==True:
            pre_line=pre_line1.loc[(pre_line1["Catalog"]=="  "+courses[1][course_no])]
        if pre_line.empty==True:
            pre_line=pre_line1.loc[(pre_line1["Catalog"]==" "+courses[1][course_no])]
        if(pre_line.empty==True):
            continue
        
        course_name=pre_line.iloc[0]['Title']
        
        fut1=pre_req.loc[(pre_req["pereq1 title "]==course_name)]
        fut1.append(pre_req.loc[(pre_req["pereq2 title "]==course_name)])
        fut1.append(pre_req.loc[(pre_req["pereq3 title "]==course_name)])
        fut1.append(pre_req.loc[(pre_req["pereq4 title "]==course_name)])
        
        lis_courses=list(fut1.iloc[:]["Title"])
        lis_courses = list(dict.fromkeys(lis_courses))
        dict1[course_name]=lis_courses
        
    return dict1 


# In[11]:


def credit_limit():
    pass


# In[12]:


def check_done(course,reg_data):
    ddf=reg_data.loc[((reg_data["Descr"])==course)]
    if(ddf.empty):
        return False
    else:
        grade=ddf.iloc[0]["Course Grade"]
        
        if grade in list(["A","B","C","D","E","A-","B-","C-",]):
            return True
    return False
    
    
    


# In[21]:


def write_file(pre_req,req_fut,student_id):
    with pd.ExcelWriter('./Output/'+student_id+'.xlsx') as writer: 
        workbook  = writer.book
        header_format = workbook.add_format({
        'bold': True,
        'text_wrap': True,
        'valign': 'top',
        'fg_color': '#48ceca',
        'border': 1})

        pre_req_pd=pd.DataFrame.from_dict(pre_req, orient='index')
        pre_req_pd.to_excel(writer, sheet_name='Sheet1',startcol=0,startrow=3)
        worksheet = writer.sheets['Sheet1']
        worksheet.write( 1,0, "PRE REQ NOT SATISFIED", header_format)
        worksheet.write( 2,0, "Course", header_format)
        
                
        pre_fut_pd=pd.DataFrame.from_dict(req_fut, orient='index')
        pre_fut_pd.to_excel(writer, sheet_name='Sheet1',startcol=0,startrow=12)
        worksheet = writer.sheets['Sheet1']
        worksheet.write( 10,0, "FUTURE COURSES REQ", header_format)
        worksheet.write( 11,0, "Current Course", header_format)
        
        
#         row=3
#         col=0
#         for course in pre_req.keys:
#             worksheet.write( row,0,course,header_format)
#             col=0
#             for i,pre in enumerate(pre_req[course]):
#                 worksheet.write( row,i+1,pre,header_format)


# In[14]:


registration_data=load_reg()

pre_req=load_pre()
# In[22]:


for student in students_id:
    loc = ("./Input/"+student) 
    # To open Workbook 
    wb = xlrd.open_workbook(loc) 
    sheet = wb.sheet_by_index(0) 
    
    student_id=student[:13]
    
    student_reg=registration_data.loc[((registration_data["Campus ID"])== student_id)]
    
    courses = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(1,sheet.nrows)]
    
    pre_not_satisfied=pre_req_satisfied(courses,student_reg)
    pre_future=pre_req_future(courses)
    
    write_file(pre_not_satisfied,pre_future,student_id)
    
# print(pre_not_satisfied)
# print(pre_future)


# In[17]:




