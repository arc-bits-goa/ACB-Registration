
# coding: utf-8

# In[3]:


import numpy as np
import pandas as pd
import csv


# In[8]:


# global degree_data=pd.DataFrame
# global degree_courses=pd.DataFrame
# global degree_students=pd.DataFrame



def load_files():
    from os import listdir
    from os.path import isfile, join

    files = [f for f in listdir("./data") if isfile(join("./data", f))]
    files.sort()
    files
    registration_data=pd.DataFrame()
    for i in files:
        data = pd.read_excel("./data/"+i,"sheet1",header=1)
        registration_data=registration_data.append(data)
    return registration_data 


# In[5]:


def units():
    excel_file = 'units.xls'
    units1 = pd.read_excel(excel_file,header=None)
    unit_dic=dict(zip(units1[0],units1[2]))
    return unit_dic


# In[6]:


def minor_load(degree):
        ############ Store Students in each minor course ###############
    
    minor_students=pd.read_excel("Minor Student list.xls",header=None)
    minor_students.columns=['ID','Name',"Minor"]
    #minor_students
    
    degree_students= minor_students.ix[(minor_students['Minor']== degree +' Minor'), ['ID','Name']]
    #print("sdfio",degree_students)
    CGPA_create(degree_students) 
    #print(minor_students)
    #print(degree_students)
    
    
    degree_courses=pd.read_csv(degree+".csv",header=None)
    degree_courses.rows=['Core/Elective','Course No',"Equivalent","Pair No"]
    
    #degree_courses
    
    
    registration_data=load_files()
    #registration_data=['Core/Elective','Course No',"Equivalent","Pair No"]
    registration_data["Course"] = registration_data["Subject"]+" "+registration_data["Catalog No."]

    #registration_data 
    #print(registration_data)
    return degree_students,degree_courses,registration_data


# In[7]:


def CGPA_create(students):
    students=students.values.tolist()
    #print(np.hape(students))
    global CGPA_courses
    for i in range(len(students)):
        temp=[]
        temp.append(students[i][0])
        CGPA_courses.append(temp)
    #print(CGPA_courses)


# In[48]:


def CGPA(degree,degree_data):
    grade_dic={"A":10,"A-":9,"B":8,"B-":7,"C":6,"C-":5,"D":4,"E":2}
    excel_file = 'units.xls'
    units1 = pd.read_excel(excel_file,header=None)
    unit_dic=dict(zip(units1[0],units1[2]))
    #var represtes total number of courses to be considered for CG of that particular degree
    var=6
    if(degree=='PEP'):
        var=7
    cnt=0
    #print(len(CGPA_courses))
    for i,data in degree_data.iterrows():
    #for i in range(len(CGPA_courses)):
        #print("a")
        #print(i)
        grades_done=0;
        total_units=0;
        #print(len(CGPA_cou/rses[i]))
        for j in range(1,len(CGPA_courses[cnt])):
            if (j==var):
                break
            units=unit_dic[CGPA_courses[cnt][j]]
            grade=degree_data.iloc[cnt][CGPA_courses[cnt][j]]
            try:
                grades_done+=grade_dic[grade]*units
            except:
                print(data)
            total_units+=units
        cg=0
        if(total_units>0):
            cg=grades_done/total_units
            #print(cg)
            #print(cg)
        degree_data.at[i, "CG"] = cg
        cnt+=1
    return degree_data


# In[9]:


def minor_main(degree,degree_students,degree_courses,registration_data):
    degree_data = pd.DataFrame(data=' ',index=degree_students,columns=degree_courses.T[1])
    global CGPA_courses
    # print(degree_data)
    for i,data in degree_data.iterrows():
        student_id = i[0]
        student=pd.DataFrame(registration_data.loc[((registration_data["Campus ID"])== (student_id))])
        a1=-1
        # abc=student.shape
        # print(abc)
        # print(len(student["Course"].values))
        for course in student["Course"].values:
            # print(course)
            a1+=1
            # if (a1>abc[0]):
            #     print('wrong')
            grade=student.iloc[a1]['Course Grade']
            
            if course in list(degree_courses.iloc[1]):
            
                # grade=student.loc[(student["Course"])==course]["Course Grade"].values
                if not(grade[0]=='W' or grade[0]=='NC' or grade[0]=='N'):
                    for a in range(len(CGPA_courses)):
                        if (CGPA_courses[a][0]==student_id):
                            if(course not in CGPA_courses[a]):
                                CGPA_courses[a].append(course)
                # try:
                #     if (grade[0]=="W" or grade[0]=="NC"):
                #         # print(grade[0],a,student_id)
                #         student.drop([a])
                #         a-=1
                # except:
                #     pass
                # if(degree_data.at[i, course]=='W' or degree_data.at[i, course]=='NC'):
                    # print('fault')
                #finance_data[student][course]=grade
                if not (degree_data.at[i, course] == None or degree_data.at[i, course] =='' or degree_data.at[i, course] == ' '):
                    if not(grade[0]=='W' or grade[0]=='NC' or grade[0]=='N'):
                        degree_data.at[i, course] = grade[0]
                else:
                    degree_data.at[i, course] = grade[0]
                    # if(grade[0]=='NC' or grade[0]=='W' or grade[0]=='E' ):
                    # print(grade[0]) 
                    
                #df[df.Letters=='C'].Letters.item()

#?degree_data
    #print(CGPA_courses)
    
    unit_dic=units()
    
    grade_dic={"A":10,"A-":9,"B":8,"B-":7,"C":6,"C-":5,"D":4,"E":2}

    for i,data in degree_data.iterrows():
        student_id=i[0]
        core_courses=0
        elective_courses=0
        total_credits=0
        total_grade=0
        equivalent_done = []
        #student.loc[(student["Course"])==course]["Course Grade"].values
        for temp,course in enumerate(degree_courses.iloc[1]):
            course_grade=degree_data.at[i, course]
            # print(course_grade)
            if (degree_courses.at[2,temp]=="No" and (course_grade in grade_dic.keys())):
                if(degree_courses.at[0,temp]=="Core"):
                    core_courses+=1
                    total_credits+=unit_dic[course]
                    total_grade+=unit_dic[course]*grade_dic[course_grade]                
                elif (degree_courses.at[0,temp]=="Elective"):
                    elective_courses+=1
                    total_credits+=unit_dic[course]
                    total_grade+=unit_dic[course]*grade_dic[course_grade]
            elif (degree_courses.at[2,temp]=="Yes" ):
                if(degree_courses.at[3,temp] in equivalent_done):
                    continue
                for tempcourseno,details in degree_courses.iteritems():
                    # print(tempcourseno,details)
                    # print(details[3])
                    # print(details[3])
                    # print(degree_courses.at[3,temp])
                    # print(tempcourseno,temp)
                    # print()
                    if(tempcourseno!=temp and details[3]==degree_courses.at[3,temp]):
                        # print('sd')
                        # print(degree_data.at[i,details[1]])
                        tempcoursegrade = degree_data.at[i,details[1]]
                        # print(tempcoursegrade)
                        # if(tempcoursegrade!=' '):
                        #     print(tempcoursegrade)
                        if (tempcoursegrade in grade_dic.keys()):
                            # print('sd')
                            # print(tempcoursegrade)
                            equivalent_done += [details[3]]
                            # finalgrade = max(grade_dic[course_grade], grade_dic[tempcoursegrade])
                            if(details[0]=='Elective'):
                                elective_courses+=1
                            else:
                                core_courses+=1
#                             total_credits+=unit_dic[course]
#                             total_grade+=finalgrade
        # break
#         if(total_credits>0):
#             cg=total_grade/total_credits        
#         else:
#             cg=0
            
        if(degree=='PEP'):
            a=3
            b=3
        elif(degree=='Physics'):
            a=3
            b=2
        elif(degree=='English'):
            a=2
            b=3
        elif(degree=='Finance'):
            a=2
            b=3
        else:
            a=2
            b=3
            
        degree_data.at[i, "#REQD_Core_courses"] = a
        degree_data.at[i, "#REQD_Elective_courses"] = b

        degree_data.at[i, "#DONE_Core Courses"] = core_courses
        degree_data.at[i, "#DONE_Elective Courses"] = elective_courses

        degree_data.at[i, "#REMAINING_Core_Courses"] = degree_data.at[i, "#REQD_Core_courses"]-degree_data.at[i, "#DONE_Core Courses"]
        degree_data.at[i, "#REMAINING_Elective Courses"] = degree_data.at[i, "#REQD_Elective_courses"]-degree_data.at[i, "#DONE_Elective Courses"]

#         degree_data.at[i, "CG"] = cg
        #print(core_courses,elective_courses,cg)  

        # print(degree_data)
    return degree_data


# In[10]:


def minor_write(degreearr,degree_dataarr,n):
    with pd.ExcelWriter('Minor_Data.xlsx') as writer: 
        workbook  = writer.book
        header_format = workbook.add_format({
        'bold': True,
        'text_wrap': True,
        'valign': 'top',
        'fg_color': '#48ceca',
        'border': 1})

        for i in range(n):
            # print(degree_dataarr[i])
            degree_dataarr[i].to_excel(writer, sheet_name=degreearr[i]+" Minor",startcol=1,startrow=2)
            worksheet = writer.sheets[degreearr[i]+' Minor']
            worksheet.write( 2,0, "ID NO", header_format)
            worksheet.write( 2,1, "NAME", header_format)
            for row_num, value in enumerate(degree_dataarr[i].index.values):
                worksheet.write( row_num + 3,0, value[0], header_format)
                worksheet.write( row_num + 3,1, value[1], header_format)


# In[11]:


def main(degree):
    degree_students,degree_courses,registration_data=minor_load(degree)
    degree_students_cp=degree_students
    degree_data1=minor_main(degree,degree_students,degree_courses,registration_data)
    #print(degree_data1)
    return degree,degree_data1,degree_students


# In[2]:


import sys
lell=len(sys.argv)
degreearr=[]
degree_dataarr = []
for i in range(1,lell):
    CGPA_courses=[]
    degree,degree_data,degree_students = main(sys.argv[i])
    degree_data=CGPA(degree,degree_data)
    degreearr.append(degree)
    degree_dataarr.append(degree_data)
    #degreestudnts=CGPA_courses(degree_students)

minor_write(degreearr,degree_dataarr,lell-1)

