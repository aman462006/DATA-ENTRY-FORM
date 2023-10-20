from tkinter import *
from tkinter import ttk
from tkinter import messagebox
import os
import openpyxl

window = Tk()
window.title("Data Entry Form")

frame = Frame(window)
frame.pack()

#USER INFORMATION
user_info_frame=LabelFrame(frame,text="User Information")
user_info_frame.grid(row=0 , column=0,padx=20,pady=20)

first_name_label= Label(user_info_frame,text="First Name")
first_name_label.grid(row=0,column=1)
first_name_entry= Entry(user_info_frame,text="First Name")
first_name_entry.grid(row=1,column=1)

last_name_label= Label(user_info_frame,text="Last Name")
last_name_label.grid(row=0,column=2)
last_name_entry= Entry(user_info_frame,text="Last Name")
last_name_entry.grid(row=1,column=2)


Title=Label(user_info_frame,text="Title")
Title.grid(row=0,column=0)
Title_combobox=ttk.Combobox(user_info_frame , values=["Mr.","Mrs.","Dr."])
Title_combobox.grid(row=1,column=0)

Age=Label(user_info_frame,text="Standard")
Age.grid(row=2,column=0)
Age_spinbox=Spinbox(user_info_frame , from_=1, to= 12, wrap = True)
Age_spinbox.grid(row=3,column=0)

Nationality=Label(user_info_frame,text="Nationality")
Nationality.grid(row=2,column=1)
Nationality_combobox=ttk.Combobox(user_info_frame , values=["India","USA","Japan","Korea"])
Nationality_combobox.grid(row=3,column=1)

for widget in user_info_frame.winfo_children():
    widget.grid_configure(padx=10,pady=5)

#MARKS INFO

marks_frame=LabelFrame(frame)
marks_frame.grid(row=1 , column=0,sticky="news",padx=20,pady=20)

TestName=Label(marks_frame,text="Term")
TestName.grid(row=0,column=1)
TestName_combobox=ttk.Combobox(marks_frame, values=["Board","Half Yearly","Periodic Test","Weekly Test"])
TestName_combobox.grid(row=1,column=1,padx=100)

Subject_label=Label(marks_frame,text="Subject")
Subject_label.grid(row=2,column=0)
Marks_label=Label(marks_frame,text="Marks")
Marks_label.grid(row=2,column=1)
Max=Label(marks_frame,text="MAX")
Max.grid(row=2,column=2)

Subject1_combobox=ttk.Combobox(marks_frame, values=["Physics","Maths","Chemistry","PE","English","EVS","Gujarati","Hindi","Science","SS","Sanskrit","Accountancy","Economics","B.S."])
Subject1_combobox.grid(row=3,column=0)
Subject2_combobox=ttk.Combobox(marks_frame, values=["Physics","Maths","Chemistry","PE","English","EVS","Gujarati","Hindi","Science","SS","Sanskrit","Accountancy","Economics","B.S."])
Subject2_combobox.grid(row=4,column=0)
Subject3_combobox=ttk.Combobox(marks_frame, values=["Physics","Maths","Chemistry","PE","English","EVS","Gujarati","Hindi","Science","SS","Sanskrit","Accountancy","Economics","B.S."])
Subject3_combobox.grid(row=5,column=0)

Subject1_marks=Entry(marks_frame)
Subject1_marks.grid(row=3,column=1)
Subject2_marks=Entry(marks_frame)
Subject2_marks.grid(row=4,column=1)
Subject3_marks=Entry(marks_frame)
Subject3_marks.grid(row=5,column=1)

Subject1_maxmarks=Entry(marks_frame)
Subject1_maxmarks.grid(row=3,column=2)
Subject2_maxmarks=Entry(marks_frame)
Subject2_maxmarks.grid(row=4,column=2)
Subject3_maxmarks=Entry(marks_frame)
Subject3_maxmarks.grid(row=5,column=2)


for widget in marks_frame.winfo_children():
    widget.grid_configure(padx=10,pady=5)



def percent(x,y):
    x=int(x)
    y=int(y)
    per = (x/y)*100
    return round(per,2)


def grade_cal(per):
    if per <= 100 and per > 90:
        return "A1"
    elif per <= 90 and per >=80:
        return "A2"
    elif per < 80 and per >=70:
        return "B1"
    elif per < 70 and per >=60:
        return "B2"
    elif per < 60 and per >= 50:
        return "C"
    elif per < 50 and per >= 40:
        return "D"
    else:
        return "FAIL"
    
def result():
    if first_name_entry.get()=="" or last_name_entry.get() == "" or Title_combobox.get()=="" or Nationality_combobox=="":
        messagebox.showwarning(title="Error",message="All information is compulsory.")
    elif Subject1_combobox.get()==""or Subject2_combobox.get()==""or Subject3_combobox.get()==""or Subject3_marks.get()=="" or TestName_combobox.get()=="" or Subject2_marks.get()=="" or Subject1_marks.get()=="" or Subject3_maxmarks.get()=="" or Subject2_maxmarks.get()=="" or Subject1_maxmarks.get()=="" :
        messagebox.showwarning(title="Error",message="Subjects and marks are compulsory.")
    elif Subject1_combobox.get()==Subject2_combobox.get() or Subject1_combobox.get()==Subject3_combobox.get() or Subject3_combobox.get()==Subject2_combobox.get():
        messagebox.showwarning(title="Error",message="Please choose different subjects.")
    else:
        sub1mar=int(Subject1_marks.get())
        sub2mar=int(Subject2_marks.get())
        sub3mar=int(Subject3_marks.get())

        sub1maxmar=int(Subject1_maxmarks.get())
        sub2maxmar=int(Subject2_maxmarks.get())
        sub3maxmar=int(Subject3_maxmarks.get())   

        totalmar=sub1mar+sub2mar+sub3mar
        totalmaxmar=sub1maxmar+sub2maxmar+sub3maxmar
        totalper=percent(totalmar,totalmaxmar)
        totalgrade=grade_cal(totalper)

        percentagesub1=percent(sub1mar,sub1maxmar)
        percentagesub2=percent(sub2mar,sub2maxmar)                                              
        percentagesub3=percent(sub3mar,sub3maxmar)

                                                    
        gradesub1=grade_cal(percentagesub1)
        gradesub2=grade_cal(percentagesub2)
        gradesub3=grade_cal(percentagesub3)

        Total_label=Label(marks_frame,text="Total:")
        Total_label.grid(row=6,column=0)
        Totalmar_label=Label(marks_frame,text=str(totalmar))
        Totalmar_label.grid(row=6,column=1)
        Totalmaxmar_label=Label(marks_frame,text=str(totalmaxmar))
        Totalmaxmar_label.grid(row=6,column=2)
        Totalper_label=Label(marks_frame,text=str(totalper))
        Totalper_label.grid(row=6,column=3)
        Totalgrade_label=Label(marks_frame,text=str(totalgrade))
        Totalgrade_label.grid(row=6,column=4)

        grade_label=Label(marks_frame,text="Grade")
        grade_label.grid(row=2,column=4)
        Subject1_grade=Label(marks_frame,text= gradesub1)
        Subject1_grade.grid(row=3,column=4)
        Subject2_grade=Label(marks_frame,text= gradesub2)
        Subject2_grade.grid(row=4,column=4)
        Subject3_grade=Label(marks_frame,text= gradesub3)
        Subject3_grade.grid(row=5,column=4)
        
        percent_label=Label(marks_frame,text="Percent")
        percent_label.grid(row=2,column=3)
        Subject1_percent=Label(marks_frame,text = str(percentagesub1))
        Subject1_percent.grid(row=3,column=3)
        Subject2_percent=Label(marks_frame,text = str(percentagesub2))
        Subject2_percent.grid(row=4,column=3)
        Subject3_percent=Label(marks_frame,text = str(percentagesub3))
        Subject3_percent.grid(row=5,column=3)

        for widget in marks_frame.winfo_children():
            widget.grid_configure(padx=10,pady=5)


        

    
        filepath = "S:\DATA ENTRY\data.xlsx"
        if not os.path.exists(filepath):
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            heading = ["Title","First Name","Last Name","Standard","Nationality","Term","Subject","Marks","MAX","Percent","Grade"]
            sheet.append(heading)
            workbook.save(filepath)

        workbook = openpyxl.load_workbook(filepath)
        sheet = workbook.active
        sheet.append([Title_combobox.get(),first_name_entry.get(),last_name_entry.get(),Age_spinbox.get(),Nationality_combobox.get(),TestName_combobox.get(),Subject1_combobox.get(),sub1mar,sub1maxmar,percentagesub1,gradesub1])
        sheet.append(["","","","","","",Subject2_combobox.get(),sub2mar,sub2maxmar,percentagesub2,gradesub2])
        sheet.append(["","","","","","",Subject3_combobox.get(),sub3mar,sub3maxmar,percentagesub3,gradesub3])
        sheet.append(["","","","","","","Total:",totalmar,totalmaxmar,totalper,totalgrade])
        workbook.save(filepath)


    
Enter_frame=LabelFrame(frame)
Enter_frame.grid(row=2 , column=0,sticky="news",padx=20,pady=20)
Label(Enter_frame,text="").grid(row=0,column=0,padx=85)
Enter_Button=Button(Enter_frame,text="ENTER",padx=50,command=lambda: result())
Enter_Button.grid(row=0,column=1,pady=10)






window.mainloop()
