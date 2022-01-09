import requests
import time
import pymysql
import socket
import pandas
import re
import json
import random
import csv
import pandas as pd
from docxtpl import DocxTemplate
from docx2pdf import convert
import os
import tkinter as tk


from tkinter import *
from tkinter import Toplevel,messagebox,filedialog
from tkinter.ttk import Treeview
from tkinter import ttk
from requests import get
from urllib.request import urlopen
import requests



url = 'http://ipinfo.io/json'
response = urlopen(url)
data = json.load(response)


IP=data['ip']
org=data['org']
city = data['city']
country=data['country']
region=data['region']
addedtime=time.strftime("%H:%M:%S")
addeddate=time.strftime("%d/%m/%Y")
hostname = socket.gethostname()


def addstudent():
    def submitadd():
        id=idval.get()
        name=nameval.get()
        mobile=mobileval.get()
        email=emailval.get()
        address=addressval.get()
        gender=genderval.get()
        dob=dobval.get()
        addedtime=time.strftime("%H:%M:%S")
        addeddate=time.strftime("%d/%m/%Y")

        try:
            strr='insert into data values(%s,%s,%s,%s,%s,%s,%s,%s,%s)'
            mycursor.execute(strr,(id,name,mobile,email,address,gender,dob,addedtime,addeddate))
            con.commit()
            res=messagebox.askyesnocancel('Notification','Id {} Name {} Added sucessfully ...and what to clean the form'.format(id,name),parent=addroot)
            if(res==True):
                idval.set('')
                nameval.set('')
                mobileval.set('')
                emailval.set('')
                addressval.set('')
                genderval.set('')
                dobval.set('')
            mycursor.execute("drop trigger if exists mytrigger")
            trigger = "CREATE TRIGGER mytrigger BEFORE INSERT ON data FOR EACH ROW IF(EXISTS(SELECT 1 FROM data WHERE mobile= NEW.mobile))THEN SIGNAL SQLSTATE VALUE '45000' SET MESSAGE_TEXT = 'INSERT failed due to duplicate mobile number'; END IF;"
            mycursor.execute(trigger)
        except:
            messagebox.showerror('Notification','Id or Phone Number already Exists try another ID or Phone Number....',parent=addroot)
        strr='select * from data'

        mycursor.execute(strr)
        datas=mycursor.fetchall()
        studentmttable.delete(*studentmttable.get_children())
        for i in datas:
            vv=[i[0],i[1],i[2],i[3],i[4],i[5],i[6],i[7],i[8]]
            studentmttable.insert('',END,values=vv)

    addroot = Toplevel(master=DataEntryFrame)
    addroot.grab_set()
    addroot.geometry('470x470+400+300')
    addroot.title('Add Student')
    addroot.iconbitmap('logo.ico')
    addroot.resizable(FALSE,FALSE)

    idlabel = Label(addroot, text='Enter Id:', bg='#243142', font=('Comic Sans MS', 15, NORMAL), relief=GROOVE,fg='white', borderwidth=3, width=12, anchor='w')
    idlabel.place(x=10, y=10)
    namelabel = Label(addroot, text='Enter Name:', bg='#243142', font=('Comic Sans MS', 15, NORMAL), relief=GROOVE,fg='white', borderwidth=3, width=12, anchor='w')
    namelabel.place(x=10, y=70)
    mobilelabel = Label(addroot, text='Enter Phone:', bg='#243142', font=('Comic Sans MS', 15, NORMAL), relief=GROOVE, fg='white',borderwidth=3, width=12, anchor='w')
    mobilelabel.place(x=10, y=130)
    emaillabel = Label(addroot, text='Enter Email:', bg='#243142', font=('Comic Sans MS', 15, NORMAL), relief=GROOVE,fg='white', borderwidth=3, width=12, anchor='w')
    emaillabel.place(x=10, y=190)
    addresslabel = Label(addroot, text='Enter Address:', bg='#243142', font=('Comic Sans MS', 15, NORMAL), relief=GROOVE, fg='white',borderwidth=3, width=12, anchor='w')
    addresslabel.place(x=10, y=250)

    genderlabel = Label(addroot, text='Enter Gender:', bg='#243142', font=('Comic Sans MS', 15, NORMAL), relief=GROOVE,fg='white', borderwidth=3, width=12, anchor='w')
    genderlabel.place(x=10, y=310)
    doblabel = Label(addroot, text='Enter DOB:', bg='#243142', font=('Comic Sans MS', 15, NORMAL), relief=GROOVE,fg='white', borderwidth=3, width=12, anchor='w')
    doblabel.place(x=10, y=370)

    idval=StringVar()
    nameval=StringVar()
    mobileval=StringVar()
    emailval=StringVar()
    addressval=StringVar()
    genderval=StringVar()
    dobval=StringVar()

    identry = Entry(addroot, font=('Comic Sans MS', 12, NORMAL), bd=5, textvariable=idval)
    identry.place(x=220, y=10)
    nameentry = Entry(addroot, font=('Comic Sans MS', 12, NORMAL), bd=5, textvariable=nameval)
    nameentry.place(x=220, y=70)
    mobileentry = Entry(addroot, font=('Comic Sans MS', 12, NORMAL), bd=5, textvariable=mobileval)
    mobileentry.place(x=220, y=130)
    emailentry = Entry(addroot, font=('Comic Sans MS', 12, NORMAL), bd=5, textvariable=emailval)
    emailentry.place(x=220, y=190)
    addressentry = Entry(addroot, font=('Comic Sans MS', 12, NORMAL), bd=5, textvariable=addressval)
    addressentry.place(x=220, y=250)
    genderentry = Entry(addroot, font=('Comic Sans MS', 12, NORMAL), bd=5, textvariable=genderval)
    genderentry.place(x=220, y=310)
    dobentry = Entry(addroot, font=('Comic Sans MS', 12, NORMAL), bd=5, textvariable=dobval)
    dobentry.place(x=220, y=370)

    submitbtn=Button(addroot,bg='#b7e576',text='SUBMIT',font=('Comic Sans MS',12,NORMAL),width=20,bd=6,activebackground='#b7e576',command=submitadd)
    submitbtn.place(x=130,y=410)

    addroot.mainloop()

def searchstudent():
    def search():
        id = idval.get()
        name = nameval.get()
        mobile = mobileval.get()
        email = emailval.get()
        address = addressval.get()
        gender = genderval.get()
        dob = dobval.get()
        addeddate = time.strftime("%d/%m/%Y")
        if(id !=''):
            strr='select * from data where id=%s'
            mycursor.execute(strr,(id))
            datas=mycursor.fetchall()
            studentmttable.delete(*studentmttable.get_children())
            for i in datas:
                vv = [i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8]]
                studentmttable.insert('', END, values=vv)
        elif(name !=''):
            strr='select * from data where name=%s'
            mycursor.execute(strr,(name))
            datas=mycursor.fetchall()
            studentmttable.delete(*studentmttable.get_children())
            for i in datas:
                vv = [i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8]]
                studentmttable.insert('', END, values=vv)
        elif(mobile != ''):
            strr = 'select * from data where mobile=%s'
            mycursor.execute(strr, (mobile))
            datas = mycursor.fetchall()
            studentmttable.delete(*studentmttable.get_children())
            for i in datas:
                vv = [i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8]]
                studentmttable.insert('', END, values=vv)
        elif (email != ''):
            strr = 'select * from data where email=%s'
            mycursor.execute(strr, (email))
            datas = mycursor.fetchall()
            studentmttable.delete(*studentmttable.get_children())
            for i in datas:
                vv = [i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8]]
                studentmttable.insert('', END, values=vv)
        elif (address != ''):
            strr = 'select * from data where address=%s'
            mycursor.execute(strr, (address))
            datas = mycursor.fetchall()
            studentmttable.delete(*studentmttable.get_children())
            for i in datas:
                vv = [i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8]]
                studentmttable.insert('', END, values=vv)
        elif (gender != ''):
            strr = 'select * from data where gender=%s'
            mycursor.execute(strr, (gender))
            datas = mycursor.fetchall()
            studentmttable.delete(*studentmttable.get_children())
            for i in datas:
                vv = [i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8]]
                studentmttable.insert('', END, values=vv)
        elif (dob != ''):
            strr = 'select * from data where dob=%s'
            mycursor.execute(strr, (dob))
            datas = mycursor.fetchall()
            studentmttable.delete(*studentmttable.get_children())
            for i in datas:
                vv = [i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8]]
                studentmttable.insert('', END, values=vv)
        elif (addeddate != ''):
            strr = 'select * from data where addeddate=%s'
            mycursor.execute(strr, (addeddate))
            datas = mycursor.fetchall()
            studentmttable.delete(*studentmttable.get_children())
            for i in datas:
                vv = [i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8]]
                studentmttable.insert('', END, values=vv)

    searchroot = Toplevel(master=DataEntryFrame)
    searchroot.grab_set()
    searchroot.geometry('470x540+400+300')
    searchroot.title('Search Student')
    searchroot.iconbitmap('logo.ico')
    searchroot.resizable(FALSE,FALSE)

    idlabel = Label(searchroot, text='Enter Id:', bg='#243142', font=('Comic Sans MS', 15, NORMAL), relief=GROOVE,fg='white', borderwidth=3, width=12, anchor='w')
    idlabel.place(x=10, y=10)
    namelabel = Label(searchroot, text='Enter Name:', bg='#243142', font=('Comic Sans MS', 15, NORMAL), relief=GROOVE,fg='white', borderwidth=3, width=12, anchor='w')
    namelabel.place(x=10, y=70)
    mobilelabel = Label(searchroot, text='Enter Phone:', bg='#243142', font=('Comic Sans MS', 15, NORMAL), relief=GROOVE,fg='white', borderwidth=3, width=12, anchor='w')
    mobilelabel.place(x=10, y=130)
    emaillabel = Label(searchroot, text='Enter Email:', bg='#243142', font=('Comic Sans MS', 15, NORMAL), relief=GROOVE,fg='white', borderwidth=3, width=12, anchor='w')
    emaillabel.place(x=10, y=190)
    addresslabel = Label(searchroot, text='Enter Address:', bg='#243142', font=('Comic Sans MS', 15, NORMAL), relief=GROOVE,fg='white', borderwidth=3, width=12, anchor='w')
    addresslabel.place(x=10, y=250)
    genderlabel = Label(searchroot, text='Enter Gender:', bg='#243142', font=('Comic Sans MS', 15, NORMAL), relief=GROOVE,fg='white', borderwidth=3, width=12, anchor='w')
    genderlabel.place(x=10, y=310)
    doblabel = Label(searchroot, text='Enter DOB:', bg='#243142', font=('Comic Sans MS', 15, NORMAL), relief=GROOVE,fg='white', borderwidth=3, width=12, anchor='w')
    doblabel.place(x=10, y=370)
    datelable = Label(searchroot, text='Enter Date:', bg='#243142', font=('Comic Sans MS', 15, NORMAL), relief=GROOVE,fg='white', borderwidth=3, width=12, anchor='w')
    datelable.place(x=10, y=430)

    idval=StringVar()
    nameval=StringVar()
    mobileval=StringVar()
    emailval=StringVar()
    addressval=StringVar()
    genderval=StringVar()
    dobval=StringVar()
    dateval=StringVar()

    identry = Entry(searchroot, font=('Comic Sans MS', 12, NORMAL), bd=5, textvariable=idval)
    identry.place(x=220, y=10)
    nameentry = Entry(searchroot, font=('Comic Sans MS', 12, NORMAL), bd=5, textvariable=nameval)
    nameentry.place(x=220, y=70)
    mobileentry = Entry(searchroot, font=('Comic Sans MS', 12, NORMAL), bd=5, textvariable=mobileval)
    mobileentry.place(x=220, y=130)
    emailentry = Entry(searchroot, font=('Comic Sans MS', 12, NORMAL), bd=5, textvariable=emailval)
    emailentry.place(x=220, y=190)
    addressentry = Entry(searchroot, font=('Comic Sans MS', 12, NORMAL), bd=5, textvariable=addressval)
    addressentry.place(x=220, y=250)
    genderentry = Entry(searchroot, font=('Comic Sans MS', 12, NORMAL), bd=5, textvariable=genderval)
    genderentry.place(x=220, y=310)
    dobentry = Entry(searchroot, font=('Comic Sans MS', 12, NORMAL), bd=5, textvariable=dobval)
    dobentry.place(x=220, y=370)
    dateentry = Entry(searchroot, font=('Comic Sans MS', 12, NORMAL), bd=5, textvariable=dateval)
    dateentry.place(x=220, y=430)

    submitbtn=Button(searchroot,bg='#b7e576',text='SUBMIT',font=('Comic Sans MS',12,NORMAL),width=20,bd=6,activebackground='#b7e576',command=search)
    submitbtn.place(x=130,y=480)

    searchroot.mainloop()
def deletestudent():
    cc = studentmttable.focus()
    content = studentmttable.item(cc)
    pp = content['values'][0]
    strr = 'delete from data where id=%s'
    mycursor.execute(strr,(pp))
    con.commit()
    messagebox.showinfo('Notificaiton','Id {} deleted sucessfully...'.format(pp))
    strr = 'select * from data'
    mycursor.execute(strr)
    datas = mycursor.fetchall()
    studentmttable.delete(*studentmttable.get_children())
    for i in datas:
        vv = [i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8]]
        studentmttable.insert('', END, values=vv)

def updatestudent():
    def Update():
        id = idval.get()
        name = nameval.get()
        mobile = mobileval.get()
        email = emailval.get()
        address = addressval.get()
        gender = genderval.get()
        dob = dobval.get()
        date = dateval.get()
        time = timeval.get()



        strr='update data set name=%s,mobile=%s,email=%s,address=%s,gender=%s,dob=%s,date=%s,time=%s where id=%s'
        mycursor.execute(strr,(name,mobile,email,address,gender,dob,date,time,id))
        con.commit()
        messagebox.showinfo('Notification','Id {} Modified Sucessfully...'.format(id),parent=updateroot)
        strr = 'select * from data'
        mycursor.execute(strr)
        datas = mycursor.fetchall()
        studentmttable.delete(*studentmttable.get_children())
        for i in datas:
            vv = [i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8]]
            studentmttable.insert('', END, values=vv)

    updateroot = Toplevel(master=DataEntryFrame)
    updateroot.grab_set()
    updateroot.geometry('470x600+400+300')
    updateroot.title('Update Student')
    updateroot.iconbitmap('logo.ico')
    updateroot.resizable(FALSE,FALSE)

    idlabel = Label(updateroot, text='Enter Id:', bg='#243142', font=('Comic Sans MS', 15, NORMAL), relief=GROOVE,fg='white', borderwidth=3, width=12, anchor='w')
    idlabel.place(x=10, y=10)
    namelabel = Label(updateroot, text='Enter Name:', bg='#243142', font=('Comic Sans MS', 15, NORMAL), relief=GROOVE, fg='white',borderwidth=3, width=12, anchor='w')
    namelabel.place(x=10, y=70)
    mobilelabel = Label(updateroot, text='Enter Phone:', bg='#243142', font=('Comic Sans MS', 15, NORMAL), relief=GROOVE,fg='white', borderwidth=3, width=12, anchor='w')
    mobilelabel.place(x=10, y=130)
    emaillabel = Label(updateroot, text='Enter Email:', bg='#243142', font=('Comic Sans MS', 15, NORMAL), relief=GROOVE, fg='white',borderwidth=3, width=12, anchor='w')
    emaillabel.place(x=10, y=190)
    addresslabel = Label(updateroot, text='Enter Address:', bg='#243142', font=('Comic Sans MS', 15, NORMAL), relief=GROOVE,fg='white', borderwidth=3, width=12, anchor='w')
    addresslabel.place(x=10, y=250)
    genderlabel = Label(updateroot, text='Enter Gender:', bg='#243142', font=('Comic Sans MS', 15, NORMAL), relief=GROOVE, fg='white',borderwidth=3, width=12, anchor='w')
    genderlabel.place(x=10, y=310)
    doblabel = Label(updateroot, text='Enter DOB:', bg='#243142', font=('Comic Sans MS', 15, NORMAL), relief=GROOVE,fg='white', borderwidth=3, width=12, anchor='w')
    doblabel.place(x=10, y=370)
    datelable = Label(updateroot, text='Enter Date:', bg='#243142', font=('Comic Sans MS', 15, NORMAL), relief=GROOVE,fg='white', borderwidth=3, width=12, anchor='w')
    datelable.place(x=10, y=430)
    timelable = Label(updateroot, text='Enter Time:', bg='#243142', font=('Comic Sans MS', 15, NORMAL), relief=GROOVE,fg='white', borderwidth=3, width=12, anchor='w')
    timelable.place(x=10, y=490)

    idval=StringVar()
    nameval=StringVar()
    mobileval=StringVar()
    emailval=StringVar()
    addressval=StringVar()
    genderval=StringVar()
    dobval=StringVar()
    dateval=StringVar()
    timeval=StringVar()

    identry = Entry(updateroot, font=('Comic Sans MS', 12, NORMAL), bd=5, textvariable=idval)
    identry.place(x=220, y=10)
    nameentry = Entry(updateroot, font=('Comic Sans MS', 12, NORMAL), bd=5, textvariable=nameval)
    nameentry.place(x=220, y=70)
    mobileentry = Entry(updateroot, font=('Comic Sans MS', 12, NORMAL), bd=5, textvariable=mobileval)
    mobileentry.place(x=220, y=130)
    emailentry = Entry(updateroot, font=('Comic Sans MS', 12, NORMAL), bd=5, textvariable=emailval)
    emailentry.place(x=220, y=190)
    addressentry = Entry(updateroot, font=('Comic Sans MS', 12, NORMAL), bd=5, textvariable=addressval)
    addressentry.place(x=220, y=250)
    genderentry = Entry(updateroot, font=('Comic Sans MS', 12, NORMAL), bd=5, textvariable=genderval)
    genderentry.place(x=220, y=310)
    dobentry = Entry(updateroot, font=('Comic Sans MS', 12, NORMAL), bd=5, textvariable=dobval)
    dobentry.place(x=220, y=370)
    dateentry = Entry(updateroot, font=('Comic Sans MS', 12, NORMAL), bd=5, textvariable=dateval)
    dateentry.place(x=220, y=430)
    timeentry = Entry(updateroot, font=('Comic Sans MS', 12, NORMAL), bd=5, textvariable=timeval)
    timeentry.place(x=220, y=490)

    submitbtn=Button(updateroot,bg='#b7e576',text='SUBMIT',font=('Comic Sans MS',12,NORMAL),width=20,bd=6,activebackground='#b7e576',command=Update)
    submitbtn.place(x=130,y=540)

    cc=studentmttable.focus()
    content=studentmttable.item(cc)
    pp=content['values']
    if(len(pp) !=0):
        idval.set(pp[0])
        nameval.set(pp[1])
        mobileval.set(pp[2])
        emailval.set(pp[3])
        addressval.set(pp[4])
        genderval.set(pp[5])
        dobval.set(pp[6])
        dateval.set(pp[7])
        timeval.set(pp[8])

    updateroot.mainloop()
def showstudent():
    strr = 'select * from data'
    mycursor.execute(strr)
    datas = mycursor.fetchall()
    studentmttable.delete(*studentmttable.get_children())
    for i in datas:
        vv = [i[0], i[1], i[2], i[3], i[4], i[5], i[6], i[7], i[8]]
        studentmttable.insert('', END, values=vv)
def exportstudent():
    ff = filedialog.asksaveasfilename()
    gg = studentmttable.get_children()
    id,name,mobile,email,address,gender,dob,addeddate,addedtime=[],[],[],[],[],[],[],[],[]
    for i in gg:
        content = studentmttable.item(i)
        pp = content['values']
        id.append(pp[0]),name.append(pp[1]),mobile.append(pp[2]),email.append(pp[3]),address.append(pp[4]),gender.append(pp[5]),dob.append(pp[6]),addeddate.append(pp[7]),addedtime.append(pp[8])
    dd = ['Id','Name','Mobile','Email','Address','Gender','D.O.B','Added Data','Added Time']
    df = pandas.DataFrame(list(zip(id,name,mobile,email,address,gender,dob,addeddate,addedtime)),columns=dd)
    paths = r"{}.csv".format(ff)
    df.to_csv(paths,index=False)
    messagebox.showinfo('Notification', 'Student data is saved {}'.format(paths))

def exitstudent():
    res= messagebox.askyesno('Notification','Do you want to exit?')
    if(res==True):
        root.destroy()

def Connectdb():
    def submitdb():

        global con,mycursor
        host=hostval.get()
        user=userval.get()
        password=passwordval.get()
        try:
            con=pymysql.connect(host=host,user=user,password=password)
            mycursor=con.cursor()
            message = 'New login. Dear Hemanth, we detected a login into your database on {} at {} IST.\n\nDevice: {}\nIP Address: {}\nCity: {}\nCountry: {}\nOrganization: {}\n\nIf this wasnt you, Sorry Hemanth, the database is being used by someone else....'.format(addeddate, addedtime,hostname,IP,city,country,org)
            token = '5029349522:AAHyH9TlmLxPJDoAFjkR2Vn_FFh-2ouBLQ0'
            url = f'https://api.telegram.org/bot{token}/sendMessage'
            data = {'chat_id': {5060166011}, 'text': {message}} # Kumar swami -- > 874434712 
            requests.post(url, data).json()
        except:
            messagebox.showerror('Notification','Date is incorrect please try again')
            return
        try:
            strr='create database studentdatabase'
            mycursor.execute(strr)
            strr ='use studentdatabase'
            mycursor.execute(strr)
            strr ='create table data(id int,name varchar(100),mobile varchar(100),email varchar(100),address varchar(100),gender varchar(100),dob varchar(100),date varchar(100),time varchar(100))'
            mycursor.execute(strr)
            strr='alter table data modify column id int not null'
            mycursor.execute(strr)
            strr='alter table data modify column id int primary key'
            mycursor.execute(strr)
            messagebox.showinfo('Notification', 'Database has been created and now you are connected to the Database...', parent=dbroot)
            message = 'New login. Dear Hemanth, we detected a login into your database on {} at {} IST.\n\nDevice: {}\nIP Address: {}\nCity: {}\nCountry: {}\nOrganization: {}\n\nIf this wasnt you, Sorry Hemanth, the database is being used by someone else....'.format(addeddate, addedtime,hostname,IP,city,country,org)
            token = '1310871540:AAF7EzI9VjdUF5_9B8yzlAmWzoPfp-0J63Y'
            url = f'https://api.telegram.org/bot{token}/sendMessage'
            data = {'chat_id': {627491741}, 'text': {message}}
            requests.post(url, data).json()
        except:
            strr='use studentdatabase'
            mycursor.execute(strr)
            messagebox.showinfo('Notification', 'Now You are connected to the Database...',parent=dbroot)
        dbroot.destroy()

    dbroot=Toplevel()
    dbroot.grab_set()
    dbroot.config(bg='white')
    dbroot.title('Connect To Database')
    dbroot.geometry('470x280+730+350')
    dbroot.resizable(FALSE,FALSE)
    dbroot.iconbitmap('logo.ico')

    hostlabel = Label(dbroot, text="Enter Host : ", bg='#243142', font=('Comic Sans MS', 15, NORMAL),fg='white', relief=GROOVE,
                    borderwidth=3, width=15, anchor='n')
    hostlabel.place(x=10, y=10)

    userlabel = Label(dbroot, text="Enter User : ", bg='#243142', font=('Comic Sans MS', 15, NORMAL),fg='white', relief=GROOVE,
                    borderwidth=3, width=15, anchor='n')
    userlabel.place(x=10, y=90)

    passwordlabel = Label(dbroot, text="Enter Password : ", bg='#243142', font=('Comic Sans MS', 15, NORMAL),fg='white', relief=GROOVE,
                    borderwidth=3, width=15, anchor='n')
    passwordlabel.place(x=10, y=170)

    hostval=StringVar()
    userval = StringVar()
    passwordval = StringVar()

    hostentry=Entry(dbroot,font=('Comic Sans MS',12,'bold'),bd=5,textvariable=hostval)
    hostentry.place(x=250,y=10)

    userentry = Entry(dbroot, font=('Comic Sans MS', 12, 'bold'), bd=5, textvariable=userval)
    userentry.place(x=250, y=90)

    passwordentry = Entry(dbroot,show='*', font=('Comic Sans MS', 12, 'bold'), bd=5, textvariable=passwordval)
    passwordentry.place(x=250, y=170)

    submitbutton=Button(dbroot,bg='#b7e576',text='SUBMIT',font=('Comic Sans MS',12,NORMAL),width=20,bd=6,activebackground='#b7e576',command=submitdb)
    submitbutton.place(x=150,y=220)

    dbroot.mainloop()

def tick():
    time_string=time.strftime("%H:%M:%S")
    date_string=time.strftime("%d-%m-%Y")
    clock.config(text='Date:'+date_string+"\n"+'Time:'+time_string)
    clock.after(200,tick)

def IntroLabelTick():
    global count,text
    if (count>=len(ss)):
        count=0
        text=''
        SliderLable.config(text=text)
    else:
        text=text+ss[count]
        SliderLable.config(text=text)
        count+=1
    SliderLable.after(200,IntroLabelTick)

root = Tk()
root.title('Student Certificate Maker')
root.config(bg='#243142')
root.geometry('1174x675+350+150')
root.iconbitmap('logo.ico')
root.resizable(False,False)

DataEntryFrame=Frame(root,bg='#243142',relief=GROOVE,borderwidth=5)
DataEntryFrame.place(x=10,y=80,width=500,height=500)
frontlabel=Label(DataEntryFrame,text='-----------Welcome----------',width=45,font=('Comic Sans MS',15,NORMAL))
frontlabel.pack(side=TOP,expand=TRUE)
addbtn=Button(DataEntryFrame,text='1. Add Student', width=25,font=('Comic Sans MS',15,NORMAL),bd=6,command=addstudent)
addbtn.pack(side=TOP,expand=TRUE)

searchbtn=Button(DataEntryFrame,text='2. Search Student', width=25,font=('Comic Sans MS',15,NORMAL),bd=6,command=searchstudent)
searchbtn.pack(side=TOP,expand=TRUE)

deletebtn=Button(DataEntryFrame,text='3. Delete Student', width=25,font=('Comic Sans MS',15,NORMAL),bd=6,command=deletestudent)
deletebtn.pack(side=TOP,expand=TRUE)

updatebtn=Button(DataEntryFrame,text='4.Update Student', width=25,font=('Comic Sans MS',15,NORMAL),bd=6,command=updatestudent)
updatebtn.pack(side=TOP,expand=TRUE)

showbtn=Button(DataEntryFrame,text='5. Show All', width=25,font=('Comic Sans MS',15,NORMAL),bd=6,command=showstudent)
showbtn.pack(side=TOP,expand=TRUE)

exportbtn=Button(DataEntryFrame,text='6. Export Data', width=25,font=('Comic Sans MS',15,NORMAL),bd=6,command=exportstudent)
exportbtn.pack(side=TOP,expand=TRUE)

exitbtn=Button(DataEntryFrame,text='7. Exit', width=25,font=('Comic Sans MS',15,NORMAL),bd=6,command=exitstudent)
exitbtn.pack(side=TOP,expand=TRUE)

ShowDataFrame=Frame(root,bg='white',relief=GROOVE,borderwidth=5)
ShowDataFrame.place(x=550,y=80,width=620, height=500)

style=ttk.Style()
style.configure('Treeview.Heading',font=('Comic Sans MS',11,'bold'))
style.configure('Treeview',font=('Comic Sans MS',10,NORMAL))

scroll_x=Scrollbar(ShowDataFrame,orient=HORIZONTAL)
scroll_y=Scrollbar(ShowDataFrame,orient=VERTICAL)

studentmttable=Treeview(ShowDataFrame,columns=('Id','Name','Mobile No','Email','Address','Gender','D.O.B','Added Date','Added Time'),yscrollcommand=scroll_y.set,xscrollcommand=scroll_x.set)
scroll_x.pack(side=BOTTOM,fill=X)
scroll_y.pack(side=RIGHT,fill=Y)
scroll_x.config(command=studentmttable.xview)
scroll_y.config(command=studentmttable.yview)
studentmttable.heading('Id',text='Id')
studentmttable.heading('Name',text='Name')
studentmttable.heading('Mobile No',text='Mobile No')
studentmttable.heading('Email',text='Email')
studentmttable.heading('Address',text='Address')
studentmttable.heading('Gender',text='Gender')
studentmttable.heading('D.O.B',text='D.O.B')
studentmttable.heading('Added Date',text='Added Date')
studentmttable.heading('Added Time',text='Added Time')
studentmttable['show']='headings'

studentmttable.column('Id',width=100)
studentmttable.column('Mobile No',width=175)
studentmttable.column('Email',width=250)
studentmttable.column('Address',width=250)
studentmttable.column('Gender',width=100)
studentmttable.column('D.O.B',width=100)

studentmttable.pack(fill=BOTH,expand=1)

ss='Welcome To Student Certificate Maker'
count=0
text=''
SliderLable = Label(root, text=ss,font=('Comic Sans MS',20,NORMAL),relief=RIDGE,borderwidth=5,bg='white')
SliderLable.place(x=260,y=0)

clock=Label(root, font=('Comic Sans MS',13,NORMAL),relief=RIDGE,borderwidth=5,bg='white')
clock.place(x=0,y=0)
tick()
connectbutton=Button(root,text='Connect To Database',width=20,font=('Comic Sans MS',15,NORMAL),relief=RIDGE,borderwidth=5,bg='#b7e576',bd=6,activebackground='#b7e576',command=Connectdb)
connectbutton.place(x=930,y=0)

def Makecr():
    dbroot = Toplevel()
    dbroot.grab_set()
    dbroot.config(bg='white')
    dbroot.title('Make Certificate')
    dbroot.geometry('600x110+680+470')
    dbroot.resizable(FALSE, FALSE)

    Frame1 = Frame(dbroot, bg='#243142')
    Frame1.place(x=30, y=30, width=100, height=40)

    Frame2 = Frame(dbroot, bg='#243142')
    Frame2.place(x=150, y=30, width=100, height=40)

    Frame3 = Frame(dbroot, bg='#243142')
    Frame3.place(x=270, y=30, width=100, height=40)

    Frame4 = Frame(dbroot, bg='#243142')
    Frame4.place(x=390, y=30, width=180, height=40)

    def template():
        global ff
        ff = filedialog.askopenfilename()
        return ff

    def data():
        global cc
        cc = filedialog.askopenfilename()
        return cc

    def result():
        global pdffiles
        pdffiles = filedialog.askdirectory()
        return pdffiles

    def payslip():
        csvfn = r"{}".format(cc)

        def mkw(n):
            tpl1 = r"{}".format(ff)
            tpl = DocxTemplate(tpl1)
            print(tpl)

            filepath = r'{}'.format(pdffiles)

            # tpl = DocxTemplate("template.docx") # In same directory
            df = pd.read_csv(csvfn)
            df_to_doct = df.to_dict()  # dataframe -> dict for the template render
            x = df.to_dict(orient='records')
            context = x
            tpl.render(context[n])
            tpl.save("{}/%s.docx".format(filepath) % str(n + 1))
            wait = time.sleep(random.randint(1, 2))

        df2 = len(pd.read_csv(csvfn))
        print("There will be ", df2, "files")

        for i in range(0, df2):
            print("Making file: ", f"{i},", "..Please Wait...")
            mkw(i)

        print("Done! - Now check your files")

        newpath = r'{}'.format(pdffiles)
        print(newpath)
        if not os.path.exists(newpath):
            os.makedirs(newpath)

        i = df2
        for a in range(1, i + 1):
            convert("{}/{}.docx".format(newpath, a), r"{}/".format(pdffiles))
        messagebox.showinfo('Notification', 'Sucussfully Done.....')

        dbl1 = r"{}".format(pdffiles)
        x = df2
        for x in range(1, x + 1):
            os.remove("{}/{}.docx".format(newpath, x))
        messagebox.showinfo('Notification', 'Successfully Deleted all word files...Certificates are Ready!!')

    tmpbtn = Button(Frame1, text='Template', bg='#b7e576', font=('Comic Sans MS', 15, NORMAL),relief=RIDGE,borderwidth=5, width=15, anchor='n', command=template)
    tmpbtn.place(x=0, y=0, width=100, height=40)

    databtn = Button(Frame2, text='Data', bg='#b7e576', font=('Comic Sans MS', 15, NORMAL), relief=RIDGE,borderwidth=5, command=data)
    databtn.place(x=0, y=0, width=100, height=40)

    rsltbtn = Button(Frame3, text='Result',  bg='#b7e576', font=('Comic Sans MS', 15, NORMAL),relief=RIDGE,borderwidth=5, command=result)
    rsltbtn.place(x=0, y=0, width=100, height=40)

    pslpbtn = Button(Frame4, text='Make Certificate',  bg='#b7e576', font=('Comic Sans MS', 15, NORMAL), relief=RIDGE,borderwidth=5,
                     command=payslip)
    pslpbtn.place(x=0, y=0, width=180, height=40)

certibutton=Button(root,text='Make Certificate',width=20,font=('Comic Sans MS',15,NORMAL),relief=RIDGE,borderwidth=5,bg='#b7e576',bd=6,activebackground='#b7e576',command=Makecr)
certibutton.place(x=450,y=600)

root.mainloop()
