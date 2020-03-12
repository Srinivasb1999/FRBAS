import tkinter as tk
from tkinter import messagebox
import cv2,os
import numpy as np
from PIL import Image
import pandas as pd
from xlutils.copy import copy
from xlrd import open_workbook     
from xlwt import Workbook 

#quit button for main screen
def destroy():
    w1.destroy()
    cv2.destroyAllWindows()
    
#admin view dialog   
def adminview():
    global w2
    w2 = tk.Tk()
    w2.title('Admin View')
    w2.config(bg='gray21')
    window_width = 450
    window_height = 270
    
    screen_width = w2.winfo_screenwidth()
    screen_height = w2.winfo_screenheight()
    
    x_coord = (screen_width/2) - (window_width/2)
    y_coord = (screen_height/2) - (window_height/2)
    
    w2.geometry('%dx%d+%d+%d' % (window_width, window_height, x_coord, y_coord))
    
    l1 = tk.Label(w2, text = 'Password:', bg = 'gray63')
    l1.place(x=20,y=20)
    global e1
    e1 = tk.Entry(w2, bd=5, show = '*', bg = 'gray63')
    e1.place(x=110, y=20)
    bb = tk.Button(w2, text = 'Enter', command = checkpassword, width = 8, bg = 'gray63')
    bb.place(x=300, y=20)
    w2.mainloop()
    
#checks for the right password  
def checkpassword():
    if e1.get() == '1234':
        b0 = tk.Button(w2, text = 'Add days', command = add_days, bg = 'gray63')
        b0. place(x=20,y=80)
        b1 = tk.Button(w2, text='Add new user', command = add_new_user, bg = 'gray63')
        b1.place(x=20,y=140)
        b3 = tk.Button(w2, text = 'Check Attendance', command = f , bg = 'gray63')
        b3.place(x=160, y= 140)
        b2 = tk.Button(w2, text = 'Quit', command = w2.destroy, width =8, bg = 'gray63')
        b2. place(x=20,y=200)
        w2.mainloop()
    else:
        err = 'Wrong password'
        messagebox.showerror('!!!!!!!!!!!', err)
        
f = lambda: os.system("start EXCEL.EXE Attendance.xls")


#adding a new user to the database
def add_new_user():
    global w3
    w3 = tk.Tk()
    w3.config(bg='gray21')
    w3.title('Create Student Details')   
    window_width = 900
    window_height = 600
    
    screen_width = w3.winfo_screenwidth()
    screen_height = w3.winfo_screenheight()
    
    x_coord = (screen_width/2) - (window_width/2)
    y_coord = (screen_height/2) - (window_height/2)
    
    w3.geometry('%dx%d+%d+%d' % (window_width, window_height, x_coord, y_coord))
    
    lbl = tk.Label(w3, text="Enter ID" , bg = 'gray63') 
    lbl.place(x=100, y=100)
    
    global txt, txt2
    
    txt = tk.Entry(w3, bg = 'gray63')
    txt.place(x=200, y=100)

    lbl2 = tk.Label(w3, text="Enter Name", bg = 'gray63') 
    lbl2.place(x=100, y=200)

    txt2 = tk.Entry(w3, bg = 'gray63')
    txt2.place(x=200, y=200)
    
    takeImg = tk.Button(w3, text="Take Images", command=TakeImages, bg = 'gray63')
    takeImg.place(x=200, y=300)
    
    lbl3 = tk.Label(w3, text="Status : ", bg = 'gray63') 
    lbl3.place(x=100, y=400)
    
    b2 = tk.Button(w3, text = 'Quit', command = w3.destroy, width =8, bg = 'gray63')
    b2. place(x=100,y=500)
    w3.mainloop()
    
    
#checks whether details entered match the required datatype
def is_number(s):
    try:
        float(s)
        return True
    except ValueError:
        pass
 
    try:
        import unicodedata
        unicodedata.numeric(s)
        return True
    except (TypeError, ValueError):
        pass
 
    return False
    
#Clicks images and draws rectangle around the face
def TakeImages():        
    Id=(txt.get())
    name=(txt2.get())
    if(is_number(Id) and name.isalpha()):
        cam = cv2.VideoCapture(0)
        harcascadePath = "haarcascade_frontalface_default.xml"
        detector=cv2.CascadeClassifier(harcascadePath)
        sampleNum=0
        while(True):
            ret, img = cam.read()
            gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
            faces = detector.detectMultiScale(gray, 1.3, 5)
            for (x,y,w,h) in faces:
                cv2.rectangle(img,(x,y),(x+w,y+h),(255,0,0),2)        
                #incrementing sample number 
                sampleNum=sampleNum+1
                #saving the captured face in the dataset folder TrainingImage
                cv2.imwrite("TrainingImage\ "+name +"."+Id +'.'+ str(sampleNum) + ".jpg", gray[y:y+h,x:x+w])
                #display the frame
                cv2.imshow('frame',img)
            #wait for 100 miliseconds 
            if cv2.waitKey(100) & 0xFF == ord('q'):
                break
            # break if the sample number is morethan 100
            elif sampleNum>60:
                break
        cam.release()
        cv2.destroyAllWindows() 
        res = "Images Saved for ID : " + Id +" Name : "+ name
        row = [Id , name]
        df=pd.read_excel("Attendance.xls")       
        book_1 = open_workbook('Attendance.xls')
        book = copy(book_1)
        sheet1 = book.get_sheet(0)
        sheet1.write(len(df.index)+1,0,Id)
        sheet1.write(len(df.index)+1,1,name)
        book.save('Attendance.xls')
                
        recognizer = cv2.face_LBPHFaceRecognizer.create()#recognizer = cv2.face.LBPHFaceRecognizer_create()#$cv2.createLBPHFaceRecognizer()
        '''harcascadePath = "haarcascade_frontalface_default.xml"
        detector =cv2.CascadeClassifier(harcascadePath)'''
        faces,Id = getImagesAndLabels("TrainingImage")
        recognizer.train(faces, np.array(Id))
        recognizer.save("Trainer.xml")
        l1= tk.Label(w3, text=res, bg = 'gray63')  
        l1.place(x=200,y=400)

    else:
        if(is_number(Id)):
            res1 = "Enter Alphabetical Name"
            messagebox.showerror('NameError',res1)

        if(name.isalpha()):
            res2 = "Enter Numeric Id"
            messagebox.showerror('IDError',res2)
        
            
def getImagesAndLabels(path):
    #get the path of all the files in the folder
    imagePaths=[os.path.join(path,f) for f in os.listdir(path)] 
    #print(imagePaths)
    
    #create empth face list
    faces=[]
    #create empty ID list
    Ids=[]
    #now looping through all the image paths and loading the Ids and the images
    for imagePath in imagePaths:
        #loading the image and converting it to gray scale
        pilImage=Image.open(imagePath).convert('L')
        #Now we are converting the PIL image into numpy array
        imageNp=np.array(pilImage,'uint8')
        #getting the Id from the image
        Id=int(os.path.split(imagePath)[-1].split(".")[1])
        # extract the face from the training image sample
        faces.append(imageNp)
        Ids.append(Id)        
    return faces,Ids
            
 
#user view dialog
def userview():
    global w4
    w4=tk.Tk()
    w4.config(bg='gray21')
    w4.title('User View')
    window_width = 400
    window_height = 250
    
    screen_width = w4.winfo_screenwidth()
    screen_height = w4.winfo_screenheight()
    
    x_coord = (screen_width/2) - (window_width/2)
    y_coord = (screen_height/2) - (window_height/2)
    
    w4.geometry('%dx%d+%d+%d' % (window_width, window_height, x_coord, y_coord))
    
    att= tk.Button(w4, text="Take Attendance", command=TrackImages, bg = 'gray63')
    att.place(x=20, y=20)
    status= tk.Label(w4, text='Status:', bg = 'gray63')
    status.place(x=20,y=80)
    perc = tk.Button(w4, text = 'Current Percentage:', command = attendance_details, bg ='gray63')
    perc.place(x=20, y = 130 )
    qb= tk.Button(w4, text='Quit', command = w4.destroy, width = 8, bg = 'gray63')
    qb.place(x=20, y=180)
    w4.mainloop()
    
#attendance percentage
def attendance_details():
    df=pd.read_excel('Attendance.xls') 
    inc=len(df.columns)-2
    df = df.fillna(0)
    a = 0
    no = 0
    for i in df.loc[:,'Id']:
        if i == Id:
            a = df.iloc[i-1,2:inc+1].mean()
            a = a*100
            val= tk.Label(w4, text = "%f" %(a) + '%', bg = 'gray63')
            val.place(x=200, y =130)
            no = df.isin([1]).sum(i)[i-1]
            print(no)


#attendance tracking by face recognition
def TrackImages():
    recognizer = cv2.face.LBPHFaceRecognizer_create()#cv2.createLBPHFaceRecognizer()
    recognizer.read("Trainer.xml")
    harcascadePath = "haarcascade_frontalface_default.xml"
    df=pd.read_excel('Attendance.xls')    
    book_1 = open_workbook('Attendance.xls')
    book = copy(book_1)
    sheet1 = book.get_sheet(0)
    faceCascade = cv2.CascadeClassifier(harcascadePath);
    cam = cv2.VideoCapture(0)
    font = cv2.FONT_HERSHEY_SIMPLEX       
    while True:
        ret, im =cam.read()
        gray=cv2.cvtColor(im,cv2.COLOR_BGR2GRAY)
        faces=faceCascade.detectMultiScale(gray, 1.2,5)    
        for(x,y,w,h) in faces:
            cv2.rectangle(im,(x,y),(x+w,y+h),(225,0,0),2)
            global Id
            Id, conf = recognizer.predict(gray[y:y+h,x:x+w])                                   
            if(conf < 50): 
                aa=df.loc[df['Id'] == Id]['Name'].values
                tt=str(Id)+"-"+aa
                i=0
                for row in df.loc[:,'Name']:
                    i = i+1
                    if row==aa:
                        rownum=i
                
                sheet1.write(rownum,len(df.columns)-1,1)
                book.save('Attendance.xls')
                lab= tk.Label(w4, text='Attendance taken for ' +aa, bg ='gray63')
                lab.place(x=100, y= 80)
                #attendance_details()
            else:
                Id='Unknown'                
                tt=str(Id)  
            if(conf > 75):
                noOfFile=len(os.listdir("ImagesUnknown"))+1
                cv2.imwrite("ImagesUnknown\Image"+str(noOfFile) + ".jpg", im[y:y+h,x:x+w])            
            cv2.putText(im,str(tt),(x,y+h), font, 1,(255,255,255),2)  
        cv2.imshow('im',im) 
        if (cv2.waitKey(1)==ord('q')):
            cam.release()
            cv2.destroyAllWindows()
            break
    

#adding days into the Attendance sheet
def add_days():
    df = pd.read_excel('Attendance.xls')
 #   writer = pd.ExcelWriter('excel.xlsx')
    inc=len(df.columns)-1
    str1='Day'+ str(inc)
    
    df = pd.read_excel('Attendance.xls')
    book_1 = open_workbook('Attendance.xls')
    book = copy(book_1)
    sheet1 = book.get_sheet(0)
    
    sheet1.write(0,len(df.columns),str1)
    book.save('Attendance.xls')
    
#GUI for main window
w1 = tk.Tk()
#w1.geometry('400x200')
w1.title('Attendance System')

w1.config(bg='gray21')

window_width = 400
window_height = 200

screen_width = w1.winfo_screenwidth()
screen_height = w1.winfo_screenheight()

x_coord = (screen_width/2) - (window_width/2)
y_coord = (screen_height/2) - (window_height/2)

w1.geometry('%dx%d+%d+%d' % (window_width, window_height, x_coord, y_coord))

b1 = tk.Button(w1, text="Admin View", command=adminview, width = 12, bg='gray63')
b1.place(x=40, y=20)

b2 = tk.Button(w1, text="User View", command=userview, width = 12, bg='gray63')
b2.place(x=240, y=20)

b3 = tk.Button(w1, text="Quit",width = 8, command=destroy, bg='gray63')
b3.place(x=160,y=80)

label = tk.Label(w1, text='Made by the I/Openers', bg='gray21', fg = 'white')
label.config(font=('Montserrat',13))
label.place(x=70, y=140)
w1.mainloop()

#Uncomment this part of the code to create a new Attendance.xls
'''
# Workbook is created 
wb = Workbook() 
  
# add_sheet is used to create sheet. 
sheet1 = wb.add_sheet('Sheet 1')  
sheet1.write(0, 0, 'Id') 
sheet1.write(0, 1, 'Name') 
wb.save('Attendance.xls') '''
