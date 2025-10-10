import sys
from PyQt5.QtCore import *
from PyQt5.QtWidgets import * 
from PyQt5.QtGui import *
import cv2
import pickle
import numpy as np
import os
import csv
from sklearn.neighbors import KNeighborsClassifier
import time
from datetime import datetime
from win32com.client import Dispatch
import pandas as pd
from PyQt5 import QtCore, QtGui, QtWidgets


class HomePage(QWidget):
    def open1(self):
        self.frame2 = QWidget()
        self.frame2 = SecondPage()
        self.frame2.show()
        HomePage.hide(self)

    def open4(self):
        self.frame5 = QWidget()
        self.frame5 = ThirdPage()
        self.frame5.show()
        HomePage.hide(self)

    def open5(self):
        self.frame6 = QWidget()
        self.frame6 = FourthPage()
        self.frame6.show()
        HomePage.hide(self)

    def open8(self):
        self.frame9 = QWidget()
        self.frame9 = FifthPage()
        self.frame9.show()
        HomePage.hide(self)

    def __init__(self):
        super(HomePage, self).__init__()

        self.setWindowTitle("ATTENDANCE SYSTEM")
        self.setFixedWidth(1000)
        self.setStyleSheet("background: #FFFFFF;")
        self.setWindowIcon(QtGui.QIcon(r'C:\Users\shukl\Desktop\AI_Project_1000\AI_Project_1000\Files\ICON.png'))

        self.widgets = {
            "logo1": [],
            "l1": [],
            "l2": [],
            "l3": [],
            "l4": [],
            "l5": [],
            "button1": [],
            "button2": [],
            "button3": [],
            "button10": [],
            "l23": []
        }
        
        self.initUI()

    def initUI(self):

        grid = QGridLayout()
        self.setLayout(grid)

        self.image1 = QPixmap(r'C:\Users\shukl\Desktop\AI_Project_1000\AI_Project_1000\Files\PERFECT.png')
        self.logo1= QLabel()
        self.logo1.setPixmap(self.image1)
        self.logo1.setAlignment(Qt.AlignHCenter)
        self.logo1.setStyleSheet("margin-top: 10px;")
        self.widgets["logo1"].append(self.logo1)

        self.l1 = QLabel()
        self.l1.setText("Our attendance detection program makes it easy for us to track attendance accurately.\nWith features like facial recognition or GPS check-ins, we can automatically record who’s present without the hassle of manual tracking.\nIt’s perfect for us whether we’re managing a classroom, a workplace, or an event.\nThis tool helps us keep reliable records and saves us time, making attendance management simple and efficient.")
        self.l1.setFont(QFont('Bahnschrift Condensed',15))
        self.l1.setAlignment(Qt.AlignCenter)
        self.widgets["l1"].append(self.l1)
        
        self.l2 = QLabel()
        self.l2.setText(" ")
        self.l2.setFont(QFont('Arial font',11))
        self.l2.setAlignment(Qt.AlignCenter)
        self.widgets["l2"].append(self.l2)

        self.l3 = QLabel()
        self.l3.setText("Click the button to fill your details")
        self.l3.setFont(QFont('Bahnschrift Condensed',12))
        self.l3.setAlignment(Qt.AlignCenter)
        self.widgets["l3"].append(self.l3)

        self.button1 = QPushButton("Fill Details Here")
        self.button1.setFont(QFont('Bahnschrift Light Condensed',12))
        self.button1.setCursor(QCursor(Qt.PointingHandCursor))
        self.button1.setStyleSheet(
            "*{border: 4px solid '#cbad8f';" +
            "border-radius: 15px;" +
            "font-size: 20px;" +
            "color: '#796648';" +
            "padding: 5px 0;" +
            "margin: 10px 20px;}" +
            "*:hover{background: '#cbad8f';}"
        )
        self.button1.clicked.connect(self.open1)
        self.widgets["button1"].append(self.button1)

        self.l4 = QLabel("Click the button get Attendance")
        self.l4.setFont(QFont('Bahnschrift Condensed',12))
        self.l4.setAlignment(Qt.AlignCenter)
        self.widgets["l4"].append(self.l4)

        self.button2 = QPushButton("Give Attendance Here")
        self.button2.setFont(QFont('Bahnschrift Light Condensed',12))
        self.button2.setCursor(QCursor(Qt.PointingHandCursor))
        self.button2.setStyleSheet(
            "*{border: 4px solid '#cbad8f';" +
            "border-radius: 15px;" +
            "font-size: 20px;" +
            "color: '#796648';" +
            "padding: 5px 0;" +
            "margin: 10px 20px;}" +
            "*:hover{background: '#cbad8f';}"
        )

        self.button2.clicked.connect(self.open4)
        self.widgets["button2"].append(self.button2)

        self.l5 = QLabel("Click the button to see the record")
        self.l5.setFont(QFont('Bahnschrift Condensed',12))
        self.l5.setAlignment(Qt.AlignCenter)
        self.widgets["l5"].append(self.l5)

        self.button3 = QPushButton("See Attendance Here")
        self.button3.setFont(QFont('Bahnschrift Light Condensed',12))
        self.button3.setCursor(QCursor(Qt.PointingHandCursor))
        self.button3.setStyleSheet(
            "*{border: 4px solid '#cbad8f';" +
            "border-radius: 15px;" +
            "font-size: 20px;" +
            "color: '#796648';" +
            "padding: 5px 0;" +
            "margin: 10px 20px;}" +
            "*:hover{background: '#cbad8f';}"
        )
        self.button3.clicked.connect(self.open5)
        self.widgets["button3"].append(self.button3)

        self.l23 = QLabel("Click to see Developer's Information")
        self.l23.setFont(QFont('Bahnschrift Condensed',12))
        self.l23.setAlignment(Qt.AlignCenter)
        self.widgets["l23"].append(self.l23)

        self.button10 = QPushButton("Developers Info")
        self.button10.setFont(QFont('Bahnschrift Light Condensed',12))
        self.button10.setCursor(QCursor(Qt.PointingHandCursor))
        self.button10.setStyleSheet(
            "*{border: 4px solid '#cbad8f';" +
            "border-radius: 15px;" +
            "font-size: 20px;" +
            "color: '#796648';" +
            "padding: 5px 0;" +
            "margin: 10px 20px;}" +
            "*:hover{background: '#cbad8f';}"
        )
        self.button10.clicked.connect(self.open8)
        self.widgets["button10"].append(self.button10)

        grid.addWidget(self.widgets["logo1"][-1], 0, 0)
        grid.addWidget(self.widgets["l1"][-1], 1, 0)
        grid.addWidget(self.widgets["l2"][-1], 2, 0)
        grid.addWidget(self.widgets["l3"][-1], 3, 0)
        grid.addWidget(self.widgets["l4"][-1], 5, 0)
        grid.addWidget(self.widgets["l5"][-1], 7, 0)
        grid.addWidget(self.widgets["button1"][-1], 4, 0)
        grid.addWidget(self.widgets["button2"][-1], 6, 0)
        grid.addWidget(self.widgets["button3"][-1], 8, 0)
        grid.addWidget(self.widgets["button10"][-1], 10, 0)
        grid.addWidget(self.widgets["l23"][-1], 9, 0)

class SecondPage(QWidget):

    def open2(self):
        self.frame3 = QWidget()
        self.frame3 = S2_1()
        self.frame3.show()
        SecondPage.hide(self)

    def __init__(self):
        super(SecondPage, self).__init__()

        self.setWindowTitle("ATTENDANCE SYSTEM")
        self.setFixedWidth(1000)
        self.setStyleSheet("background: #FFFFFF;")
        self.setWindowIcon(QtGui.QIcon(r'C:\Users\shukl\Desktop\AI_Project_1000\AI_Project_1000\Files\ICON.png'))

        self.widgets = {
            "logo2": [],
            "l6": [],
            "l7": [],
            "l8": [],
            "l9": [],
            "l10": [],
            "fname":[],
            "lname":[],
            "rno": [],
            "button4": [],
            "l11": []
        }

        self.initUI()

    def initUI(self):

        grid = QGridLayout()
        self.setLayout(grid)

        self.image2 = QPixmap(r'C:\Users\shukl\Desktop\AI_Project_1000\AI_Project_1000\Files\PERFECT.png')
        self.logo2= QLabel()
        self.logo2.setPixmap(self.image2)
        self.logo2.setAlignment(Qt.AlignHCenter)
        self.logo2.setStyleSheet("margin-top: 10px;")
        self.widgets["logo2"].append(self.logo2)

        self.l6 = QLabel("Enter your details")
        self.l6.setFont(QFont('Bahnschrift Condensed',15))
        self.l6.setAlignment(Qt.AlignCenter)
        self.widgets["l6"].append(self.l6)

        self.l7 = QLabel(" ")
        self.l7.setFont(QFont('Bahnschrift Condensed',15))
        self.l7.setAlignment(Qt.AlignCenter)
        self.widgets["l7"].append(self.l7)

        self.l8 = QLabel("First Name: ")
        self.l8.setFont(QFont('Bahnschrift Condensed',15))
        self.l8.setAlignment(Qt.AlignLeft )
        self.widgets["l8"].append(self.l8)

        self.l9 = QLabel("Last Name: ")
        self.l9.setFont(QFont('Bahnschrift Condensed',15))
        self.l9.setAlignment(Qt.AlignLeft )
        self.widgets["l9"].append(self.l9)

        self.l10 = QLabel("Roll Number: ")
        self.l10.setFont(QFont('Bahnschrift Condensed',15))
        self.l10.setAlignment(Qt.AlignLeft )
        self.widgets["l10"].append(self.l10)

        self.fname = QLineEdit()
        self.widgets["fname"].append(self.fname)
        self.lname = QLineEdit()
        self.widgets["lname"].append(self.lname)
        self.rno = QLineEdit()
        self.widgets["rno"].append(self.rno)
            
        self.button4 = QPushButton("Enter")
        self.button4.setFont(QFont('Bahnschrift Light Condensed',12))
        self.button4.setCursor(QCursor(Qt.PointingHandCursor))
        self.button4.setFixedWidth(975)
        self.button4.setStyleSheet(
            "*{border: 4px solid '#cbad8f';" +
            "border-radius: 15px;" +
            "font-size: 20px;" +
            "color: '#796648';" +
            "padding: 5px 0;" +
            "margin: 10px 20px;}" +
            "*:hover{background: '#cbad8f';}"
        )
        self.button4.clicked.connect(self.addinfo)
        self.widgets["button4"].append(self.button4)

        grid.addWidget(self.widgets["logo2"][-1], 0, 0)
        grid.addWidget(self.widgets["l6"][-1], 1, 0)
        grid.addWidget(self.widgets["l7"][-1], 2, 0)
        grid.addWidget(self.widgets["l8"][-1], 3, 0)
        grid.addWidget(self.widgets["l9"][-1], 5, 0)
        grid.addWidget(self.widgets["l10"][-1], 7, 0)
        grid.addWidget(self.widgets["fname"][-1], 4, 0)
        grid.addWidget(self.widgets["lname"][-1], 6, 0)
        grid.addWidget(self.widgets["rno"][-1], 8, 0)
        grid.addWidget(self.widgets["button4"][-1], 9, 0)


    def addinfo(self):
        warn = QMessageBox()
        FNAME = self.fname.text()
        LNAME = self.lname.text()
        RNO = self.rno.text()
        temp = open(r"C:\Users\shukl\Desktop\AI_Project_1000\AI_Project_1000\Files\STUDENTS.csv", "a")
        tempRead = csv.reader(temp)
        info = RNO + ',' + FNAME + ',' + LNAME
        INFO = RNO + FNAME + LNAME

        if INFO != '':
            if RNO != '' or FNAME != '' or LNAME != '':
                temp.write(info + '\n')

                self.open2()

            else:
                warn.setWindowTitle("System")
                warn.setText("Student Details are required")
                warn.setIcon(QMessageBox.Warning)
                warn.exec_()
        else:
            warn.setWindowTitle("System")
            warn.setText("Student Details are required")
            warn.setIcon(QMessageBox.Warning)
            warn.exec_()

class S2_1(QWidget):

    def open3(self):
        self.frame4 = QWidget()
        self.frame4.setWindowTitle("ATTENDANCE SYSTEM")
        self.frame4.setFixedWidth(1000)
        self.frame4.setStyleSheet("background: #FFFFFF;")
        self.frame4.setWindowIcon(QtGui.QIcon(r'C:\Users\shukl\Desktop\AI_Project_1000\AI_Project_1000\Files\ICON.png'))
        self.frame4 = HomePage()
        self.frame4.show()
        S2_1.hide(self)

    def __init__(self):
        super(S2_1, self).__init__()

        self.setWindowTitle("ATTENDANCE SYSTEM")
        self.setFixedWidth(1000)
        self.setStyleSheet("background: #FFFFFF;")
        self.setWindowIcon(QtGui.QIcon(r'C:\Users\shukl\Desktop\AI_Project_1000\AI_Project_1000\Files\ICON.png'))
        

        self.widgets = {
            "logo3": [],
            "l11": [],
            "l12": [],
            "l13": [],
            "l14": [],
            "button5": [],
            "button6": []
        }

        self.initUI()

    def initUI(self):

        grid = QGridLayout()
        self.setLayout(grid)

        self.image3 = QPixmap(r'C:\Users\shukl\Desktop\AI_Project_1000\AI_Project_1000\Files\PERFECT.png')
        self.logo3= QLabel()
        self.logo3.setPixmap(self.image3)
        self.logo3.setAlignment(Qt.AlignHCenter)
        self.logo3.setStyleSheet("margin-top: 10px;")
        self.widgets["logo3"].append(self.logo3)

        self.l11 = QLabel("Uploaded Successfully...")
        self.l11.setFont(QFont('Bahnschrift Condensed',15))
        self.l11.setAlignment(Qt.AlignCenter)
        self.widgets["l11"].append(self.l11)

        self.l12 = QLabel(" ")
        self.l12.setFont(QFont('Bahnschrift Condensed',15))
        self.l12.setAlignment(Qt.AlignCenter)
        self.widgets["l12"].append(self.l12)

        self.l13 = QLabel("For attendance system, press the below button to take images of the student")
        self.l13.setFont(QFont('Bahnschrift Condensed',15))
        self.l13.setAlignment(Qt.AlignCenter)
        self.widgets["l13"].append(self.l13)

        self.button5 = QPushButton("Take Image")
        self.button5.setFont(QFont('Bahnschrift Light Condensed',12))
        self.button5.setCursor(QCursor(Qt.PointingHandCursor))
        self.button5.setFixedWidth(975)
        self.button5.setStyleSheet(
            "*{border: 4px solid '#cbad8f';" +
            "border-radius: 15px;" +
            "font-size: 20px;" +
            "color: '#796648';" +
            "padding: 5px 0;" +
            "margin: 10px 20px;}" +
            "*:hover{background: '#cbad8f';}"
        )
        self.button5.clicked.connect(self.take_image)
        self.widgets["button5"].append(self.button5)

        self.l14 = QLabel(" ")
        self.l14.setFont(QFont('Bahnschrift Condensed',15))
        self.l14.setAlignment(Qt.AlignCenter)
        self.widgets["l14"].append(self.l14)

        self.button6 = QPushButton("Home")
        self.button6.setFont(QFont('Bahnschrift Light Condensed',12))
        self.button6.setCursor(QCursor(Qt.PointingHandCursor))
        self.button6.setFixedWidth(200)
        self.button6.setStyleSheet(
            "*{border: 4px solid '#cbad8f';" +
            "border-radius: 15px;" +
            "font-size: 20px;" +
            "color: '#796648';" +
            "padding: 5px 0;" +
            "margin: 10px 20px;}" +
            "*:hover{background: '#cbad8f';}"
        )
        self.button6.clicked.connect(self.open3)
        self.widgets["button6"].append(self.button6)

        grid.addWidget(self.widgets["logo3"][-1], 0, 0)
        grid.addWidget(self.widgets["l11"][-1], 1, 0)
        grid.addWidget(self.widgets["l12"][-1], 2, 0)
        grid.addWidget(self.widgets["l13"][-1], 3, 0)
        grid.addWidget(self.widgets["button5"][-1], 4, 0)
        grid.addWidget(self.widgets["l14"][-1], 5, 0)
        grid.addWidget(self.widgets["button6"][-1], 6, 1)

    def take_image(self):
        video=cv2.VideoCapture(0)

        facedetect = cv2.CascadeClassifier(r"C:\Users\shukl\Desktop\AI_Project_1000\AI_Project_1000\Face Detection\data\haarcascade_frontalface_default.xml")
        faces_data=[]

        i=0

        def read_specific_column(file_path, column_index):
            column_data = []
            
            with open(file_path, mode='r', newline='') as csvfile:
                csvreader = csv.reader(csvfile)
                                
                for row in csvreader:
                    if len(row) > column_index:
                        column_data.append(row[column_index])
            
            return column_data

        file_path = r'C:\Users\shukl\Desktop\AI_Project_1000\AI_Project_1000\Files\STUDENTS.csv'
        column_index = 1
        data = read_specific_column(file_path, column_index)

        def Peek(data):
            top = len(data)-1
            return data[top]

        data1=Peek(data)

        def read_specific_column(file_path, column_index):
            column_data = []
            
            with open(file_path, mode='r', newline='') as csvfile:
                csvreader = csv.reader(csvfile)

                for row in csvreader:
                    if len(row) > column_index:
                        column_data.append(row[column_index])
            
            return column_data

        file_path = r'C:\Users\shukl\Desktop\AI_Project_1000\AI_Project_1000\Files\STUDENTS.csv'
        column_index = 2
        data = read_specific_column(file_path, column_index)

        def Peek(data):
            top = len(data)-1
            return data[top]

        data2=Peek(data)

        DATA = data1 + " " + data2
        
        name = DATA

        def read_specific_column(file_path, column_index):
            column_data = []
            
            with open(file_path, mode='r', newline='') as csvfile:
                csvreader = csv.reader(csvfile)
                
                for row in csvreader:
                    if len(row) > column_index:
                        column_data.append(row[column_index])
            
            return column_data

        file_path = r'C:\Users\shukl\Desktop\AI_Project_1000\AI_Project_1000\Files\STUDENTS.csv'
        column_index = 0
        data = read_specific_column(file_path, column_index)

        def Peek(data):
            top = len(data)-1
            return data[top]

        data3=Peek(data)

        roLLno = data3

        while True:
            ret,frame=video.read()
            grey=cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
            faces=facedetect.detectMultiScale(grey,1.3,5)
            for (x,y,w,h) in faces:
                crop_img=frame[y:y+h, x:x+w, :]
                resized_img=cv2.resize(crop_img, (50,50))
                if len(faces_data)<=100 and i%10==0:
                    faces_data.append(resized_img)
                i=i+1
                cv2.putText(frame, str(len(faces_data)), (50,50), cv2.FONT_HERSHEY_COMPLEX, 1, (50,50,255) ,1)
                cv2.rectangle(frame, (x,y), (x+w, y+h), (50,50,255), 1)
            cv2.imshow("Frame",frame)
            k=cv2.waitKey(1)
            if k==ord("q") or len(faces_data)==100:
                break
        video.release()
        cv2.destroyAllWindows()

        faces_data=np.asarray(faces_data)
        faces_data=faces_data.reshape(100, -1)

        if 'names.pkl' not in os.listdir(r'C:\Users\shukl\Desktop\AI_Project_1000\AI_Project_1000\Face Detection\data'):
            names=[name]*100
            with open(r'C:\Users\shukl\Desktop\AI_Project_1000\AI_Project_1000\Face Detection\data\names.pkl', 'wb') as f:
                pickle.dump(names, f)
        else:
            with open(r'C:\Users\shukl\Desktop\AI_Project_1000\AI_Project_1000\Face Detection\data\names.pkl', 'rb') as f:
                names=pickle.load(f)
            names=names+[name]*100
            with open(r'C:\Users\shukl\Desktop\AI_Project_1000\AI_Project_1000\Face Detection\data\names.pkl', 'wb') as f:
                pickle.dump(names, f)
            
        if 'faces_data.pkl' not in os.listdir(r'C:\Users\shukl\Desktop\AI_Project_1000\AI_Project_1000\Face Detection\data'):
            with open(r'C:\Users\shukl\Desktop\AI_Project_1000\AI_Project_1000\Face Detection\data\faces_data.pkl', 'wb') as f:
                pickle.dump(faces_data, f)
        else:
            with open(r'C:\Users\shukl\Desktop\AI_Project_1000\AI_Project_1000\Face Detection\data\faces_data.pkl','rb') as f:
                faces=pickle.load(f)
            faces=np.append(faces, faces_data, axis=0)
            with open(r'C:\Users\shukl\Desktop\AI_Project_1000\AI_Project_1000\Face Detection\data\faces_data.pkl', 'wb') as f:
                pickle.dump(faces, f)

        if 'roll.pkl' not in os.listdir(r'C:\Users\shukl\Desktop\AI_Project_1000\AI_Project_1000\Face Detection\data'):
            roll=[roLLno]*100
            with open(r'C:\Users\shukl\Desktop\AI_Project_1000\AI_Project_1000\Face Detection\data\roll.pkl', 'wb') as f:
                pickle.dump(roll, f)
        else:
            with open(r'C:\Users\shukl\Desktop\AI_Project_1000\AI_Project_1000\Face Detection\data\roll.pkl', 'rb') as f:
                roll=pickle.load(f)
            roll=roll+[roLLno]*100
            with open(r'C:\Users\shukl\Desktop\AI_Project_1000\AI_Project_1000\Face Detection\data\roll.pkl', 'wb') as f:
                pickle.dump(roll, f)

class ThirdPage(QWidget):

    def open4(self):
        self.frame5 = QWidget()
        self.frame5.setWindowTitle("ATTENDANCE SYSTEM")
        self.frame5.setFixedWidth(1000)
        self.frame5.setStyleSheet("background: #FFFFFF;")
        self.frame5.setWindowIcon(QtGui.QIcon(r'C:\Users\shukl\Desktop\AI_Project_1000\AI_Project_1000\Files\ICON.png'))
        self.frame5 = HomePage()
        self.frame5.show()
        ThirdPage.hide(self)
    
    def __init__(self,speak=None):
        super(ThirdPage, self).__init__() 

        self.setWindowTitle("ATTENDANCE SYSTEM")
        self.setFixedWidth(1000)
        self.setStyleSheet("background: #FFFFFF;")
        self.setWindowIcon(QtGui.QIcon(r'C:\Users\shukl\Desktop\AI_Project_1000\AI_Project_1000\Files\ICON.png'))

        self.speak = speak

        def speak(str1):
            speak=Dispatch(("SAPI.SpVoice"))
            speak.Speak(str1)

        video=cv2.VideoCapture(0)
        facedetect=cv2.CascadeClassifier(r'C:\Users\shukl\Desktop\AI_Project_1000\AI_Project_1000\Face Detection\data\haarcascade_frontalface_default.xml')

        with open(r'C:\Users\shukl\Desktop\AI_Project_1000\AI_Project_1000\Face Detection\data\names.pkl','rb') as f:
            LABELS=pickle.load(f)
        with open(r'C:\Users\shukl\Desktop\AI_Project_1000\AI_Project_1000\Face Detection\data\faces_data.pkl','rb') as f:
            FACES=pickle.load(f)
        with open(r'C:\\Users\\shukl\\Desktop\\AI_Project_1000\\AI_Project_1000\\Face Detection\\data\\roll.pkl','rb') as f:
            ROLL=pickle.load(f)

        knn=KNeighborsClassifier(n_neighbors=5)
        knn.fit(FACES, LABELS)

        imgBackground=cv2.imread(r"C:\Users\shukl\Desktop\AI_Project_1000\AI_Project_1000\Face Detection\LOGO Background.jpg")

        COL_NAMES = ['ROLL NO.', 'NAME', 'TIME']

        while True:
            ret,frame=video.read()
            grey=cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
            faces=facedetect.detectMultiScale(grey, 1.3 , 5)
            for (x,y,w,h) in faces:
                crop_img=frame[y:y+h, x:x+w, :]
                resized_img=cv2.resize(crop_img, (50,50)).flatten().reshape(1,-1)
                output=knn.predict(resized_img)
                ts=time.time()
                date=datetime.fromtimestamp(ts).strftime("%d-%m-%Y")
                timestamp=datetime.fromtimestamp(ts).strftime("%H:%M:%S")
                exist=os.path.isfile(r"C:\Users\shukl\Desktop\AI_Project_1000\AI_Project_1000\Face Detection\Attendance\Attendance_" + date + r".csv")
                cv2.rectangle(frame, (x,y), (x+w,y+h), (0,0,255), 1)
                cv2.rectangle(frame, (x,y), (x+w,y+h), (50,50,255), 2)
                cv2.rectangle(frame, (x,y-40), (x+w,y), (50,50,255), -1)
                cv2.putText(frame, str(output[0]), (x,y-15),cv2.FONT_HERSHEY_COMPLEX, 1, (255,255,255), 1)
                cv2.rectangle(frame, (x,y), (x+w, y+h), (50,50,255), 1)
                attendance=[str(ROLL[0]), str(output[0]), str(timestamp)]
            imgBackground[162:162 + 480, 55:55 + 640] = frame
            cv2.imshow("Frame",imgBackground)
            k=cv2.waitKey(1)
            if k==ord("A"):
                speak("Attendance Taken..")
                time.sleep(1)
                if exist:
                    with open(r"C:\Users\shukl\Desktop\AI_Project_1000\AI_Project_1000\Face Detection\Attendance\Attendance_" + date + r".csv", r"+a" , w) as csvfile:
                        writer=csv.writer(csvfile)
                        writer.writerow(attendance)
                    csvfile.close()
                else:
                    with open(r"C:\Users\shukl\Desktop\AI_Project_1000\AI_Project_1000\Face Detection\Attendance\Attendance_" + date + r".csv", r"+a" , w) as csvfile:
                        writer=csv.writer(csvfile)
                        writer.writerow(COL_NAMES)
                        writer.writerow(attendance)
                    csvfile.close()
            if k==ord("q"):
                break
        video.release()
        cv2.destroyAllWindows()

        self.widgets = {
            "logo4": [],
            "button7": []
        }

        self.initUI()

    def initUI(self):

        grid = QGridLayout()
        self.setLayout(grid)

        self.image4 = QPixmap(r'C:\Users\shukl\Desktop\AI_Project_1000\AI_Project_1000\Files\PERFECT.png')
        self.logo4= QLabel()
        self.logo4.setPixmap(self.image4)
        self.logo4.setAlignment(Qt.AlignHCenter)
        self.logo4.setStyleSheet("margin-top: 10px;")
        self.widgets["logo4"].append(self.logo4)

        self.button7 = QPushButton("Home")
        self.button7.setFont(QFont('Bahnschrift Light Condensed',12))
        self.button7.setCursor(QCursor(Qt.PointingHandCursor))
        self.button7.setFixedWidth(200)
        self.button7.setStyleSheet(
            "*{border: 4px solid '#cbad8f';" +
            "border-radius: 15px;" +
            "font-size: 20px;" +
            "color: '#796648';" +
            "padding: 5px 0;" +
            "margin: 10px 20px;}" +
            "*:hover{background: '#cbad8f';}"
        )
        self.button7.clicked.connect(self.open4)
        self.widgets["button7"].append(self.button7)

        grid.addWidget(self.widgets["logo4"][-1], 0, 0)
        grid.addWidget(self.widgets["button7"][-1], 1, 0)

class FourthPage(QWidget):

    def open6(self):
        self.frame7 = QWidget()
        self.frame7.setWindowTitle("ATTENDANCE SYSTEM")
        self.frame7.setFixedWidth(1000)
        self.frame7.setStyleSheet("background: #FFFFFF;")
        self.frame7.setWindowIcon(QtGui.QIcon(r'C:\Users\shukl\Desktop\AI_Project_1000\AI_Project_1000\Files\ICON.png'))
        self.frame7 = CSVOPEN()
        self.frame7.show()


    def open7(self):
        self.frame8 = QWidget()
        self.frame8.setWindowTitle("ATTENDANCE SYSTEM")
        self.frame8.setFixedWidth(1000)
        self.frame8.setStyleSheet("background: #FFFFFF;")
        self.frame8.setWindowIcon(QtGui.QIcon(r'C:\Users\shukl\Desktop\AI_Project_1000\AI_Project_1000\Files\ICON.png'))
        self.frame8 = HomePage()
        self.frame8.show()
        CSVOPEN.hide(self)
     
    def __init__(self):
        super(FourthPage, self).__init__()

        self.setWindowTitle("ATTENDANCE SYSTEM")
        self.setFixedWidth(1000)
        self.setStyleSheet("background: #FFFFFF;")
        self.setWindowIcon(QtGui.QIcon(r'C:\Users\shukl\Desktop\AI_Project_1000\AI_Project_1000\Files\ICON.png'))

        self.widgets = {
            "logo5": [],
            "l15": [],
            "dAtE": [],
            "button8": [],
            "button9": []
        }

        self.initUI()

    def initUI(self):

        grid = QGridLayout()
        self.setLayout(grid)

        self.image5 = QPixmap(r'C:\Users\shukl\Desktop\AI_Project_1000\AI_Project_1000\Files\PERFECT.png')
        self.logo5= QLabel()
        self.logo5.setPixmap(self.image5)
        self.logo5.setAlignment(Qt.AlignHCenter)
        self.logo5.setStyleSheet("margin-top: 10px;")
        self.widgets["logo5"].append(self.logo5)

        self.l15 = QLabel("Enter the date -----> [DD-MM-YYYY]")
        self.l15.setFont(QFont('Bahnschrift Condensed',15))
        self.l15.setAlignment(Qt.AlignCenter)
        self.widgets["l15"].append(self.l15)

        self.dAtE = QLineEdit()
        self.widgets["dAtE"].append(self.dAtE)

        self.button8 = QPushButton("Enter")
        self.button8.setFont(QFont('Bahnschrift Light Condensed',12))
        self.button8.setCursor(QCursor(Qt.PointingHandCursor))
        self.button8.setFixedWidth(475)
        self.button8.setStyleSheet(
            "*{border: 4px solid '#cbad8f';" +
            "border-radius: 15px;" +
            "font-size: 20px;" +
            "color: '#796648';" +
            "padding: 5px 0;" +
            "margin: 10px 20px;}" +
            "*:hover{background: '#cbad8f';}"
        )
        self.button8.clicked.connect(self.addinfo1)
        self.widgets["button8"].append(self.button8)

        self.button9 = QPushButton("Home")
        self.button9.setFont(QFont('Bahnschrift Light Condensed',12))
        self.button9.setCursor(QCursor(Qt.PointingHandCursor))
        self.button9.setFixedWidth(200)
        self.button9.setStyleSheet(
            "*{border: 4px solid '#cbad8f';" +
            "border-radius: 15px;" +
            "font-size: 20px;" +
            "color: '#796648';" +
            "padding: 5px 0;" +
            "margin: 10px 20px;}" +
            "*:hover{background: '#cbad8f';}"
        )
        self.button9.clicked.connect(self.open7)
        self.widgets["button9"].append(self.button9)

        grid.addWidget(self.widgets["logo5"][-1], 0, 0)
        grid.addWidget(self.widgets["l15"][-1], 1, 0)
        grid.addWidget(self.widgets["dAtE"][-1], 2, 0)
        grid.addWidget(self.widgets["button8"][-1], 3, 0)
        grid.addWidget(self.widgets["button9"][-1], 4, 1)

    def addinfo1(self):
        warn1 = QMessageBox()
        DaTe = self.dAtE.text()

        if DaTe != '':
            with open(r"C:\Users\shukl\Desktop\AI_Project_1000\AI_Project_1000\Files\DAte.csv", "a", newline='') as temp:
                temp.write(DaTe + '\n')
            self.open6()

        else:
            warn1.setWindowTitle("System")
            warn1.setText("Date required")
            warn1.setIcon(QMessageBox.Warning)
            warn1.exec_()

class CSVOPEN(QMainWindow):
    def __init__(self):
        super(CSVOPEN, self).__init__()

        def read_specific_column(file_path, column_index):
            column_data = []
            
            with open(file_path, mode='r', newline='') as csvfile:
                csvreader = csv.reader(csvfile)
                
                for row in csvreader:
                    if len(row) > column_index:
                        column_data.append(row[column_index])
            
            return column_data

        file_path = r"C:\Users\shukl\Desktop\AI_Project_1000\AI_Project_1000\Files\DAte.csv"
        column_index = 0
        data = read_specific_column(file_path, column_index)

        def Peek(data):
            top = len(data)-1
            return data[top]

        date=Peek(data)

        self.table = QtWidgets.QTableView()

        Data=[]

        with open(r"C:\Users\shukl\Desktop\AI_Project_1000\AI_Project_1000\Face Detection\Attendance\Attendance_" + date + r".csv", newline='\r\n') as fh:
            ereader = csv.reader(fh)
            for rec in ereader:
                Data.append(rec)
                
        self.model = TableModel(Data)
        self.table.setModel(self.model)

        self.setCentralWidget(self.table)

class TableModel(QtCore.QAbstractTableModel):
    def __init__(self, data):
        super(TableModel, self).__init__()
        self._data = data

    def data(self, index, role):
        if role == Qt.DisplayRole:
            return self._data[index.row()][index.column()]

    def rowCount(self, index):
        return len(self._data)

    def columnCount(self, index):
        return len(self._data[0])
    
class FifthPage(QWidget):

    def open9(self):
        self.frame10 = QWidget()
        self.frame10.setWindowTitle("ATTENDANCE SYSTEM")
        self.frame10.setFixedWidth(1000)
        self.frame10.setStyleSheet("background: #FFFFFF;")
        self.frame10.setWindowIcon(QtGui.QIcon(r'C:\Users\shukl\Desktop\AI_Project_1000\AI_Project_1000\Files\ICON.png'))
        self.frame10 = HomePage()
        self.frame10.show()
        FifthPage.hide(self)

    def __init__(self):
        super(FifthPage, self).__init__()

        self.setWindowTitle("ATTENDANCE SYSTEM")
        #self.setFixedWidth(1000)
        self.setStyleSheet("background: #FFFFFF;")
        self.setWindowIcon(QtGui.QIcon(r'C:\Users\shukl\Desktop\AI_Project_1000\AI_Project_1000\Files\ICON.png'))

        self.widgets = {
            "logo6": [],
            "l16": [],
            "l17": [],
            "l18": [],
            "l19": [],
            "l20": [],
            "l21": [],
            "l22": [],
            "l25": [],
            "l24": [],
            "logo7": [],
            "logo8": [],
            "button10": [],
            "logo9": [],
            "l27": [],
            "l28": [],
            "l29": [],
            "l30": [],
            "l26": [],
            "l31": [],
            "l32": [],
            "l33": [],
            "logo10": [],
            "l34": [],
            "l35": [],
            "l36": [],
            "l37": []
            
        }

        self.initUI()

    def initUI(self):

        grid = QGridLayout()
        self.setLayout(grid)

        self.image6 = QPixmap(r'C:\Users\shukl\Desktop\AI_Project_1000\AI_Project_1000\Files\PERFECT.png')
        self.logo6= QLabel()
        self.logo6.setPixmap(self.image6)
        self.logo6.setAlignment(Qt.AlignHCenter)
        self.logo6.setStyleSheet("margin-top: 10px;")
        self.widgets["logo6"].append(self.logo6)

        self.l24 = QLabel(" ")
        self.l24.setFont(QFont('Bahnschrift Condensed',15))
        self.l24.setAlignment(Qt.AlignCenter)
        self.widgets["l24"].append(self.l24)

        self.l16 = QLabel("Developers Data")
        self.l16.setFont(QFont('Bahnschrift Condensed',20))
        self.l16.setAlignment(Qt.AlignCenter)
        self.widgets["l16"].append(self.l16)

        self.l25 = QLabel(" ")
        self.l25.setFont(QFont('Bahnschrift Condensed',15))
        self.l25.setAlignment(Qt.AlignCenter)
        self.widgets["l25"].append(self.l25)

        self.image7 = QPixmap(r'C:\Users\shukl\Desktop\AI_Project_1000\AI_Project_1000\Files\Ashutosh.jpeg')
        self.new_size = self.image7.size() * 0.15 
        self.resized_image = self.image7.scaled(self.new_size, aspectRatioMode=1)
        self.logo7= QLabel()
        self.logo7.setPixmap(self.resized_image)
        self.logo7.setAlignment(Qt.AlignHCenter)
        self.logo7.setStyleSheet("margin-top: 10px;")
        self.widgets["logo7"].append(self.logo7)

        self.l17 = QLabel("ASHUTOSH SAHU")
        self.l17.setFont(QFont('Bahnschrift Condensed',15))
        self.l17.setAlignment(Qt.AlignCenter)
        self.widgets["l17"].append(self.l17)

        self.l18 = QLabel("Class - XII 'A'")
        self.l18.setFont(QFont('Bahnschrift Condensed',15))
        self.l18.setAlignment(Qt.AlignCenter)
        self.widgets["l18"].append(self.l18)

        self.l19 = QLabel("Roll No. 6")
        self.l19.setFont(QFont('Bahnschrift Condensed',15))
        self.l19.setAlignment(Qt.AlignCenter)
        self.widgets["l19"].append(self.l19)
        
        self.l33 = QLabel("Role: Programmer & GUI Designer ")
        self.l33.setFont(QFont('Bahnschrift Condensed',15))
        self.l33.setAlignment(Qt.AlignCenter)
        self.widgets["l33"].append(self.l33)

        self.image8 = QPixmap(r'C:\Users\shukl\Desktop\AI_Project_1000\AI_Project_1000\Files\Sukrut.jpeg')
        self.new_size = self.image8.size() * 0.25
        self.resized_image = self.image8.scaled(self.new_size, aspectRatioMode=1)
        self.logo8= QLabel()
        self.logo8.setPixmap(self.resized_image)
        self.logo8.setAlignment(Qt.AlignHCenter)
        self.logo8.setStyleSheet("margin-top: 10px;")
        self.widgets["logo8"].append(self.logo8)

        self.l20 = QLabel("SUKRUT DUSANE")
        self.l20.setFont(QFont('Bahnschrift Condensed',15))
        self.l20.setAlignment(Qt.AlignCenter)
        self.widgets["l20"].append(self.l20)

        self.l21 = QLabel("Class - XII 'A'")
        self.l21.setFont(QFont('Bahnschrift Condensed',15))
        self.l21.setAlignment(Qt.AlignCenter)
        self.widgets["l21"].append(self.l21)

        self.l22 = QLabel("Roll No. 22")
        self.l22.setFont(QFont('Bahnschrift Condensed',15))
        self.l22.setAlignment(Qt.AlignCenter)
        self.widgets["l22"].append(self.l22)
        
        self.l32 = QLabel("Role: Group Leader, User Interface Designer ")
        self.l32.setFont(QFont('Bahnschrift Condensed',15))
        self.l32.setAlignment(Qt.AlignCenter)
        self.widgets["l32"].append(self.l32)

        self.l26 = QLabel(" ")
        self.l26.setFont(QFont('Bahnschrift Condensed',15))
        self.l26.setAlignment(Qt.AlignCenter)
        self.widgets["l26"].append(self.l26)

        self.image9 = QPixmap(r'C:\Users\shukl\Desktop\AI_Project_1000\AI_Project_1000\Files\Atharv.jpg')
        self.new_size = self.image9.size() * 0.25  # Resize to 50% of original size
        self.resized_image = self.image9.scaled(self.new_size, aspectRatioMode=1)
        self.logo9= QLabel()
        self.logo9.setPixmap(self.resized_image)
        self.logo9.setAlignment(Qt.AlignHCenter)
        self.logo9.setStyleSheet("margin-top: 10px;")
        self.widgets["logo9"].append(self.logo9)

        self.l27 = QLabel("ATHARV SHUKLA")
        self.l27.setFont(QFont('Bahnschrift Condensed',15))
        self.l27.setAlignment(Qt.AlignCenter)
        self.widgets["l27"].append(self.l27)

        self.l28 = QLabel("Class - XII 'A'")
        self.l28.setFont(QFont('Bahnschrift Condensed',15))
        self.l28.setAlignment(Qt.AlignCenter)
        self.widgets["l28"].append(self.l28)

        self.l29 = QLabel("Roll No. 7")
        self.l29.setFont(QFont('Bahnschrift Condensed',15))
        self.l29.setAlignment(Qt.AlignCenter)
        self.widgets["l29"].append(self.l29)

        self.l30 = QLabel("Role: Logbook, Coding and Structuring ")
        self.l30.setFont(QFont('Bahnschrift Condensed',15))
        self.l30.setAlignment(Qt.AlignCenter)
        self.widgets["l30"].append(self.l30)

        self.image10 = QPixmap(r'C:\Users\shukl\Desktop\AI_Project_1000\AI_Project_1000\Files\Samyak.jpg')
        self.new_size = self.image10.size() * 0.25  # Resize to 50% of original size
        self.resized_image = self.image10.scaled(self.new_size, aspectRatioMode=1)
        self.logo10= QLabel()
        self.logo10.setPixmap(self.resized_image)
        self.logo10.setAlignment(Qt.AlignHCenter)
        self.logo10.setStyleSheet("margin-top: 10px;")
        self.widgets["logo10"].append(self.logo10)

        self.l34 = QLabel("SAMYAK MALU")
        self.l34.setFont(QFont('Bahnschrift Condensed',15))
        self.l34.setAlignment(Qt.AlignCenter)
        self.widgets["l34"].append(self.l34)

        self.l35 = QLabel("Class - XII 'A'")
        self.l35.setFont(QFont('Bahnschrift Condensed',15))
        self.l35.setAlignment(Qt.AlignCenter)
        self.widgets["l35"].append(self.l35)

        self.l36 = QLabel("Roll No. 19")
        self.l36.setFont(QFont('Bahnschrift Condensed',15))
        self.l36.setAlignment(Qt.AlignCenter)
        self.widgets["l36"].append(self.l36)

        self.l37 = QLabel("Role: Videography ")
        self.l37.setFont(QFont('Bahnschrift Condensed',15))
        self.l37.setAlignment(Qt.AlignCenter)
        self.widgets["l37"].append(self.l37)

        self.l31 = QLabel(" ")
        self.l31.setFont(QFont('Bahnschrift Condensed',15))
        self.l31.setAlignment(Qt.AlignCenter)
        self.widgets["l31"].append(self.l31)

        self.button10 = QPushButton("Home")
        self.button10.setFont(QFont('Bahnschrift Light Condensed',12))
        self.button10.setCursor(QCursor(Qt.PointingHandCursor))
        self.button10.setFixedWidth(200)
        self.button10.setStyleSheet(
            "*{border: 4px solid '#cbad8f';" +
            "border-radius: 15px;" +
            "font-size: 20px;" +
            "color: '#796648';" +
            "padding: 5px 0;" +
            "margin: 10px 20px;}" +
            "*:hover{background: '#cbad8f';}"
        )
        self.button10.clicked.connect(self.open9)
        self.widgets["button10"].append(self.button10)


        grid.addWidget(self.widgets["logo6"][-1], 0, 1)
        grid.addWidget(self.widgets["l24"][-1], 1, 0)
        grid.addWidget(self.widgets["l16"][-1], 2, 1)
        grid.addWidget(self.widgets["l25"][-1], 3, 0)
        grid.addWidget(self.widgets["logo7"][-1], 4, 2)
        grid.addWidget(self.widgets["l17"][-1], 5, 2)
        grid.addWidget(self.widgets["l18"][-1], 6, 2)
        grid.addWidget(self.widgets["l19"][-1], 7, 2)
        grid.addWidget(self.widgets["logo8"][-1], 4, 0)
        grid.addWidget(self.widgets["l20"][-1], 5, 0)
        grid.addWidget(self.widgets["l21"][-1], 6, 0)
        grid.addWidget(self.widgets["l22"][-1], 7, 0)
        grid.addWidget(self.widgets["l26"][-1], 9, 0)
        grid.addWidget(self.widgets["logo9"][-1], 10, 0)
        grid.addWidget(self.widgets["l27"][-1], 11, 0)
        grid.addWidget(self.widgets["l28"][-1], 12, 0)
        grid.addWidget(self.widgets["l29"][-1], 13, 0)
        grid.addWidget(self.widgets["l30"][-1], 14, 0)
        grid.addWidget(self.widgets["button10"][-1], 16, 1)
        grid.addWidget(self.widgets["l31"][-1], 15, 0)
        grid.addWidget(self.widgets["l32"][-1], 8, 0)
        grid.addWidget(self.widgets["l33"][-1], 8, 2)
        grid.addWidget(self.widgets["logo10"][-1], 10, 2)
        grid.addWidget(self.widgets["l34"][-1], 11, 2)
        grid.addWidget(self.widgets["l35"][-1], 12, 2)
        grid.addWidget(self.widgets["l36"][-1], 13, 2)
        grid.addWidget(self.widgets["l37"][-1], 14, 2)


def main():
    app = QApplication(sys.argv)
    window = QWidget()
    window = HomePage()
    window.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()