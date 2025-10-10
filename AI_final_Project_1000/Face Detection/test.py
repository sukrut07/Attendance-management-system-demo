from sklearn.neighbors import KNeighborsClassifier
import cv2
import pickle
import numpy as np
import os
import csv
import time
from datetime import datetime

from win32com.client import Dispatch

def speak(str1):
    speak=Dispatch(("SAPI.SpVoice"))
    speak.Speak(str1)

video=cv2.VideoCapture(0)
facedetect=cv2.CascadeClassifier(r'C:\Users\shukl\Desktop\AI_Project_1000\AI_Project_1000\Face Detection\data\haarcascade_frontalface_default.xml')

with open(r'C:\Users\shukl\Desktop\AI_Project_1000\AI_Project_1000\Face Detection\data\names.pkl','rb') as f:
    LABELS=pickle.load(f)
with open(r'C:\Users\shukl\Desktop\AI_Project_1000\AI_Project_1000\Face Detection\data\faces_data.pkl','rb') as f:
    FACES=pickle.load(f)

knn=KNeighborsClassifier(n_neighbors=5)
knn.fit(FACES, LABELS)

imgBackground=cv2.imread(r"C:\Users\shukl\Desktop\AI_Project_1000\AI_Project_1000\Face Detection\LOGO Background.jpg")

COL_NAMES = ['NAME', 'TIME']

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
        timestamp=datetime.fromtimestamp(ts).strftime("%H-%M-%S")
        exist=os.path.isfile("Attendance/Attendance_" + date + ".csv")
        cv2.rectangle(frame, (x,y), (x+w,y+h), (0,0,255), 1)
        cv2.rectangle(frame, (x,y), (x+w,y+h), (50,50,255), 2)
        cv2.rectangle(frame, (x,y-40), (x+w,y), (50,50,255), -1)
        cv2.putText(frame, str(output[0]), (x,y-15),cv2.FONT_HERSHEY_COMPLEX, 1, (255,255,255), 1)
        cv2.rectangle(frame, (x,y), (x+w, y+h), (50,50,255), 1)
        attendance=[str(output[0]), str(timestamp)]
    imgBackground[162:162 + 480, 55:55 + 640] = frame
    cv2.imshow("Frame",imgBackground)
    k=cv2.waitKey(1)
    if k==ord("A"):
        speak("Attendance Taken..")
        time.sleep(1)
        if exist:
            with open(r"C:\Users\shukl\Desktop\AI_Project_1000\AI_Project_1000\Face Detection\Attendance/Attendance_" + date + r".csv", r"+a" , w) as csvfile:
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

