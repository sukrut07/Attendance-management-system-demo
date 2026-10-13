import cv2
import pickle
import numpy as np
import os
import csv
import time
from datetime import datetime
from sklearn.neighbors import KNeighborsClassifier

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_DIR = os.path.dirname(BASE_DIR)
DATA_DIR = os.path.join(BASE_DIR, "data")
ATTENDANCE_DIR = os.path.join(BASE_DIR, "Attendance")
FILES_DIR = os.path.join(PROJECT_DIR, "Files")
os.makedirs(ATTENDANCE_DIR, exist_ok=True)


def speak(text):
    """Cross-platform text-to-speech speaker."""
    try:
        import pyttsx3
        engine = pyttsx3.init()
        engine.say(text)
        engine.runAndWait()
        return
    except Exception:
        pass

    if os.name != 'nt':
        try:
            import subprocess
            subprocess.Popen(['say', text])
            return
        except Exception:
            pass

    if os.name == 'nt':
        try:
            from win32com.client import Dispatch
            speaker = Dispatch("SAPI.SpVoice")
            speaker.Speak(text)
            return
        except Exception:
            pass

    print(f"[Voice Alert]: {text}")


# Load Haar Cascade
cascade_path = os.path.join(DATA_DIR, "haarcascade_frontalface_default.xml")
if not os.path.exists(cascade_path):
    cascade_path = os.path.join(cv2.data.haarcascades, "haarcascade_frontalface_default.xml")

facedetect = cv2.CascadeClassifier(cascade_path)
if facedetect.empty():
    raise RuntimeError(f"Could not load Haar Cascade from {cascade_path}")

# Load dataset
names_path = os.path.join(DATA_DIR, "names.pkl")
faces_path = os.path.join(DATA_DIR, "faces_data.pkl")
roll_path = os.path.join(DATA_DIR, "roll.pkl")

if not os.path.exists(names_path) or not os.path.exists(faces_path):
    print("Error: Face database not found. Please run add_faces.py to register faces first.")
    exit(1)

with open(names_path, 'rb') as f:
    LABELS = pickle.load(f)
with open(faces_path, 'rb') as f:
    FACES = pickle.load(f)

if len(LABELS) == 0 or len(FACES) == 0:
    print("Error: Face database is empty. Please register faces first using add_faces.py.")
    exit(1)

# Ensure sample count alignment
min_samples = min(len(LABELS), len(FACES))
LABELS = LABELS[:min_samples]
FACES = FACES[:min_samples]

# Build Roll Number mapping
name_to_roll = {}
if os.path.exists(roll_path):
    try:
        with open(roll_path, 'rb') as f:
            ROLL = pickle.load(f)
        for idx, n in enumerate(LABELS):
            if idx < len(ROLL):
                name_to_roll[n] = ROLL[idx]
    except Exception:
        pass

# Also supplement roll mapping from STUDENTS.csv if available
students_csv = os.path.join(FILES_DIR, "STUDENTS.csv")
if os.path.exists(students_csv):
    try:
        with open(students_csv, 'r', newline='', encoding='utf-8') as f:
            for row in csv.reader(f):
                if len(row) >= 3 and row[0] != "ROLL NO.":
                    full_name = f"{row[1]} {row[2]}".strip()
                    name_to_roll[full_name] = row[0]
    except Exception:
        pass

# Train KNN classifier
n_neighbors = min(5, len(LABELS))
knn = KNeighborsClassifier(n_neighbors=n_neighbors)
knn.fit(FACES, LABELS)

# Background image
bg_path = os.path.join(BASE_DIR, "LOGO Background.jpg")
imgBackground_orig = cv2.imread(bg_path)
if imgBackground_orig is None:
    imgBackground_orig = np.zeros((720, 960, 3), dtype=np.uint8)

COL_NAMES = ['ROLL NO.', 'NAME', 'TIME']

video = cv2.VideoCapture(0)
if not video.isOpened():
    print("Error: Could not open webcam. Check camera connection or permissions.")
    exit(1)

attendance = None
last_marked_person = None
last_marked_time = 0

window_title = "Attendance System - Face Recognition"
cv2.namedWindow(window_title, cv2.WINDOW_NORMAL)
cv2.resizeWindow(window_title, 1100, 650)

print("Face Recognition Attendance System is active.")
print("Press 'a' or 'A' to mark attendance.")
print("Press 'q' or 'Q' to quit.\n")

while True:
    ret, frame = video.read()
    if not ret:
        continue

    grey = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
    faces = facedetect.detectMultiScale(grey, 1.3, 5)

    current_person = None
    for (x, y, w, h) in faces:
        y_start = max(0, y)
        y_end = min(frame.shape[0], y + h)
        x_start = max(0, x)
        x_end = min(frame.shape[1], x + w)

        crop_img = frame[y_start:y_end, x_start:x_end, :]
        if crop_img.size == 0 or crop_img.shape[0] == 0 or crop_img.shape[1] == 0:
            continue

        resized_img = cv2.resize(crop_img, (50, 50)).flatten().reshape(1, -1)
        output = knn.predict(resized_img)
        current_person = str(output[0])

        timestamp = datetime.now().strftime("%H:%M:%S")
        person_roll = str(name_to_roll.get(current_person, "N/A"))

        cv2.rectangle(frame, (x_start, y_start), (x_end, y_end), (50, 50, 255), 2)
        cv2.rectangle(frame, (x_start, max(0, y_start - 35)), (x_end, y_start), (50, 50, 255), -1)
        cv2.putText(frame, f"{current_person} ({person_roll})", (x_start, max(15, y_start - 10)),
                    cv2.FONT_HERSHEY_COMPLEX, 0.65, (255, 255, 255), 1)

        attendance = [person_roll, current_person, timestamp]

    # Resize camera frame to 640x480 for background template
    frame_resized = cv2.resize(frame, (640, 480))
    imgBackground = imgBackground_orig.copy()
    if imgBackground.shape[0] >= 642 and imgBackground.shape[1] >= 695:
        imgBackground[162:162 + 480, 55:55 + 640] = frame_resized
        cv2.imshow(window_title, imgBackground)
    else:
        cv2.imshow(window_title, frame)

    k = cv2.waitKey(1) & 0xFF
    if k in [ord("a"), ord("A")]:
        if attendance is None:
            print("No face detected! Look at camera and try again.")
            speak("No face detected")
            continue

        now = time.time()
        # Prevent rapid duplicate logging within 2 seconds
        if last_marked_person == attendance[1] and (now - last_marked_time) < 2.0:
            print(f"Attendance already marked for {attendance[1]} recently.")
            continue

        date_str = datetime.now().strftime("%d-%m-%Y")
        csv_file_path = os.path.join(ATTENDANCE_DIR, f"Attendance_{date_str}.csv")
        file_exists = os.path.isfile(csv_file_path)

        with open(csv_file_path, "a", newline="", encoding="utf-8") as csvfile:
            writer = csv.writer(csvfile)
            if not file_exists:
                writer.writerow(COL_NAMES)
            writer.writerow(attendance)

        last_marked_person = attendance[1]
        last_marked_time = now

        print(f"Attendance recorded: Roll {attendance[0]} | {attendance[1]} at {attendance[2]}")
        speak("Attendance Taken..")

    if k in [ord("q"), ord("Q")]:
        break

video.release()
cv2.destroyAllWindows()
