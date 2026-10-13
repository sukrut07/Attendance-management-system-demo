import cv2
import pickle
import numpy as np
import os
import csv
import time
from datetime import datetime
from sklearn.neighbors import KNeighborsClassifier

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
ATTENDANCE_DIR = os.path.join(BASE_DIR, "Attendance")
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

    # macOS native TTS
    if os.name != 'nt':
        try:
            import subprocess
            subprocess.Popen(['say', text])
            return
        except Exception:
            pass

    # Windows SAPI fallback
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
    print("Error: Face data not found. Please run add_faces.py to register faces first.")
    exit(1)

with open(names_path, 'rb') as f:
    LABELS = pickle.load(f)
with open(faces_path, 'rb') as f:
    FACES = pickle.load(f)

# Optional Roll Number mapping
ROLL = None
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

# Train KNN classifier
n_neighbors = min(5, len(LABELS))
knn = KNeighborsClassifier(n_neighbors=n_neighbors)
knn.fit(FACES, LABELS)

# Background image
bg_path = os.path.join(BASE_DIR, "LOGO Background.jpg")
imgBackground_orig = cv2.imread(bg_path)
if imgBackground_orig is None:
    # Fallback canvas if background image is missing
    imgBackground_orig = np.zeros((720, 960, 3), dtype=np.uint8)

# CSV Header
has_roll = len(name_to_roll) > 0
COL_NAMES = ['ROLL NO.', 'NAME', 'TIME'] if has_roll else ['NAME', 'TIME']

video = cv2.VideoCapture(0)
if not video.isOpened():
    print("Error: Could not open webcam.")
    exit(1)

attendance = None
print("Face Recognition Attendance System running...")
print("Press 'a' or 'A' to take attendance.")
print("Press 'q' or 'Q' to quit.")

while True:
    ret, frame = video.read()
    if not ret:
        print("Warning: Failed to capture frame from webcam.")
        continue

    grey = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
    faces = facedetect.detectMultiScale(grey, 1.3, 5)

    current_person = None
    for (x, y, w, h) in faces:
        crop_img = frame[y:y+h, x:x+w, :]
        resized_img = cv2.resize(crop_img, (50, 50)).flatten().reshape(1, -1)
        output = knn.predict(resized_img)
        current_person = str(output[0])

        ts = time.time()
        timestamp = datetime.fromtimestamp(ts).strftime("%H:%M:%S")

        # Bounding boxes and labels
        cv2.rectangle(frame, (x, y), (x+w, y+h), (50, 50, 255), 2)
        cv2.rectangle(frame, (x, y-40), (x+w, y), (50, 50, 255), -1)
        cv2.putText(frame, current_person, (x, y-15),
                    cv2.FONT_HERSHEY_COMPLEX, 0.8, (255, 255, 255), 1)

        if has_roll:
            person_roll = str(name_to_roll.get(current_person, "N/A"))
            attendance = [person_roll, current_person, timestamp]
        else:
            attendance = [current_person, timestamp]

    # Resize camera frame to 640x480 to fit into background template
    frame_resized = cv2.resize(frame, (640, 480))
    imgBackground = imgBackground_orig.copy()
    if imgBackground.shape[0] >= 642 and imgBackground.shape[1] >= 695:
        imgBackground[162:162 + 480, 55:55 + 640] = frame_resized
        cv2.imshow("Attendance System", imgBackground)
    else:
        cv2.imshow("Attendance System", frame)

    k = cv2.waitKey(1) & 0xFF
    if k in [ord("a"), ord("A")]:
        if attendance is None:
            print("No face detected! Look at camera and try again.")
            speak("No face detected")
            continue

        date_str = datetime.now().strftime("%d-%m-%Y")
        csv_file_path = os.path.join(ATTENDANCE_DIR, f"Attendance_{date_str}.csv")
        file_exists = os.path.isfile(csv_file_path)

        with open(csv_file_path, "a", newline="", encoding="utf-8") as csvfile:
            writer = csv.writer(csvfile)
            if not file_exists:
                writer.writerow(COL_NAMES)
            writer.writerow(attendance)

        print(f"Attendance recorded for: {attendance}")
        speak("Attendance Taken..")
        time.sleep(0.5)

    if k in [ord("q"), ord("Q")]:
        break

video.release()
cv2.destroyAllWindows()
