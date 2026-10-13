import cv2
import pickle
import numpy as np
import os
import csv

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_DIR = os.path.dirname(BASE_DIR)
DATA_DIR = os.path.join(BASE_DIR, "data")
FILES_DIR = os.path.join(PROJECT_DIR, "Files")
os.makedirs(DATA_DIR, exist_ok=True)
os.makedirs(FILES_DIR, exist_ok=True)

cascade_path = os.path.join(DATA_DIR, "haarcascade_frontalface_default.xml")
if not os.path.exists(cascade_path):
    cascade_path = os.path.join(cv2.data.haarcascades, "haarcascade_frontalface_default.xml")

facedetect = cv2.CascadeClassifier(cascade_path)
if facedetect.empty():
    raise RuntimeError(f"Failed to load Haar Cascade classifier from {cascade_path}")

video = cv2.VideoCapture(0)
if not video.isOpened():
    print("Error: Could not open webcam. Check camera connection or permissions.")
    exit(1)

name = input("Enter Student Full Name: ").strip()
while not name:
    name = input("Name cannot be empty. Enter Student Full Name: ").strip()

roll = input("Enter Student Roll Number: ").strip()
while not roll:
    roll = input("Roll Number cannot be empty. Enter Roll Number: ").strip()

faces_data = []
i = 0
TOTAL_SAMPLES = 100

print(f"\nCollecting face samples for '{name}' (Roll: {roll}).")
print("Look directly at the camera. Press 'q' to quit early.\n")

while True:
    ret, frame = video.read()
    if not ret:
        print("Warning: Failed to grab frame from webcam.")
        continue

    grey = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
    faces = facedetect.detectMultiScale(grey, 1.3, 5)

    for (x, y, w, h) in faces:
        # Safe bounding box slicing to prevent negative or out-of-frame coordinates
        y_start = max(0, y)
        y_end = min(frame.shape[0], y + h)
        x_start = max(0, x)
        x_end = min(frame.shape[1], x + w)

        crop_img = frame[y_start:y_end, x_start:x_end, :]
        if crop_img.size == 0 or crop_img.shape[0] == 0 or crop_img.shape[1] == 0:
            continue

        resized_img = cv2.resize(crop_img, (50, 50))
        if len(faces_data) < TOTAL_SAMPLES and i % 10 == 0:
            faces_data.append(resized_img)
        i = i + 1

        cv2.putText(frame, f"Samples: {len(faces_data)}/{TOTAL_SAMPLES}", (30, 50),
                    cv2.FONT_HERSHEY_COMPLEX, 0.8, (50, 50, 255), 2)
        cv2.rectangle(frame, (x_start, y_start), (x_end, y_end), (50, 50, 255), 2)

    cv2.imshow("Collecting Face Data", frame)
    k = cv2.waitKey(1) & 0xFF
    if k == ord("q") or len(faces_data) >= TOTAL_SAMPLES:
        break

video.release()
cv2.destroyAllWindows()

if len(faces_data) == 0:
    print("No face samples collected. Exiting without saving.")
    exit(0)

print(f"Collected {len(faces_data)} samples.")
faces_data = np.asarray(faces_data)
faces_data = faces_data.reshape(len(faces_data), -1)

# Save or append names.pkl
names_file = os.path.join(DATA_DIR, "names.pkl")
current_names = [name] * len(faces_data)
if not os.path.exists(names_file):
    with open(names_file, "wb") as f:
        pickle.dump(current_names, f)
else:
    with open(names_file, "rb") as f:
        existing_names = pickle.load(f)
    existing_names = existing_names + current_names
    with open(names_file, "wb") as f:
        pickle.dump(existing_names, f)

# Save or append faces_data.pkl
faces_file = os.path.join(DATA_DIR, "faces_data.pkl")
if not os.path.exists(faces_file):
    with open(faces_file, "wb") as f:
        pickle.dump(faces_data, f)
else:
    with open(faces_file, "rb") as f:
        existing_faces = pickle.load(f)
    if len(existing_faces.shape) == 1:
        existing_faces = existing_faces.reshape(-1, 7500)
    existing_faces = np.append(existing_faces, faces_data, axis=0)
    with open(faces_file, "wb") as f:
        pickle.dump(existing_faces, f)

# Save or append roll.pkl
roll_file = os.path.join(DATA_DIR, "roll.pkl")
current_roll = [roll] * len(faces_data)
if not os.path.exists(roll_file):
    with open(roll_file, "wb") as f:
        pickle.dump(current_roll, f)
else:
    with open(roll_file, "rb") as f:
        existing_roll = pickle.load(f)
    existing_roll = existing_roll + current_roll
    with open(roll_file, "wb") as f:
        pickle.dump(existing_roll, f)

# Also sync to STUDENTS.csv if available
students_csv = os.path.join(FILES_DIR, "STUDENTS.csv")
file_exists = os.path.exists(students_csv)
name_parts = name.split(None, 1)
fname = name_parts[0]
lname = name_parts[1] if len(name_parts) > 1 else ""

with open(students_csv, "a", newline="", encoding="utf-8") as f:
    writer = csv.writer(f)
    if not file_exists:
        writer.writerow(["ROLL NO.", "FIRST NAME", "LAST NAME"])
    writer.writerow([roll, fname, lname])

print(f"Successfully registered student '{name}' (Roll: {roll}).")