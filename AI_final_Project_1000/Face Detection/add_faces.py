import cv2
import pickle
import numpy as np
import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
os.makedirs(DATA_DIR, exist_ok=True)

cascade_path = os.path.join(DATA_DIR, "haarcascade_frontalface_default.xml")
if not os.path.exists(cascade_path):
    cascade_path = os.path.join(cv2.data.haarcascades, "haarcascade_frontalface_default.xml")

facedetect = cv2.CascadeClassifier(cascade_path)
if facedetect.empty():
    raise RuntimeError(f"Failed to load Haar Cascade classifier from {cascade_path}")

video = cv2.VideoCapture(0)
if not video.isOpened():
    print("Error: Could not open webcam.")
    exit(1)

name = input("Enter Your Name: ").strip()
while not name:
    name = input("Name cannot be empty. Enter Your Name: ").strip()

roll = input("Enter Your Roll Number: ").strip()

faces_data = []
i = 0
TOTAL_SAMPLES = 100

print("Collecting face samples. Look at the camera. Press 'q' to quit early.")

while True:
    ret, frame = video.read()
    if not ret:
        print("Error: Failed to grab frame from webcam.")
        break

    grey = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
    faces = facedetect.detectMultiScale(grey, 1.3, 5)

    for (x, y, w, h) in faces:
        crop_img = frame[y:y+h, x:x+w, :]
        resized_img = cv2.resize(crop_img, (50, 50))
        if len(faces_data) < TOTAL_SAMPLES and i % 10 == 0:
            faces_data.append(resized_img)
        i = i + 1

        cv2.putText(frame, f"Samples: {len(faces_data)}/{TOTAL_SAMPLES}", (30, 50),
                    cv2.FONT_HERSHEY_COMPLEX, 0.8, (50, 50, 255), 2)
        cv2.rectangle(frame, (x, y), (x+w, y+h), (50, 50, 255), 2)

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

# Save or append names
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

# Save or append faces_data
faces_file = os.path.join(DATA_DIR, "faces_data.pkl")
if not os.path.exists(faces_file):
    with open(faces_file, "wb") as f:
        pickle.dump(faces_data, f)
else:
    with open(faces_file, "rb") as f:
        existing_faces = pickle.load(f)
    existing_faces = np.append(existing_faces, faces_data, axis=0)
    with open(faces_file, "wb") as f:
        pickle.dump(existing_faces, f)

# Save or append roll
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

print(f"Successfully saved face data for '{name}' (Roll: {roll}).")