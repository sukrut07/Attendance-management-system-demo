import sys
import os
import csv
import time
import pickle
from datetime import datetime

from PyQt5.QtCore import Qt, QAbstractTableModel
from PyQt5.QtWidgets import (
    QApplication, QWidget, QMainWindow, QLabel, QPushButton,
    QLineEdit, QGridLayout, QVBoxLayout, QHBoxLayout, QMessageBox,
    QTableView
)
from PyQt5.QtGui import QFont, QIcon, QPixmap, QCursor
import cv2
import numpy as np
from sklearn.neighbors import KNeighborsClassifier

# --- Project Paths ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
FILES_DIR = os.path.join(BASE_DIR, "Files")
FACE_DETECTION_DIR = os.path.join(BASE_DIR, "Face Detection")
DATA_DIR = os.path.join(FACE_DETECTION_DIR, "data")
ATTENDANCE_DIR = os.path.join(FACE_DETECTION_DIR, "Attendance")

os.makedirs(FILES_DIR, exist_ok=True)
os.makedirs(DATA_DIR, exist_ok=True)
os.makedirs(ATTENDANCE_DIR, exist_ok=True)

ICON_PATH = os.path.join(FILES_DIR, "ICON.png")
PERFECT_PATH = os.path.join(FILES_DIR, "PERFECT.png")
STUDENTS_CSV = os.path.join(FILES_DIR, "STUDENTS.csv")
DATE_CSV = os.path.join(FILES_DIR, "DAte.csv")

HAAR_CASCADE_PATH = os.path.join(DATA_DIR, "haarcascade_frontalface_default.xml")
if not os.path.exists(HAAR_CASCADE_PATH):
    HAAR_CASCADE_PATH = os.path.join(cv2.data.haarcascades, "haarcascade_frontalface_default.xml")

NAMES_PKL = os.path.join(DATA_DIR, "names.pkl")
FACES_PKL = os.path.join(DATA_DIR, "faces_data.pkl")
ROLL_PKL = os.path.join(DATA_DIR, "roll.pkl")
BG_IMAGE_PATH = os.path.join(FACE_DETECTION_DIR, "LOGO Background.jpg")


def speak(text):
    """Cross-platform text-to-speech engine."""
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
            spk = Dispatch("SAPI.SpVoice")
            spk.Speak(text)
            return
        except Exception:
            pass

    print(f"[Speech]: {text}")


def get_pixmap_or_fallback(file_path, fallback_path=None):
    """Returns a QPixmap for a file if it exists, otherwise falls back gracefully."""
    if os.path.exists(file_path):
        return QPixmap(file_path)
    if fallback_path and os.path.exists(fallback_path):
        return QPixmap(fallback_path)
    # Blank 100x100 pixmap
    blank = QPixmap(100, 100)
    blank.fill(Qt.transparent)
    return blank


class HomePage(QWidget):
    def __init__(self):
        super(HomePage, self).__init__()
        self.setWindowTitle("ATTENDANCE SYSTEM")
        self.setFixedWidth(1000)
        self.setStyleSheet("background: #FFFFFF;")
        if os.path.exists(ICON_PATH):
            self.setWindowIcon(QIcon(ICON_PATH))

        self.child_window = None
        self.initUI()

    def open1(self):
        self.child_window = SecondPage()
        self.child_window.show()
        self.hide()

    def open4(self):
        self.child_window = ThirdPage()
        self.child_window.show()
        self.hide()

    def open5(self):
        self.child_window = FourthPage()
        self.child_window.show()
        self.hide()

    def open8(self):
        self.child_window = FifthPage()
        self.child_window.show()
        self.hide()

    def initUI(self):
        grid = QGridLayout()
        self.setLayout(grid)

        self.logo1 = QLabel()
        self.logo1.setPixmap(get_pixmap_or_fallback(PERFECT_PATH))
        self.logo1.setAlignment(Qt.AlignHCenter)
        self.logo1.setStyleSheet("margin-top: 10px;")

        self.l1 = QLabel()
        self.l1.setText(
            "Our attendance detection program makes it easy for us to track attendance accurately.\n"
            "With facial recognition, we automatically record who's present without the hassle of manual tracking.\n"
            "Perfect for managing a classroom, a workplace, or an event.\n"
            "This tool keeps reliable records and saves time, making attendance management simple and efficient."
        )
        self.l1.setFont(QFont('Segoe UI, Bahnschrift, Arial', 13))
        self.l1.setAlignment(Qt.AlignCenter)

        self.l2 = QLabel(" ")
        self.l2.setFont(QFont('Arial', 11))

        self.l3 = QLabel("Click the button to register student details")
        self.l3.setFont(QFont('Segoe UI, Bahnschrift, Arial', 12))
        self.l3.setAlignment(Qt.AlignCenter)

        btn_style = (
            "QPushButton { border: 4px solid #cbad8f; border-radius: 15px; font-size: 18px; "
            "color: #796648; padding: 8px 0; margin: 8px 20px; font-weight: bold; background: #FFFFFF; } "
            "QPushButton:hover { background: #cbad8f; color: #FFFFFF; }"
        )

        self.button1 = QPushButton("Fill Details Here")
        self.button1.setCursor(QCursor(Qt.PointingHandCursor))
        self.button1.setStyleSheet(btn_style)
        self.button1.clicked.connect(self.open1)

        self.l4 = QLabel("Click the button to start Face Recognition Attendance")
        self.l4.setFont(QFont('Segoe UI, Bahnschrift, Arial', 12))
        self.l4.setAlignment(Qt.AlignCenter)

        self.button2 = QPushButton("Give Attendance Here")
        self.button2.setCursor(QCursor(Qt.PointingHandCursor))
        self.button2.setStyleSheet(btn_style)
        self.button2.clicked.connect(self.open4)

        self.l5 = QLabel("Click the button to see attendance records")
        self.l5.setFont(QFont('Segoe UI, Bahnschrift, Arial', 12))
        self.l5.setAlignment(Qt.AlignCenter)

        self.button3 = QPushButton("See Attendance Here")
        self.button3.setCursor(QCursor(Qt.PointingHandCursor))
        self.button3.setStyleSheet(btn_style)
        self.button3.clicked.connect(self.open5)

        self.l23 = QLabel("Click to view developers information")
        self.l23.setFont(QFont('Segoe UI, Bahnschrift, Arial', 12))
        self.l23.setAlignment(Qt.AlignCenter)

        self.button10 = QPushButton("Developers Info")
        self.button10.setCursor(QCursor(Qt.PointingHandCursor))
        self.button10.setStyleSheet(btn_style)
        self.button10.clicked.connect(self.open8)

        grid.addWidget(self.logo1, 0, 0)
        grid.addWidget(self.l1, 1, 0)
        grid.addWidget(self.l2, 2, 0)
        grid.addWidget(self.l3, 3, 0)
        grid.addWidget(self.button1, 4, 0)
        grid.addWidget(self.l4, 5, 0)
        grid.addWidget(self.button2, 6, 0)
        grid.addWidget(self.l5, 7, 0)
        grid.addWidget(self.button3, 8, 0)
        grid.addWidget(self.l23, 9, 0)
        grid.addWidget(self.button10, 10, 0)


class SecondPage(QWidget):
    def __init__(self):
        super(SecondPage, self).__init__()
        self.setWindowTitle("ATTENDANCE SYSTEM - Register Student")
        self.setFixedWidth(1000)
        self.setStyleSheet("background: #FFFFFF;")
        if os.path.exists(ICON_PATH):
            self.setWindowIcon(QIcon(ICON_PATH))

        self.child_window = None
        self.initUI()

    def open2(self):
        self.child_window = S2_1()
        self.child_window.show()
        self.hide()

    def open_home(self):
        self.child_window = HomePage()
        self.child_window.show()
        self.hide()

    def initUI(self):
        grid = QGridLayout()
        self.setLayout(grid)

        self.logo2 = QLabel()
        self.logo2.setPixmap(get_pixmap_or_fallback(PERFECT_PATH))
        self.logo2.setAlignment(Qt.AlignHCenter)
        self.logo2.setStyleSheet("margin-top: 10px;")

        self.l6 = QLabel("Enter Student Details")
        self.l6.setFont(QFont('Segoe UI, Bahnschrift, Arial', 16, QFont.Bold))
        self.l6.setAlignment(Qt.AlignCenter)

        label_font = QFont('Segoe UI, Bahnschrift, Arial', 13)
        input_style = "QLineEdit { border: 2px solid #cbad8f; border-radius: 8px; padding: 6px; font-size: 15px; margin: 4px 20px; }"

        self.l8 = QLabel("First Name:")
        self.l8.setFont(label_font)
        self.l8.setStyleSheet("margin-left: 20px;")

        self.fname = QLineEdit()
        self.fname.setStyleSheet(input_style)
        self.fname.setPlaceholderText("Enter First Name")

        self.l9 = QLabel("Last Name:")
        self.l9.setFont(label_font)
        self.l9.setStyleSheet("margin-left: 20px;")

        self.lname = QLineEdit()
        self.lname.setStyleSheet(input_style)
        self.lname.setPlaceholderText("Enter Last Name")

        self.l10 = QLabel("Roll Number:")
        self.l10.setFont(label_font)
        self.l10.setStyleSheet("margin-left: 20px;")

        self.rno = QLineEdit()
        self.rno.setStyleSheet(input_style)
        self.rno.setPlaceholderText("Enter Roll Number")

        btn_style = (
            "QPushButton { border: 4px solid #cbad8f; border-radius: 15px; font-size: 18px; "
            "color: #796648; padding: 8px 0; margin: 10px 20px; font-weight: bold; background: #FFFFFF; } "
            "QPushButton:hover { background: #cbad8f; color: #FFFFFF; }"
        )

        self.button4 = QPushButton("Submit Details")
        self.button4.setCursor(QCursor(Qt.PointingHandCursor))
        self.button4.setStyleSheet(btn_style)
        self.button4.clicked.connect(self.addinfo)

        self.button_back = QPushButton("Back to Home")
        self.button_back.setCursor(QCursor(Qt.PointingHandCursor))
        self.button_back.setStyleSheet(btn_style)
        self.button_back.clicked.connect(self.open_home)

        grid.addWidget(self.logo2, 0, 0)
        grid.addWidget(self.l6, 1, 0)
        grid.addWidget(self.l8, 2, 0)
        grid.addWidget(self.fname, 3, 0)
        grid.addWidget(self.l9, 4, 0)
        grid.addWidget(self.lname, 5, 0)
        grid.addWidget(self.l10, 6, 0)
        grid.addWidget(self.rno, 7, 0)
        grid.addWidget(self.button4, 8, 0)
        grid.addWidget(self.button_back, 9, 0)

    def addinfo(self):
        fname = self.fname.text().strip()
        lname = self.lname.text().strip()
        rno = self.rno.text().strip()

        if not fname or not lname or not rno:
            QMessageBox.warning(self, "Missing Information", "All fields (First Name, Last Name, Roll Number) are required!")
            return

        # Append to STUDENTS.csv
        file_exists = os.path.exists(STUDENTS_CSV)
        with open(STUDENTS_CSV, "a", newline="", encoding="utf-8") as temp:
            writer = csv.writer(temp)
            if not file_exists:
                writer.writerow(["ROLL NO.", "FIRST NAME", "LAST NAME"])
            writer.writerow([rno, fname, lname])

        self.open2()


class S2_1(QWidget):
    def __init__(self):
        super(S2_1, self).__init__()
        self.setWindowTitle("ATTENDANCE SYSTEM - Capture Images")
        self.setFixedWidth(1000)
        self.setStyleSheet("background: #FFFFFF;")
        if os.path.exists(ICON_PATH):
            self.setWindowIcon(QIcon(ICON_PATH))

        self.child_window = None
        self.initUI()

    def open3(self):
        self.child_window = HomePage()
        self.child_window.show()
        self.hide()

    def initUI(self):
        grid = QGridLayout()
        self.setLayout(grid)

        self.logo3 = QLabel()
        self.logo3.setPixmap(get_pixmap_or_fallback(PERFECT_PATH))
        self.logo3.setAlignment(Qt.AlignHCenter)
        self.logo3.setStyleSheet("margin-top: 10px;")

        self.l11 = QLabel("Details Uploaded Successfully!")
        self.l11.setFont(QFont('Segoe UI, Bahnschrift, Arial', 16, QFont.Bold))
        self.l11.setAlignment(Qt.AlignCenter)
        self.l11.setStyleSheet("color: #2e7d32;")

        self.l13 = QLabel("Press the button below to capture 100 face images of the student via webcam.\nMake sure your face is clearly visible.")
        self.l13.setFont(QFont('Segoe UI, Bahnschrift, Arial', 13))
        self.l13.setAlignment(Qt.AlignCenter)

        btn_style = (
            "QPushButton { border: 4px solid #cbad8f; border-radius: 15px; font-size: 18px; "
            "color: #796648; padding: 8px 0; margin: 10px 20px; font-weight: bold; background: #FFFFFF; } "
            "QPushButton:hover { background: #cbad8f; color: #FFFFFF; }"
        )

        self.button5 = QPushButton("Take Images Now")
        self.button5.setCursor(QCursor(Qt.PointingHandCursor))
        self.button5.setStyleSheet(btn_style)
        self.button5.clicked.connect(self.take_image)

        self.button6 = QPushButton("Home")
        self.button6.setCursor(QCursor(Qt.PointingHandCursor))
        self.button6.setStyleSheet(btn_style)
        self.button6.clicked.connect(self.open3)

        grid.addWidget(self.logo3, 0, 0)
        grid.addWidget(self.l11, 1, 0)
        grid.addWidget(self.l13, 2, 0)
        grid.addWidget(self.button5, 3, 0)
        grid.addWidget(self.button6, 4, 0)

    def take_image(self):
        # Read the latest student from STUDENTS.csv
        if not os.path.exists(STUDENTS_CSV):
            QMessageBox.warning(self, "Error", "No student details found! Please enter details first.")
            return

        rows = []
        with open(STUDENTS_CSV, mode='r', newline='', encoding='utf-8') as f:
            reader = csv.reader(f)
            for r in reader:
                if r:
                    rows.append(r)

        # Filter out header if present
        data_rows = [r for r in rows if r[0] != "ROLL NO."]
        if not data_rows:
            QMessageBox.warning(self, "Error", "No student details found in file! Please enter details first.")
            return

        last_student = data_rows[-1]
        roll_no = last_student[0]
        student_name = f"{last_student[1]} {last_student[2]}".strip()

        # Load Haar Cascade
        facedetect = cv2.CascadeClassifier(HAAR_CASCADE_PATH)
        if facedetect.empty():
            QMessageBox.critical(self, "Error", f"Failed to load Haar Cascade from {HAAR_CASCADE_PATH}")
            return

        video = cv2.VideoCapture(0)
        if not video.isOpened():
            QMessageBox.critical(self, "Error", "Could not open webcam. Please verify camera permissions.")
            return

        faces_data = []
        i = 0
        total_samples = 100

        while True:
            ret, frame = video.read()
            if not ret:
                break

            grey = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
            faces = facedetect.detectMultiScale(grey, 1.3, 5)

            for (x, y, w, h) in faces:
                crop_img = frame[y:y+h, x:x+w, :]
                resized_img = cv2.resize(crop_img, (50, 50))
                if len(faces_data) < total_samples and i % 10 == 0:
                    faces_data.append(resized_img)
                i += 1

                cv2.putText(frame, f"Samples: {len(faces_data)}/{total_samples}", (30, 50),
                            cv2.FONT_HERSHEY_COMPLEX, 0.8, (50, 50, 255), 2)
                cv2.rectangle(frame, (x, y), (x+w, y+h), (50, 50, 255), 2)

            cv2.imshow("Capturing Student Face - Press 'q' to Quit", frame)
            k = cv2.waitKey(1) & 0xFF
            if k == ord("q") or len(faces_data) >= total_samples:
                break

        video.release()
        cv2.destroyAllWindows()

        if len(faces_data) == 0:
            QMessageBox.warning(self, "Warning", "No face samples were collected.")
            return

        faces_data = np.asarray(faces_data)
        faces_data = faces_data.reshape(len(faces_data), -1)

        # Append names.pkl
        current_names = [student_name] * len(faces_data)
        if not os.path.exists(NAMES_PKL):
            with open(NAMES_PKL, 'wb') as f:
                pickle.dump(current_names, f)
        else:
            with open(NAMES_PKL, 'rb') as f:
                names = pickle.load(f)
            names = names + current_names
            with open(NAMES_PKL, 'wb') as f:
                pickle.dump(names, f)

        # Append faces_data.pkl
        if not os.path.exists(FACES_PKL):
            with open(FACES_PKL, 'wb') as f:
                pickle.dump(faces_data, f)
        else:
            with open(FACES_PKL, 'rb') as f:
                faces = pickle.load(f)
            faces = np.append(faces, faces_data, axis=0)
            with open(FACES_PKL, 'wb') as f:
                pickle.dump(faces, f)

        # Append roll.pkl
        current_rolls = [roll_no] * len(faces_data)
        if not os.path.exists(ROLL_PKL):
            with open(ROLL_PKL, 'wb') as f:
                pickle.dump(current_rolls, f)
        else:
            with open(ROLL_PKL, 'rb') as f:
                rolls = pickle.load(f)
            rolls = rolls + current_rolls
            with open(ROLL_PKL, 'wb') as f:
                pickle.dump(rolls, f)

        QMessageBox.information(
            self, "Success",
            f"Successfully captured {len(faces_data)} images for {student_name} (Roll: {roll_no})!"
        )


class ThirdPage(QWidget):
    def __init__(self):
        super(ThirdPage, self).__init__()
        self.setWindowTitle("ATTENDANCE SYSTEM - Live Attendance")
        self.setFixedWidth(1000)
        self.setStyleSheet("background: #FFFFFF;")
        if os.path.exists(ICON_PATH):
            self.setWindowIcon(QIcon(ICON_PATH))

        self.child_window = None
        self.initUI()

    def open_home(self):
        self.child_window = HomePage()
        self.child_window.show()
        self.hide()

    def initUI(self):
        grid = QGridLayout()
        self.setLayout(grid)

        self.logo4 = QLabel()
        self.logo4.setPixmap(get_pixmap_or_fallback(PERFECT_PATH))
        self.logo4.setAlignment(Qt.AlignHCenter)
        self.logo4.setStyleSheet("margin-top: 10px;")

        self.title = QLabel("Live Facial Recognition Attendance")
        self.title.setFont(QFont('Segoe UI, Bahnschrift, Arial', 16, QFont.Bold))
        self.title.setAlignment(Qt.AlignCenter)

        self.instructions = QLabel(
            "Instructions:\n"
            "1. Click 'Start Live Attendance' to open camera.\n"
            "2. Face the camera until your name is recognized.\n"
            "3. Press 'A' or 'a' on your keyboard to record attendance.\n"
            "4. Press 'Q' or 'q' to close camera and return."
        )
        self.instructions.setFont(QFont('Segoe UI, Bahnschrift, Arial', 13))
        self.instructions.setAlignment(Qt.AlignCenter)

        btn_style = (
            "QPushButton { border: 4px solid #cbad8f; border-radius: 15px; font-size: 18px; "
            "color: #796648; padding: 8px 0; margin: 10px 20px; font-weight: bold; background: #FFFFFF; } "
            "QPushButton:hover { background: #cbad8f; color: #FFFFFF; }"
        )

        self.btn_start = QPushButton("Start Live Attendance")
        self.btn_start.setCursor(QCursor(Qt.PointingHandCursor))
        self.btn_start.setStyleSheet(btn_style)
        self.btn_start.clicked.connect(self.run_live_attendance)

        self.btn_home = QPushButton("Back to Home")
        self.btn_home.setCursor(QCursor(Qt.PointingHandCursor))
        self.btn_home.setStyleSheet(btn_style)
        self.btn_home.clicked.connect(self.open_home)

        grid.addWidget(self.logo4, 0, 0)
        grid.addWidget(self.title, 1, 0)
        grid.addWidget(self.instructions, 2, 0)
        grid.addWidget(self.btn_start, 3, 0)
        grid.addWidget(self.btn_home, 4, 0)

    def run_live_attendance(self):
        # Verify required datasets exist
        if not os.path.exists(NAMES_PKL) or not os.path.exists(FACES_PKL):
            QMessageBox.warning(self, "No Data", "Face database is empty! Please register faces first.")
            return

        try:
            with open(NAMES_PKL, 'rb') as f:
                LABELS = pickle.load(f)
            with open(FACES_PKL, 'rb') as f:
                FACES = pickle.load(f)
        except Exception as e:
            QMessageBox.critical(self, "Data Error", f"Failed to load face data: {e}")
            return

        if len(LABELS) == 0 or len(FACES) == 0:
            QMessageBox.warning(self, "No Data", "Face database is empty! Please register faces first.")
            return

        # Build name -> roll mapping
        name_to_roll = {}
        if os.path.exists(ROLL_PKL):
            try:
                with open(ROLL_PKL, 'rb') as f:
                    ROLL = pickle.load(f)
                for idx, n in enumerate(LABELS):
                    if idx < len(ROLL):
                        name_to_roll[n] = ROLL[idx]
            except Exception:
                pass

        # Also pull from STUDENTS.csv if available
        if os.path.exists(STUDENTS_CSV):
            try:
                with open(STUDENTS_CSV, 'r', newline='', encoding='utf-8') as f:
                    for row in csv.reader(f):
                        if len(row) >= 3 and row[0] != "ROLL NO.":
                            full = f"{row[1]} {row[2]}".strip()
                            name_to_roll[full] = row[0]
            except Exception:
                pass

        # Train KNN
        n_neighbors = min(5, len(LABELS))
        knn = KNeighborsClassifier(n_neighbors=n_neighbors)
        knn.fit(FACES, LABELS)

        facedetect = cv2.CascadeClassifier(HAAR_CASCADE_PATH)
        if facedetect.empty():
            QMessageBox.critical(self, "Error", f"Failed to load Haar Cascade from {HAAR_CASCADE_PATH}")
            return

        imgBackground_orig = cv2.imread(BG_IMAGE_PATH)
        if imgBackground_orig is None:
            imgBackground_orig = np.zeros((720, 960, 3), dtype=np.uint8)

        video = cv2.VideoCapture(0)
        if not video.isOpened():
            QMessageBox.critical(self, "Error", "Could not open webcam.")
            return

        COL_NAMES = ['ROLL NO.', 'NAME', 'TIME']
        attendance = None

        while True:
            ret, frame = video.read()
            if not ret:
                continue

            grey = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
            faces = facedetect.detectMultiScale(grey, 1.3, 5)

            current_name = None
            for (x, y, w, h) in faces:
                crop_img = frame[y:y+h, x:x+w, :]
                resized_img = cv2.resize(crop_img, (50, 50)).flatten().reshape(1, -1)
                output = knn.predict(resized_img)
                current_name = str(output[0])

                timestamp = datetime.now().strftime("%H:%M:%S")
                student_roll = str(name_to_roll.get(current_name, "N/A"))

                cv2.rectangle(frame, (x, y), (x+w, y+h), (50, 50, 255), 2)
                cv2.rectangle(frame, (x, y-40), (x+w, y), (50, 50, 255), -1)
                cv2.putText(frame, f"{current_name} ({student_roll})", (x, y-15),
                            cv2.FONT_HERSHEY_COMPLEX, 0.7, (255, 255, 255), 1)

                attendance = [student_roll, current_name, timestamp]

            frame_resized = cv2.resize(frame, (640, 480))
            imgBackground = imgBackground_orig.copy()
            if imgBackground.shape[0] >= 642 and imgBackground.shape[1] >= 695:
                imgBackground[162:162 + 480, 55:55 + 640] = frame_resized
                cv2.imshow("Live Attendance - Press 'A' to Mark, 'Q' to Quit", imgBackground)
            else:
                cv2.imshow("Live Attendance - Press 'A' to Mark, 'Q' to Quit", frame)

            k = cv2.waitKey(1) & 0xFF
            if k in [ord("a"), ord("A")]:
                if attendance is None:
                    speak("No face detected")
                    continue

                date_str = datetime.now().strftime("%d-%m-%Y")
                csv_path = os.path.join(ATTENDANCE_DIR, f"Attendance_{date_str}.csv")
                file_exists = os.path.isfile(csv_path)

                with open(csv_path, "a", newline="", encoding="utf-8") as csvfile:
                    writer = csv.writer(csvfile)
                    if not file_exists:
                        writer.writerow(COL_NAMES)
                    writer.writerow(attendance)

                speak("Attendance Taken..")
                time.sleep(0.5)

            if k in [ord("q"), ord("Q")]:
                break

        video.release()
        cv2.destroyAllWindows()


class FourthPage(QWidget):
    def __init__(self):
        super(FourthPage, self).__init__()
        self.setWindowTitle("ATTENDANCE SYSTEM - Select Date")
        self.setFixedWidth(1000)
        self.setStyleSheet("background: #FFFFFF;")
        if os.path.exists(ICON_PATH):
            self.setWindowIcon(QIcon(ICON_PATH))

        self.child_window = None
        self.initUI()

    def open6(self):
        self.child_window = CSVOPEN()
        self.child_window.show()

    def open7(self):
        self.child_window = HomePage()
        self.child_window.show()
        self.hide()

    def initUI(self):
        grid = QGridLayout()
        self.setLayout(grid)

        self.logo5 = QLabel()
        self.logo5.setPixmap(get_pixmap_or_fallback(PERFECT_PATH))
        self.logo5.setAlignment(Qt.AlignHCenter)
        self.logo5.setStyleSheet("margin-top: 10px;")

        self.l15 = QLabel("Enter Date to View Attendance [DD-MM-YYYY]:")
        self.l15.setFont(QFont('Segoe UI, Bahnschrift, Arial', 14, QFont.Bold))
        self.l15.setAlignment(Qt.AlignCenter)

        self.dAtE = QLineEdit()
        self.dAtE.setText(datetime.now().strftime("%d-%m-%Y"))
        self.dAtE.setStyleSheet(
            "QLineEdit { border: 2px solid #cbad8f; border-radius: 8px; padding: 6px; "
            "font-size: 15px; margin: 4px 20px; }"
        )

        btn_style = (
            "QPushButton { border: 4px solid #cbad8f; border-radius: 15px; font-size: 18px; "
            "color: #796648; padding: 8px 0; margin: 10px 20px; font-weight: bold; background: #FFFFFF; } "
            "QPushButton:hover { background: #cbad8f; color: #FFFFFF; }"
        )

        self.button8 = QPushButton("View Records")
        self.button8.setCursor(QCursor(Qt.PointingHandCursor))
        self.button8.setStyleSheet(btn_style)
        self.button8.clicked.connect(self.addinfo1)

        self.button9 = QPushButton("Back to Home")
        self.button9.setCursor(QCursor(Qt.PointingHandCursor))
        self.button9.setStyleSheet(btn_style)
        self.button9.clicked.connect(self.open7)

        grid.addWidget(self.logo5, 0, 0)
        grid.addWidget(self.l15, 1, 0)
        grid.addWidget(self.dAtE, 2, 0)
        grid.addWidget(self.button8, 3, 0)
        grid.addWidget(self.button9, 4, 0)

    def addinfo1(self):
        date_text = self.dAtE.text().strip()
        if not date_text:
            QMessageBox.warning(self, "Required", "Date is required!")
            return

        with open(DATE_CSV, "a", newline="", encoding="utf-8") as temp:
            temp.write(date_text + '\n')

        self.open6()


class CSVOPEN(QMainWindow):
    def __init__(self):
        super(CSVOPEN, self).__init__()
        self.setWindowTitle("ATTENDANCE RECORDS")
        self.resize(800, 500)
        if os.path.exists(ICON_PATH):
            self.setWindowIcon(QIcon(ICON_PATH))

        # Read last requested date
        date_val = datetime.now().strftime("%d-%m-%Y")
        if os.path.exists(DATE_CSV):
            with open(DATE_CSV, mode='r', newline='', encoding='utf-8') as f:
                dates = [line.strip() for line in f if line.strip()]
                if dates:
                    date_val = dates[-1]

        target_csv = os.path.join(ATTENDANCE_DIR, f"Attendance_{date_val}.csv")

        data = []
        if os.path.exists(target_csv):
            with open(target_csv, 'r', newline='', encoding='utf-8') as fh:
                reader = csv.reader(fh)
                for rec in reader:
                    if rec:
                        data.append(rec)
        else:
            QMessageBox.information(
                self, "Not Found",
                f"No attendance records found for date: {date_val}\nFile: {target_csv}"
            )

        central_widget = QWidget()
        layout = QVBoxLayout()
        central_widget.setLayout(layout)

        lbl_header = QLabel(f"Attendance Record for {date_val}")
        lbl_header.setFont(QFont('Segoe UI, Bahnschrift, Arial', 14, QFont.Bold))
        lbl_header.setAlignment(Qt.AlignCenter)
        layout.addWidget(lbl_header)

        self.table = QTableView()
        self.model = TableModel(data)
        self.table.setModel(self.model)
        layout.addWidget(self.table)

        btn_close = QPushButton("Close")
        btn_close.setStyleSheet(
            "QPushButton { border: 2px solid #cbad8f; border-radius: 8px; font-size: 15px; "
            "color: #796648; padding: 6px; font-weight: bold; background: #FFFFFF; } "
            "QPushButton:hover { background: #cbad8f; color: #FFFFFF; }"
        )
        btn_close.clicked.connect(self.close)
        layout.addWidget(btn_close)

        self.setCentralWidget(central_widget)


class TableModel(QAbstractTableModel):
    def __init__(self, data):
        super(TableModel, self).__init__()
        self._data = data

    def data(self, index, role):
        if role == Qt.DisplayRole:
            if index.row() < len(self._data) and index.column() < len(self._data[index.row()]):
                return self._data[index.row()][index.column()]
        return None

    def rowCount(self, index=None):
        return len(self._data)

    def columnCount(self, index=None):
        if not self._data:
            return 0
        return len(self._data[0])


class FifthPage(QWidget):
    def __init__(self):
        super(FifthPage, self).__init__()
        self.setWindowTitle("ATTENDANCE SYSTEM - Developers Info")
        self.setFixedWidth(1000)
        self.setStyleSheet("background: #FFFFFF;")
        if os.path.exists(ICON_PATH):
            self.setWindowIcon(QIcon(ICON_PATH))

        self.child_window = None
        self.initUI()

    def open9(self):
        self.child_window = HomePage()
        self.child_window.show()
        self.hide()

    def initUI(self):
        grid = QGridLayout()
        self.setLayout(grid)

        logo6 = QLabel()
        logo6.setPixmap(get_pixmap_or_fallback(PERFECT_PATH))
        logo6.setAlignment(Qt.AlignHCenter)
        logo6.setStyleSheet("margin-top: 10px;")

        l16 = QLabel("Developers Data")
        l16.setFont(QFont('Segoe UI, Bahnschrift, Arial', 20, QFont.Bold))
        l16.setAlignment(Qt.AlignCenter)

        # Developer 1: Ashutosh Sahu
        ashutosh_img = get_pixmap_or_fallback(os.path.join(FILES_DIR, "Ashutosh.jpeg"), ICON_PATH)
        resized_ashutosh = ashutosh_img.scaled(120, 150, Qt.KeepAspectRatio, Qt.SmoothTransformation)
        logo_ashutosh = QLabel()
        logo_ashutosh.setPixmap(resized_ashutosh)
        logo_ashutosh.setAlignment(Qt.AlignCenter)

        # Developer 2: Sukrut Dusane
        sukrut_img = get_pixmap_or_fallback(os.path.join(FILES_DIR, "Sukrut.jpeg"), ICON_PATH)
        resized_sukrut = sukrut_img.scaled(120, 150, Qt.KeepAspectRatio, Qt.SmoothTransformation)
        logo_sukrut = QLabel()
        logo_sukrut.setPixmap(resized_sukrut)
        logo_sukrut.setAlignment(Qt.AlignCenter)

        # Developer 3: Atharv Shukla
        atharv_img = get_pixmap_or_fallback(os.path.join(FILES_DIR, "Atharv.jpg"), ICON_PATH)
        resized_atharv = atharv_img.scaled(120, 150, Qt.KeepAspectRatio, Qt.SmoothTransformation)
        logo_atharv = QLabel()
        logo_atharv.setPixmap(resized_atharv)
        logo_atharv.setAlignment(Qt.AlignCenter)

        # Developer 4: Samyak Malu
        samyak_img = get_pixmap_or_fallback(os.path.join(FILES_DIR, "Samyak.jpg"), ICON_PATH)
        resized_samyak = samyak_img.scaled(120, 150, Qt.KeepAspectRatio, Qt.SmoothTransformation)
        logo_samyak = QLabel()
        logo_samyak.setPixmap(resized_samyak)
        logo_samyak.setAlignment(Qt.AlignCenter)

        info_font_bold = QFont('Segoe UI, Bahnschrift, Arial', 13, QFont.Bold)
        info_font = QFont('Segoe UI, Bahnschrift, Arial', 12)

        def make_card(name, roll, role, pic_label):
            card = QWidget()
            v = QVBoxLayout()
            card.setLayout(v)
            v.addWidget(pic_label)
            lbl_name = QLabel(name)
            lbl_name.setFont(info_font_bold)
            lbl_name.setAlignment(Qt.AlignCenter)
            v.addWidget(lbl_name)

            lbl_cls = QLabel("Class - XII 'A'")
            lbl_cls.setFont(info_font)
            lbl_cls.setAlignment(Qt.AlignCenter)
            v.addWidget(lbl_cls)

            lbl_roll = QLabel(f"Roll No. {roll}")
            lbl_roll.setFont(info_font)
            lbl_roll.setAlignment(Qt.AlignCenter)
            v.addWidget(lbl_roll)

            lbl_role = QLabel(role)
            lbl_role.setFont(info_font)
            lbl_role.setAlignment(Qt.AlignCenter)
            v.addWidget(lbl_role)
            return card

        card_ashutosh = make_card("ASHUTOSH SAHU", "6", "Role: Programmer & GUI Designer", logo_ashutosh)
        card_sukrut = make_card("SUKRUT DUSANE", "22", "Role: Group Leader, UI Designer", logo_sukrut)
        card_atharv = make_card("ATHARV SHUKLA", "7", "Role: Logbook, Coding & Structuring", logo_atharv)
        card_samyak = make_card("SAMYAK MALU", "19", "Role: Videography", logo_samyak)

        btn_style = (
            "QPushButton { border: 4px solid #cbad8f; border-radius: 15px; font-size: 18px; "
            "color: #796648; padding: 8px 0; margin: 10px 20px; font-weight: bold; background: #FFFFFF; } "
            "QPushButton:hover { background: #cbad8f; color: #FFFFFF; }"
        )

        btn_home = QPushButton("Home")
        btn_home.setCursor(QCursor(Qt.PointingHandCursor))
        btn_home.setStyleSheet(btn_style)
        btn_home.clicked.connect(self.open9)

        grid.addWidget(logo6, 0, 0, 1, 2)
        grid.addWidget(l16, 1, 0, 1, 2)
        grid.addWidget(card_ashutosh, 2, 0)
        grid.addWidget(card_sukrut, 2, 1)
        grid.addWidget(card_atharv, 3, 0)
        grid.addWidget(card_samyak, 3, 1)
        grid.addWidget(btn_home, 4, 0, 1, 2)


def main():
    app = QApplication(sys.argv)
    if os.path.exists(ICON_PATH):
        app.setWindowIcon(QIcon(ICON_PATH))
    window = HomePage()
    window.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()