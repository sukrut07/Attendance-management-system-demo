# Face Recognition Attendance Management System

A Python-based smart Attendance Management System leveraging **OpenCV**, **Haar Cascades**, **K-Nearest Neighbors (KNN)**, and **PyQt5** for automated facial recognition and real-time attendance logging.

---

## 🌟 Key Features

- **Automated Face Recognition**: Real-time face detection via webcam and classification using KNN.
- **Audio Feedback**: Text-to-speech voice announcements when attendance is recorded (cross-platform support for Windows, macOS, and Linux).
- **GUI Desktop Application**: Intuitive PyQt5 desktop interface for registering students, taking live attendance, viewing attendance records, and developer information.
- **CLI Tools**: Standalone scripts to capture facial samples and run standalone attendance detection.
- **Automated CSV Logging**: Records student Roll Number, Name, Date, and Time in dated CSV files (`Attendance_DD-MM-YYYY.csv`).
- **Cross-Platform**: Dynamically resolves asset and dataset paths, eliminating hardcoded filesystem dependencies.

---

## 📁 Project Structure

```text
Attendance-management-system-demo/
│
├── AI_final_Project_1000/
│   ├── 2-edited.py                # Main PyQt5 GUI Application
│   │
│   ├── Face Detection/
│   │   ├── add_faces.py           # CLI script to register and capture face data
│   │   ├── test.py                # CLI script to test recognition & log attendance
│   │   ├── LOGO Background.jpg    # UI frame background template
│   │   ├── data/
│   │   │   ├── haarcascade_frontalface_default.xml # Haar Cascade face detector
│   │   │   ├── faces_data.pkl     # Serialized facial features array
│   │   │   ├── names.pkl          # Serialized student names
│   │   │   └── roll.pkl           # Serialized roll numbers
│   │   └── Attendance/            # Generated attendance records (CSV)
│   │
│   └── Files/                     # Images, icons, and student metadata CSVs
│       ├── STUDENTS.csv           # Registered student details
│       ├── DAte.csv               # Historical date query logs
│       ├── ICON.png               # Application icon
│       ├── PERFECT.png            # Application logo header
│       └── *.jpeg / *.jpg         # Developer profiles
│
├── requirements.txt               # Python package dependencies
├── .gitignore                     # Git ignore rules
└── README.md                      # Project documentation
```

---

## 🚀 Getting Started

### 1. Prerequisites

- Python 3.8 to 3.12+
- A working webcam connected to your machine

### 2. Clone the Repository

```bash
git clone https://github.com/sukrut07/Attendance-management-system-demo.git
cd Attendance-management-system-demo
```

### 3. Setup Virtual Environment (Recommended)

```bash
python3 -m venv venv
source venv/bin/activate       # On macOS/Linux
# .\venv\Scripts\activate      # On Windows
```

### 4. Install Dependencies

```bash
pip install -r requirements.txt
```

---

## 🖥️ Usage

### Option 1: Run the Desktop GUI Application

Launch the full-featured desktop interface:

```bash
python AI_final_Project_1000/2-edited.py
```

From the GUI, you can:
1. **Fill Details Here**: Register a new student's First Name, Last Name, and Roll Number, and automatically capture 100 face images.
2. **Give Attendance Here**: Open the camera, detect faces in real-time, press `A` or `a` to log attendance, and press `Q` or `q` to quit.
3. **See Attendance Here**: Select a date (`DD-MM-YYYY`) and view recorded attendance in a table.
4. **Developers Info**: View details of the project contributors.

---

### Option 2: Command Line Scripts

#### Add / Register Face Samples

Collect 100 face samples for a student:

```bash
python "AI_final_Project_1000/Face Detection/add_faces.py"
```
Follow the prompt to enter the student's Name and Roll Number. The webcam will collect 100 face samples and save them to `data/`.

#### Run Face Recognition & Attendance Marking

Run live detection in standalone mode:

```bash
python "AI_final_Project_1000/Face Detection/test.py"
```
- Press **`A`** or **`a`**: Mark and save attendance with voice confirmation.
- Press **`Q`** or **`q`**: Exit the camera stream.

---

## 🛠️ Tech Stack

- **Python**
- **OpenCV (`opencv-python`)** – Computer vision & video capture
- **Scikit-Learn** – K-Nearest Neighbors (KNN) classifier
- **NumPy** – Numerical array processing
- **PyQt5** – Graphical user interface
- **pyttsx3** – Text-to-speech audio alerts

---

## 📄 License

This repository is licensed for educational and demonstration purposes.
