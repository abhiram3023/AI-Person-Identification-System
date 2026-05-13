# AI Person Identification System

A real-time AI-based person identification and face recognition system developed using Python, OpenCV, and Tkinter.

The system detects and recognizes individuals from live camera streams using computer vision and machine learning techniques. It provides a simple graphical interface for registering users, training facial data, and identifying individuals in real time.

---

# Features

- Real-time face detection
- Real-time person identification
- Face recognition using LBPH algorithm
- GUI-based application using Tkinter
- Live webcam integration
- Person registration system
- Automatic model training
- Identification logs generation
- Excel-based data storage
- Password-protected model training
- Real-time identification logs update

---

# Technologies Used

- Python
- OpenCV
- Tkinter
- NumPy
- Pandas
- OpenPyXL
- Pillow (PIL)

---

# Project Architecture

1. User Registration
2. Face Image Collection
3. Face Data Preprocessing
4. Model Training
5. Real-Time Face Detection
6. Face Recognition
7. Identification Logging

---

# How the System Works

## 1. Person Registration

The user enters:
- Person ID
- Person Name

The system captures multiple face images through the webcam.

These images are stored inside:

TrainingImage/

---

## 2. Face Detection

The project uses:

haarcascade_frontalface_default.xml

This is OpenCV’s Haar Cascade classifier used for detecting faces in real time.

The webcam frame is converted into grayscale and the classifier identifies face regions.

---

## 3. Image Collection

The system captures approximately 100 facial samples for every person.

Each image is saved in the format:

Name.Serial.ID.SampleNumber.jpg

Example:

John.1.101.45.jpg

---

## 4. Model Training

The project uses:

LBPH Face Recognizer

LBPH = Local Binary Pattern Histogram

The recognizer:
- extracts facial texture features
- converts them into histograms
- learns patterns for each registered person

The trained model is stored as:

TrainingImageLabel/Trainer.yml

---

## 5. Real-Time Identification

During live webcam streaming:
- Faces are detected
- Features are extracted
- The trained model predicts the identity

If confidence is high:
- Person name is displayed
- Entry is added to identification logs

---

## 6. Identification Logs

All identified persons are stored inside:

IdentificationRecords/IdentificationRecords.xlsx

The logs contain:
- Serial Number
- ID
- Name
- Date
- Time

---

# Folder Structure

AI-Person-Identification-System/
│
├── main.py
├── requirements.txt
├── README.md
├── .gitignore
├── haarcascade_frontalface_default.xml
│
├── screenshots/
│   ├── ui.png
│   ├── detection.png
│
├── TrainingImage/
├── TrainingImageLabel/
├── IdentificationRecords/
└── PersonDetails/

---

# Installation

## Clone Repository

git clone https://github.com/abhiram3023/AI-Person-Identification-System.git

## Move Into Folder

cd AI-Person-Identification-System

## Install Dependencies

pip install -r requirements.txt

## Run Project

python main.py

---

# Required Libraries

- opencv-contrib-python
- pillow
- numpy
- pandas
- openpyxl

---

# Generate requirements.txt

pip freeze > requirements.txt

---

# GUI Overview

The GUI contains:

- Person Registration Panel
- Real-Time Identification Panel
- Live Identification Logs
- Model Training Controls
- Delete Utilities
- System Status Messages

---

# Advantages

- Fast real-time recognition
- Lightweight model
- Simple GUI interface
- Easy to use
- Efficient for small-scale identification systems

---

# Future Enhancements

- Deep learning-based recognition
- Cloud database integration
- Multiple camera support
- Anti-spoofing detection
- Mobile application integration

---

# Screenshots

## Main GUI

![GUI](screenshots/Project_output_GUI(2).png)
![GUI](screenshots/Project_output_GUI(3).png)

## Real-Time Detection

(Add screenshot here)

---

# Author

Developed as an AI and Computer Vision project using Python and OpenCV.
Abhiram Majeti
