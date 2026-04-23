# 🖐️ Gesture-Controlled Presentation System

A touch-free presentation tool that lets you control PowerPoint 
slides using simple hand movements — no clicker, no keyboard, 
just your hands and a webcam.

---

## 💡 What Problem Does This Solve?

Traditional slide controllers can interrupt your flow during 
presentations. This system gives you a natural, hands-free way 
to navigate slides in real time using just a standard laptop 
webcam — making presentations smoother, more accessible, and 
more hygienic.

---

## 🛠️ Technologies Used

- Python, OpenCV, MediaPipe
- Win32com, PyAutoGUI
- NumPy

---

## 🤚 Supported Gestures

| Gesture | Action |
|---|---|
| Swipe Right | Next slide |
| Swipe Left | Previous slide |
| Pinch | Trigger action |
| Fist | Stop/Pause |
| Zoom | Scale adjustment |

---

## 📊 Performance

| Metric | Result |
|---|---|
| Swipe Accuracy | ~90% |
| Fist Accuracy | ~92% |
| Pinch Accuracy | ~80% |
| Response Latency | 30–50 ms |
| Continuous Usage | 45+ minutes stable |

---

## 🚀 How to Run

**1. Clone the repository**
```bash
git clone https://github.com/uzairkhan101/GESTURE_CONTROL.git
```

**2. Install dependencies**
```bash
pip install -r requirements.txt
```
⚠️ *MediaPipe only supports Python 3.7–3.10. 
Make sure you are NOT using Python 3.11 or above.*

**3. Run the application**
```bash
python main.py
```
Make sure PowerPoint is open before running.

---

## ✨ Features

- 🎥 Real-time hand gesture recognition via webcam
- 🖥️ Direct PowerPoint slide control
- ⚡ Low latency response (30–50 ms)
- 💡 Works under various lighting conditions
- 🔄 Cooldown logic to prevent accidental triggers

---

## ⚠️ Known Limitations

- MediaPipe does not support Python 3.11 or above — 
  use Python 3.10
- Heavy background interference may affect detection accuracy
- Lighting conditions can impact gesture recognition

---

## 🔮 Future Improvements

- Voice and gesture hybrid control
- Support for Google Slides and macOS Keynote
- Custom gesture training
- Mobile camera integration

---

## 👤 Author

**Uzair Ahmed Khan**
Graduate Student — Data Science & Artificial Intelligence
University of Central Missouri
📧 uzairahmedk10@gmail.com
