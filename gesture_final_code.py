"""
Hand gesture -> PowerPoint controller (Fixed)

Gestures:
1. NAVIGATION: Point Index Finger (☝️) and move it Left or Right.
2. ZOOM:       Pinch (👌) and move hand Closer (In) or Away (Out).
3. RESET:      Make a Fist (✊).

Requirements:
    pip install opencv-python mediapipe numpy pywin32
"""

import cv2
import mediapipe as mp
import math
import time

# --- Configuration ---
ZOOM_SENSITIVITY = 0.05    # Sensitivity for depth zoom
SWIPE_THRESHOLD = 60       # Pixels finger must move to trigger slide
SWIPE_COOLDOWN = 1.5       # Seconds to wait between slides

try:
    import win32com.client
    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False
try:
    import pyautogui
    HAS_PYAUTOGUI = True
except ImportError:
    HAS_PYAUTOGUI = False

# --- PPT Controller ---
class PPTController:
    def __init__(self):
        self.app = None
        self.last_slide_time = 0
        if HAS_WIN32:
            try:
                self.app = win32com.client.Dispatch("PowerPoint.Application")
            except:
                print("Error connecting to PowerPoint.")

    def _can_trigger(self):
        # Prevents double-firing (cooldown)
        if time.time() - self.last_slide_time > SWIPE_COOLDOWN:
            self.last_slide_time = time.time()
            return True
        return False

    def get_slide_count(self):
        try:
            if self.app.ActivePresentation:
                return self.app.ActivePresentation.Slides.Count
        except:
            return 0
        return 0

    def get_current_slide_index(self):
        try:
            # Try Slideshow view first
            if self.app.SlideShowWindows.Count > 0:
                return self.app.SlideShowWindows(1).View.Slide.SlideIndex
            # Try Normal view
            return self.app.ActiveWindow.View.Slide.SlideIndex
        except:
            return 1

    def next_slide(self):
        if not self._can_trigger(): return "Cooldown"
        
        total = self.get_slide_count()
        current = self.get_current_slide_index()
        
        if current >= total and total > 0:
            return "End of Slides" # Don't go past end

        try:
            if self.app.SlideShowWindows.Count > 0:
                self.app.SlideShowWindows(1).View.Next()
            else:
                self.app.ActiveWindow.View.Next()
            return "Next Slide >"
        except:
            if HAS_PYAUTOGUI: pyautogui.press('right')
            return "Key: Right"

    def prev_slide(self):
        if not self._can_trigger(): return "Cooldown"

        current = self.get_current_slide_index()
        if current <= 1:
            return "Start of Slides" # Don't go before start

        try:
            if self.app.SlideShowWindows.Count > 0:
                self.app.SlideShowWindows(1).View.Previous()
            else:
                self.app.ActiveWindow.View.Previous()
            return "< Prev Slide"
        except:
            if HAS_PYAUTOGUI: pyautogui.press('left')
            return "Key: Left"

    def zoom(self, direction):
        # Zoom usually only works in Edit mode, not Slideshow
        try:
            if self.app.SlideShowWindows.Count == 0:
                view = self.app.ActiveWindow.View
                current = view.Zoom
                if direction == 'in':
                    view.Zoom = min(400, current + 5)
                else:
                    view.Zoom = max(10, current - 5)
        except:
            pass

    def reset_zoom(self):
        try:
            if self.app.SlideShowWindows.Count == 0:
                self.app.ActiveWindow.View.ZoomToFit = True
        except:
            pass

# --- Main Script ---
ppt = PPTController()
mp_hands = mp.solutions.hands
hands = mp_hands.Hands(model_complexity=0, max_num_hands=1, min_detection_confidence=0.7, min_tracking_confidence=0.7)
mp_drawing = mp.solutions.drawing_utils

cap = cv2.VideoCapture(0)

# State Variables
state = {
    'zoom_base': None,      # Hand size when pinch started
    'swipe_anchor': None,   # X position when pointer started
    'status_text': "Ready"
}

def get_dist(p1, p2, w, h):
    return math.hypot((p1.x - p2.x) * w, (p1.y - p2.y) * h)

while cap.isOpened():
    success, frame = cap.read()
    if not success: continue

    # Flip for mirror effect
    frame = cv2.flip(frame, 1)
    h, w, _ = frame.shape
    rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
    results = hands.process(rgb)

    active_gesture = "None"
    
    if results.multi_hand_landmarks:
        for hand_lm in results.multi_hand_landmarks:
            lm = hand_lm.landmark
            mp_drawing.draw_landmarks(frame, hand_lm, mp_hands.HAND_CONNECTIONS)

            # 1. GEOMETRY CHECKS
            # Pinch distance (Thumb tip 4 to Index tip 8)
            pinch_dist = get_dist(lm[4], lm[8], w, h)
            # Hand Size (Wrist 0 to Middle Finger Knuckle 9) - for Depth Zoom
            hand_size = get_dist(lm[0], lm[9], w, h)
            
            # Finger States (Check if tips are above pip joints)
            # Note: Y decreases upwards
            index_up = lm[8].y < lm[6].y
            middle_up = lm[12].y < lm[10].y
            ring_up = lm[16].y < lm[14].y
            pinky_up = lm[20].y < lm[18].y
            
            fingers_count = sum([index_up, middle_up, ring_up, pinky_up])
            
            # 2. LOGIC TREE
            
            # --- A. FIST (RESET) ---
            # Strict check: 0 fingers up
            if fingers_count == 0:
                active_gesture = "FIST (Reset)"
                ppt.reset_zoom()
                state['zoom_base'] = None
                state['swipe_anchor'] = None

            # --- B. PINCH (ZOOM) ---
            # Priority over pointer if thumb and index are touching
            elif pinch_dist < 40: 
                active_gesture = "PINCH (Zoom)"
                state['swipe_anchor'] = None # Cancel swipe
                
                if state['zoom_base'] is None:
                    state['zoom_base'] = hand_size # Lock initial size
                else:
                    # Calculate Ratio
                    ratio = hand_size / state['zoom_base']
                    if ratio > (1 + ZOOM_SENSITIVITY):
                        ppt.zoom('in')
                        state['status_text'] = "Zooming In"
                        state['zoom_base'] = hand_size * 0.99 # Update ref smoothly
                    elif ratio < (1 - ZOOM_SENSITIVITY):
                        ppt.zoom('out')
                        state['status_text'] = "Zooming Out"
                        state['zoom_base'] = hand_size * 1.01

            # --- C. POINTER (SWIPE) ---
            # Check: Index is UP, others are DOWN (or mostly down)
            # We removed 'and not is_fist' because the first 'if' catches the fist case.
            elif fingers_count <= 2 and index_up:
                active_gesture = "POINTER (Nav)"
                state['zoom_base'] = None # Cancel zoom
                
                current_x = lm[8].x * w # Track Index Tip
                
                if state['swipe_anchor'] is None:
                    state['swipe_anchor'] = current_x # Lock starting position
                else:
                    diff = current_x - state['swipe_anchor']
                    
                    # Draw a visual line to show drag
                    cv2.line(frame, (int(state['swipe_anchor']), int(lm[8].y*h)), 
                             (int(current_x), int(lm[8].y*h)), (0, 255, 255), 3)

                    if diff > SWIPE_THRESHOLD: # Moved Right
                        res = ppt.next_slide()
                        state['status_text'] = res
                        state['swipe_anchor'] = None # Reset anchor
                    elif diff < -SWIPE_THRESHOLD: # Moved Left
                        res = ppt.prev_slide()
                        state['status_text'] = res
                        state['swipe_anchor'] = None # Reset anchor

            else:
                # Neutral
                state['zoom_base'] = None
                state['swipe_anchor'] = None

    # --- UI Overlay ---
    # Info bar background
    cv2.rectangle(frame, (0, 0), (w, 60), (50, 50, 50), -1)
    
    # Status Text
    color = (0, 255, 0) if "Cooldown" not in state['status_text'] else (0, 0, 255)
    cv2.putText(frame, f"CMD: {state['status_text']}", (20, 40), 
                cv2.FONT_HERSHEY_SIMPLEX, 1, color, 2)
    
    # Gesture Debug
    cv2.putText(frame, f"Mode: {active_gesture}", (w - 250, 40), 
                cv2.FONT_HERSHEY_SIMPLEX, 0.7, (200, 200, 200), 1)

    cv2.imshow("PPT Controller", frame)
    
    if cv2.waitKey(1) & 0xFF == ord('q'):
        break

cap.release()
cv2.destroyAllWindows()
hands.close()