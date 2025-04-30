import cv2
import numpy as np
import math
import win32com.client
import time
import pyautogui
import win32gui

# Disable PyAutoGUI fail-safe temporarily
pyautogui.FAILSAFE = False

# Initialize PowerPoint
try:
    Application = win32com.client.Dispatch("PowerPoint.Application")
    Presentation = Application.Presentations.Open(r"C:\Users\ADITYA\Desktop\hand-gesture-control\presentation.pptx")
    Presentation.SlideShowSettings.Run()
    slide_show = Presentation.SlideShowWindow.View
    slide_show.PointerType = 4  # Make pointer always visible
except Exception as e:
    print(f"Error initializing PowerPoint: {e}")
    exit()

# Set up webcam
cap = cv2.VideoCapture(0)
if not cap.isOpened():
    print("Error: Could not open video stream")
    exit()

# Screen dimensions for pointer mapping
screen_width, screen_height = pyautogui.size()
frame_width = int(cap.get(cv2.CAP_PROP_FRAME_WIDTH))
frame_height = int(cap.get(cv2.CAP_PROP_FRAME_HEIGHT))

# Variables for cooldowns
last_operation = time.time()
swipe_cooldown = 2.0
zoom_cooldown = 3.0

current_zoom_level = 0
max_zoom_level = 8
min_zoom_level = 0

# Skin color range in HSV for finger detection
lower_skin = np.array([0, 30, 60], dtype=np.uint8)
upper_skin = np.array([20, 150, 255], dtype=np.uint8)

# Set the ROI width for zoom control (tripled width)
region_width = 1800  # Tripled from 600 pixels

def detect_swipe(flow):
    flow_x = flow[..., 0]
    avg_flow_x = np.mean(flow_x)
    if avg_flow_x > 1:
        return 'right'
    elif avg_flow_x < -1:
        return 'left'
    return None

def detect_scroll(flow, zoom_threshold=1.5):
    flow_y = flow[..., 1]
    avg_flow_y = np.mean(flow_y)
    print("avg_flow_y:", avg_flow_y)  # Debug: output the vertical flow average

    if avg_flow_y > zoom_threshold:
        return "zoom-out"
    elif avg_flow_y < -zoom_threshold:
        return 'zoom-in'
    return None

def activate_presentation():
    hwnd = win32gui.FindWindow(None, Presentation.Name + " - PowerPoint Slide Show")
    if hwnd:
        win32gui.ShowWindow(hwnd, 5)  # Ensure window is visible
        win32gui.SetForegroundWindow(hwnd)
        print("Activated PowerPoint presentation window.")
    else:
        print("PowerPoint presentation window not found.")

# Initialize previous frames for optical flow
ret, frame = cap.read()
if not ret:
    print("Failed to grab frame")
    exit()

frame = cv2.flip(frame, 1)
prev_frame_gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
prev_right_region_gray = prev_frame_gray[:, frame_width - region_width:]

# Main loop
frame_counter = 0

while True:
    ret, frame = cap.read()
    if not ret:
        print("Failed to grab frame")
        break

    frame = cv2.flip(frame, 1)
    hsv = cv2.cvtColor(frame, cv2.COLOR_BGR2HSV)
    mask = cv2.inRange(hsv, lower_skin, upper_skin)

    # Dilation and blurring for noise reduction
    mask = cv2.dilate(mask, np.ones((3, 3), np.uint8), iterations=2)
    mask = cv2.GaussianBlur(mask, (5, 5), 100)

    # Draw a line indicating the ROI boundary for zoom control
    right_edge_x = frame_width - region_width
    cv2.line(frame, (right_edge_x, 0), (right_edge_x, frame_height), (0, 255, 0), 2)

    # Find contours in the mask
    contours, _ = cv2.findContours(mask, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)
    print(f"Contours detected: {len(contours)}")  # Debugging contour count

    fingertip = None
    if contours:
        cnt = max(contours, key=lambda x: cv2.contourArea(x))
        epsilon = 0.0005 * cv2.arcLength(cnt, True)
        approx = cv2.approxPolyDP(cnt, epsilon, True)

        if len(approx) >= 3:
            hull = cv2.convexHull(approx, returnPoints=False)
            defects = cv2.convexityDefects(approx, hull)

            if defects is not None:
                finger_count = 0
                for i in range(defects.shape[0]):
                    s, e, f, d = defects[i, 0]
                    start = tuple(approx[s][0])
                    end = tuple(approx[e][0])
                    far = tuple(approx[f][0])

                    # Calculate angle and distance to identify fingers
                    a = math.dist(end, start)
                    b = math.dist(far, start)
                    c = math.dist(end, far)
                    angle = math.acos((b ** 2 + c ** 2 - a ** 2) / (2 * b * c)) * 57

                    if angle <= 90:
                        finger_count += 1

                finger_count += 1  # Include the thumb

                # Find the fingertip
                fingertip = tuple(cnt[cnt[:, :, 1].argmin()][0])
                print("Fingertip detected at:", fingertip)  # Debug fingertip position

                # Map fingertip to screen
                if fingertip and finger_count == 1:
                    screen_x = int(fingertip[0] * screen_width / frame_width)
                    screen_y = int(fingertip[1] * screen_height / frame_height)

                    # Debug pointer position
                    print(f"Pointer moved to ({screen_x}, {screen_y})")

                    activate_presentation()
                    pyautogui.moveTo(screen_x, screen_y)

    # Process every 3rd frame for optical flow
    frame_counter += 1
    if frame_counter % 3 == 0:
        curr_frame_gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
        curr_right_region_gray = curr_frame_gray[:, frame_width - region_width:]

        # Optical flow for slide navigation (entire frame)
        flow = cv2.calcOpticalFlowFarneback(prev_frame_gray, curr_frame_gray, None, 0.5, 3, 15, 3, 5, 1.2, 0)

        if time.time() - last_operation > swipe_cooldown:
            swipe_direction = detect_swipe(flow)
            if swipe_direction == 'left':
                slide_show.Previous()
                last_swipe_time = time.time()
                print("Swipe Left - Previous Slide")
            elif swipe_direction == 'right':
                slide_show.Next()
                last_swipe_time = time.time()
                print("Swipe Right - Next Slide")

        # Optical flow for zoom control (right region)
        flow_right = cv2.calcOpticalFlowFarneback(prev_right_region_gray, curr_right_region_gray, None, 0.5, 3, 15, 3, 5, 1.2, 0)

        if time.time() - last_operation > zoom_cooldown:
            scroll_direction = detect_scroll(flow_right)
            if scroll_direction == "zoom-in" and current_zoom_level < max_zoom_level:
                activate_presentation()
                pyautogui.hotkey('ctrl', '+')
                last_zoom_time = time.time()
                current_zoom_level += 1
                print("Zooming In")
            elif scroll_direction == "zoom-out" and current_zoom_level > min_zoom_level:
                activate_presentation()
                pyautogui.hotkey('ctrl', '-')
                last_zoom_time = time.time()
                current_zoom_level -= 1
                print("Zooming Out")

        # Update previous frames
        prev_frame_gray = curr_frame_gray.copy()
        prev_right_region_gray = curr_right_region_gray.copy()

    # Display mask and frame
    cv2.imshow("mask", mask)
    cv2.imshow("frame", frame)

    # Exit on 'q' key
    if cv2.waitKey(1) & 0xFF == ord('q'):
        break

cap.release()
cv2.destroyAllWindows()
