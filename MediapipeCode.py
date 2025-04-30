import cv2
import mediapipe as mp
import pyautogui
import time
import subprocess
import math

# PowerPoint file path
ppt_file_path = r"C:\Users\ADITYA\Desktop\hand-gesture-control\presentation.pptx"  # Update this path to your PPT file

# Open PowerPoint presentation
subprocess.Popen([r"C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE", ppt_file_path])  # Adjust path if needed

# Initialize MediaPipe Hands
mp_hands = mp.solutions.hands
hands = mp_hands.Hands(min_detection_confidence=0.7, min_tracking_confidence=0.5)

# Initialize video capture
cap = cv2.VideoCapture(0)

# Slide change variables
slide_change_delay = 2  # Minimum time between slide changes
previous_slide_time = time.time()

# Pointer settings
screen_width, screen_height = pyautogui.size()

# Marker trail settings
previous_pointer_position = None  # Stores the previous finger position for trail drawing
marker_enabled = False
marker_thickness = 9  # Approximate thickness for 3mm trail (9 pixels)

# Zoom control variables
initial_pinch_distance = None  # Track the initial pinch distance when both fingers are detected
zoom_threshold = 50  # Threshold to trigger zoom in/out
zoom_delay = 1.5  # Time delay between consecutive zoom actions (in seconds)
previous_zoom_time = time.time()  # Track the last time a zoom action occurred
current_zoom_level = 0  # Initialize zoom level
min_zoom_level = 0  # Minimum zoom-out level for presentation
max_zoom_level = 8  # Maximum zoom-in level (adjust based on how far you want to zoom in)

def calculate_distance(point1, point2):
    """Calculate the Euclidean distance between two points."""
    return math.sqrt((point1[0] - point2[0])**2 + (point1[1] - point2[1])**2)

def is_index_finger_extended(hand_landmarks):
    """Check if only the index finger is extended."""
    index_tip = hand_landmarks.landmark[mp_hands.HandLandmark.INDEX_FINGER_TIP]
    index_dip = hand_landmarks.landmark[mp_hands.HandLandmark.INDEX_FINGER_DIP]

    # Ensure index finger is extended
    index_extended = index_tip.y < index_dip.y  # Index finger tip is above its DIP joint

    # Ensure thumb, middle, ring, and pinky fingers are not extended
    thumb_tip = hand_landmarks.landmark[mp_hands.HandLandmark.THUMB_TIP]
    thumb_ip = hand_landmarks.landmark[mp_hands.HandLandmark.THUMB_IP]
    middle_tip = hand_landmarks.landmark[mp_hands.HandLandmark.MIDDLE_FINGER_TIP]
    middle_pip = hand_landmarks.landmark[mp_hands.HandLandmark.MIDDLE_FINGER_PIP]
    ring_tip = hand_landmarks.landmark[mp_hands.HandLandmark.RING_FINGER_TIP]
    ring_pip = hand_landmarks.landmark[mp_hands.HandLandmark.RING_FINGER_PIP]
    pinky_tip = hand_landmarks.landmark[mp_hands.HandLandmark.PINKY_TIP]
    pinky_pip = hand_landmarks.landmark[mp_hands.HandLandmark.PINKY_PIP]

    thumb_not_extended = thumb_tip.y > thumb_ip.y
    middle_not_extended = middle_tip.y > middle_pip.y
    ring_not_extended = ring_tip.y > ring_pip.y
    pinky_not_extended = pinky_tip.y > pinky_pip.y

    # Return True if only the index finger is extended
    return index_extended and thumb_not_extended and middle_not_extended and ring_not_extended and pinky_not_extended

def is_thumb_and_index_finger_extended(hand_landmarks):
    """Check if only the thumb and index finger are extended for zoom detection."""
    thumb_tip = hand_landmarks.landmark[mp_hands.HandLandmark.THUMB_TIP]
    thumb_ip = hand_landmarks.landmark[mp_hands.HandLandmark.THUMB_IP]
    index_tip = hand_landmarks.landmark[mp_hands.HandLandmark.INDEX_FINGER_TIP]
    index_dip = hand_landmarks.landmark[mp_hands.HandLandmark.INDEX_FINGER_DIP]

    # Check if thumb and index are extended
    thumb_extended = thumb_tip.y < thumb_ip.y  # Thumb tip is above thumb IP joint
    index_extended = index_tip.y < index_dip.y  # Index finger tip is above its DIP joint

    # Ensure that middle, ring, and pinky fingers are not extended
    middle_tip = hand_landmarks.landmark[mp_hands.HandLandmark.MIDDLE_FINGER_TIP]
    middle_pip = hand_landmarks.landmark[mp_hands.HandLandmark.MIDDLE_FINGER_PIP]
    ring_tip = hand_landmarks.landmark[mp_hands.HandLandmark.RING_FINGER_TIP]
    ring_pip = hand_landmarks.landmark[mp_hands.HandLandmark.RING_FINGER_PIP]
    pinky_tip = hand_landmarks.landmark[mp_hands.HandLandmark.PINKY_TIP]
    pinky_pip = hand_landmarks.landmark[mp_hands.HandLandmark.PINKY_PIP]

    middle_not_extended = middle_tip.y > middle_pip.y
    ring_not_extended = ring_tip.y > ring_pip.y
    pinky_not_extended = pinky_tip.y > pinky_pip.y

    # Zoom condition: Thumb and index are extended, others are not
    return thumb_extended and index_extended and middle_not_extended and ring_not_extended and pinky_not_extended

while True:
    success, image = cap.read()
    if not success:
        break

    # Flip the image horizontally for a later selfie-view display
    image = cv2.flip(image, 1)
    image_rgb = cv2.cvtColor(image, cv2.COLOR_BGR2RGB)

    # Process the image and get hand landmarks
    results = hands.process(image_rgb)

    if results.multi_hand_landmarks:
        for hand_landmarks in results.multi_hand_landmarks:
            # Detect if only the index finger is extended
            if is_index_finger_extended(hand_landmarks):
                finger_tip = hand_landmarks.landmark[mp_hands.HandLandmark.INDEX_FINGER_TIP]
                pointer_x = int(finger_tip.x * screen_width)
                pointer_y = int(finger_tip.y * screen_height)

                # Move the system mouse pointer using pyautogui
                pyautogui.moveTo(pointer_x, pointer_y)

                # Disable marker when only one finger is raised
                if marker_enabled:
                    pyautogui.mouseUp()  # Stop drawing
                    marker_enabled = False
                previous_pointer_position = None

            # Detect zoom gestures if both thumb and index finger are extended
            elif is_thumb_and_index_finger_extended(hand_landmarks):
                thumb_tip = hand_landmarks.landmark[mp_hands.HandLandmark.THUMB_TIP]
                index_finger_tip = hand_landmarks.landmark[mp_hands.HandLandmark.INDEX_FINGER_TIP]

                # Calculate the distance between thumb and index finger (for pinch detection)
                thumb_coords = (int(thumb_tip.x * image.shape[1]), int(thumb_tip.y * image.shape[0]))
                index_finger_coords = (int(index_finger_tip.x * image.shape[1]), int(index_finger_tip.y * image.shape[0]))
                pinch_distance = calculate_distance(thumb_coords, index_finger_coords)

                # Gesture waiting: Detect if pinch distance changes over time
                if initial_pinch_distance is None:
                    initial_pinch_distance = pinch_distance  # Set initial pinch distance when fingers are detected
                else:
                    pinch_diff = pinch_distance - initial_pinch_distance
                    current_time = time.time()

                    # Ensure the zoom action happens only after the zoom_delay and based on pinch gesture
                    if current_time - previous_zoom_time > zoom_delay:
                        if pinch_diff > zoom_threshold:  # Pinch out (zoom in)
                            if current_zoom_level < max_zoom_level:
                                pyautogui.hotkey('ctrl', '=')
                                current_zoom_level += 1
                                previous_zoom_time = current_time
                                print(f"Zooming In: Level {current_zoom_level}")
                        elif pinch_diff < -zoom_threshold:  # Pinch in (zoom out)
                            if current_zoom_level > min_zoom_level:
                                pyautogui.hotkey('ctrl', '-')
                                current_zoom_level -= 1
                                previous_zoom_time = current_time
                                print(f"Zooming Out: Level {current_zoom_level}")

                initial_pinch_distance = pinch_distance  # Update the initial distance for comparison

            # Detect thumb direction for slide changes
            else:
                thumb_tip = hand_landmarks.landmark[mp_hands.HandLandmark.THUMB_TIP]
                thumb_cmc = hand_landmarks.landmark[mp_hands.HandLandmark.THUMB_CMC]

                # Check if thumb is pointing left or right
                if thumb_tip.x > thumb_cmc.x:  # Thumb pointing right
                    if time.time() - previous_slide_time > slide_change_delay:
                        pyautogui.press('right')
                        previous_slide_time = time.time()
                        print("Moving to next slide.")
                elif thumb_tip.x < thumb_cmc.x:  # Thumb pointing left
                    if time.time() - previous_slide_time > slide_change_delay:
                        pyautogui.press('left')
                        previous_slide_time = time.time()
                        print("Moving to previous slide.")

            # Draw landmarks on the image for visual feedback (optional, for debugging)
            mp.solutions.drawing_utils.draw_landmarks(image, hand_landmarks, mp_hands.HAND_CONNECTIONS)

    # Display the frame (useful for debugging, but you can comment it out if needed)
    cv2.imshow('Hand Gesture Control with Zoom and Marker', image)

    # Break the loop on 'q' key press
    if cv2.waitKey(5) & 0xFF == ord('q'):
        break

# Release the video capture and close the window
cap.release()
cv2.destroyAllWindows()
