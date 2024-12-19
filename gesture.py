import cv2
import mediapipe as mp
import pyautogui
import numpy as np
from scipy.spatial import distance
import speech_recognition as sr
import threading

def eye_aspect_ratio(eye_landmarks):
    """Calculate the Eye Aspect Ratio (EAR)."""
    vertical_1 = distance.euclidean(eye_landmarks[1], eye_landmarks[5])
    vertical_2 = distance.euclidean(eye_landmarks[2], eye_landmarks[4])
    horizontal = distance.euclidean(eye_landmarks[0], eye_landmarks[3])
    return (vertical_1 + vertical_2) / (2.0 * horizontal)

def listen_for_stop(stop_event):
    """Listen for the 'stop tracking' voice command."""
    recognizer = sr.Recognizer()
    mic = sr.Microphone()
    with mic as source:
        recognizer.adjust_for_ambient_noise(source)
    while not stop_event.is_set():
        try:
            with mic as source:
                audio = recognizer.listen(source, timeout=5)
            command = recognizer.recognize_google(audio).lower()
            if "stop tracking" in command:
                stop_event.set()
                break
        except sr.WaitTimeoutError:
            continue
        except sr.UnknownValueError:
            continue

def track_gesture(speak_callback):
    """Track hand gestures and eye blinks for cursor control, scrolling, zooming, and clicking."""
    speak_callback("Starting gesture tracking. Say 'stop tracking' to quit.")
    stop_event = threading.Event()

    # Start listening for the stop command in a separate thread
    listener_thread = threading.Thread(target=listen_for_stop, args=(stop_event,))
    listener_thread.start()

    # Initialize Mediapipe Hands and Face Mesh
    mp_hands = mp.solutions.hands
    mp_face = mp.solutions.face_mesh
    hands = mp_hands.Hands(min_detection_confidence=0.7, min_tracking_confidence=0.7)
    face_mesh = mp_face.FaceMesh(min_detection_confidence=0.7, min_tracking_confidence=0.7)
    mp_drawing = mp.solutions.drawing_utils
    cap = cv2.VideoCapture(0)

    # Screen dimensions
    screen_width, screen_height = pyautogui.size()

    # Cursor variables
    prev_cursor_x, prev_cursor_y = None, None
    smoothing_factor = 0.7

    # Eye blink detection
    eye_blink_threshold = 0.2  # EAR threshold
    blink_cooldown = 15  # Frames to wait after a blink
    blink_timer = 0

    while cap.isOpened() and not stop_event.is_set():
        ret, frame = cap.read()
        if not ret:
            break

        # Flip and preprocess the frame
        frame = cv2.flip(frame, 1)
        rgb_frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        frame = cv2.resize(frame, (screen_width, screen_height))


        # Process frame with Mediapipe Hands and Face Mesh
        hand_results = hands.process(rgb_frame)
        face_results = face_mesh.process(rgb_frame)
        frame_height, frame_width, _ = frame.shape

        # Check for face and eye blink
        if face_results.multi_face_landmarks:
            for face_landmarks in face_results.multi_face_landmarks:
                # Extract left eye landmarks
                left_eye = [
                    (int(face_landmarks.landmark[i].x * frame_width),
                     int(face_landmarks.landmark[i].y * frame_height))
                    for i in [33, 160, 158, 133, 153, 144]
                ]
                ear = eye_aspect_ratio(left_eye)
                if ear < eye_blink_threshold and blink_timer == 0:
                    pyautogui.click()
                    blink_timer = blink_cooldown  # Start cooldown timer

        if blink_timer > 0:
            blink_timer -= 1

        # Check for hand landmarks and gestures
        if hand_results.multi_hand_landmarks:
            for hand_landmarks in hand_results.multi_hand_landmarks:
                # Detect the number of raised fingers
                raised_fingers = 0
                for tip_id, mcp_id in zip(
                    [mp_hands.HandLandmark.INDEX_FINGER_TIP,
                     mp_hands.HandLandmark.MIDDLE_FINGER_TIP,
                     mp_hands.HandLandmark.RING_FINGER_TIP,
                     mp_hands.HandLandmark.PINKY_TIP],
                    [mp_hands.HandLandmark.INDEX_FINGER_MCP,
                     mp_hands.HandLandmark.MIDDLE_FINGER_MCP,
                     mp_hands.HandLandmark.RING_FINGER_MCP,
                     mp_hands.HandLandmark.PINKY_MCP]):
                    if hand_landmarks.landmark[tip_id].y < hand_landmarks.landmark[mcp_id].y:
                        raised_fingers += 1

                # Move cursor with index finger
                index_tip = hand_landmarks.landmark[mp_hands.HandLandmark.INDEX_FINGER_TIP]
                cursor_x = int(index_tip.x * screen_width)
                cursor_y = int(index_tip.y * screen_height)

                # Smooth cursor movement
                smoothed_x = int(smoothing_factor * prev_cursor_x + (1 - smoothing_factor) * cursor_x) if prev_cursor_x else cursor_x
                smoothed_y = int(smoothing_factor * prev_cursor_y + (1 - smoothing_factor) * cursor_y) if prev_cursor_y else cursor_y
                pyautogui.moveTo(smoothed_x, smoothed_y)
                prev_cursor_x, prev_cursor_y = smoothed_x, smoothed_y

                # Gestures for actions
                if raised_fingers == 2:  # Scroll up
                    pyautogui.scroll(1)
                elif raised_fingers == 3:  # Scroll down
                    pyautogui.scroll(-1)
                elif raised_fingers == 4:  # Zoom out
                    pyautogui.hotkey("ctrl", "-")  # Simulate zoom out
                elif raised_fingers == 5:  # Zoom in
                    pyautogui.hotkey("ctrl", "+")  # Simulate zoom in

                # Draw landmarks for visualization
                mp_drawing.draw_landmarks(frame, hand_landmarks, mp_hands.HAND_CONNECTIONS)

        # Display the output frame
        cv2.putText(frame, "Say 'stop tracking' to quit", (10, 30), cv2.FONT_HERSHEY_SIMPLEX, 1, (255, 0, 0), 2)
        cv2.imshow("Gesture Tracking", frame)

        # Exit on pressing 'q'
        if cv2.waitKey(5) & 0xFF == ord('q'):
            stop_event.set()
            break

    # Release resources
    cap.release()
    cv2.destroyAllWindows()
    speak_callback("Gesture tracking stopped.")
