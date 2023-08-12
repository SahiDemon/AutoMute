import cv2
import numpy as np
import pyaudio
import time
import win32api
import win32gui



# Set up the microphone
p = pyaudio.PyAudio()
stream = p.open(format=pyaudio.paInt16, channels=1, rate=48000, input=True, frames_per_buffer=1024)
print("Microphone set up")

# Set up the camera
cap = cv2.VideoCapture(0)
print("Camera set up")

# Check if the connection was successful
if not cap.isOpened():
    print("Failed to open camera")
else:
    print("Camera opened successfully")
        # Capture a frame from the camera
    ret, frame = cap.read()

    # Save the captured frame to an image file
    cv2.imwrite("image.png", frame)
    print("Image saved")
    



# Set the initial frame to None
prev_frame = None

while True:
    # Read a frame from the camera
    ret, frame = cap.read()
    
    # Convert the frame to grayscale
    gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
    
    # If this is the first frame, set it as the previous frame and continue
    if prev_frame is None:
        prev_frame = gray
        continue
    
    # Calculate the absolute difference between the current frame and the previous frame
    diff = cv2.absdiff(gray, prev_frame)
    
    # Threshold the difference image to create a binary image
    thresh = cv2.threshold(diff, 100, 100, cv2.THRESH_BINARY)[1]
    
    # Find the contours in the thresholded image
    contours, _ = cv2.findContours(thresh.copy(), cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    
    # If there are no contours (i.e. no movement), mute the microphone
    if len(contours) == 0:
        WM_APPCOMMAND = 0x319
        APPCOMMAND_MICROPHONE_VOLUME_MUTE = 0x180000

        hwnd_active = win32gui.GetForegroundWindow()
        win32api.SendMessage(hwnd_active, WM_APPCOMMAND, None, APPCOMMAND_MICROPHONE_VOLUME_MUTE)
        print("No movement detected. Microphone muted.")
        time.sleep(5)
    else:
        stream.start_stream()
        print("Movement detected. Holding..")
        time.sleep(5)
        
    
    # Set the current frame as the previous frame for the next iteration
    prev_frame = gray

# Release the camera and microphone resources
cap.release()
stream.stop_stream()
stream.close()
p.terminate()
print("Camera and microphone resources released")
