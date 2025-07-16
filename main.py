import cv2
import face_recognition
import time
import openpyxl
import datetime
import os

# Directory containing known face images
known_faces_dir = r"A:\open_cv\photo"

# Load known face encodings and names
known_face_encodings = []
known_face_names = []

# Loop through each image in the directory
for filename in os.listdir(known_faces_dir):
    if filename.endswith(".jpg") or filename.endswith(".jpeg") or filename.endswith(".png"):
        image_path = os.path.join(known_faces_dir, filename)
        image = face_recognition.load_image_file(image_path)
        encodings = face_recognition.face_encodings(image)
        if encodings:
            encoding = encodings[0]
            known_face_encodings.append(encoding)
            known_face_names.append(os.path.splitext(filename)[0])  # Use the filename (without extension) as the person's name

# Initialize webcam
video_capture = cv2.VideoCapture(0)

# Path to the Excel file
excel_file_path = r"A:\open_cv\attendance.xlsx"
# Create or load the Excel file
if not os.path.exists(excel_file_path):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Attendance1"
    sheet.append(["Name", "Date", "Time",])
    workbook.save(excel_file_path)
else:
    try:
        workbook = openpyxl.load_workbook(excel_file_path)
        sheet = workbook.active
    except PermissionError as e:
        print(f"PermissionError: {e}. Please ensure the file is not open in another application and try again.")
        exit()

# Function to log attendance
def log_attendance(name):
    now = datetime.datetime.now()
    current_date = now.strftime("%Y-%m-%d")
    current_time = now.strftime("%H:%M:%S")
    sheet.append([name, current_date, current_time])
    workbook.save(excel_file_path)

# Initialize an empty set to track recognized names
recognized_names = set()
pTime = 0

while True:
    # Capture frame-by-frame
    ret, frame = video_capture.read()
    
    # Find all face locations in the current frame
    face_locations = face_recognition.face_locations(frame)
    face_encodings = face_recognition.face_encodings(frame, face_locations)
    
    # Loop through each face found in the frame
    for (top, right, bottom, left), face_encoding in zip(face_locations, face_encodings):
        # Check if the face matches any known faces
        distances = face_recognition.face_distance(known_face_encodings, face_encoding)
        best_match_index = None
        if len(distances) > 0:
            best_match_index = distances.argmin()
        name = "Unknown"
        
        if best_match_index is not None and distances[best_match_index] < 0.6:
            name = known_face_names[best_match_index]
            print(name)
            
            # Log attendance if the name is recognized and not already logged in this session
            if name not in recognized_names:
                log_attendance(name)
                recognized_names.add(name)
        
        # Draw a box around the face and label with the name
        cv2.rectangle(frame, (left, top), (right, bottom), (0, 0, 255), 2)
        cv2.putText(frame, name, (left, top - 10), cv2.FONT_HERSHEY_SIMPLEX, 0.9, (0, 0, 255), 2)
    
    # Display the resulting frame
    cv2.imshow("Video", frame)
    cTime = time.time()
    fps = 1 / (cTime - pTime)
    pTime = cTime

    cv2.putText(frame, f'FPS: {int(fps)}', (400, 70), cv2.FONT_HERSHEY_PLAIN,
                3, (255, 0, 0), 3)

    # Break the loop when 'q' is pressed
    if cv2.waitKey(1) & 0xFF == ord('q'):
        break

# When everything is done, release the capture and destroy windows
video_capture.release()
cv2.destroyAllWindows() 
    
