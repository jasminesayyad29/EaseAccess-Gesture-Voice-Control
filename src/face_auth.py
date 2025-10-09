# face_auth.py
import cv2
import numpy as np
import time
import os
import glob

# ---------- User settings ----------
AUTHORIZED_IMAGES_FOLDER = "known_faces"  # Changed from single image path to folder
CAMERA_INDEX = 0
FRAME_WIDTH = 640
FRAME_HEIGHT = 480

# More lenient thresholds for testing
MATCH_RATIO_THRESHOLD = 0.10
GOOD_MATCHS_MIN = 4
AUTH_FRAMES_REQUIRED = 4
CHECK_INTERVAL = 5.0
MIN_FACE_TIME = 1.5
# ------------------------------------

class FaceAuthenticator:
    def __init__(self):
        self.face_cascade = cv2.CascadeClassifier(cv2.data.haarcascades + "haarcascade_frontalface_default.xml")
        if self.face_cascade.empty():
            print("[ERROR] Could not load face cascade classifier")
            return

        self.orb = cv2.ORB_create(nfeatures=1000)
        self.bf = cv2.BFMatcher(cv2.NORM_HAMMING, crossCheck=False)
        
        # Authentication state
        self.authorized = False
        self.last_check_time = 0
        self.face_detected_start_time = 0
        self.continuous_face_detected = False
        self.last_match_result = "No check yet"
        self.last_match_time = 0
        self.last_face_detection_time = 0
        
        # Load authorized faces (multiple)
        self.auth_faces = []  # List to store face images for display
        self.auth_des_list = []  # List to store descriptors for all authorized faces
        self.load_authorized_faces()  # Changed method name
    
    def load_authorized_faces(self):
        """Load all face images from the known_faces folder"""
        if not os.path.exists(AUTHORIZED_IMAGES_FOLDER):
            print(f"[ERROR] Authorized images folder not found at {AUTHORIZED_IMAGES_FOLDER}")
            return False
        
        # Get all image files from the folder
        image_extensions = ['*.jpg', '*.jpeg', '*.png', '*.bmp']
        image_paths = []
        for extension in image_extensions:
            image_paths.extend(glob.glob(os.path.join(AUTHORIZED_IMAGES_FOLDER, extension)))
        
        if not image_paths:
            print(f"[ERROR] No image files found in {AUTHORIZED_IMAGES_FOLDER}")
            return False
        
        loaded_count = 0
        for image_path in image_paths:
            try:
                face_img = self.load_face_from_image(image_path)
                _, des = self.compute_orb_descriptors(face_img)
                if des is not None:
                    self.auth_faces.append(face_img)
                    self.auth_des_list.append(des)
                    loaded_count += 1
                    print(f"[FACE AUTH] Loaded authorized face: {os.path.basename(image_path)}")
                else:
                    print(f"[WARNING] Couldn't compute descriptors for: {os.path.basename(image_path)}")
            except Exception as e:
                print(f"[WARNING] Failed to load face from {os.path.basename(image_path)}: {e}")
        
        if loaded_count > 0:
            print(f"[FACE AUTH] Successfully loaded {loaded_count} authorized faces")
            return True
        else:
            print("[ERROR] No valid faces could be loaded from the folder")
            return False
    
    def load_face_from_image(self, path):
        img = cv2.imread(path)
        if img is None:
            raise FileNotFoundError(f"Image not found at: {path}")
        
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        faces = self.face_cascade.detectMultiScale(gray, scaleFactor=1.05, minNeighbors=3, minSize=(50,50))
        
        if len(faces) == 0:
            raise ValueError(f"No face detected in image: {path}")
        
        x,y,w,h = max(faces, key=lambda r: r[2]*r[3])
        face = gray[y:y+h, x:x+w]
        face = cv2.equalizeHist(face)
        face = cv2.resize(face, (200, 200))
        return face
    
    def compute_orb_descriptors(self, img_gray):
        kp, des = self.orb.detectAndCompute(img_gray, None)
        return kp, des
    
    def good_matches_ratio(self, des1, des2):
        if des1 is None or des2 is None:
            return 0, 0.0, []
        
        matches = self.bf.knnMatch(des1, des2, k=2)
        good = []
        
        for m_n in matches:
            if len(m_n) != 2:
                continue
            m, n = m_n
            if m.distance < 0.75 * n.distance:
                good.append(m)
        
        denom = max(1, min(len(des1), len(des2)))
        ratio = len(good) / denom
        return len(good), ratio, good
    
    def authenticate_against_all_faces(self, des_live):
        """Check if live face matches any of the authorized faces"""
        if not self.auth_des_list or des_live is None:
            return False, "No authorized faces loaded", 0, 0.0
        
        best_match_score = 0
        best_good_matches = 0
        best_ratio = 0.0
        
        for i, auth_des in enumerate(self.auth_des_list):
            n_good, ratio, _ = self.good_matches_ratio(auth_des, des_live)
            match_score = n_good * ratio  # Simple scoring metric
            
            if match_score > best_match_score:
                best_match_score = match_score
                best_good_matches = n_good
                best_ratio = ratio
        
        is_match = (best_good_matches >= GOOD_MATCHS_MIN) and (best_ratio >= MATCH_RATIO_THRESHOLD)
        match_text = f"MATCH: {best_good_matches} good, {best_ratio:.3f} ratio" if is_match else f"NO MATCH: {best_good_matches} good, {best_ratio:.3f} ratio"
        
        return is_match, match_text, best_good_matches, best_ratio
    
    def authenticate_frame(self, frame):
        current_time = time.time()
        gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
        
        # Face detection
        faces = self.face_cascade.detectMultiScale(gray, scaleFactor=1.1, minNeighbors=5, minSize=(80,80))
        face_detected = len(faces) > 0
        
        # Update face detection timing
        if face_detected:
            self.last_face_detection_time = current_time
            if not self.continuous_face_detected:
                self.continuous_face_detected = True
                self.face_detected_start_time = current_time
        else:
            self.continuous_face_detected = False
            self.face_detected_start_time = 0
        
        # Check authentication if conditions are met
        time_since_last_check = current_time - self.last_check_time
        face_detected_duration = current_time - self.face_detected_start_time if self.continuous_face_detected else 0
        
        if (time_since_last_check >= CHECK_INTERVAL and 
            face_detected_duration >= MIN_FACE_TIME and
            face_detected and self.auth_des_list):
            
            x,y,w,h = max(faces, key=lambda r: r[2]*r[3])
            face_roi = gray[y:y+h, x:x+w]
            face_roi = cv2.equalizeHist(face_roi)
            face_roi = cv2.resize(face_roi, (200,200))
            
            _, des_live = self.compute_orb_descriptors(face_roi)
            
            if des_live is not None:
                is_match, match_text, n_good, ratio = self.authenticate_against_all_faces(des_live)
                
                if is_match:
                    if not self.authorized:
                        self.authorized = True
                        print("[FACE AUTH] Authorized - Gesture control ENABLED")
                    self.last_match_result = match_text
                else:
                    if self.authorized:
                        self.authorized = False
                        print("[FACE AUTH] Unauthorized - Gesture control DISABLED")
                    self.last_match_result = match_text
                
                self.last_match_time = current_time
                self.last_check_time = current_time
        
        # Auto-unauthorize if no face detected for 10 seconds
        if self.authorized and not face_detected and (current_time - self.last_face_detection_time > 10):
            self.authorized = False
            print("[FACE AUTH] No face detected for 10s - UNAUTHORIZED")
        
        return face_detected, faces

    def get_auth_status(self):
        return self.authorized
    
    def get_last_match_result(self):
        return self.last_match_result
    
    def get_last_match_time(self):
        return self.last_match_time

# Standalone face authentication function
def run_face_auth_standalone():
    """Run face authentication as a standalone application"""
    print("Starting Face Authentication Standalone...")
    print("Press 'q' to quit.")
    
    face_auth = FaceAuthenticator()
    cap = cv2.VideoCapture(CAMERA_INDEX)
    cap.set(cv2.CAP_PROP_FRAME_WIDTH, FRAME_WIDTH)
    cap.set(cv2.CAP_PROP_FRAME_HEIGHT, FRAME_HEIGHT)
    
    if not cap.isOpened():
        print("[ERROR] Cannot open webcam.")
        return
    
    while True:
        ret, frame = cap.read()
        if not ret:
            print("[ERROR] Failed to read frame from camera.")
            break
        
        # Run face authentication
        face_detected, faces = face_auth.authenticate_frame(frame)
        is_authorized = face_auth.get_auth_status()
        
        # Draw face bounding box and status
        if face_detected:
            x,y,w,h = max(faces, key=lambda r: r[2]*r[3])
            color = (0, 255, 0) if is_authorized else (0, 0, 255)
            cv2.rectangle(frame, (x,y), (x+w, y+h), color, 2)
        
        # Display authorization status
        auth_status = "AUTHORIZED" if is_authorized else "UNAUTHORIZED"
        status_color = (0, 255, 0) if is_authorized else (0, 0, 255)
        cv2.putText(frame, f"Status: {auth_status}", (10, 30), 
                   cv2.FONT_HERSHEY_SIMPLEX, 0.7, status_color, 2)
        
        # Show last match result if recent
        current_time = time.time()
        if current_time - face_auth.get_last_match_time() < 10:
            cv2.putText(frame, face_auth.get_last_match_result(), (10, 60), 
                       cv2.FONT_HERSHEY_SIMPLEX, 0.5, (255, 255, 255), 1)
        
        # Show timing info
        if face_detected and face_auth.continuous_face_detected:
            face_time = current_time - face_auth.face_detected_start_time
            timing_text = f"Face time: {face_time:.1f}s"
            cv2.putText(frame, timing_text, (10, 90), cv2.FONT_HERSHEY_SIMPLEX, 0.5, (255, 255, 255), 1)
        
        # Show authorized faces for reference (first 4 faces)
        for i, auth_face in enumerate(face_auth.auth_faces[:4]):
            cv2.imshow(f"Authorized Face {i+1}", auth_face)
        
        cv2.imshow('Face Authentication - Standalone', frame)
        
        # Quit on 'q'
        if cv2.waitKey(5) & 0xFF == ord('q'):
            break
    
    cap.release()
    cv2.destroyAllWindows()
    print("Face Authentication stopped.")

# Run standalone if this file is executed directly
if __name__ == "__main__":
    run_face_auth_standalone()