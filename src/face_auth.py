# face_auth.py
import cv2
import numpy as np
import time
import os
import glob
from collections import deque

# ---------- User settings ----------
AUTHORIZED_IMAGES_FOLDER = "known_faces"  # folder of authorized face images
CAMERA_INDEX = 0
FRAME_WIDTH = 640
FRAME_HEIGHT = 480

# Tuned thresholds (works well with ORB; adjust slightly after testing)
MATCH_RATIO_THRESHOLD = 0.12
GOOD_MATCHS_MIN = 8

# Temporal / smoothing
CHECK_INTERVAL_UNAUTHORIZED = 1.0  # faster checks while trying to authorize
CHECK_INTERVAL_AUTHORIZED = 2.0    # slower checks once authorized
MIN_FACE_TIME = 1.0                # must see a face this long before first check
AUTO_UNAUTH_SECONDS = 8.0          # auto-unauthorize after this many seconds of no face

# Rolling score buffer
RECENT_SCORES_LEN = 5
RECENT_DECISIONS_LEN = 7
PASS_RATIO_THRESHOLD = 0.60
FAIL_RATIO_THRESHOLD = 0.70

# Streaks: require consecutive successes/fails to change state
SUCCESS_STREAK_REQUIRED = 2
FAIL_STREAK_REQUIRED = 3
# ------------------------------------

class FaceAuthenticator:
    def __init__(self):
        # Load Haar cascade
        self.face_cascade = cv2.CascadeClassifier(
            cv2.data.haarcascades + "haarcascade_frontalface_default.xml"
        )
        if self.face_cascade.empty():
            raise RuntimeError("[ERROR] Could not load face cascade classifier")

        # ORB + BF matcher
        self.orb = cv2.ORB_create(nfeatures=1000)
        self.bf = cv2.BFMatcher(cv2.NORM_HAMMING, crossCheck=False)

        # CLAHE for better local contrast
        self.clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))

        # Auth state
        self.authorized = False
        self.last_check_time = 0.0
        self.face_detected_start_time = 0.0
        self.continuous_face_detected = False
        self.last_match_result = "No check yet"
        self.last_match_time = 0.0
        self.last_face_detection_time = 0.0

        # Streaks & rolling scores
        self.success_streak = 0
        self.fail_streak = 0
        self.recent_scores = []
        self.recent_decisions = deque(maxlen=RECENT_DECISIONS_LEN)

        # Authorized faces storage
        # auth_faces_images -> list of preprocessed grayscale faces (for display/comparison)
        # auth_descriptors  -> list of ORB descriptors for each authorized face image
        self.auth_faces_images = []
        self.auth_descriptors = []

        # Load authorized faces from folder
        self._load_authorized_faces()

    def _enhance_gray(self, gray):
        """Normalize illumination/noise before detection and ORB extraction."""
        gray = cv2.GaussianBlur(gray, (3, 3), 0)
        gray = self.clahe.apply(gray)
        return gray

    def _detect_faces(self, gray):
        """
        Detect faces with a small fallback sweep for robustness across light/angle.
        Returns list of rectangles (x, y, w, h).
        """
        enhanced = self._enhance_gray(gray)

        faces = self.face_cascade.detectMultiScale(
            enhanced, scaleFactor=1.06, minNeighbors=5, minSize=(60, 60)
        )
        if len(faces) > 0:
            return faces

        # Fallback: slightly more permissive params
        faces = self.face_cascade.detectMultiScale(
            enhanced, scaleFactor=1.10, minNeighbors=4, minSize=(50, 50)
        )
        if len(faces) > 0:
            return faces

        # Last fallback on original grayscale (helps rare over-equalized frames)
        faces = self.face_cascade.detectMultiScale(
            gray, scaleFactor=1.08, minNeighbors=4, minSize=(50, 50)
        )
        return faces

    def _load_authorized_faces(self):
        """Load face images from the authorized folder and compute ORB descriptors."""
        if not os.path.exists(AUTHORIZED_IMAGES_FOLDER):
            print(f"[ERROR] Authorized images folder not found at: {AUTHORIZED_IMAGES_FOLDER}")
            return

        image_extensions = ("*.jpg", "*.jpeg", "*.png", "*.bmp")
        image_paths = []
        for ext in image_extensions:
            image_paths.extend(glob.glob(os.path.join(AUTHORIZED_IMAGES_FOLDER, ext)))

        if not image_paths:
            print(f"[ERROR] No image files found in: {AUTHORIZED_IMAGES_FOLDER}")
            return

        loaded = 0
        for path in image_paths:
            try:
                face_img = self._extract_and_preprocess_face_from_image(path)
                if face_img is None:
                    print(f"[WARNING] No face extracted from: {os.path.basename(path)}")
                    continue

                # Compute descriptors
                _, des = self.compute_orb_descriptors(face_img)
                if des is None:
                    print(f"[WARNING] Couldn't compute descriptors for: {os.path.basename(path)}")
                    continue

                self.auth_faces_images.append(face_img)
                self.auth_descriptors.append(des)
                loaded += 1
                print(f"[FACE AUTH] Loaded authorized face: {os.path.basename(path)}")
            except Exception as e:
                print(f"[WARNING] Failed to load {os.path.basename(path)}: {e}")

        if loaded > 0:
            print(f"[FACE AUTH] Successfully loaded {loaded} authorized face images.")
        else:
            print("[ERROR] No valid authorized face images loaded.")

    def _extract_and_preprocess_face_from_image(self, path):
        """Read image, detect largest face, preprocess and return 200x200 grayscale face or None."""
        img = cv2.imread(path)
        if img is None:
            raise FileNotFoundError(f"Image not found at: {path}")

        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        faces = self._detect_faces(gray)
        if len(faces) == 0:
            return None

        x, y, w, h = max(faces, key=lambda r: r[2] * r[3])
        face = gray[y:y + h, x:x + w]

        # Preprocessing: resize -> normalize illumination/noise
        face = cv2.resize(face, (200, 200))
        face = self._enhance_gray(face)

        return face

    def compute_orb_descriptors(self, img_gray):
        """Return keypoints, descriptors (may be None)."""
        if img_gray is None:
            return None, None
        kp, des = self.orb.detectAndCompute(img_gray, None)
        return kp, des

    def good_matches_ratio(self, des1, des2):
        """Return number of good matches, ratio, and list of good matches."""
        if des1 is None or des2 is None:
            return 0, 0.0, []

        # knnMatch may throw if descriptors are empty shapes; guard
        try:
            matches = self.bf.knnMatch(des1, des2, k=2)
        except Exception:
            return 0, 0.0, []

        good = []
        for pair in matches:
            if len(pair) != 2:
                continue
            m, n = pair
            if m.distance < 0.75 * n.distance:
                good.append(m)

        denom = max(1, min(len(des1), len(des2)))
        ratio = len(good) / denom
        return len(good), ratio, good

    def authenticate_against_all_faces(self, des_live):
        """
        Compare live descriptors against all authorized descriptors.
        Returns:
            is_match (bool),
            match_text (str),
            best_good_matches (int),
            best_ratio (float),
            best_score (float)
        """
        if not self.auth_descriptors or des_live is None:
            return False, "No authorized faces loaded or no live descriptors", 0, 0.0, 0.0

        best_score = 0.0
        best_good = 0
        best_ratio = 0.0
        best_idx = -1

        for idx, auth_des in enumerate(self.auth_descriptors):
            n_good, ratio, _ = self.good_matches_ratio(auth_des, des_live)

            # Weighted score: gives balance between match count and ratio
            score = (n_good * 0.7) + (ratio * 100.0)

            if score > best_score:
                best_score = score
                best_good = n_good
                best_ratio = ratio
                best_idx = idx

        # Use both absolute good-match minimum and ratio threshold (but derived from score)
        is_match = (best_good >= GOOD_MATCHS_MIN) and (best_ratio >= MATCH_RATIO_THRESHOLD)

        match_text = (
            f"MATCH idx={best_idx}: {best_good} good, ratio={best_ratio:.3f}, score={best_score:.2f}"
            if is_match
            else f"NO MATCH best={best_good} good, ratio={best_ratio:.3f}, score={best_score:.2f}"
        )

        return is_match, match_text, best_good, best_ratio, best_score

    def authenticate_frame(self, frame):
        """
        Process a single frame: detect face, and occasionally run authentication logic.
        Returns: face_detected (bool), faces (list of rects)
        """
        current_time = time.time()

        # Convert and detect faces
        gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
        faces = self._detect_faces(gray)
        face_detected = len(faces) > 0

        # Update face detection timing
        if face_detected:
            self.last_face_detection_time = current_time
            if not self.continuous_face_detected:
                self.continuous_face_detected = True
                self.face_detected_start_time = current_time
        else:
            self.continuous_face_detected = False
            self.face_detected_start_time = 0.0

        # When to attempt an authentication check
        time_since_last_check = current_time - self.last_check_time
        face_detected_duration = (
            current_time - self.face_detected_start_time if self.continuous_face_detected else 0.0
        )

        check_interval = CHECK_INTERVAL_AUTHORIZED if self.authorized else CHECK_INTERVAL_UNAUTHORIZED

        if (
            face_detected
            and self.auth_descriptors
            and time_since_last_check >= check_interval
            and face_detected_duration >= MIN_FACE_TIME
        ):
            # Crop the largest face
            x, y, w, h = max(faces, key=lambda r: r[2] * r[3])
            face_roi = gray[y:y + h, x:x + w]

            # Preprocess ROI similar to authorized samples
            face_roi = cv2.resize(face_roi, (200, 200))
            face_roi = self._enhance_gray(face_roi)

            # Compute descriptors
            _, des_live = self.compute_orb_descriptors(face_roi)

            if des_live is not None:
                is_match, match_text, n_good, ratio, score = self.authenticate_against_all_faces(des_live)

                # Maintain rolling score buffer
                self.recent_scores.append(score)
                if len(self.recent_scores) > RECENT_SCORES_LEN:
                    self.recent_scores.pop(0)
                avg_score = float(sum(self.recent_scores) / len(self.recent_scores))

                # Use average score to smooth decisions (and keep streak logic)
                # Derive boolean decision from avg_score and match conditions
                final_is_match = (avg_score > 15.0) and (n_good >= GOOD_MATCHS_MIN and ratio >= MATCH_RATIO_THRESHOLD)
                self.recent_decisions.append(1 if final_is_match else 0)
                pass_ratio = float(sum(self.recent_decisions) / len(self.recent_decisions))
                stable_pass = pass_ratio >= PASS_RATIO_THRESHOLD
                stable_fail = pass_ratio <= (1.0 - FAIL_RATIO_THRESHOLD)

                # Update streaks
                if final_is_match:
                    self.success_streak += 1
                    self.fail_streak = 0
                    if self.success_streak >= SUCCESS_STREAK_REQUIRED and stable_pass:
                        if not self.authorized:
                            self.authorized = True
                            print("[FACE AUTH] Authorized - Gesture control ENABLED")
                else:
                    self.fail_streak += 1
                    self.success_streak = 0
                    if self.fail_streak >= FAIL_STREAK_REQUIRED and stable_fail:
                        if self.authorized:
                            self.authorized = False
                            print("[FACE AUTH] Unauthorized - Gesture control DISABLED")

                # Update textual/debug info and timers
                self.last_match_result = f"{match_text} | avg_score={avg_score:.2f}, pass_ratio={pass_ratio:.2f}"
                self.last_match_time = current_time
                self.last_check_time = current_time
            else:
                # No descriptors for live face (rare). Count as fail.
                self.fail_streak += 1
                self.success_streak = 0
                self.recent_decisions.append(0)
                if self.fail_streak >= FAIL_STREAK_REQUIRED and self.authorized:
                    self.authorized = False
                    print("[FACE AUTH] Unauthorized (no descriptors) - Gesture control DISABLED")
                self.last_match_result = "No live descriptors"
                self.last_match_time = current_time
                self.last_check_time = current_time

        # Auto-unauthorize if no face detected for a while
        if self.authorized and (not face_detected) and (current_time - self.last_face_detection_time > AUTO_UNAUTH_SECONDS):
            self.authorized = False
            self.success_streak = 0
            self.fail_streak = 0
            self.recent_decisions.clear()
            print("[FACE AUTH] No face detected for {:.1f}s - UNAUTHORIZED".format(AUTO_UNAUTH_SECONDS))

        return face_detected, faces

    def get_auth_status(self):
        return self.authorized

    def get_last_match_result(self):
        return self.last_match_result

    def get_last_match_time(self):
        return self.last_match_time


def run_face_auth_standalone():
    """Run face authentication as a standalone application."""
    print("Starting Face Authentication Standalone...")
    print("Press 'q' to quit.")

    face_auth = FaceAuthenticator()

    cap = cv2.VideoCapture(CAMERA_INDEX)
    cap.set(cv2.CAP_PROP_FRAME_WIDTH, FRAME_WIDTH)
    cap.set(cv2.CAP_PROP_FRAME_HEIGHT, FRAME_HEIGHT)

    if not cap.isOpened():
        print("[ERROR] Cannot open webcam.")
        return

    try:
        while True:
            ret, frame = cap.read()
            if not ret:
                print("[ERROR] Failed to read frame from camera.")
                break

            face_detected, faces = face_auth.authenticate_frame(frame)
            is_authorized = face_auth.get_auth_status()

            # Draw face bounding box and status
            if face_detected:
                x, y, w, h = max(faces, key=lambda r: r[2] * r[3])
                color = (0, 255, 0) if is_authorized else (0, 0, 255)
                cv2.rectangle(frame, (x, y), (x + w, y + h), color, 2)

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
            for i, auth_face in enumerate(face_auth.auth_faces_images[:4]):
                # Convert grayscale to BGR for consistent display size and easier visualization
                disp = cv2.cvtColor(auth_face, cv2.COLOR_GRAY2BGR)
                disp = cv2.resize(disp, (120, 120))
                cv2.imshow(f"Authorized Face {i + 1}", disp)

            cv2.imshow('Face Authentication - Standalone', frame)

            # Quit on 'q'
            if cv2.waitKey(5) & 0xFF == ord('q'):
                break

    finally:
        cap.release()
        cv2.destroyAllWindows()
        print("Face Authentication stopped.")


if __name__ == "__main__":
    run_face_auth_standalone()
