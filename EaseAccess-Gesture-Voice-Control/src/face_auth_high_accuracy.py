import cv2
import numpy as np
import time
import os
import glob
from collections import deque

# ---------- User settings ----------
AUTHORIZED_IMAGES_FOLDER = "known_faces"
CAMERA_INDEX = 0
FRAME_WIDTH = 640
FRAME_HEIGHT = 480

# LBPH thresholds (lower confidence is better)
LBPH_CONFIDENCE_MAX = 62.0
LBPH_STRONG_MATCH_MAX = 50.0

# ORB thresholds (fallback confirmation)
ORB_MATCH_RATIO_THRESHOLD = 0.12
ORB_GOOD_MATCHES_MIN = 10

# Temporal smoothing
CHECK_INTERVAL_UNAUTHORIZED = 0.6
CHECK_INTERVAL_AUTHORIZED = 1.2
MIN_FACE_TIME = 0.8
AUTO_UNAUTH_SECONDS = 10.0

# Decision stability
SUCCESS_STREAK_REQUIRED = 2
FAIL_STREAK_REQUIRED = 3
AUTHORIZED_FAIL_STREAK_REQUIRED = 6
RECENT_DECISIONS_LEN = 7
PASS_RATIO_THRESHOLD = 0.65
FAIL_RATIO_THRESHOLD = 0.75

# Authorized-mode tolerance: brief motion blur or partial pose change should not instantly deauthorize.
SOFT_FAIL_LBPH_MARGIN = 8.0
SOFT_FAIL_ORB_RATIO_MARGIN = 0.03
SOFT_FAIL_ORB_GOOD_MATCH_MARGIN = 2
# ------------------------------------


class FaceAuthenticatorHighAccuracy:
    """
    High-reliability face authenticator using two-stage verification:
    1) LBPH identity prediction
    2) ORB feature consistency check on predicted identity

    Notes:
    - No vision system can guarantee 100% accuracy in real-world conditions.
    - This implementation prioritizes stricter acceptance criteria to reduce false accepts.
    """

    def __init__(self, authorized_images_folder=None):
        self.face_cascade = cv2.CascadeClassifier(
            cv2.data.haarcascades + "haarcascade_frontalface_default.xml"
        )
        if self.face_cascade.empty():
            raise RuntimeError("[ERROR] Could not load face cascade classifier")

        self.orb = cv2.ORB_create(nfeatures=1200)
        self.bf = cv2.BFMatcher(cv2.NORM_HAMMING, crossCheck=False)
        self.clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))

        # OpenCV contrib recognizer; gracefully handle environments where it is unavailable.
        self.lbph_enabled = hasattr(cv2, "face") and hasattr(cv2.face, "LBPHFaceRecognizer_create")
        self.lbph = cv2.face.LBPHFaceRecognizer_create() if self.lbph_enabled else None

        self.authorized = False
        self.last_check_time = 0.0
        self.face_detected_start_time = 0.0
        self.continuous_face_detected = False
        self.last_match_result = "No check yet"
        self.last_match_time = 0.0
        self.last_face_detection_time = 0.0

        self.success_streak = 0
        self.fail_streak = 0
        self.recent_decisions = deque(maxlen=RECENT_DECISIONS_LEN)

        self.auth_faces_images = []
        self.auth_descriptors = []
        self.identity_labels = []
        self.identity_name_by_label = {}
        self.auth_descriptors_by_label = {}

        script_dir = os.path.dirname(os.path.abspath(__file__))
        folder_candidate = authorized_images_folder or AUTHORIZED_IMAGES_FOLDER
        if os.path.isabs(folder_candidate):
            self.authorized_images_folder = folder_candidate
        else:
            self.authorized_images_folder = os.path.join(script_dir, folder_candidate)
            if not os.path.exists(self.authorized_images_folder):
                self.authorized_images_folder = os.path.join(os.getcwd(), folder_candidate)

        self._load_and_train()

    def _enhance_gray(self, gray):
        gray = cv2.GaussianBlur(gray, (3, 3), 0)
        gray = self.clahe.apply(gray)
        return gray

    def _detect_faces(self, gray):
        enhanced = self._enhance_gray(gray)

        faces = self.face_cascade.detectMultiScale(
            enhanced, scaleFactor=1.06, minNeighbors=5, minSize=(60, 60)
        )
        if len(faces) > 0:
            return faces

        faces = self.face_cascade.detectMultiScale(
            enhanced, scaleFactor=1.10, minNeighbors=4, minSize=(50, 50)
        )
        if len(faces) > 0:
            return faces

        faces = self.face_cascade.detectMultiScale(
            gray, scaleFactor=1.08, minNeighbors=4, minSize=(50, 50)
        )
        return faces

    def _extract_face_from_image(self, path):
        img = cv2.imread(path)
        if img is None:
            raise FileNotFoundError(f"Image not found at: {path}")

        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        faces = self._detect_faces(gray)
        if len(faces) == 0:
            return None

        x, y, w, h = max(faces, key=lambda r: r[2] * r[3])
        face = gray[y:y + h, x:x + w]
        if face.size == 0:
            return None

        face = cv2.resize(face, (200, 200))
        face = self._enhance_gray(face)
        return face

    def _augment_face(self, face):
        # Small augmentations improve recognition robustness from limited samples.
        variants = [face]
        variants.append(cv2.flip(face, 1))

        brighter = cv2.convertScaleAbs(face, alpha=1.05, beta=8)
        darker = cv2.convertScaleAbs(face, alpha=0.95, beta=-8)
        variants.append(brighter)
        variants.append(darker)

        noisy = cv2.GaussianBlur(face, (3, 3), 0.4)
        variants.append(noisy)

        return variants

    def _compute_orb(self, img_gray):
        if img_gray is None:
            return None, None
        kp, des = self.orb.detectAndCompute(img_gray, None)
        return kp, des

    def _good_matches_ratio(self, des1, des2):
        if des1 is None or des2 is None:
            return 0, 0.0

        try:
            matches = self.bf.knnMatch(des1, des2, k=2)
        except Exception:
            return 0, 0.0

        good = []
        for pair in matches:
            if len(pair) != 2:
                continue
            m, n = pair
            if m.distance < 0.75 * n.distance:
                good.append(m)

        denom = max(1, min(len(des1), len(des2)))
        ratio = len(good) / denom
        return len(good), ratio

    def _load_and_train(self):
        self.auth_faces_images.clear()
        self.auth_descriptors.clear()
        self.identity_labels.clear()
        self.identity_name_by_label = {}
        self.auth_descriptors_by_label = {}

        if not os.path.exists(self.authorized_images_folder):
            print(f"[ERROR] Authorized images folder not found at: {self.authorized_images_folder}")
            return

        image_extensions = ("*.jpg", "*.jpeg", "*.png", "*.bmp")
        image_paths = []
        for ext in image_extensions:
            image_paths.extend(glob.glob(os.path.join(self.authorized_images_folder, ext)))

        if not image_paths:
            print(f"[ERROR] No image files found in: {self.authorized_images_folder}")
            return

        label_by_identity = {}
        train_images = []
        train_labels = []

        loaded = 0
        for path in sorted(image_paths):
            try:
                face = self._extract_face_from_image(path)
                if face is None:
                    print(f"[WARNING] No face extracted from: {os.path.basename(path)}")
                    continue

                identity_name = os.path.splitext(os.path.basename(path))[0].lower()
                if identity_name not in label_by_identity:
                    new_label = len(label_by_identity)
                    label_by_identity[identity_name] = new_label
                    self.identity_name_by_label[new_label] = identity_name
                    self.auth_descriptors_by_label[new_label] = []

                label = label_by_identity[identity_name]
                variants = self._augment_face(face)

                for v in variants:
                    _, des = self._compute_orb(v)
                    if des is not None:
                        self.auth_descriptors.append(des)
                        self.auth_descriptors_by_label[label].append(des)

                    self.auth_faces_images.append(v)
                    self.identity_labels.append(label)
                    train_images.append(v)
                    train_labels.append(label)

                loaded += 1
                print(f"[FACE AUTH HA] Loaded authorized face: {os.path.basename(path)}")
            except Exception as e:
                print(f"[WARNING] Failed to load {os.path.basename(path)}: {e}")

        if self.lbph_enabled and train_images:
            labels_np = np.array(train_labels, dtype=np.int32)
            self.lbph.train(train_images, labels_np)

        if loaded > 0:
            mode = "LBPH+ORB" if self.lbph_enabled else "ORB-only"
            print(f"[FACE AUTH HA] Trained with {loaded} identity images ({mode}).")
        else:
            print("[ERROR] No valid authorized face images loaded.")

    def _predict_with_lbph(self, face):
        if not self.lbph_enabled or self.lbph is None:
            return None, float("inf")
        try:
            label, confidence = self.lbph.predict(face)
            return label, float(confidence)
        except Exception:
            return None, float("inf")

    def _verify_orb_for_label(self, live_des, label):
        if live_des is None or label not in self.auth_descriptors_by_label:
            return 0, 0.0

        best_good = 0
        best_ratio = 0.0
        for auth_des in self.auth_descriptors_by_label[label]:
            n_good, ratio = self._good_matches_ratio(auth_des, live_des)
            if n_good > best_good or (n_good == best_good and ratio > best_ratio):
                best_good = n_good
                best_ratio = ratio

        return best_good, best_ratio

    def authenticate_frame(self, frame):
        current_time = time.time()

        gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
        faces = self._detect_faces(gray)
        face_detected = len(faces) > 0

        if face_detected:
            self.last_face_detection_time = current_time
            if not self.continuous_face_detected:
                self.continuous_face_detected = True
                self.face_detected_start_time = current_time
        else:
            self.continuous_face_detected = False
            self.face_detected_start_time = 0.0

        time_since_last_check = current_time - self.last_check_time
        face_detected_duration = (
            current_time - self.face_detected_start_time if self.continuous_face_detected else 0.0
        )
        check_interval = CHECK_INTERVAL_AUTHORIZED if self.authorized else CHECK_INTERVAL_UNAUTHORIZED

        can_check = (
            face_detected
            and len(self.auth_descriptors) > 0
            and time_since_last_check >= check_interval
            and face_detected_duration >= MIN_FACE_TIME
        )

        if can_check:
            x, y, w, h = max(faces, key=lambda r: r[2] * r[3])
            face_roi = gray[y:y + h, x:x + w]
            if face_roi.size == 0:
                self.last_match_result = "Invalid face ROI"
                self.last_match_time = current_time
                self.last_check_time = current_time
                return face_detected, faces

            face_roi = cv2.resize(face_roi, (200, 200))
            face_roi = self._enhance_gray(face_roi)

            _, live_des = self._compute_orb(face_roi)

            lbph_label, lbph_conf = self._predict_with_lbph(face_roi)
            lbph_pass = lbph_conf <= LBPH_CONFIDENCE_MAX
            lbph_strong = lbph_conf <= LBPH_STRONG_MATCH_MAX

            # If LBPH is unavailable, evaluate ORB across all labels.
            if lbph_label is None:
                best_label = None
                best_good = 0
                best_ratio = 0.0
                for label in self.auth_descriptors_by_label.keys():
                    n_good, ratio = self._verify_orb_for_label(live_des, label)
                    if n_good > best_good or (n_good == best_good and ratio > best_ratio):
                        best_good = n_good
                        best_ratio = ratio
                        best_label = label
                orb_good = best_good
                orb_ratio = best_ratio
                predicted_label = best_label
                orb_pass = (orb_good >= ORB_GOOD_MATCHES_MIN) and (orb_ratio >= ORB_MATCH_RATIO_THRESHOLD)
                final_match = orb_pass
            else:
                orb_good, orb_ratio = self._verify_orb_for_label(live_des, lbph_label)
                orb_pass = (orb_good >= ORB_GOOD_MATCHES_MIN) and (orb_ratio >= ORB_MATCH_RATIO_THRESHOLD)

                # Strict acceptance rule for higher precision.
                final_match = (lbph_pass and orb_pass) or (lbph_strong and orb_pass)
                predicted_label = lbph_label

            # While already authorized, treat near-threshold misses as soft failures.
            # This prevents rapid auth/unauth flips during small head movement.
            near_lbph = (lbph_label is not None) and (lbph_conf <= (LBPH_CONFIDENCE_MAX + SOFT_FAIL_LBPH_MARGIN))
            near_orb = (
                orb_ratio >= max(0.0, ORB_MATCH_RATIO_THRESHOLD - SOFT_FAIL_ORB_RATIO_MARGIN)
                and orb_good >= max(1, ORB_GOOD_MATCHES_MIN - SOFT_FAIL_ORB_GOOD_MATCH_MARGIN)
            )
            soft_fail = self.authorized and (not final_match) and (near_lbph or near_orb)

            if final_match:
                self.recent_decisions.append(1.0)
            elif soft_fail:
                self.recent_decisions.append(0.5)
            else:
                self.recent_decisions.append(0.0)

            pass_ratio = float(sum(self.recent_decisions) / len(self.recent_decisions))
            stable_pass = pass_ratio >= PASS_RATIO_THRESHOLD
            stable_fail = pass_ratio <= (1.0 - FAIL_RATIO_THRESHOLD)

            if final_match:
                self.success_streak += 1
                self.fail_streak = 0
                if self.success_streak >= SUCCESS_STREAK_REQUIRED and stable_pass and not self.authorized:
                    self.authorized = True
                    print("[FACE AUTH HA] Authorized - Gesture control ENABLED")
            elif soft_fail:
                # Keep state unchanged on soft failures and avoid fail streak growth.
                self.success_streak = max(0, self.success_streak - 1)
            else:
                self.fail_streak += 1
                self.success_streak = 0
                required_fails = AUTHORIZED_FAIL_STREAK_REQUIRED if self.authorized else FAIL_STREAK_REQUIRED
                if self.fail_streak >= required_fails and stable_fail and self.authorized:
                    self.authorized = False
                    print("[FACE AUTH HA] Unauthorized - Gesture control DISABLED")

            identity = self.identity_name_by_label.get(predicted_label, "unknown")
            self.last_match_result = (
                f"id={identity} lbph={lbph_conf:.1f} orb={orb_good}/{orb_ratio:.3f} "
                f"final={'MATCH' if final_match else ('SOFT_FAIL' if soft_fail else 'NO_MATCH')} pass={pass_ratio:.2f}"
            )
            self.last_match_time = current_time
            self.last_check_time = current_time

        if self.authorized and (not face_detected) and (current_time - self.last_face_detection_time > AUTO_UNAUTH_SECONDS):
            self.authorized = False
            self.success_streak = 0
            self.fail_streak = 0
            self.recent_decisions.clear()
            print("[FACE AUTH HA] No face detected for {:.1f}s - UNAUTHORIZED".format(AUTO_UNAUTH_SECONDS))

        return face_detected, faces

    def get_auth_status(self):
        return self.authorized

    def get_last_match_result(self):
        return self.last_match_result

    def get_last_match_time(self):
        return self.last_match_time


def run_high_accuracy_face_auth_standalone():
    print("Starting High-Accuracy Face Authentication...")
    print("Press 'q' to quit.")

    face_auth = FaceAuthenticatorHighAccuracy()

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

            if face_detected:
                x, y, w, h = max(faces, key=lambda r: r[2] * r[3])
                color = (0, 255, 0) if is_authorized else (0, 0, 255)
                cv2.rectangle(frame, (x, y), (x + w, y + h), color, 2)

            auth_status = "AUTHORIZED" if is_authorized else "UNAUTHORIZED"
            status_color = (0, 255, 0) if is_authorized else (0, 0, 255)
            cv2.putText(frame, f"Status: {auth_status}", (10, 30),
                        cv2.FONT_HERSHEY_SIMPLEX, 0.7, status_color, 2)

            now = time.time()
            if now - face_auth.get_last_match_time() < 10:
                cv2.putText(frame, face_auth.get_last_match_result(), (10, 60),
                            cv2.FONT_HERSHEY_SIMPLEX, 0.5, (255, 255, 255), 1)

            cv2.imshow("Face Authentication - High Accuracy", frame)
            if cv2.waitKey(5) & 0xFF == ord('q'):
                break

    finally:
        cap.release()
        cv2.destroyAllWindows()
        print("High-accuracy Face Authentication stopped.")


if __name__ == "__main__":
    run_high_accuracy_face_auth_standalone()
