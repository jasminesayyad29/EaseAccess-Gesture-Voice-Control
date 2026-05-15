import cv2
import numpy as np
import time
import os
import glob
import re
import shutil
import urllib.request
from collections import deque

# ---------- User settings ----------
AUTHORIZED_IMAGES_FOLDER = "known_faces"
CAMERA_INDEX = 0
FRAME_WIDTH = 640
FRAME_HEIGHT = 480

# Modern OpenCV DNN face authentication.
# YuNet detects faces and landmarks; SFace creates face embeddings for fast matching.
DNN_MODEL_FOLDER = "models/face_auth"
YUNET_MODEL_NAME = "face_detection_yunet_2023mar.onnx"
SFACE_MODEL_NAME = "face_recognition_sface_2021dec.onnx"
YUNET_MODEL_URL = (
    "https://raw.githubusercontent.com/opencv/opencv_zoo/main/models/"
    "face_detection_yunet/face_detection_yunet_2023mar.onnx"
)
SFACE_MODEL_URL = (
    "https://raw.githubusercontent.com/opencv/opencv_zoo/main/models/"
    "face_recognition_sface/face_recognition_sface_2021dec.onnx"
)
AUTO_DOWNLOAD_DNN_MODELS = True
DNN_DOWNLOAD_TIMEOUT = 25
DNN_DETECTION_SCORE_THRESHOLD = 0.88
DNN_NMS_THRESHOLD = 0.30
DNN_TOP_K = 5000

# SFace cosine similarity: higher is better. OpenCV's public LFW threshold is 0.363;
# authentication uses a stricter value to reduce false accepts.
SFACE_COSINE_THRESHOLD = 0.42
SFACE_STRONG_COSINE_THRESHOLD = 0.50
SFACE_SOFT_FAIL_MARGIN = 0.05

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


def _resolve_authorized_images_folder(folder_candidate):
    script_dir = os.path.dirname(os.path.abspath(__file__))
    if os.path.isabs(folder_candidate):
        return folder_candidate

    # Prefer sibling of .py folder (src/known_faces) for this repository layout.
    candidates = [
        os.path.join(os.path.dirname(script_dir), folder_candidate),
        os.path.join(script_dir, folder_candidate),
        os.path.join(os.getcwd(), folder_candidate),
    ]
    for path in candidates:
        if os.path.exists(path):
            return path
    return candidates[0]


def _resolve_model_folder(folder_candidate):
    script_dir = os.path.dirname(os.path.abspath(__file__))
    if os.path.isabs(folder_candidate):
        return folder_candidate

    return os.path.join(os.path.dirname(script_dir), folder_candidate)


def _identity_name_from_path(path):
    name = os.path.splitext(os.path.basename(path))[0].lower()
    # face_auth_register.py saves name_01.jpg, name_02.jpg, etc.
    return re.sub(r"[_-]\d+$", "", name)


def _rect_from_dnn_face(face_row):
    x, y, w, h = [int(round(v)) for v in face_row[:4]]
    return (x, y, w, h)


class FaceAuthenticatorHighAccuracy:
    """
    High-reliability face authenticator using the best available local backend:
    1) OpenCV DNN YuNet + SFace embeddings when ONNX models are available
    2) LBPH + ORB fallback when DNN models are unavailable

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
        self.sface_features_by_label = {}
        self.dnn_enabled = False
        self.dnn_status = "not initialized"
        self.dnn_detector = None
        self.sface = None

        folder_candidate = authorized_images_folder or AUTHORIZED_IMAGES_FOLDER
        self.authorized_images_folder = _resolve_authorized_images_folder(folder_candidate)
        self.model_folder = _resolve_model_folder(DNN_MODEL_FOLDER)

        self._init_dnn_backend()
        self._load_and_train()

    def _model_path(self, model_name):
        return os.path.join(self.model_folder, model_name)

    def _ensure_model_file(self, model_name, url):
        path = self._model_path(model_name)
        if os.path.exists(path) and os.path.getsize(path) > 0:
            return path

        if not AUTO_DOWNLOAD_DNN_MODELS:
            return None

        os.makedirs(self.model_folder, exist_ok=True)
        tmp_path = path + ".download"
        try:
            print(f"[FACE AUTH HA] Downloading {model_name}...")
            with urllib.request.urlopen(url, timeout=DNN_DOWNLOAD_TIMEOUT) as response:
                with open(tmp_path, "wb") as tmp_file:
                    shutil.copyfileobj(response, tmp_file)
            if not os.path.exists(tmp_path) or os.path.getsize(tmp_path) == 0:
                raise RuntimeError("downloaded file is empty")
            os.replace(tmp_path, path)
            return path
        except Exception as e:
            try:
                if os.path.exists(tmp_path):
                    os.remove(tmp_path)
            except Exception:
                pass
            print(f"[FACE AUTH HA] Could not download {model_name}: {e}")
            return None

    def _init_dnn_backend(self):
        if not hasattr(cv2, "FaceDetectorYN_create") or not hasattr(cv2, "FaceRecognizerSF_create"):
            self.dnn_status = "OpenCV DNN face APIs unavailable"
            print(f"[FACE AUTH HA] {self.dnn_status}; using LBPH+ORB fallback.")
            return

        yunet_path = self._ensure_model_file(YUNET_MODEL_NAME, YUNET_MODEL_URL)
        sface_path = self._ensure_model_file(SFACE_MODEL_NAME, SFACE_MODEL_URL)
        if not yunet_path or not sface_path:
            self.dnn_status = "DNN model files unavailable"
            print(f"[FACE AUTH HA] {self.dnn_status}; using LBPH+ORB fallback.")
            return

        try:
            self.dnn_detector = cv2.FaceDetectorYN_create(
                yunet_path,
                "",
                (FRAME_WIDTH, FRAME_HEIGHT),
                DNN_DETECTION_SCORE_THRESHOLD,
                DNN_NMS_THRESHOLD,
                DNN_TOP_K,
            )
            self.sface = cv2.FaceRecognizerSF_create(sface_path, "")
            self.dnn_enabled = True
            self.dnn_status = "YuNet+SFace"
            print("[FACE AUTH HA] Modern DNN backend enabled: YuNet+SFace.")
        except Exception as e:
            self.dnn_enabled = False
            self.dnn_status = f"DNN init failed: {e}"
            print(f"[FACE AUTH HA] {self.dnn_status}; using LBPH+ORB fallback.")

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

    def _detect_faces_dnn(self, frame):
        if not self.dnn_enabled or self.dnn_detector is None:
            return None

        h, w = frame.shape[:2]
        self.dnn_detector.setInputSize((w, h))
        try:
            _, faces = self.dnn_detector.detect(frame)
        except Exception:
            return None

        if faces is None or len(faces) == 0:
            return np.empty((0, 4), dtype=np.int32), []

        face_rows = sorted(faces, key=lambda r: float(r[2] * r[3]), reverse=True)
        rects = np.array([_rect_from_dnn_face(row) for row in face_rows], dtype=np.int32)
        return rects, face_rows

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

    def _extract_dnn_feature_from_image(self, path):
        if not self.dnn_enabled or self.sface is None:
            return None

        img = cv2.imread(path)
        if img is None:
            raise FileNotFoundError(f"Image not found at: {path}")

        detection = self._detect_faces_dnn(img)
        if detection is None:
            return None
        _, face_rows = detection
        if not face_rows:
            return None

        try:
            aligned = self.sface.alignCrop(img, np.asarray(face_rows[0], dtype=np.float32))
            feature = self.sface.feature(aligned)
            return feature.copy()
        except Exception:
            return None

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
        self.sface_features_by_label = {}

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

                identity_name = _identity_name_from_path(path)
                if identity_name not in label_by_identity:
                    new_label = len(label_by_identity)
                    label_by_identity[identity_name] = new_label
                    self.identity_name_by_label[new_label] = identity_name
                    self.auth_descriptors_by_label[new_label] = []
                    self.sface_features_by_label[new_label] = []

                label = label_by_identity[identity_name]
                dnn_feature = self._extract_dnn_feature_from_image(path)
                if dnn_feature is not None:
                    self.sface_features_by_label[label].append(dnn_feature)

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
            if self.dnn_enabled and any(self.sface_features_by_label.values()):
                mode = "YuNet+SFace with LBPH+ORB fallback"
            else:
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

    def _match_sface_feature(self, live_feature):
        if live_feature is None or self.sface is None:
            return None, -1.0

        best_label = None
        best_score = -1.0
        for label, features in self.sface_features_by_label.items():
            for auth_feature in features:
                try:
                    score = float(self.sface.match(
                        auth_feature,
                        live_feature,
                        cv2.FaceRecognizerSF_FR_COSINE,
                    ))
                except Exception:
                    continue
                if score > best_score:
                    best_score = score
                    best_label = label

        return best_label, best_score

    def _authenticate_frame_dnn(self, frame, face_rows, current_time):
        if not face_rows or self.sface is None or not any(self.sface_features_by_label.values()):
            return False

        try:
            face_row = np.asarray(face_rows[0], dtype=np.float32)
            aligned = self.sface.alignCrop(frame, face_row)
            live_feature = self.sface.feature(aligned)
        except Exception as e:
            self.last_match_result = f"DNN align/feature failed: {e}"
            self.last_match_time = current_time
            self.last_check_time = current_time
            return False

        predicted_label, score = self._match_sface_feature(live_feature)
        final_match = score >= SFACE_COSINE_THRESHOLD
        strong_match = score >= SFACE_STRONG_COSINE_THRESHOLD
        soft_fail = self.authorized and (not final_match) and score >= (SFACE_COSINE_THRESHOLD - SFACE_SOFT_FAIL_MARGIN)

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
            required_successes = 1 if strong_match else SUCCESS_STREAK_REQUIRED
            if self.success_streak >= required_successes and stable_pass and not self.authorized:
                self.authorized = True
                print("[FACE AUTH HA] Authorized - Gesture control ENABLED")
        elif soft_fail:
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
            f"id={identity} sface={score:.3f} "
            f"final={'MATCH' if final_match else ('SOFT_FAIL' if soft_fail else 'NO_MATCH')} pass={pass_ratio:.2f}"
        )
        self.last_match_time = current_time
        self.last_check_time = current_time
        return True

    def authenticate_frame(self, frame):
        current_time = time.time()

        gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
        dnn_face_rows = []
        dnn_detection = self._detect_faces_dnn(frame)
        if dnn_detection is not None:
            faces, dnn_face_rows = dnn_detection
        else:
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
            and (len(self.auth_descriptors) > 0 or any(self.sface_features_by_label.values()))
            and time_since_last_check >= check_interval
            and face_detected_duration >= MIN_FACE_TIME
        )

        if can_check:
            if self.dnn_enabled and dnn_face_rows and any(self.sface_features_by_label.values()):
                if self._authenticate_frame_dnn(frame, dnn_face_rows, current_time):
                    return face_detected, faces

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
