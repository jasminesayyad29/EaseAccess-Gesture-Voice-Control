import argparse
import os
import time
from typing import Optional

import cv2
import numpy as np

# Register images in the same format expected by face_auth_high_accuracy.py
CAMERA_INDEX = 0
FRAME_WIDTH = 640
FRAME_HEIGHT = 480
TARGET_FACE_SIZE = (200, 200)

DEFAULT_SAVE_COUNT = 18
MIN_SECONDS_BETWEEN_SAVES = 0.35

# Quality gates tuned for stable LBPH + ORB training samples
MIN_FACE_SIZE_RATIO = 0.11
MIN_BRIGHTNESS = 55.0
MAX_BRIGHTNESS = 200.0
MIN_SHARPNESS = 85.0
MIN_EYE_PAIRS = 1

PROMPT_LINES = [
    "Position your face inside the guide box.",
    "Look straight, then slightly left/right/up/down.",
    "Keep your expression natural and avoid blur.",
]


class FaceRegister:
    def __init__(self, known_faces_folder: str):
        self.known_faces_folder = known_faces_folder
        self.face_cascade = cv2.CascadeClassifier(
            cv2.data.haarcascades + "haarcascade_frontalface_default.xml"
        )
        self.eye_cascade = cv2.CascadeClassifier(cv2.data.haarcascades + "haarcascade_eye.xml")
        self.clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))

        if self.face_cascade.empty() or self.eye_cascade.empty():
            raise RuntimeError("Could not load Haar cascades for face/eye detection")

    def _enhance_gray(self, gray: np.ndarray) -> np.ndarray:
        gray = cv2.GaussianBlur(gray, (3, 3), 0)
        gray = self.clahe.apply(gray)
        return gray

    def _detect_faces(self, gray: np.ndarray):
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

        faces = self.face_cascade.detectMultiScale(gray, scaleFactor=1.08, minNeighbors=4, minSize=(50, 50))
        return faces

    def _quality_metrics(self, face_gray: np.ndarray, face_w: int, face_h: int, frame_w: int, frame_h: int):
        brightness = float(np.mean(face_gray))
        sharpness = float(cv2.Laplacian(face_gray, cv2.CV_64F).var())

        face_area_ratio = (face_w * face_h) / float(max(1, frame_w * frame_h))

        eyes = self.eye_cascade.detectMultiScale(face_gray, scaleFactor=1.1, minNeighbors=4, minSize=(18, 18))
        eye_pairs = len(eyes) // 2 if len(eyes) >= 2 else len(eyes)

        return {
            "brightness": brightness,
            "sharpness": sharpness,
            "face_area_ratio": face_area_ratio,
            "eye_pairs": eye_pairs,
        }

    def _passes_quality(self, metrics):
        return (
            metrics["face_area_ratio"] >= MIN_FACE_SIZE_RATIO
            and MIN_BRIGHTNESS <= metrics["brightness"] <= MAX_BRIGHTNESS
            and metrics["sharpness"] >= MIN_SHARPNESS
            and metrics["eye_pairs"] >= MIN_EYE_PAIRS
        )

    def _expand_face_rect(self, x: int, y: int, w: int, h: int, frame_w: int, frame_h: int):
        # Haar detections are often tight; expand asymmetrically to preserve full head area.
        pad_left = int(w * 0.22)
        pad_right = int(w * 0.22)
        pad_top = int(h * 0.35)
        pad_bottom = int(h * 0.25)

        x1 = max(0, x - pad_left)
        y1 = max(0, y - pad_top)
        x2 = min(frame_w, x + w + pad_right)
        y2 = min(frame_h, y + h + pad_bottom)

        # Keep ROI roughly square before resizing so face geometry is less distorted.
        roi_w = x2 - x1
        roi_h = y2 - y1
        if roi_w > 0 and roi_h > 0:
            if roi_w > roi_h:
                add = (roi_w - roi_h) // 2
                y1 = max(0, y1 - add)
                y2 = min(frame_h, y2 + add)
            elif roi_h > roi_w:
                add = (roi_h - roi_w) // 2
                x1 = max(0, x1 - add)
                x2 = min(frame_w, x2 + add)

        return x1, y1, x2, y2

    def capture(self, person_name: str, save_count: int = DEFAULT_SAVE_COUNT):
        os.makedirs(self.known_faces_folder, exist_ok=True)

        cap = cv2.VideoCapture(CAMERA_INDEX, cv2.CAP_DSHOW)
        if not cap.isOpened():
            cap = cv2.VideoCapture(CAMERA_INDEX)

        cap.set(cv2.CAP_PROP_FRAME_WIDTH, FRAME_WIDTH)
        cap.set(cv2.CAP_PROP_FRAME_HEIGHT, FRAME_HEIGHT)

        if not cap.isOpened():
            raise RuntimeError("Cannot open webcam")

        window_name = "Face Register (Q: Quit | C: Force Capture)"
        cv2.namedWindow(window_name, cv2.WINDOW_NORMAL)

        saved = 0
        last_save_time = 0.0
        sample_index = 1

        print(f"[REGISTER] Saving {save_count} samples for '{person_name}' to: {self.known_faces_folder}")
        print("[REGISTER] Automatic capture will save only good-quality frames.")

        try:
            while True:
                ret, frame = cap.read()
                if not ret:
                    print("[REGISTER] Failed to read frame.")
                    continue

                display = frame.copy()
                gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
                faces = self._detect_faces(gray)

                h, w = display.shape[:2]
                guide_size = int(min(w, h) * 0.62)
                gx1 = (w - guide_size) // 2
                gy1 = (h - guide_size) // 2
                gx2 = gx1 + guide_size
                gy2 = gy1 + guide_size
                cv2.rectangle(display, (gx1, gy1), (gx2, gy2), (220, 220, 220), 1)

                status_color = (0, 0, 255)
                status_text = "No face"
                quality_line = ""

                selected = None
                metrics = None

                if len(faces) > 0:
                    x, y, fw, fh = max(faces, key=lambda r: r[2] * r[3])
                    ex1, ey1, ex2, ey2 = self._expand_face_rect(x, y, fw, fh, w, h)
                    selected = (ex1, ey1, ex2 - ex1, ey2 - ey1)

                    face_gray = gray[ey1:ey2, ex1:ex2]
                    if face_gray.size > 0:
                        face_gray_enhanced = self._enhance_gray(face_gray)
                        metrics = self._quality_metrics(
                            face_gray_enhanced,
                            ex2 - ex1,
                            ey2 - ey1,
                            w,
                            h,
                        )

                        if self._passes_quality(metrics):
                            status_color = (0, 200, 0)
                            status_text = "Quality OK"
                        else:
                            status_color = (0, 165, 255)
                            status_text = "Adjust pose/light"

                        quality_line = (
                            f"bri={metrics['brightness']:.0f} sharp={metrics['sharpness']:.0f} "
                            f"size={metrics['face_area_ratio']:.2f} eyes={metrics['eye_pairs']}"
                        )

                if selected is not None:
                    x, y, fw, fh = selected
                    cv2.rectangle(display, (x, y), (x + fw, y + fh), status_color, 2)

                    centered = x >= gx1 and y >= gy1 and (x + fw) <= gx2 and (y + fh) <= gy2

                    if metrics is not None and self._passes_quality(metrics) and centered:
                        now = time.time()
                        if now - last_save_time >= MIN_SECONDS_BETWEEN_SAVES and saved < save_count:
                            sx, sy, sw, sh = selected
                            face_gray = gray[sy : sy + sh, sx : sx + sw]
                            face_norm = cv2.resize(face_gray, TARGET_FACE_SIZE)
                            face_norm = self._enhance_gray(face_norm)

                            save_name = f"{person_name}_{sample_index:02d}.jpg"
                            save_path = os.path.join(self.known_faces_folder, save_name)
                            cv2.imwrite(save_path, face_norm)

                            saved += 1
                            sample_index += 1
                            last_save_time = now
                            status_text = f"Captured {saved}/{save_count}"
                            status_color = (0, 220, 0)

                for idx, line in enumerate(PROMPT_LINES):
                    cv2.putText(
                        display,
                        line,
                        (12, 26 + idx * 22),
                        cv2.FONT_HERSHEY_SIMPLEX,
                        0.57,
                        (242, 242, 242),
                        1,
                        cv2.LINE_AA,
                    )

                cv2.putText(
                    display,
                    f"Person: {person_name}",
                    (12, h - 48),
                    cv2.FONT_HERSHEY_SIMPLEX,
                    0.62,
                    (255, 255, 255),
                    2,
                    cv2.LINE_AA,
                )
                cv2.putText(
                    display,
                    f"{status_text} | Saved: {saved}/{save_count}",
                    (12, h - 24),
                    cv2.FONT_HERSHEY_SIMPLEX,
                    0.58,
                    status_color,
                    2,
                    cv2.LINE_AA,
                )

                if quality_line:
                    cv2.putText(
                        display,
                        quality_line,
                        (12, min(h - 70, 110)),
                        cv2.FONT_HERSHEY_SIMPLEX,
                        0.5,
                        (235, 235, 235),
                        1,
                        cv2.LINE_AA,
                    )

                cv2.imshow(window_name, display)

                if saved >= save_count:
                    print("[REGISTER] Capture completed.")
                    break

                key = cv2.waitKey(1) & 0xFF
                if key == ord("q"):
                    print("[REGISTER] Stopped by user.")
                    break
                if key == ord("c") and selected is not None:
                    x, y, fw, fh = selected
                    face_gray = gray[y : y + fh, x : x + fw]
                    if face_gray.size > 0:
                        face_norm = cv2.resize(face_gray, TARGET_FACE_SIZE)
                        face_norm = self._enhance_gray(face_norm)
                        save_name = f"{person_name}_{sample_index:02d}.jpg"
                        save_path = os.path.join(self.known_faces_folder, save_name)
                        cv2.imwrite(save_path, face_norm)
                        saved += 1
                        sample_index += 1
                        print(f"[REGISTER] Forced capture {saved}/{save_count}: {save_name}")

        finally:
            cap.release()
            cv2.destroyAllWindows()

        return saved


def _resolve_known_faces_folder(folder_arg: Optional[str]) -> str:
    script_dir = os.path.dirname(os.path.abspath(__file__))
    folder_candidate = folder_arg or "known_faces"

    if os.path.isabs(folder_candidate):
        return folder_candidate

    local = os.path.join(script_dir, folder_candidate)
    if os.path.exists(local):
        return local

    return os.path.join(os.getcwd(), folder_candidate)


def main():
    parser = argparse.ArgumentParser(
        description="Capture high-quality authorized face images for face_auth_high_accuracy"
    )
    parser.add_argument("name", help="Identity name prefix for saved files (e.g. aadar)")
    parser.add_argument("--count", type=int, default=DEFAULT_SAVE_COUNT, help="Number of samples to save")
    parser.add_argument(
        "--folder",
        default=None,
        help="Known faces folder path (default: src/known_faces or ./known_faces)",
    )

    args = parser.parse_args()
    person_name = args.name.strip().lower().replace(" ", "_")

    if not person_name:
        raise ValueError("Name cannot be empty")

    folder = _resolve_known_faces_folder(args.folder)
    registrar = FaceRegister(folder)
    saved = registrar.capture(person_name=person_name, save_count=max(1, args.count))

    print(f"[REGISTER] Done. Saved {saved} images for '{person_name}' in: {folder}")
    print("[REGISTER] You can now run face_auth_high_accuracy.py to verify recognition quality.")


if __name__ == "__main__":
    main()
