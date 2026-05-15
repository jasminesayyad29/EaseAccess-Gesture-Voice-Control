import cv2
import mediapipe as mp
import pyautogui
import math
import argparse
import sys
import ctypes
from concurrent.futures import ThreadPoolExecutor
from enum import IntEnum
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from google.protobuf.json_format import MessageToDict
import screen_brightness_control as sbcontrol
import time
import traceback
import win32con
import win32gui

from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume


def _enable_high_dpi_awareness():
    if sys.platform != "win32":
        return

    try:
        shcore = ctypes.windll.shcore
        try:
            shcore.SetProcessDpiAwareness(2)
        except Exception:
            ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass


_enable_high_dpi_awareness()

FaceAuthenticator = None
try:
    from face_auth_high_accuracy import FaceAuthenticatorHighAccuracy as FaceAuthenticator
except Exception:
    FaceAuthenticator = None

try:
    from gest_auth_indicator import DesktopAuthIndicator
except Exception:
    DesktopAuthIndicator = None

pyautogui.FAILSAFE = False
pyautogui.PAUSE = 0
if hasattr(pyautogui, "MINIMUM_DURATION"):
    pyautogui.MINIMUM_DURATION = 0
mp_drawing = mp.solutions.drawing_utils
mp_hands = mp.solutions.hands

# Debug/runtime controls
VERBOSE_FRAME_LOGS = False
AUTHORIZATION_HOLD_SECONDS = 3.5
FACE_AUTH_FAILURE_EXIT_CODE = 86

# Preview UI controls
PREVIEW_WINDOW_NAME = 'EaseAccess Preview (Press Q to quit)'
PIP_WIDTH = 468
PIP_MARGIN = 20
CAMERA_WIDTH = 1280
CAMERA_HEIGHT = 720
CAMERA_FPS = 60
CURSOR_DEADZONE_PX = 4
CURSOR_MIN_GAIN = 0.30
CURSOR_MAX_GAIN = 2.40
CURSOR_SMOOTHING = 0.28
GESTURE_STABILITY_FRAMES = 3
PRESENTATION_ACTION_COOLDOWN = 0.42
WINDOW_SWITCH_COOLDOWN = 1.1
WINDOW_SWITCH_MAX_HISTORY = 22
WINDOW_SWITCH_MIN_STEP = 0.025
WINDOW_SWITCH_MIN_SIDE = 0.22
WINDOW_SWITCH_MAX_SIDE_RATIO = 1.9
WINDOW_SWITCH_CLOSURE_TOLERANCE = 0.12
WINDOW_SWITCH_AXIS_RATIO = 1.18
WINDOW_SWITCH_STEP_COOLDOWN = 0.05
WINDOW_SWITCH_POST_SWITCH_HOLD_SECONDS = 0.5
WINDOW_SWITCH_SESSION_IDLE_TIMEOUT = 0.85
WINDOW_SWITCH_FINGER_PATTERN = 14
DESKTOP_SWITCH_POSE_HOLD_FRAMES = 4
DESKTOP_SWITCH_STEP_COOLDOWN = 0.18
DESKTOP_SWITCH_POST_SWITCH_HOLD_SECONDS = 0.9
HAND_DETECTION_CONFIDENCE = 0.7
HAND_TRACKING_CONFIDENCE = 0.8
FINGER_EXTENSION_ANGLE_DEGREES = 160.0
FINGER_EXTENSION_WRIST_RATIO = 1.04
WINDOW_SWITCH_HORIZONTAL_THRESHOLD = 0.004
WINDOW_SWITCH_SMOOTHING_ALPHA = 0.35
WINDOW_SWITCH_DIRECTION_HOLD_FRAMES = 2
WINDOW_SWITCH_VERTICAL_THRESHOLD = 0.0025

print("=== GESTURE CONTROLLER DEBUG START ===")

# Gesture Encodings 
class Gest(IntEnum):
    FIST = 0
    PINKY = 1
    RING = 2
    MID = 4
    LAST3 = 7
    INDEX = 8
    FIRST2 = 12
    LAST4 = 15
    THUMB = 16    
    PALM = 31
    V_GEST = 33
    TWO_FINGER_CLOSED = 34
    PINCH_MAJOR = 35
    PINCH_MINOR = 36
    SLIDE_NEXT = 37
    SLIDE_PREV = 38
    ZOOM = 39

class HLabel(IntEnum):
    MINOR = 0
    MAJOR = 1

class HandRecog:
    def __init__(self, hand_label):
        self.finger = 0
        self.ori_gesture = Gest.PALM
        self.prev_gesture = Gest.PALM
        self.frame_count = 0
        self.hand_result = None
        self.hand_label = hand_label
        self.last_presentation_action_time = 0.0
        self.window_switch_points = []
        self.last_window_switch_action_time = 0.0
        self.window_switch_alt_held = False
        self.window_switch_last_axis = None
        self.window_switch_last_sign = 0
        self.window_switch_vertical_last_sign = 0
        self.window_switch_last_x = None
        self.window_switch_last_y = None
        self.window_switch_smoothed_x = None
        self.window_switch_smoothed_y = None
        self.window_switch_direction_hold = 0
        self.window_switch_last_motion_time = 0.0
        self.window_switch_turn_count = 0
        self.desktop_switch_last_sign = 0
        self.desktop_switch_last_x = None
        self.desktop_switch_smoothed_x = None
        self.desktop_switch_direction_hold = 0
        self.desktop_switch_last_motion_time = 0.0
        self.desktop_switch_last_action_time = 0.0
        self.desktop_switch_pose_frames = 0
    
    def update_hand_result(self, hand_result):
        self.hand_result = hand_result

    def get_signed_dist(self, point):
        sign = -1
        if self.hand_result.landmark[point[0]].y < self.hand_result.landmark[point[1]].y:
            sign = 1
        dist = (self.hand_result.landmark[point[0]].x - self.hand_result.landmark[point[1]].x)**2
        dist += (self.hand_result.landmark[point[0]].y - self.hand_result.landmark[point[1]].y)**2
        dist = math.sqrt(dist)
        return dist*sign
    
    def get_dist(self, point):
        dist = (self.hand_result.landmark[point[0]].x - self.hand_result.landmark[point[1]].x)**2
        dist += (self.hand_result.landmark[point[0]].y - self.hand_result.landmark[point[1]].y)**2
        dist = math.sqrt(dist)
        return dist
    
    def get_dz(self,point):
        return abs(self.hand_result.landmark[point[0]].z - self.hand_result.landmark[point[1]].z)

    def _joint_angle_degrees(self, first_point, joint_point, third_point):
        first = self.hand_result.landmark[first_point]
        joint = self.hand_result.landmark[joint_point]
        third = self.hand_result.landmark[third_point]

        first_vector_x = first.x - joint.x
        first_vector_y = first.y - joint.y
        third_vector_x = third.x - joint.x
        third_vector_y = third.y - joint.y

        first_length = math.hypot(first_vector_x, first_vector_y)
        third_length = math.hypot(third_vector_x, third_vector_y)
        if first_length == 0.0 or third_length == 0.0:
            return 0.0

        cosine = (first_vector_x * third_vector_x + first_vector_y * third_vector_y) / (first_length * third_length)
        cosine = max(-1.0, min(1.0, cosine))
        return math.degrees(math.acos(cosine))

    def _finger_is_extended(self, mcp_point, pip_point, tip_point):
        angle = self._joint_angle_degrees(mcp_point, pip_point, tip_point)
        wrist = self.hand_result.landmark[0]
        pip = self.hand_result.landmark[pip_point]
        tip = self.hand_result.landmark[tip_point]

        pip_distance = math.hypot(pip.x - wrist.x, pip.y - wrist.y)
        tip_distance = math.hypot(tip.x - wrist.x, tip.y - wrist.y)

        return angle >= FINGER_EXTENSION_ANGLE_DEGREES and tip_distance >= (pip_distance * FINGER_EXTENSION_WRIST_RATIO)

    def _window_switch_begin_session(self, now, axis, sign):
        self.window_switch_alt_held = True

        self.window_switch_last_axis = axis
        self.window_switch_last_sign = sign
        self.window_switch_turn_count = 0
        self.window_switch_last_motion_time = now

    def _window_switch_end_session(self):
        if self.window_switch_alt_held:
            try:
                pyautogui.keyUp('alt')
            except Exception:
                pass

        self.window_switch_alt_held = False
        self.window_switch_last_axis = None
        self.window_switch_last_sign = 0
        self.window_switch_vertical_last_sign = 0
        self.window_switch_last_x = None
        self.window_switch_last_y = None
        self.window_switch_smoothed_x = None
        self.window_switch_smoothed_y = None
        self.window_switch_direction_hold = 0
        self.window_switch_last_motion_time = 0.0
        self.window_switch_turn_count = 0

    def _window_switch_pose_active(self):
        if self.hand_result is None:
            return False

        return self.finger == WINDOW_SWITCH_FINGER_PATTERN

    def _window_switch_point(self):
        landmarks = self.hand_result.landmark
        centroid_x = (landmarks[8].x + landmarks[12].x + landmarks[16].x) / 3.0
        centroid_y = (landmarks[8].y + landmarks[12].y + landmarks[16].y) / 3.0
        palm_span = math.hypot(landmarks[9].x - landmarks[0].x, landmarks[9].y - landmarks[0].y)
        palm_span = max(0.05, palm_span)
        return (centroid_x / palm_span, centroid_y / palm_span)

    def _window_switch_horizontal_point(self):
        landmarks = self.hand_result.landmark
        return (landmarks[8].x + landmarks[12].x + landmarks[16].x) / 3.0

    def _window_switch_vertical_point(self):
        landmarks = self.hand_result.landmark
        return (landmarks[8].y + landmarks[12].y + landmarks[16].y) / 3.0

    def _desktop_switch_pose_active(self):
        if self.hand_result is None:
            return False

        return self.finger == Gest.LAST4

    def _desktop_switch_horizontal_point(self):
        landmarks = self.hand_result.landmark
        return (landmarks[8].x + landmarks[12].x + landmarks[16].x + landmarks[20].x) / 4.0

    def _desktop_switch_end_session(self):
        self.desktop_switch_last_sign = 0
        self.desktop_switch_last_x = None
        self.desktop_switch_smoothed_x = None
        self.desktop_switch_direction_hold = 0
        self.desktop_switch_last_motion_time = 0.0
        self.desktop_switch_last_action_time = 0.0
        self.desktop_switch_pose_frames = 0

    def _window_switch_minimize_all_windows(self):
        try:
            def minimize_window(hwnd, _extra):
                try:
                    if not win32gui.IsWindowVisible(hwnd):
                        return
                    win32gui.ShowWindow(hwnd, win32con.SW_FORCEMINIMIZE)
                except Exception:
                    pass

            win32gui.EnumWindows(minimize_window, None)
            return
        except Exception:
            pass

        if hasattr(ctypes, "windll") and hasattr(ctypes.windll, "user32"):
            user32 = ctypes.windll.user32
            VK_LWIN = 0x5B
            VK_D = 0x44
            KEYEVENTF_KEYUP = 0x0002
            user32.keybd_event(VK_LWIN, 0, 0, 0)
            user32.keybd_event(VK_D, 0, 0, 0)
            user32.keybd_event(VK_D, 0, KEYEVENTF_KEYUP, 0)
            user32.keybd_event(VK_LWIN, 0, KEYEVENTF_KEYUP, 0)
            return

        pyautogui.hotkey('win', 'd')

    def _compress_window_switch_path(self, points):
        compressed = [points[0]]
        for point in points[1:]:
            if math.hypot(point[0] - compressed[-1][0], point[1] - compressed[-1][1]) >= WINDOW_SWITCH_MIN_STEP:
                compressed.append(point)
        return compressed

    def _direction_from_segment(self, start_point, end_point):
        delta_x = end_point[0] - start_point[0]
        delta_y = end_point[1] - start_point[1]
        magnitude = math.hypot(delta_x, delta_y)
        if magnitude < WINDOW_SWITCH_MIN_STEP:
            return None

        if abs(delta_x) >= abs(delta_y) * WINDOW_SWITCH_AXIS_RATIO:
            return ("h", 1 if delta_x > 0 else -1, magnitude)
        if abs(delta_y) >= abs(delta_x) * WINDOW_SWITCH_AXIS_RATIO:
            return ("v", 1 if delta_y > 0 else -1, magnitude)
        return None

    def _is_window_switch_pattern(self):
        if len(self.window_switch_points) < 6:
            return False

        points = self._compress_window_switch_path(self.window_switch_points[-WINDOW_SWITCH_MAX_HISTORY:])
        if len(points) < 6:
            return False

        xs = [point[0] for point in points]
        ys = [point[1] for point in points]
        span_x = max(xs) - min(xs)
        span_y = max(ys) - min(ys)
        if span_x < WINDOW_SWITCH_MIN_SIDE or span_y < WINDOW_SWITCH_MIN_SIDE:
            return False

        side_ratio = span_x / max(span_y, 1e-6)
        if side_ratio > WINDOW_SWITCH_MAX_SIDE_RATIO or side_ratio < (1.0 / WINDOW_SWITCH_MAX_SIDE_RATIO):
            return False

        if math.hypot(points[-1][0] - points[0][0], points[-1][1] - points[0][1]) > WINDOW_SWITCH_CLOSURE_TOLERANCE:
            return False

        directions = []
        for start_point, end_point in zip(points, points[1:]):
            segment = self._direction_from_segment(start_point, end_point)
            if segment is None:
                return False

            axis, sign, magnitude = segment
            if magnitude < WINDOW_SWITCH_MIN_STEP:
                continue

            if not directions or directions[-1][:2] != (axis, sign):
                directions.append((axis, sign))

        if len(directions) < 4:
            return False

        first_four = directions[:4]
        axes = [item[0] for item in first_four]
        signs = [item[1] for item in first_four]

        if axes[0] == axes[1] or axes[1] == axes[2] or axes[2] == axes[3]:
            return False

        if signs[0] == signs[2] or signs[1] == signs[3]:
            return False

        return True
    
    def set_finger_state(self):
        if self.hand_result == None:
            return

        finger_points = [[5, 6, 8], [9, 10, 12], [13, 14, 16], [17, 18, 20]]
        finger_state = 0

        for mcp_point, pip_point, tip_point in finger_points:
            finger_state = finger_state << 1
            if self._finger_is_extended(mcp_point, pip_point, tip_point):
                finger_state = finger_state | 1

        self.finger = finger_state

    def get_gesture(self):
        if self.hand_result == None:
            return Gest.PALM

        current_gesture = Gest.PALM
        if self.finger in [Gest.LAST3,Gest.LAST4] and self.get_dist([8,4]) < 0.05:
            if self.hand_label == HLabel.MINOR :
                current_gesture = Gest.PINCH_MINOR
            else:
                current_gesture = Gest.PINCH_MAJOR

        elif Gest.FIRST2 == self.finger :
            point = [[8,12],[5,9]]
            dist1 = self.get_dist(point[0])
            dist2 = self.get_dist(point[1])
            ratio = dist1/dist2
            if ratio > 1.7:
                current_gesture = Gest.V_GEST
            else:
                if self.get_dz([8,12]) < 0.1:
                    current_gesture =  Gest.TWO_FINGER_CLOSED
                else:
                    current_gesture =  Gest.MID
        else:
            current_gesture =  self.finger
        
        if current_gesture == self.prev_gesture:
            self.frame_count += 1
        else:
            self.frame_count = 0

        self.prev_gesture = current_gesture

        if self.finger in (WINDOW_SWITCH_FINGER_PATTERN, Gest.LAST4):
            self.ori_gesture = Gest.PALM
            return self.ori_gesture

        if self.frame_count >= GESTURE_STABILITY_FRAMES:
            self.ori_gesture = current_gesture
        return self.ori_gesture
    
    def perform_presentation_action(self):
    
        if self.hand_result is None:
            return

        now = time.monotonic()
        if now - self.last_presentation_action_time < PRESENTATION_ACTION_COOLDOWN:
            return self.ori_gesture

        thumb_tip = self.hand_result.landmark[4]
        thumb_ip = self.hand_result.landmark[3]

        # 1️⃣ Next Slide (Thumb pointing right)
        if (thumb_tip.x > thumb_ip.x) and self.finger == Gest.THUMB:
            pyautogui.press('right')
            self.last_presentation_action_time = now
            print("Next Slide")

        # 2️⃣ Previous Slide (Thumb pointing left)
        elif (thumb_tip.x < thumb_ip.x) and self.finger == Gest.THUMB:
            pyautogui.press('left')
            self.last_presentation_action_time = now
            print("Previous Slide")

        # 3️⃣ Zoom In (Pinch gesture where thumb and index move apart)
        # pinch_distance = math.sqrt(
        #     (thumb_tip.x - index_tip.x)**2 +
        #     (thumb_tip.y - index_tip.y)**2
        # )

        # # track pinch motion — start when close, then zoom when moving apart
        # if pinch_distance < 0.03:
        #     self.pinch_start = pinch_distance
        # elif hasattr(self, 'pinch_start') and pinch_distance - self.pinch_start > 0.04:
        #     pyautogui.hotkey('ctrl', '+')
        #     print("Zoom In")
        #     del self.pinch_start
        #     time.sleep(0.5)

        return self.ori_gesture

    def perform_window_switch_action(self):
        if self.hand_result is None:
            self.window_switch_points.clear()
            self._window_switch_end_session()
            return

        now = time.monotonic()
        if not self._window_switch_pose_active():
            self.window_switch_points.clear()
            self._window_switch_end_session()
            return

        current_x = self._window_switch_horizontal_point()

        if self.window_switch_last_x is None or self.window_switch_smoothed_x is None:
            self.window_switch_last_x = current_x
            self.window_switch_smoothed_x = current_x
            self.window_switch_last_y = self._window_switch_vertical_point()
            self.window_switch_smoothed_y = self.window_switch_last_y
            self.window_switch_direction_hold = 0
            if not self.window_switch_alt_held:
                pyautogui.keyDown('alt')
                self.window_switch_alt_held = True
            self.window_switch_last_motion_time = now
            return

        current_y = self._window_switch_vertical_point()
        self.window_switch_smoothed_x = (
            WINDOW_SWITCH_SMOOTHING_ALPHA * current_x
            + (1.0 - WINDOW_SWITCH_SMOOTHING_ALPHA) * self.window_switch_smoothed_x
        )
        self.window_switch_smoothed_y = (
            WINDOW_SWITCH_SMOOTHING_ALPHA * current_y
            + (1.0 - WINDOW_SWITCH_SMOOTHING_ALPHA) * self.window_switch_smoothed_y
        )
        delta_x = self.window_switch_smoothed_x - self.window_switch_last_x
        delta_y = self.window_switch_smoothed_y - self.window_switch_last_y

        if abs(delta_y) >= WINDOW_SWITCH_VERTICAL_THRESHOLD and abs(delta_y) >= abs(delta_x) * 1.05:
            direction_sign = 1 if delta_y > 0 else -1
            if direction_sign != self.window_switch_vertical_last_sign:
                self.window_switch_vertical_last_sign = direction_sign
                self.window_switch_direction_hold = 0
                self.window_switch_last_x = self.window_switch_smoothed_x
                self.window_switch_last_y = self.window_switch_smoothed_y
                self.window_switch_last_motion_time = now

                if direction_sign > 0:
                    self._window_switch_minimize_all_windows()
                    print("Minimize All Windows")
                
            return

        if abs(delta_x) < WINDOW_SWITCH_HORIZONTAL_THRESHOLD:
            if self.window_switch_alt_held and (now - self.window_switch_last_motion_time > WINDOW_SWITCH_SESSION_IDLE_TIMEOUT):
                self.window_switch_points.clear()
                self._window_switch_end_session()
            return

        direction_sign = 1 if delta_x > 0 else -1
        if direction_sign != self.window_switch_last_sign:
            self.window_switch_last_sign = direction_sign
            self.window_switch_direction_hold = 1
            self.window_switch_last_x = self.window_switch_smoothed_x
            self.window_switch_last_motion_time = now
            return

        self.window_switch_direction_hold += 1
        if self.window_switch_direction_hold < WINDOW_SWITCH_DIRECTION_HOLD_FRAMES:
            self.window_switch_last_x = self.window_switch_smoothed_x
            self.window_switch_last_motion_time = now
            return

        if now - self.last_window_switch_action_time < WINDOW_SWITCH_POST_SWITCH_HOLD_SECONDS:
            self.window_switch_last_x = self.window_switch_smoothed_x
            self.window_switch_last_motion_time = now
            self.window_switch_direction_hold = 0
            return

        if now - self.last_window_switch_action_time < WINDOW_SWITCH_STEP_COOLDOWN:
            self.window_switch_last_x = self.window_switch_smoothed_x
            self.window_switch_last_motion_time = now
            return

        if direction_sign > 0:
            pyautogui.hotkey('alt', 'tab')
        else:
            pyautogui.hotkey('alt', 'shift', 'tab')

        self.last_window_switch_action_time = now
        self.window_switch_last_x = self.window_switch_smoothed_x
        self.window_switch_last_y = self.window_switch_smoothed_y
        self.window_switch_last_motion_time = now
        self.window_switch_direction_hold = 0
        print("Window switch step")

    def perform_desktop_switch_action(self):
        if self.hand_result is None:
            self._desktop_switch_end_session()
            return

        if not self._desktop_switch_pose_active():
            self._desktop_switch_end_session()
            return

        self.desktop_switch_pose_frames += 1
        if self.desktop_switch_pose_frames < DESKTOP_SWITCH_POSE_HOLD_FRAMES:
            self.desktop_switch_last_motion_time = time.monotonic()
            return

        now = time.monotonic()
        current_x = self._desktop_switch_horizontal_point()

        if self.desktop_switch_last_x is None or self.desktop_switch_smoothed_x is None:
            self.desktop_switch_last_x = current_x
            self.desktop_switch_smoothed_x = current_x
            self.desktop_switch_direction_hold = 0
            self.desktop_switch_last_motion_time = now
            return

        self.desktop_switch_smoothed_x = (
            WINDOW_SWITCH_SMOOTHING_ALPHA * current_x
            + (1.0 - WINDOW_SWITCH_SMOOTHING_ALPHA) * self.desktop_switch_smoothed_x
        )
        delta_x = self.desktop_switch_smoothed_x - self.desktop_switch_last_x

        if abs(delta_x) < WINDOW_SWITCH_HORIZONTAL_THRESHOLD:
            if now - self.desktop_switch_last_motion_time > WINDOW_SWITCH_SESSION_IDLE_TIMEOUT:
                self._desktop_switch_end_session()
            return

        direction_sign = 1 if delta_x > 0 else -1
        if direction_sign != self.desktop_switch_last_sign:
            self.desktop_switch_last_sign = direction_sign
            self.desktop_switch_direction_hold = 1
            self.desktop_switch_last_x = self.desktop_switch_smoothed_x
            self.desktop_switch_last_motion_time = now
            return

        self.desktop_switch_direction_hold += 1
        if self.desktop_switch_direction_hold < WINDOW_SWITCH_DIRECTION_HOLD_FRAMES:
            self.desktop_switch_last_x = self.desktop_switch_smoothed_x
            self.desktop_switch_last_motion_time = now
            return

        if now - self.desktop_switch_last_action_time < DESKTOP_SWITCH_POST_SWITCH_HOLD_SECONDS:
            self.desktop_switch_last_x = self.desktop_switch_smoothed_x
            self.desktop_switch_last_motion_time = now
            self.desktop_switch_direction_hold = 0
            return

        if now - self.desktop_switch_last_action_time < DESKTOP_SWITCH_STEP_COOLDOWN:
            self.desktop_switch_last_x = self.desktop_switch_smoothed_x
            self.desktop_switch_last_motion_time = now
            return

        if direction_sign > 0:
            pyautogui.hotkey('win', 'ctrl', 'right')
        else:
            pyautogui.hotkey('win', 'ctrl', 'left')

        self.desktop_switch_last_action_time = now
        self.desktop_switch_last_x = self.desktop_switch_smoothed_x
        self.desktop_switch_last_motion_time = now
        self.desktop_switch_direction_hold = 0
        print("Desktop switch step")



class Controller:
    tx_old = 0
    ty_old = 0
    trial = True
    flag = False
    grabflag = False
    pinchmajorflag = False
    pinchminorflag = False
    pinchstartxcoord = None
    pinchstartycoord = None
    pinchdirectionflag = None
    prevpinchlv = 0
    pinchlv = 0
    framecount = 0
    prev_hand = None
    active_hand_key = None
    prev_hand_positions = {}
    cursor_positions = {}
    pinch_threshold = 0.3
    dpi_level = 50

    @staticmethod
    def set_dpi_level(level):
        try:
            Controller.dpi_level = max(1, min(100, int(level)))
        except Exception:
            Controller.dpi_level = 50

    @staticmethod
    def _dpi_profile():
        normalized = (Controller.dpi_level - 1) / 99.0
        speed_curve = normalized ** 0.55
        deadzone = max(1, int(round(6 - (5 * speed_curve))))
        min_gain = 1.60 + (98.40 * speed_curve)
        max_gain = 32.00 + (588.00 * speed_curve)
        smoothing = 0.22 - (0.20 * speed_curve)
        accel_divisor = 260.0 - (259.5 * speed_curve)
        return deadzone, min_gain, max_gain, smoothing, accel_divisor

    @staticmethod
    def reset_control_state():
        if Controller.grabflag:
            try:
                pyautogui.mouseUp(button="left")
            except Exception:
                pass
        Controller.flag = False
        Controller.grabflag = False
        Controller.pinchmajorflag = False
        Controller.pinchminorflag = False
        Controller.pinchstartxcoord = None
        Controller.pinchstartycoord = None
        Controller.pinchdirectionflag = None
        Controller.prevpinchlv = 0
        Controller.pinchlv = 0
        Controller.framecount = 0

    @staticmethod
    def reset_motion_state(hand_key=None):
        if hand_key is None:
            Controller.prev_hand_positions.clear()
            Controller.cursor_positions.clear()
            return

        Controller.prev_hand_positions.pop(hand_key, None)
        Controller.cursor_positions.pop(hand_key, None)
    
    @staticmethod
    def getpinchylv(hand_result):
        dist = round((Controller.pinchstartycoord - hand_result.landmark[8].y)*10,1)
        return dist
    
    @staticmethod
    def getpinchxlv(hand_result):
        dist = round((hand_result.landmark[8].x - Controller.pinchstartxcoord)*10,1)
        return dist
    
    @staticmethod
    def changesystembrightness():
        currentBrightnessLv = sbcontrol.get_brightness(display=0)/100.0
        currentBrightnessLv += Controller.pinchlv/50.0
        if currentBrightnessLv > 1.0:
            currentBrightnessLv = 1.0
        elif currentBrightnessLv < 0.0:
            currentBrightnessLv = 0.0       
        sbcontrol.fade_brightness(int(100*currentBrightnessLv) , start = sbcontrol.get_brightness(display=0))
    
    @staticmethod
    def changesystemvolume():
        devices = AudioUtilities.GetSpeakers()
        interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        volume = cast(interface, POINTER(IAudioEndpointVolume))
        currentVolumeLv = volume.GetMasterVolumeLevelScalar()
        currentVolumeLv += Controller.pinchlv/50.0
        if currentVolumeLv > 1.0:
            currentVolumeLv = 1.0
        elif currentVolumeLv < 0.0:
            currentVolumeLv = 0.0
        volume.SetMasterVolumeLevelScalar(currentVolumeLv, None)
    
    @staticmethod
    def scrollVertical():
        pyautogui.scroll(120 if Controller.pinchlv>0.0 else -120)
        
    @staticmethod
    def scrollHorizontal():
        pyautogui.keyDown('shift')
        pyautogui.keyDown('ctrl')
        pyautogui.scroll(-120 if Controller.pinchlv>0.0 else 120)
        pyautogui.keyUp('ctrl')
        pyautogui.keyUp('shift')


        # ---------- Presentation Controls ----------
    @staticmethod
    def next_slide():
        """Move to the next presentation slide"""
        pyautogui.press('right')
        print("Next Slide Triggered")

    @staticmethod
    def previous_slide():
        """Move to the previous presentation slide"""
        pyautogui.press('left')
        print("Previous Slide Triggered")

    # def zoom_in():
    #     """Zoom in on the current presentation"""
    #     pyautogui.hotkey('ctrl', '+')
    #     print("Zoom In Triggered")

    # def zoom_out():
    #     """Zoom out on the current presentation"""
    #     pyautogui.hotkey('ctrl', '-')
    #     print("Zoom Out Triggered")


    @staticmethod
    def get_position(hand_result, hand_key="default"):
        point = 9
        position = [hand_result.landmark[point].x ,hand_result.landmark[point].y]
        sx,sy = pyautogui.size()
        x_old,y_old = pyautogui.position()
        x = int(position[0] * sx)
        y = int(position[1] * sy)
        deadzone, min_gain, max_gain, smoothing, accel_divisor = Controller._dpi_profile()

        previous_hand = Controller.prev_hand_positions.get(hand_key)
        if previous_hand is None:
            Controller.prev_hand_positions[hand_key] = (x, y)
            Controller.cursor_positions[hand_key] = (x_old, y_old)
            return (x_old, y_old)

        delta_x = x - previous_hand[0]
        delta_y = y - previous_hand[1]
        distance = math.hypot(delta_x, delta_y)
        Controller.prev_hand_positions[hand_key] = (x, y)

        if distance <= deadzone:
            Controller.cursor_positions[hand_key] = (x_old, y_old)
            return (x_old, y_old)

        gain = min(max_gain, max(min_gain, min_gain + (distance / accel_divisor)))
        target_x = x_old + delta_x * gain
        target_y = y_old + delta_y * gain

        previous_cursor = Controller.cursor_positions.get(hand_key, (x_old, y_old))
        smoothing = smoothing if distance < 220 else max(0.18, smoothing - 0.06)
        x_pos = previous_cursor[0] + (target_x - previous_cursor[0]) * smoothing
        y_pos = previous_cursor[1] + (target_y - previous_cursor[1]) * smoothing

        x_pos = max(0, min(sx - 1, x_pos))
        y_pos = max(0, min(sy - 1, y_pos))
        Controller.cursor_positions[hand_key] = (x_pos, y_pos)
        return (int(x_pos), int(y_pos))

    @staticmethod
    def pinch_control_init(hand_result):
        Controller.pinchstartxcoord = hand_result.landmark[8].x
        Controller.pinchstartycoord = hand_result.landmark[8].y
        Controller.pinchlv = 0
        Controller.prevpinchlv = 0
        Controller.framecount = 0

    def pinch_control(hand_result, controlHorizontal, controlVertical):
        if Controller.framecount == 5:
            Controller.framecount = 0
            Controller.pinchlv = Controller.prevpinchlv
            if Controller.pinchdirectionflag == True:
                controlHorizontal()
            elif Controller.pinchdirectionflag == False:
                controlVertical()
        lvx =  Controller.getpinchxlv(hand_result)
        lvy =  Controller.getpinchylv(hand_result)
        if abs(lvy) > abs(lvx) and abs(lvy) > Controller.pinch_threshold:
            Controller.pinchdirectionflag = False
            if abs(Controller.prevpinchlv - lvy) < Controller.pinch_threshold:
                Controller.framecount += 1
            else:
                Controller.prevpinchlv = lvy
                Controller.framecount = 0
        elif abs(lvx) > Controller.pinch_threshold:
            Controller.pinchdirectionflag = True
            if abs(Controller.prevpinchlv - lvx) < Controller.pinch_threshold:
                Controller.framecount += 1
            else:
                Controller.prevpinchlv = lvx
                Controller.framecount = 0

    def handle_controls(gesture, hand_result, is_authorized, hand_key="default"):  
        """Only execute gestures if authorized"""
        if not is_authorized:
            Controller.active_hand_key = None
            Controller.reset_control_state()
            Controller.reset_motion_state(hand_key)
            return

        if Controller.active_hand_key != hand_key:
            Controller.reset_control_state()
            Controller.reset_motion_state(hand_key)
            Controller.active_hand_key = hand_key
        
        x,y = None,None
        if gesture != Gest.PALM :
            x,y = Controller.get_position(hand_result, hand_key)
        
        # flag reset
        if gesture != Gest.FIST and Controller.grabflag:
            Controller.grabflag = False
            pyautogui.mouseUp(button = "left")
        if gesture != Gest.PINCH_MAJOR and Controller.pinchmajorflag:
            Controller.pinchmajorflag = False
        if gesture != Gest.PINCH_MINOR and Controller.pinchminorflag:
            Controller.pinchminorflag = False

        # implementation
        if gesture == Gest.V_GEST:
            Controller.flag = True
            pyautogui.moveTo(x, y)
        elif gesture == Gest.FIST:
            if not Controller.grabflag : 
                Controller.grabflag = True
                pyautogui.mouseDown(button = "left")
            pyautogui.moveTo(x, y)
        elif gesture == Gest.MID and Controller.flag:
            pyautogui.click()
            Controller.flag = False
        elif gesture == Gest.INDEX and Controller.flag:
            pyautogui.click(button='right')
            Controller.flag = False
        elif gesture == Gest.TWO_FINGER_CLOSED and Controller.flag:
            pyautogui.doubleClick()
            Controller.flag = False
        elif gesture == Gest.PINCH_MINOR:
            if Controller.pinchminorflag == False:
                Controller.pinch_control_init(hand_result)
                Controller.pinchminorflag = True
            Controller.pinch_control(hand_result,Controller.scrollHorizontal, Controller.scrollVertical)
        elif gesture == Gest.PINCH_MAJOR:
            if Controller.pinchmajorflag == False:
                Controller.pinch_control_init(hand_result)
                Controller.pinchmajorflag = True
            Controller.pinch_control(hand_result,Controller.changesystembrightness, Controller.changesystemvolume)
        
class GestureController:
    gc_mode = 0
    cap = None
    CAM_HEIGHT = None
    CAM_WIDTH = None
    hr_major = None
    hr_minor = None
    dom_hand = True
    face_auth = None
    auth_gate_authorized = False
    last_authorized_time = 0.0
    auth_error_count = 0
    max_auth_errors = 20

    @staticmethod
    def _gesture_priority(gesture):
        if gesture in (Gest.FIST, Gest.V_GEST):
            return 3
        if gesture in (Gest.PINCH_MAJOR, Gest.PINCH_MINOR):
            return 2
        if gesture in (Gest.MID, Gest.INDEX, Gest.TWO_FINGER_CLOSED):
            return 1
        return 0

    @staticmethod
    def _configure_preview_window(frame_width, frame_height):
        cv2.namedWindow(PREVIEW_WINDOW_NAME, cv2.WINDOW_NORMAL)
        cv2.setWindowProperty(PREVIEW_WINDOW_NAME, cv2.WND_PROP_TOPMOST, 1)
        screen_w, screen_h = pyautogui.size()
        pip_width = max(320, int(screen_w * 0.30))
        pip_height = max(180, int(screen_h * 0.30))
        cv2.resizeWindow(PREVIEW_WINDOW_NAME, pip_width, pip_height)

        pos_x = max(0, int((screen_w - pip_width) / 2))
        pos_y = max(0, int((screen_h - pip_height) / 2))
        cv2.moveWindow(PREVIEW_WINDOW_NAME, pos_x, pos_y)

    @staticmethod
    def _draw_preview_ui(image, is_authorized, raw_is_authorized, face_auth, log_mode=False):
        h, w = image.shape[:2]
        status_text = "AUTHORIZED" if is_authorized else "UNAUTHORIZED"
        if is_authorized and not raw_is_authorized:
            status_text = "AUTHORIZED (HOLD)"
        status_color = (42, 165, 95) if is_authorized else (52, 64, 212)

        overlay = image.copy()
        cv2.rectangle(overlay, (0, 0), (w, 52), (16, 20, 25), -1)
        cv2.rectangle(overlay, (0, h - 34), (w, h), (16, 20, 25), -1)
        image = cv2.addWeighted(overlay, 0.52, image, 0.48, 0)

        cv2.circle(image, (22, 26), 8, status_color, -1)
        cv2.putText(
            image,
            f"Status: {status_text}",
            (38, 31),
            cv2.FONT_HERSHEY_SIMPLEX,
            0.85,
            (245, 245, 245),
            2,
            cv2.LINE_AA,
        )
        cv2.putText(
            image,
            "EaseAccess | Press Q to quit",
            (12, h - 12),
            cv2.FONT_HERSHEY_SIMPLEX,
            .5,
            (220, 220, 220),
            1,
            cv2.LINE_AA,
        )

        if log_mode:
            cv2.putText(
                image,
                "LOG MODE",
                (w - 94, 31),
                cv2.FONT_HERSHEY_SIMPLEX,
                0.52,
                (255, 210, 96),
                1,
                cv2.LINE_AA,
            )

        if face_auth is not None and (time.time() - face_auth.get_last_match_time() < 10):
            cv2.putText(
                image,
                face_auth.get_last_match_result(),
                (12, min(60, h - 36)),
                cv2.FONT_HERSHEY_SIMPLEX,
                0.58,
                (230, 230, 230),
                1,
                cv2.LINE_AA,
            )

        return image

    def __init__(self, disable_face_auth=False, dpi_level=50, log_preview=False):
        print("Initializing GestureController...")
        self.disable_face_auth = bool(disable_face_auth)
        self.log_preview_enabled = bool(log_preview)
        Controller.set_dpi_level(dpi_level)
        print(f"Cursor DPI control set to {Controller.dpi_level}/100")
        self.preview_window_visible = False
        self.desktop_indicator = DesktopAuthIndicator() if DesktopAuthIndicator is not None else None
        self.prev_raw_authorized = False
        self.minimize_after_auth_done = False
        
        # Initialize camera with error handling
        try:
            self.cap = cv2.VideoCapture(0, cv2.CAP_DSHOW)  # Use CAP_DSHOW for Windows
            if not self.cap.isOpened():
                print("Failed to open camera with index 0, trying index 1...")
                self.cap = cv2.VideoCapture(1, cv2.CAP_DSHOW)
            
            if not self.cap.isOpened():
                print("ERROR: Could not open any camera!")
                return
            
            # Set camera properties
            self.cap.set(cv2.CAP_PROP_FRAME_WIDTH, CAMERA_WIDTH)
            self.cap.set(cv2.CAP_PROP_FRAME_HEIGHT, CAMERA_HEIGHT)
            self.cap.set(cv2.CAP_PROP_FPS, CAMERA_FPS)
            self.cap.set(cv2.CAP_PROP_BUFFERSIZE, 1)
            
            self.CAM_HEIGHT = self.cap.get(cv2.CAP_PROP_FRAME_HEIGHT)
            self.CAM_WIDTH = self.cap.get(cv2.CAP_PROP_FRAME_WIDTH)
            
            print(f"Camera initialized: {self.CAM_WIDTH}x{self.CAM_HEIGHT}")
            
        except Exception as e:
            print(f"Camera initialization error: {e}")
            return
        
        # Initialize face authentication
        if self.disable_face_auth:
            self.face_auth = None
            print("Face authentication disabled (-fd mode)")
        else:
            if FaceAuthenticator is None:
                raise RuntimeError("face auth module unavailable")
            try:
                self.face_auth = FaceAuthenticator()
                print("Face authentication initialized")
            except Exception as e:
                print(f"Face auth initialization error: {e}")
                raise RuntimeError("face auth initialization failed") from e
        
        self.gc_mode = 1
        self.auth_gate_authorized = False
        self.last_authorized_time = 0.0
        print("GestureController initialized successfully")

    def _show_preview_window(self):
        if self.preview_window_visible:
            return
        self._configure_preview_window(int(self.CAM_WIDTH or 640), int(self.CAM_HEIGHT or 480))
        self.preview_window_visible = True

    def _hide_preview_window(self):
        if not self.preview_window_visible:
            return
        try:
            cv2.destroyWindow(PREVIEW_WINDOW_NAME)
        except cv2.error:
            pass
        self.preview_window_visible = False

    def classify_hands(self, results):
        left, right = None, None
        if not results.multi_hand_landmarks or not results.multi_handedness:
            self.hr_major = None
            self.hr_minor = None
            return

        for hand_landmarks, handedness in zip(results.multi_hand_landmarks, results.multi_handedness):
            try:
                handedness_dict = MessageToDict(handedness)
                label = handedness_dict['classification'][0]['label']
            except Exception:
                continue

            if label == 'Right':
                right = hand_landmarks
            else:
                left = hand_landmarks
        
        if self.dom_hand:
            self.hr_major = right
            self.hr_minor = left
        else:
            self.hr_major = left
            self.hr_minor = right

    def start(self):
        if self.cap is None or not self.cap.isOpened():
            print("ERROR: Camera not available. Cannot start.")
            return

        # Start in auth mode with centered preview visible.
        self._show_preview_window()
        
        print("Gesture Controller Starting...")
        print("Press 'q' to quit, 'Enter' was causing issues")
        
        handmajor = HandRecog(HLabel.MAJOR)
        handminor = HandRecog(HLabel.MINOR)

        try:
            with mp_hands.Hands(max_num_hands=2, min_detection_confidence=HAND_DETECTION_CONFIDENCE, min_tracking_confidence=HAND_TRACKING_CONFIDENCE) as hands, ThreadPoolExecutor(max_workers=2, thread_name_prefix="gesture-frame") as frame_executor:
                print("MediaPipe Hands model loaded")

                def _run_face_auth(frame):
                    if self.disable_face_auth:
                        return False, None, True, None

                    try:
                        face_detected, faces = self.face_auth.authenticate_frame(frame)
                        raw_is_authorized = self.face_auth.get_auth_status()
                        return face_detected, faces, raw_is_authorized, None
                    except Exception as exc:
                        return None, None, None, exc

                def _run_hand_processing(frame):
                    try:
                        mirrored_for_processing = cv2.flip(frame, 1)
                        image_rgb = cv2.cvtColor(mirrored_for_processing, cv2.COLOR_BGR2RGB)
                        image_rgb.flags.writeable = False
                        results = hands.process(image_rgb)
                        image_rgb.flags.writeable = True
                        processed_image = cv2.cvtColor(image_rgb, cv2.COLOR_RGB2BGR)
                        return results, processed_image, None
                    except Exception as exc:
                        return None, None, exc
                
                frame_count = 0
                
                while self.gc_mode and self.cap.isOpened():
                    success, image = self.cap.read()
                    frame_count += 1
                    
                    if not success:
                        print(f"Frame {frame_count}: Failed to read frame")
                        continue

                    if VERBOSE_FRAME_LOGS and frame_count % 30 == 0:
                        print(f"Frame {frame_count}: Processing...")

                    face_future = frame_executor.submit(_run_face_auth, image.copy())
                    hand_future = frame_executor.submit(_run_hand_processing, image.copy())
                    

                    face_detected, faces, raw_is_authorized, face_error = face_future.result()

                    if face_error is not None:
                        self.auth_error_count += 1
                        print(f"Face auth error ({self.auth_error_count}/{self.max_auth_errors}): {face_error}")
                        if self.auth_error_count >= self.max_auth_errors:
                            raise RuntimeError("face auth runtime failed repeatedly") from face_error
                        # Do not instantly cut off gestures due to a transient auth exception.
                        is_authorized = (time.time() - self.last_authorized_time) <= AUTHORIZATION_HOLD_SECONDS
                        raw_is_authorized = is_authorized
                        self.prev_raw_authorized = raw_is_authorized
                    else:
                        now = time.time()

                        # Keep gestures enabled for a short grace window to prevent frame-level auth flicker.
                        if raw_is_authorized:
                            self.last_authorized_time = now
                            self.auth_gate_authorized = True
                        elif now - self.last_authorized_time <= AUTHORIZATION_HOLD_SECONDS:
                            self.auth_gate_authorized = True
                        else:
                            self.auth_gate_authorized = False

                        is_authorized = self.auth_gate_authorized
                        self.auth_error_count = 0

                        if raw_is_authorized and not self.prev_raw_authorized:
                            if self.desktop_indicator is not None and not self.minimize_after_auth_done:
                                self.desktop_indicator.minimize_all_windows()
                                self.minimize_after_auth_done = True
                        self.prev_raw_authorized = raw_is_authorized

                    results, processed_image, hand_error = hand_future.result()

                    if hand_error is not None:
                        print(f"Gesture processing error: {hand_error}")

                    display_image = processed_image if (is_authorized and processed_image is not None) else image
                    if face_detected and faces is not None:
                        x, y, w, h = max(faces, key=lambda r: r[2] * r[3])
                        color = (0, 255, 0) if raw_is_authorized else (0, 0, 255)
                        cv2.rectangle(display_image, (x, y), (x + w, y + h), color, 2)
                    image = display_image
                    
                    # Only process gestures if authorized
                    if is_authorized:
                        if self.preview_window_visible and not self.log_preview_enabled:
                            self._hide_preview_window()
                        if self.desktop_indicator is not None:
                            self.desktop_indicator.show()
                        try:
                            if results is not None and results.multi_hand_landmarks:
                                print(f"=== Frame {frame_count} MediaPipe Results ===")
                                print(f"Hands detected: {len(results.multi_hand_landmarks)}")
                                
                                for i, hand_landmarks in enumerate(results.multi_hand_landmarks):
                                    print(f"Hand {i+1}: {len(hand_landmarks.landmark)} landmarks")
                                    
                                    # Print first few landmark coordinates
                                    print(f"  Wrist (0): x={hand_landmarks.landmark[0].x:.3f}, y={hand_landmarks.landmark[0].y:.3f}")
                                    print(f"  Index Tip (8): x={hand_landmarks.landmark[8].x:.3f}, y={hand_landmarks.landmark[8].y:.3f}")
                                    print(f"  Thumb Tip (4): x={hand_landmarks.landmark[4].x:.3f}, y={hand_landmarks.landmark[4].y:.3f}")
                            elif VERBOSE_FRAME_LOGS:
                                print(f"Frame {frame_count}: No hands detected")

                            if results is not None and results.multi_hand_landmarks:
                                if VERBOSE_FRAME_LOGS:
                                    print(f"Frame {frame_count}: Hands detected")
                                self.classify_hands(results)
                                hand_states = []

                                if self.hr_major is not None:
                                    handmajor.update_hand_result(self.hr_major)
                                    handmajor.set_finger_state()
                                    major_gesture = handmajor.get_gesture()
                                    hand_states.append(("right", handmajor, major_gesture))
                                else:
                                    handmajor.update_hand_result(None)
                                    handmajor.window_switch_points.clear()
                                    handmajor._window_switch_end_session()

                                if self.hr_minor is not None:
                                    handminor.update_hand_result(self.hr_minor)
                                    handminor.set_finger_state()
                                    minor_gesture = handminor.get_gesture()
                                    hand_states.append(("left", handminor, minor_gesture))
                                else:
                                    handminor.update_hand_result(None)
                                    handminor.window_switch_points.clear()
                                    handminor._window_switch_end_session()

                                control_candidates = [
                                    (GestureController._gesture_priority(gesture), hand_key, tracker, gesture)
                                    for hand_key, tracker, gesture in hand_states
                                    if gesture != Gest.PALM
                                ]

                                if control_candidates:
                                    _, hand_key, tracker, gesture = max(
                                        control_candidates,
                                        key=lambda item: (item[0], item[1] == Controller.active_hand_key),
                                    )
                                    Controller.handle_controls(gesture, tracker.hand_result, is_authorized, hand_key=hand_key)
                                else:
                                    Controller.active_hand_key = None
                                    Controller.reset_control_state()

                                for hand_key, tracker, _gesture in hand_states:
                                    tracker.perform_presentation_action()
                                    if tracker.finger == WINDOW_SWITCH_FINGER_PATTERN:
                                        tracker.perform_window_switch_action()
                                    elif tracker.finger == Gest.LAST4:
                                        tracker.perform_desktop_switch_action()

                                
                                for hand_landmarks in results.multi_hand_landmarks:
                                    mp_drawing.draw_landmarks(image, hand_landmarks, mp_hands.HAND_CONNECTIONS)
                            else:
                                Controller.active_hand_key = None
                                Controller.reset_control_state()
                                Controller.reset_motion_state()
                                handmajor.window_switch_points.clear()
                                handmajor._window_switch_end_session()
                                handminor.window_switch_points.clear()
                                handminor._window_switch_end_session()
                                if VERBOSE_FRAME_LOGS:
                                    print(f"Frame {frame_count}: No hands detected")
                        
                        except Exception as e:
                            print(f"Gesture processing error: {e}")
                    else:
                        if self.desktop_indicator is not None:
                            self.desktop_indicator.hide()
                        handmajor.window_switch_points.clear()
                        handmajor._window_switch_end_session()
                        handminor.window_switch_points.clear()
                        handminor._window_switch_end_session()
                        if not self.preview_window_visible:
                            self._show_preview_window()
                        cv2.putText(image, "Authenticating...", 
                                   (50, image.shape[0]//2), cv2.FONT_HERSHEY_SIMPLEX, 1, (0, 0, 255), 2)

                    if is_authorized:
                        image = cv2.flip(image, 1)

                    if self.preview_window_visible:
                        if self.disable_face_auth:
                            image = self._draw_preview_ui(image, True, True, None, log_mode=self.log_preview_enabled)
                        else:
                            image = self._draw_preview_ui(image, is_authorized, raw_is_authorized, self.face_auth, log_mode=self.log_preview_enabled)
                        cv2.imshow(PREVIEW_WINDOW_NAME, image)
                    
                    # Use 'q' instead of Enter for quitting
                    if cv2.waitKey(1) & 0xFF == ord('q'):
                        print("Quit signal received")
                        break
                        
                    # # Limit frames for testing
                    # if frame_count >= 100:  # Remove this in production
                    #     print("Frame limit reached for testing")
                    #     break
                    
        except Exception as e:
            print(f"Main loop error: {e}")
            traceback.print_exc()
            if "face auth" in str(e).lower():
                raise
        
        finally:
            if self.desktop_indicator is not None:
                self.desktop_indicator.hide()
                self.desktop_indicator.close()
            handmajor.window_switch_points.clear()
            handmajor._window_switch_end_session()
            handminor.window_switch_points.clear()
            handminor._window_switch_end_session()
            if self.cap:
                self.cap.release()
            cv2.destroyAllWindows()
            print("Gesture Controller stopped.")

if __name__ == "__main__":
    print("=== MAIN EXECUTION START ===")
    parser = argparse.ArgumentParser(description="Gesture controller")
    parser.add_argument("-fd", "--face-disabled", action="store_true", help="Disable face auth gate")
    parser.add_argument("-log", "--log-preview", action="store_true", help="Keep the small hand-landmark preview visible in authorized mode")
    parser.add_argument("--dpi", type=int, default=50, choices=range(1, 101), metavar="1-100", help="Cursor sensitivity / DPI control (1-100)")
    args = parser.parse_args()

    try:
        gc = GestureController(disable_face_auth=args.face_disabled, dpi_level=args.dpi, log_preview=args.log_preview)
        if gc.cap and gc.cap.isOpened():
            gc.start()
        else:
            print("Failed to initialize GestureController")
            sys.exit(1)
    except RuntimeError as e:
        err = str(e).lower()
        if "face auth" in err:
            print(f"Exiting with face-auth failure code: {FACE_AUTH_FAILURE_EXIT_CODE}")
            sys.exit(FACE_AUTH_FAILURE_EXIT_CODE)
        print(f"Runtime error: {e}")
        traceback.print_exc()
        sys.exit(1)
    except Exception as e:
        print(f"Main execution error: {e}")
        traceback.print_exc()
        sys.exit(1)
    print("=== PROGRAM END ===")