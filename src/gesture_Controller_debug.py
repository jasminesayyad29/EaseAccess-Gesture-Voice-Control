import cv2
import mediapipe as mp
import pyautogui
import math
from enum import IntEnum
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from google.protobuf.json_format import MessageToDict
import screen_brightness_control as sbcontrol
import time
import traceback

from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
try:
    from face_auth_high_accuracy import FaceAuthenticatorHighAccuracy as FaceAuthenticator
except Exception:
    from face_auth import FaceAuthenticator

pyautogui.FAILSAFE = False
mp_drawing = mp.solutions.drawing_utils
mp_hands = mp.solutions.hands

# Debug/runtime controls
VERBOSE_FRAME_LOGS = False
AUTHORIZATION_HOLD_SECONDS = 3.5

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
    
    def set_finger_state(self):
        if self.hand_result == None:
            return

        points = [[8,5,0],[12,9,0],[16,13,0],[20,17,0]]
        self.finger = 0
        self.finger = self.finger | 0
        for idx,point in enumerate(points):
            dist = self.get_signed_dist(point[:2])
            dist2 = self.get_signed_dist(point[1:])
            try:
                ratio = round(dist/dist2,1)
            except:
                ratio = round(dist/0.01,1)
            self.finger = self.finger << 1
            if ratio > 0.5 :
                self.finger = self.finger | 1

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

        if self.frame_count > 4 :
            self.ori_gesture = current_gesture
        return self.ori_gesture
    
    def perform_presentation_action(self):
    
        if self.hand_result is None:
            return

        # Get landmark references for thumb and index
        thumb_tip = self.hand_result.landmark[4]
        thumb_ip = self.hand_result.landmark[3]
        index_tip = self.hand_result.landmark[8]
        wrist = self.hand_result.landmark[0]

        # 1️⃣ Next Slide (Thumb pointing right)
        if (thumb_tip.x > thumb_ip.x) and self.finger == Gest.THUMB:
            pyautogui.press('right')
            print("Next Slide")
            time.sleep(0.7)  # small delay to avoid multiple triggers

        # 2️⃣ Previous Slide (Thumb pointing left)
        elif (thumb_tip.x < thumb_ip.x) and self.finger == Gest.THUMB:
            pyautogui.press('left')
            print("Previous Slide")
            time.sleep(0.7)

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
    pinch_threshold = 0.3
    
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
    def get_position(hand_result):
        point = 9
        position = [hand_result.landmark[point].x ,hand_result.landmark[point].y]
        sx,sy = pyautogui.size()
        x_old,y_old = pyautogui.position()
        x = int(position[0]*sx)
        y = int(position[1]*sy)
        if Controller.prev_hand is None:
            Controller.prev_hand = x,y
        delta_x = x - Controller.prev_hand[0]
        delta_y = y - Controller.prev_hand[1]
        distsq = delta_x**2 + delta_y**2
        ratio = 1
        Controller.prev_hand = [x,y]
        if distsq <= 25:
            ratio = 0
        elif distsq <= 900:
            ratio = 0.07 * (distsq ** (1/2))
        else:
            ratio = 2.1
        x , y = x_old + delta_x*ratio , y_old + delta_y*ratio
        return (x,y)

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

    def handle_controls(gesture, hand_result, is_authorized):  
        """Only execute gestures if authorized"""
        if not is_authorized:
            Controller.prev_hand = None
            return
        
        x,y = None,None
        if gesture != Gest.PALM :
            x,y = Controller.get_position(hand_result)
        
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
            pyautogui.moveTo(x, y, duration = 0.1)
        elif gesture == Gest.FIST:
            if not Controller.grabflag : 
                Controller.grabflag = True
                pyautogui.mouseDown(button = "left")
            pyautogui.moveTo(x, y, duration = 0.1)
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

    def __init__(self):
        print("Initializing GestureController...")
        
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
            self.cap.set(cv2.CAP_PROP_FRAME_WIDTH, 640)
            self.cap.set(cv2.CAP_PROP_FRAME_HEIGHT, 480)
            
            self.CAM_HEIGHT = self.cap.get(cv2.CAP_PROP_FRAME_HEIGHT)
            self.CAM_WIDTH = self.cap.get(cv2.CAP_PROP_FRAME_WIDTH)
            
            print(f"Camera initialized: {self.CAM_WIDTH}x{self.CAM_HEIGHT}")
            
        except Exception as e:
            print(f"Camera initialization error: {e}")
            return
        
        # Initialize face authentication
        try:
            self.face_auth = FaceAuthenticator()
            print("Face authentication initialized")
        except Exception as e:
            print(f"Face auth initialization error: {e}")
            return
        
        self.gc_mode = 1
        self.auth_gate_authorized = False
        self.last_authorized_time = 0.0
        print("GestureController initialized successfully")

    def classify_hands(self, results):
        left, right = None, None
        try:
            handedness_dict = MessageToDict(results.multi_handedness[0])
            if handedness_dict['classification'][0]['label'] == 'Right':
                right = results.multi_hand_landmarks[0]
            else:
                left = results.multi_hand_landmarks[0]
        except:
            pass

        try:
            handedness_dict = MessageToDict(results.multi_handedness[1])
            if handedness_dict['classification'][0]['label'] == 'Right':
                right = results.multi_hand_landmarks[1]
            else:
                left = results.multi_hand_landmarks[1]
        except:
            pass
        
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
        
        print("Gesture Controller Starting...")
        print("Press 'q' to quit, 'Enter' was causing issues")
        
        handmajor = HandRecog(HLabel.MAJOR)
        handminor = HandRecog(HLabel.MINOR)

        try:
            with mp_hands.Hands(max_num_hands=2, min_detection_confidence=0.5, min_tracking_confidence=0.5) as hands:
                print("MediaPipe Hands model loaded")
                
                frame_count = 0
                
                while self.gc_mode and self.cap.isOpened():
                    success, image = self.cap.read()
                    frame_count += 1
                    
                    if not success:
                        print(f"Frame {frame_count}: Failed to read frame")
                        continue

                    if VERBOSE_FRAME_LOGS and frame_count % 30 == 0:
                        print(f"Frame {frame_count}: Processing...")
                    
                    # Face authentication
                    try:
                        face_detected, faces = self.face_auth.authenticate_frame(image)
                        raw_is_authorized = self.face_auth.get_auth_status()
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
                        
                        # Draw face bounding box and status
                        if face_detected:
                            x,y,w,h = max(faces, key=lambda r: r[2]*r[3])
                            color = (0, 255, 0) if raw_is_authorized else (0, 0, 255)
                            cv2.rectangle(image, (x,y), (x+w, y+h), color, 2)
                        
                        # Display authorization status
                        auth_status = "AUTHORIZED" if is_authorized else "UNAUTHORIZED"
                        if is_authorized and not raw_is_authorized:
                            auth_status = "AUTHORIZED (HOLD)"
                        status_color = (0, 255, 0) if is_authorized else (0, 0, 255)
                        cv2.putText(image, f"Status: {auth_status}", (10, 30), 
                                   cv2.FONT_HERSHEY_SIMPLEX, 0.7, status_color, 2)

                        if time.time() - self.face_auth.get_last_match_time() < 10:
                            cv2.putText(image, self.face_auth.get_last_match_result(), (10, 60),
                                       cv2.FONT_HERSHEY_SIMPLEX, 0.5, (255, 255, 255), 1)
                        
                    except Exception as e:
                        print(f"Face auth error: {e}")
                        # Do not instantly cut off gestures due to a transient auth exception.
                        is_authorized = (time.time() - self.last_authorized_time) <= AUTHORIZATION_HOLD_SECONDS
                    
                    # Only process gestures if authorized
                    if is_authorized:
                        try:
                            image_rgb = cv2.cvtColor(cv2.flip(image, 1), cv2.COLOR_BGR2RGB)
                            image_rgb.flags.writeable = False
                            results = hands.process(image_rgb)
                            image_rgb.flags.writeable = True
                            image = cv2.cvtColor(image_rgb, cv2.COLOR_RGB2BGR)

                            if VERBOSE_FRAME_LOGS and results.multi_hand_landmarks:
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

                            if results.multi_hand_landmarks:
                                if VERBOSE_FRAME_LOGS:
                                    print(f"Frame {frame_count}: Hands detected")
                                self.classify_hands(results)
                                handmajor.update_hand_result(self.hr_major)
                                handminor.update_hand_result(self.hr_minor)

                                handmajor.set_finger_state()
                                handminor.set_finger_state()
                                # Update gestures
                                gest_name_minor = handminor.get_gesture()
                                gest_name_major = handmajor.get_gesture()

                                # Handle standard mouse/volume/brightness gestures
                                if gest_name_minor == Gest.PINCH_MINOR:
                                    Controller.handle_controls(gest_name_minor, handminor.hand_result, is_authorized)
                                else:
                                    Controller.handle_controls(gest_name_major, handmajor.hand_result, is_authorized)

                                # ---------- NEW: Presentation Controls ----------
                                # Call perform_presentation_action() for both hands
                                handmajor.perform_presentation_action()
                                handminor.perform_presentation_action()

                                
                                for hand_landmarks in results.multi_hand_landmarks:
                                    mp_drawing.draw_landmarks(image, hand_landmarks, mp_hands.HAND_CONNECTIONS)
                            else:
                                Controller.prev_hand = None
                                if VERBOSE_FRAME_LOGS:
                                    print(f"Frame {frame_count}: No hands detected")
                        
                        except Exception as e:
                            print(f"Gesture processing error: {e}")
                    else:
                        cv2.putText(image, "UNAUTHORIZED - Gestures disabled", 
                                   (50, image.shape[0]//2), cv2.FONT_HERSHEY_SIMPLEX, 1, (0, 0, 255), 2)
                    
                    cv2.imshow('Gesture Controller - Press Q to quit', image)
                    
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
        
        finally:
            if self.cap:
                self.cap.release()
            cv2.destroyAllWindows()
            print("Gesture Controller stopped.")

if __name__ == "__main__":
    print("=== MAIN EXECUTION START ===")
    try:
        gc = GestureController()
        if gc.cap and gc.cap.isOpened():
            gc.start()
        else:
            print("Failed to initialize GestureController")
    except Exception as e:
        print(f"Main execution error: {e}")
        traceback.print_exc()
    print("=== PROGRAM END ===")