# Gesture Controller Debug Documentation

## Scope
This document explains how src/gesture_Controller_debug.py works, step by step, including:
- Libraries used and why they are needed
- Startup and runtime flow
- Face-authorization gating logic
- Gesture recognition pipeline
- Gesture to action mapping
- Current camera preview UI behavior

## Main File
- src/gesture_Controller_debug.py

## Related Files
- src/face_auth_high_accuracy.py
- src/face_auth.py
- requirements.txt

## Libraries Used

### Core computer vision and hand tracking
- cv2 (OpenCV): camera capture, drawing overlays, window handling, and image conversion
- mediapipe: hand landmark detection and tracking
- google.protobuf.json_format.MessageToDict: converts MediaPipe handedness metadata to dictionary form

### System control
- pyautogui: mouse movement, clicks, drag, keyboard shortcuts, scrolling, and presentation key presses
- screen_brightness_control: changes system brightness via pinch gesture
- pycaw + comtypes + ctypes: changes system volume via pinch gesture

### Utility
- math: distance and ratio calculations for gesture inference
- time: gesture debouncing, auth hold timing, and UI recency checks
- traceback: debug stack traces when runtime errors occur
- enum.IntEnum: stable numeric gesture labels and hand role labels

### Face authentication module
- FaceAuthenticator is imported from:
  - src/face_auth_high_accuracy.py (preferred)
  - src/face_auth.py (fallback)

## Runtime Constants (Important)
- VERBOSE_FRAME_LOGS = False
  - Enables additional frame-by-frame debug logs when set True
- AUTHORIZATION_HOLD_SECONDS = 3.5
  - Keeps gestures enabled briefly after a temporary auth drop
- PREVIEW_WINDOW_NAME = EaseAccess Preview (Press Q to quit)
- PIP_WIDTH = 420
- PIP_MARGIN = 24

## High-Level Architecture
The file is organized into three main classes:

1. HandRecog
- Converts hand landmarks into stable gesture labels
- Handles presentation slide gestures from thumb orientation

2. Controller
- Executes system actions based on recognized gestures
- Handles cursor movement, click gestures, scrolling, brightness, and volume

3. GestureController
- Owns camera loop and orchestrates face auth + hand tracking
- Enforces authorization gating before action execution
- Draws compact picture-in-picture style preview UI

## Step-by-Step Execution Flow

### 1) Program entry
- In main block, GestureController() is created
- If camera opens successfully, start() runs

### 2) Initialization in GestureController
- Attempts to open webcam with index 0 (CAP_DSHOW)
- Falls back to index 1 if needed
- Sets capture resolution to 640x480
- Initializes FaceAuthenticator
- Sets internal auth-gate state variables

### 3) Preview window setup
At start of start():
- Creates resizable top-most OpenCV window
- Computes PiP height using camera aspect ratio and PIP_WIDTH
- Moves window to top-right corner using screen size and PIP_MARGIN

### 4) Main frame loop
For each frame:
- Reads frame from camera
- Runs face authentication for current frame
- Updates authorization gate state:
  - Immediate authorized when face auth succeeds
  - Keeps authorized for AUTHORIZATION_HOLD_SECONDS grace period

### 5) Face status rendering
- If a face is detected, draws a face rectangle
- Rectangle color:
  - Green when raw face auth is authorized
  - Red when raw face auth is unauthorized

### 6) Hand processing (only when authorized)
- Frame is mirrored for MediaPipe processing (preserves existing gesture behavior)
- Hands are detected and classified into major/minor
- Finger state is computed for both hands
- Stable gesture is inferred for each hand
- Gesture actions are executed

### 7) Gesture action dispatch
- If minor hand gesture is PINCH_MINOR, minor-hand pinch action path is used
- Otherwise major hand gesture path is used
- Presentation actions are checked on both hands each frame

### 8) Unauthorized behavior
- Gesture actions are not executed
- Center warning text is shown: UNAUTHORIZED - Gestures disabled

### 9) Final UI render and exit
- Authorized preview is flipped back so the displayed view is not mirrored
- Professional overlay (status bar and footer) is drawn
- Window is shown as PiP preview
- Press q to exit

## Hand Classification and Gesture Inference

### Hand assignment
- MediaPipe handedness is read from results.multi_handedness
- If dom_hand is True:
  - Right hand => major
  - Left hand => minor

### Finger state encoding
- Uses relative distances between fingertip and lower joints
- Encodes open/closed state into integer bit pattern
- Additional logic detects:
  - V_GEST
  - TWO_FINGER_CLOSED
  - PINCH_MAJOR / PINCH_MINOR

### Temporal smoothing
- A gesture must persist several frames before becoming active
- This reduces accidental single-frame misclassifications

## Authorization Gating Logic
Two auth states are used:

1. raw_is_authorized
- Direct result from FaceAuthenticator

2. is_authorized (gated)
- True immediately when raw auth is true
- Remains true for AUTHORIZATION_HOLD_SECONDS after raw auth drops
- Prevents rapid enable/disable flicker during brief detection instability

Effect:
- All control actions only run when is_authorized is true

## Gesture to Action Mapping

### Emoji Legend
- ✊ = FIST
- ✌️ = V_GEST
- ☝️ = INDEX
- 🖕 = MID
- 🤏 = PINCH
- 👍 = THUMB orientation gesture
- 🖱️ = Mouse action
- 🧭 = Cursor movement
- 🔊 = Volume change
- 🔆 = Brightness change
- ↕️ = Vertical scroll
- ↔️ = Horizontal scroll
- ⏭️ = Next slide
- ⏮️ = Previous slide

### Core control gestures (Controller.handle_controls)
## Gesture Controls

| Gesture | Condition | Action |
|--------|----------|--------|
| ✌️ V_GEST | Recognized and authorized | 🧭 Enable pointer-control mode and move cursor |
| ✊ FIST | Recognized and authorized | 🖱️ Hold left mouse button (drag) while moving cursor |
| 🖕 MID | Only if pointer-control mode is active | 🖱️ Left click |
| ☝️ INDEX | Only if pointer-control mode is active | 🖱️ Right click |
| ✌️🤏 TWO_FINGER_CLOSED | Only if pointer-control mode is active | 🖱️ Double click |
| 🤏 PINCH_MINOR | Recognized on minor hand | ↕️ / ↔️ Scroll control (vertical or horizontal based on pinch direction) |
| 🤏 PINCH_MAJOR | Recognized on major hand | 🔆 / 🔊 Control brightness or volume based on pinch direction |

Notes:
- Pointer-control mode is latched by V_GEST using Controller.flag
- Drag is released when gesture is no longer FIST

### Presentation gestures (HandRecog.perform_presentation_action)

| Gesture logic | Condition | Action |
|---|---|---|
| 👍➡️ Thumb points right | self.finger == THUMB and thumb_tip.x > thumb_ip.x | ⏭️ Next slide (Right Arrow key) |
| 👍⬅️ Thumb points left | self.finger == THUMB and thumb_tip.x < thumb_ip.x | ⏮️ Previous slide (Left Arrow key) |

Debounce:
- Each slide action sleeps for about 0.7 seconds to avoid repeated triggers

## Cursor Motion Behavior
- Cursor is controlled using landmark point 9
- Motion delta is smoothed using distance-based gain
- Small movement is damped, medium movement is scaled, large movement is capped

## Pinch Control Behavior
Two-axis interpretation:
- Predominantly vertical pinch movement selects vertical function
- Predominantly horizontal pinch movement selects horizontal function

Function mapping:
- Minor hand pinch:
  - Horizontal: ↔️ horizontal scroll
  - Vertical: ↕️ vertical scroll
- Major hand pinch:
  - Horizontal: 🔆 brightness change
  - Vertical: 🔊 volume change

## Current Preview UI Behavior
- Compact PiP camera window at top-right corner
- Top-most window so it stays visible
- Professional overlay with:
  - Status indicator dot
  - Status text (AUTHORIZED / UNAUTHORIZED / AUTHORIZED (HOLD))
  - Footer hint text
  - Recent face-auth detail line
- Display shown non-mirrored in authorized state

## Dependencies Checklist
From requirements.txt, the most important packages for this file are:
- opencv-python
- opencv-contrib-python
- mediapipe
- pyautogui
- screen-brightness-control
- pycaw
- comtypes
- protobuf
- numpy

## How To Run
From project root:

python src/gesture_Controller_debug.py

## Troubleshooting
- Camera open fails:
  - Check webcam permissions in Windows Privacy settings
  - Verify no other app exclusively holds camera
- Face auth always unauthorized:
  - Ensure known face images are available in src/known_faces or configured folder
  - Use good lighting and frontal face orientation
- Volume control not working:
  - Confirm pycaw/comtypes are installed correctly on Windows
- Brightness control not working:
  - Some external monitors may not support software brightness control

## Summary
src/gesture_Controller_debug.py combines:
- Face-auth gated safety
- MediaPipe hand landmark tracking
- Gesture-driven desktop and presentation controls
- A small professional PiP preview window

The logic ensures gestures are only active for an authorized user, with temporal smoothing and auth hold behavior to keep control stable in real-world use.
