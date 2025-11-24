import pyttsx3
import speech_recognition as sr
from datetime import date
import time
import webbrowser
import datetime
from pynput.keyboard import Key, Controller
import pyautogui
import sys
import os
from os import listdir
from os.path import isfile, join, exists
import wikipedia
import app
from threading import Thread
import re
import json
from difflib import SequenceMatcher
import numpy as np
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.naive_bayes import MultinomialNB
from sklearn.pipeline import make_pipeline
import joblib
import warnings
import subprocess
import pygetwindow as gw
warnings.filterwarnings('ignore')
import glob
import subprocess
import pytesseract
from PIL import ImageGrab, Image
import pyautogui
import cv2
import numpy as np
import winreg
from difflib import get_close_matches
import win32gui
import win32con
import win32api
import pythoncom


# Keep track of where we came from so we can return if needed
last_context = None   # possible values: None, 'presentation', 'browser', ...
AUTO_RETURN_AFTER_SEARCH = False # Set True if you want automatic return after search

pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# -------------Object Initialization---------------
today = date.today()
r = sr.Recognizer()
keyboard = Controller()

# Initialize text-to-speech engine with error handling
try:
    engine = pyttsx3.init('sapi5')
except:
    try:
        engine = pyttsx3.init()
    except Exception as e:
        print(f"TTS initialization failed: {e}")
        engine = None

if engine:
    try:
        voices = engine.getProperty('voices')
        if voices:
            engine.setProperty('voice', voices[0].id)
    except:
        pass

# ----------------Variables------------------------
file_exp_status = False
files = []
path = ''
is_awake = True  # Bot status
current_directory = 'C://'  # Track current directory
import winreg
import glob



def get_all_installed_apps():
    """
    Returns a dict mapping normalized app names -> launch target.
    Launch target may be:
      - .lnk shortcut path
      - .exe absolute path
      - install folder (we will search .exe inside)
      - built-in executable name (e.g., 'notepad.exe')
    """
    apps = {}

    # A. Start Menu Shortcuts (.lnk)
    start_menu_paths = [
        os.path.expandvars(r"%APPDATA%\Microsoft\Windows\Start Menu\Programs"),
        os.path.expandvars(r"%PROGRAMDATA%\Microsoft\Windows\Start Menu\Programs"),
    ]
    for sm_path in start_menu_paths:
        try:
            for shortcut in glob.glob(os.path.join(sm_path, "**", "*.lnk"), recursive=True):
                name = os.path.splitext(os.path.basename(shortcut))[0].lower()
                apps[name] = shortcut
        except Exception:
            pass

    # B. Registry Installed Apps
    reg_paths = [
        r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall",
    ]
    for reg_path in reg_paths:
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, reg_path) as root:
                for i in range(winreg.QueryInfoKey(root)[0]):
                    try:
                        subkey_name = winreg.EnumKey(root, i)
                        with winreg.OpenKey(root, subkey_name) as subkey:
                            try:
                                display_name = winreg.QueryValueEx(subkey, "DisplayName")[0]
                            except Exception:
                                display_name = None
                            try:
                                install_loc = winreg.QueryValueEx(subkey, "InstallLocation")[0]
                            except Exception:
                                install_loc = None

                            if display_name:
                                key = display_name.lower()
                                # prefer install location if available
                                if install_loc:
                                    apps[key] = install_loc
                                else:
                                    # keep what we have (maybe later Start Menu will override)
                                    apps.setdefault(key, "") 
                    except Exception:
                        pass
        except Exception:
            pass

    # C. WindowsApps (UWP) - find exe files in package folders
    uwp_dir = r"C:\Program Files\WindowsApps"
    if os.path.isdir(uwp_dir):
        try:
            for folder in os.listdir(uwp_dir):
                folder_path = os.path.join(uwp_dir, folder)
                if os.path.isdir(folder_path):
                    try:
                        for f in os.listdir(folder_path):
                            if f.lower().endswith(".exe"):
                                name = os.path.splitext(f)[0].lower()
                                apps[name] = os.path.join(folder_path, f)
                    except Exception:
                        pass
        except Exception:
            pass

    # D. System32 tools (safe built-in exe names)
    system_tools = ["notepad.exe", "calc.exe", "explorer.exe", "cmd.exe", "powershell.exe", "mspaint.exe"]
    sys32 = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "System32")
    for exe in system_tools:
        exe_path = os.path.join(sys32, exe)
        if os.path.exists(exe_path):
            apps[os.path.splitext(exe)[0].lower()] = exe_path
        else:
            # fallback to name only (os.startfile will handle these)
            apps.setdefault(os.path.splitext(exe)[0].lower(), exe)

    # E. Portable EXEs in common user folders
    user_home = os.path.expanduser("~")
    portable_search = [
        os.path.join(user_home, "Downloads"),
        os.path.join(user_home, "Desktop"),
        os.path.join(user_home, "Documents"),
    ]
    for d in portable_search:
        if os.path.isdir(d):
            try:
                for exe in glob.glob(os.path.join(d, "**", "*.exe"), recursive=True):
                    name = os.path.splitext(os.path.basename(exe))[0].lower()
                    apps[name] = exe
            except Exception:
                pass

    # Normalize: remove empty install locations for registry-only entries if Start Menu found later
    # (Start Menu/WindowsApps entries added earlier will override)
    return apps


def match_app(user_input, app_dict):
    """
    Return the matched key from app_dict (or None).
    Matching order: exact -> startswith -> fuzzy
    """
    if not user_input:
        return None

    name = user_input.lower().strip()
    keys = list(app_dict.keys())

    # Exact
    if name in app_dict:
        return name

    # startswith
    for k in keys:
        if k.startswith(name):
            return k

    # fuzzy (reasonable cutoff)
    match = get_close_matches(name, keys, n=1, cutoff=0.65)
    if match:
        return match[0]

    return None

# ---------- UNIVERSAL 3-LAYER TASKBAR + ICON OPENER ----------
def get_taskbar_buttons():
    """
    Universal collector of candidate taskbar icon HWNDs / centers.
    Returns list of (hwnd, center_x, center_y).
    This is a best-effort collector: some Windows 11 systems return wrapper HWNDs;
    we still return center points which we will hover and OCR or click.
    """
    pythoncom.CoInitialize()
    result = []

    try:
        taskbar = win32gui.FindWindow("Shell_TrayWnd", None)
        if not taskbar:
            return result

        # Candidate container class names (try all)
        possible = [
            "MSTaskListWClass", "MSTaskSwWClass", "ToolbarWindow32",
            "ReBarWindow32", "Windows.UI.Core.CoreWindow",
            "Windows.UI.Composition.DesktopWindowContentBridge"
        ]

        # Enumerate children of Shell_TrayWnd
        children = []
        def enum_child(hwnd, _):
            children.append(hwnd)
        win32gui.EnumChildWindows(taskbar, enum_child, None)

        # For each descendant container, enumerate its children and collect rect centers
        for cont in children:
            try:
                cls = win32gui.GetClassName(cont)
            except:
                cls = ""
            if cls not in possible and not cls.lower().startswith("tool"):
                # still attempt to pull children from everything (broad)
                pass

            # Enumerate grandchildren (the actual icon buttons often are grandchildren)
            def enum_grandchild(hwnd, _):
                try:
                    rect = win32gui.GetWindowRect(hwnd)
                    # Filter tiny/invalid rects
                    w = rect[2] - rect[0]
                    h = rect[3] - rect[1]
                    if w <= 8 or h <= 8:
                        return
                    cx = (rect[0] + rect[2]) // 2
                    cy = (rect[1] + rect[3]) // 2
                    result.append((hwnd, cx, cy))
                except Exception:
                    pass

            win32gui.EnumChildWindows(cont, enum_grandchild, None)

        # If still empty, fall back to sampling along the taskbar rectangle
        if not result:
            try:
                rect = win32gui.GetWindowRect(taskbar)
                left, top, right, bottom = rect
                # sample horizontally across the taskbar (20 points)
                steps = 20
                for i in range(steps):
                    cx = left + int((i + 0.5) * (right - left) / steps)
                    cy = top + (bottom - top) // 2
                    result.append((None, cx, cy))
            except:
                pass

    except Exception:
        pass

    # Deduplicate by center coordinates (keep unique centers)
    seen = set()
    filtered = []
    for item in result:
        _, cx, cy = item
        key = (cx, cy)
        if key not in seen:
            seen.add(key)
            filtered.append(item)

    return filtered


def _hover_and_ocr_matches(app_name, centers, ocr_timeout=1.2):
    """
    Hover over each center point in 'centers' (list of (hwnd, x, y)).
    After hovering, capture a quick screenshot and run pytesseract to see tooltip text.
    Returns (x,y) of matched center or None.
    """
    app_name = app_name.lower().strip()
    for item in centers:
        _, cx, cy = item
        try:
            pyautogui.moveTo(cx, cy, duration=0.12)
            # small pause for tooltip to appear
            time.sleep(0.35)
            # capture a region around mouse for speed — tooltips usually near cursor
            screen = ImageGrab.grab()
            text = pytesseract.image_to_string(screen).lower()
            if app_name in text:
                return (cx, cy)
            # optional: small scroll to reveal hidden tooltips (skip by default)
        except Exception as e:
            # continue to next candidate
            print("hover_ocr error:", e)
            continue
    return None


def _match_icon_templates_on_screen(app_name, templates, threshold=0.65):
    """
    Try multi-template multi-scale matching on the full screen.
    templates: list of file paths for the same icon (different variants).
    Returns center (x,y) on screen if match found, else None.
    """
    try:
        screen = np.array(ImageGrab.grab())
        screen_gray = cv2.cvtColor(screen, cv2.COLOR_BGR2GRAY)
    except Exception as e:
        print("screenshot error:", e)
        return None

    for template_path in templates:
        try:
            if not os.path.exists(template_path):
                continue
            tpl = cv2.imread(template_path, 0)
            if tpl is None:
                continue
            (tH, tW) = tpl.shape[:2]
            # edge fallback for robustness
            tpl_edges = cv2.Canny(tpl, 50, 150)

            best_val = 0
            best_loc = None
            best_size = (tW, tH)

            # scale variations
            for scale in np.linspace(0.5, 1.6, 18):
                rw = int(tW * scale)
                rh = int(tH * scale)
                if rw < 8 or rh < 8 or rw > screen.shape[1] or rh > screen.shape[0]:
                    continue
                resized = cv2.resize(tpl, (rw, rh))
                try:
                    res = cv2.matchTemplate(screen_gray, resized, cv2.TM_CCOEFF_NORMED)
                    min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(res)
                    if max_val > best_val:
                        best_val = max_val
                        best_loc = max_loc
                        best_size = (rw, rh)
                except Exception:
                    pass

                # edge-match
                try:
                    resized_edges = cv2.resize(tpl_edges, (rw, rh))
                    screen_edges = cv2.Canny(screen_gray, 50, 150)
                    res_e = cv2.matchTemplate(screen_edges, resized_edges, cv2.TM_CCOEFF_NORMED)
                    _, max_val_e, _, max_loc_e = cv2.minMaxLoc(res_e)
                    if max_val_e > best_val:
                        best_val = max_val_e
                        best_loc = max_loc_e
                        best_size = (rw, rh)
                except Exception:
                    pass

            if best_loc and best_val >= threshold:
                x = best_loc[0] + best_size[0] // 2
                y = best_loc[1] + best_size[1] // 2
                return (x, y)
        except Exception as e:
            print("template error:", e)
            continue

    return None


def open_taskbar_app(app_name):
    """
    Universal 3-layer launcher:
      1) Direct system launch via open_system_app()
      2) Hover + OCR over taskbar icons
      3) Template matching on full screen (multi-scale)
    Returns True if opened, False otherwise.
    """
    name = app_name.lower().strip()

    # --------- Layer 1: Direct launch (fast & reliable if installed) ----------
    try:
        # open_system_app returns a human message; consider success if it indicates opening
        sys_result = open_system_app(name)
        if isinstance(sys_result, str):
            low = sys_result.lower()
            if "opening" in low and "could not" not in low:
                reply(f"Opened {name} directly.")
                return True
    except Exception:
        pass

    # --------- Layer 2: Hover + OCR (universal for pinned taskbar icons) -----
    try:
        centers = get_taskbar_buttons()
        if centers:
            # centers is list of (hwnd, x, y) or similar — ensure format
            # if returned (hwnd,x,y), keep as-is; if (hwnd, x, y) where hwnd may be None
            match = _hover_and_ocr_matches(name, centers)
            if match:
                mx, my = match
                pyautogui.click(mx, my)
                reply(f"Opening {name} from taskbar (tooltip match).")
                return True
    except Exception as e:
        print("hover layer error:", e)

    # --------- Layer 3: Template matching (icon vision) ---------------------
    # templates: map of appname -> list of template file paths
    TEMPLATES = {
        # Local files you uploaded (adjust or add more file paths if you like)
        "whatsapp": [r"/mnt/data/whatsapp icon.png", r"/mnt/data/whatsapp.png"],
        # Add other apps here, e.g.
        # "chrome": [r"C:\path\to\chrome_template.png"],
        # "edge": [r"C:\path\to\edge_template.png"],
    }

    try:
        templates = TEMPLATES.get(name, [])
        if templates:
            pos = _match_icon_templates_on_screen(name, templates, threshold=0.66)
            if pos:
                px, py = pos
                pyautogui.moveTo(px, py, duration=0.18)
                pyautogui.click()
                reply(f"Opening {name} from screen match.")
                return True
    except Exception as e:
        print("template layer error:", e)

    # Nothing found
    reply(f"I could not find {name} on your taskbar or system.")
    return False

def open_system_app(app_name):
    """
    Universal app launcher:
    - auto-discovers installed apps and UWP exes
    - launches .lnk or .exe or finds exe in folder
    """
    if not app_name:
        return "No application name provided."

    try:
        apps = get_all_installed_apps()
        match = match_app(app_name, apps)
        if not match:
            # try some common aliases: 'code' -> 'visual studio code', 'vscode' etc
            # quick alias table (extend if you want)
            aliases = {
                "vscode": "visual studio code",
                "visual studio": "visual studio",
                "vs code": "visual studio code",
                "whatsapp": "whatsapp",
                "one note": "onenote",
                "onenote": "onenote",
                "snipping tool": "snippingtool",
                "snip": "snippingtool",
                "file explorer": "explorer",
                "explorer": "explorer"
            }
            for a, b in aliases.items():
                if a in app_name and b in apps:
                    match = b
                    break

        if not match:
            return f"I could not find an application named '{app_name}'."

        target = apps.get(match)

        # If it's a .lnk shortcut, launch it
        if target and target.lower().endswith(".lnk"):
            os.startfile(target)
            return f"Opening {match}"

        # If it's an exe path, launch it
        if target and target.lower().endswith(".exe") and os.path.exists(target):
            os.startfile(target)
            return f"Opening {match}"

        # If target is a folder (install location), search for exe inside
        if target and os.path.isdir(target):
            try:
                for f in os.listdir(target):
                    if f.lower().endswith(".exe"):
                        exe_path = os.path.join(target, f)
                        os.startfile(exe_path)
                        return f"Opening {match}"
            except Exception:
                pass

        # If nothing else, attempt a generic start by name (e.g., builtin exe names or PATH commands)
        try:
            os.startfile(match)
            return f"Opening {match}"
        except Exception:
            pass

        # Last-resort: try fuzzy-executing the match string via subprocess
        try:
            subprocess.Popen(match)
            return f"Opening {match}"
        except Exception as e:
            return f"I found {match}, but could not start it: {e}"

    except Exception as e:
        return f"Error launching {app_name}: {e}"


def minimize_browser_windows():
    """
    Minimize open browser windows like Chrome, Edge, or Firefox.
    Useful before returning to PowerPoint so slides become visible.
    """
    import pygetwindow as gw

    # Common browser window titles
    browsers = ["Google Chrome", "Chrome", "Microsoft Edge", "Edge", "Firefox", "Mozilla Firefox"]

    try:
        all_titles = gw.getAllTitles()
        for title in all_titles:
            if any(b.lower() in title.lower() for b in browsers):
                try:
                    win = gw.getWindowsWithTitle(title)[0]
                    win.minimize()
                    print(f"🪟 Minimized browser: {title}")
                except Exception as e:
                    print(f"⚠️ Could not minimize {title}: {e}")
    except Exception as e:
        print(f"⚠️ Error minimizing browser windows: {e}")

# place these module-level cache vars near the top of your module/file
_ocr_image_cache = {"image_hash": None, "data": None}

_ocr_cache = {"img_hash": None, "data": None}

def find_text_boxes(target, fuzzy_threshold=0.65):
    """
    Accurate OCR text finder for both:
    - SINGLE WORDS (primary, unchanged behavior)
    - MULTI-WORD PHRASES (added, but isolated so single words are never affected)

    No architecture changes.
    Just accuracy fixes.
    """

    from difflib import SequenceMatcher
    import re

    # ----------------------------
    # Normalize spoken numbers
    # ----------------------------
    number_words = {
        "zero":"0","one":"1","two":"2","three":"3","four":"4","five":"5",
        "six":"6","seven":"7","eight":"8","nine":"9","ten":"10"
    }
    parts = target.lower().strip().split()
    target = " ".join([number_words.get(w, w) for w in parts]).strip()
    target_words = target.split()

    single_word_mode = (len(target_words) == 1)
    multi_word_mode  = (len(target_words) >  1)

    # Raise fuzzy threshold for short UI labels
    if single_word_mode and len(target) <= 5:
        fuzzy_threshold = max(fuzzy_threshold, 0.85)

    if not target:
        return []

    # ----------------------------
    # 1. OCR CAPTURE
    # ----------------------------
    img = ImageGrab.grab()
    data = pytesseract.image_to_data(img, output_type=pytesseract.Output.DICT)

    # ----------------------------
    # Helper: Fix OCR errors
    # ----------------------------
    def fix_ocr(w):
        return (
            w.replace("1", "i")
             .replace("l", "i")
             .replace("|", "i")
             .replace("0", "o")
             .replace("5", "s")
             .replace("$", "s")
        )

    # ==========================================================
    # MODE 1: HIGH-ACCURACY SINGLE WORD MATCHING
    # ==========================================================
    word_hits = []
    if single_word_mode:

        for i, raw_word in enumerate(data["text"]):
            if not raw_word.strip():
                continue

            word = raw_word.lower().strip()
            fixed = fix_ocr(word)

            x = data["left"][i]
            y = data["top"][i]
            w = data["width"][i]
            h = data["height"][i]

            # direct or substring
            if target == word or target in word or word in target:
                word_hits.append((x, y, w, h))
                continue

            # fuzzy raw
            if SequenceMatcher(None, target, word).ratio() >= fuzzy_threshold:
                word_hits.append((x, y, w, h))
                continue

            # fuzzy corrected
            if SequenceMatcher(None, target, fixed).ratio() >= fuzzy_threshold:
                word_hits.append((x, y, w, h))
                continue

        if word_hits:
            return word_hits

    # ==========================================================
    # 2. ORIGINAL LINE GROUPING (UNCHANGED STRUCTURE)
    # ==========================================================
    lines = {}
    for i, raw_word in enumerate(data["text"]):
        if not raw_word.strip():
            continue

        block = data["block_num"][i]
        ln = data["line_num"][i]
        key = (block, ln)

        x = data["left"][i]
        y = data["top"][i]
        w = data["width"][i]
        h = data["height"][i]

        if key not in lines:
            lines[key] = {"words": [], "boxes": []}

        lines[key]["words"].append(raw_word.lower())
        lines[key]["boxes"].append((x, y, w, h))

    matches = []

    # ==========================================================
    # MODE 2: MULTI-WORD PHRASE MATCHING
    # ==========================================================
    for key, info in lines.items():

        # cluster split (same as before)
        entries = list(zip(info["words"], info["boxes"]))
        entries.sort(key=lambda e: e[1][0])

        clusters = []
        cluster = [entries[0]]

        for i in range(1, len(entries)):
            prev_x, prev_w = entries[i-1][1][0], entries[i-1][1][2]
            curr_x = entries[i][1][0]
            gap = curr_x - (prev_x + prev_w)

            if gap < 50:
                cluster.append(entries[i])
            else:
                clusters.append(cluster)
                cluster = [entries[i]]

        clusters.append(cluster)

        # Evaluate clusters
        for cluster in clusters:

            text = " ".join([w for w, _ in cluster])

            xs = [b[0] for _, b in cluster]
            ys = [b[1] for _, b in cluster]
            ws = [b[2] for _, b in cluster]
            hs = [b[3] for _, b in cluster]

            x1 = min(xs)
            y1 = min(ys)
            x2 = max(xs[i] + ws[i] for i in range(len(ws)))
            y2 = max(ys[i] + hs[i] for i in range(len(hs)))
            box = (x1, y1, x2 - x1, y2 - y1)

            # ---------------------
            # MULTI-WORD LOGIC
            if multi_word_mode:

                # direct phrase substring
                if target in text:
                    matches.append(box)
                    continue

                # fuzzy entire phrase
                if SequenceMatcher(None, target, text).ratio() >= fuzzy_threshold:
                    matches.append(box)
                    continue

                # token-level majority match
                token_hits = sum(1 for t in target_words if t in text)
                if token_hits >= max(1, len(target_words)//2):
                    matches.append(box)
                    continue

            # ---------------------
            # FALLBACK SINGLE-WORD LOGIC
            if single_word_mode:
                if target in text:
                    matches.append(box)
                    continue
                if SequenceMatcher(None, target, text).ratio() >= fuzzy_threshold:
                    matches.append(box)

    return matches

def scroll_and_find(target, max_scrolls=20):
    """Scroll the screen and search for text."""
    for _ in range(max_scrolls):
        boxes = find_text_boxes(target)
        if boxes:
            return boxes
        pyautogui.scroll(-800)  # scroll down
        time.sleep(0.3)
    return []

def perform_click_action(entity, is_double=False, full_cmd=""):
    target = entity.lower().strip()

    # Normalize number words
    num_map = {
        "zero":"0","one":"1","two":"2","three":"3","four":"4","five":"5",
        "six":"6","seven":"7","eight":"8","nine":"9","ten":"10"
    }
    parts = target.split()
    target = " ".join([num_map.get(w, w) for w in parts])

    reply(f"Searching for {target}...")

    # 1. Try exact/partial text match
    boxes = find_text_boxes(target)

    # Detect nth occurrence
    index = 0
    mapping = {
        "first":0,"1st":0,
        "second":1,"2nd":1,
        "third":2,"3rd":2,
        "fourth":3,"4th":3
    }

    for w in full_cmd.lower().split():
        if w in mapping:
            index = mapping[w]

    # 2. If not visible → scroll
    if not boxes:
        boxes = scroll_and_find(target)

    if not boxes:
        reply(f"I could not find {entity} on the screen.")
        return

    # 3. Region targeting
    cmd = full_cmd.lower()
    if "top right" in cmd:
        boxes = sorted(boxes, key=lambda b: (b[1], -b[0]))
    elif "top left" in cmd:
        boxes = sorted(boxes, key=lambda b: (b[1], b[0]))
    elif "bottom right" in cmd:
        boxes = sorted(boxes, key=lambda b: (-b[1], -b[0]))
    elif "bottom left" in cmd:
        boxes = sorted(boxes, key=lambda b: (-b[1], b[0]))

    # 6. Choose match by index
    try:
        x, y, w, h = boxes[index]
    except:
        x, y, w, h = boxes[0]

    # Final click
    pyautogui.moveTo(x + w//2, y + h//2, duration=0.25)
    pyautogui.doubleClick() if is_double else pyautogui.click()

    reply("Done.")

def focus_powerpoint():
    """Bring PowerPoint slideshow (not editor) into focus."""
    import pygetwindow as gw
    import pyautogui
    import time

    pyautogui.press("esc")

    try:
        keywords_slideshow = ["Slide Show", "PowerPoint Slide Show"]
        keywords_editor = ["PowerPoint", "Microsoft PowerPoint", "Presentation"]

        # Get all open windows
        all_titles = gw.getAllTitles()
        print(f"All open windows: {all_titles}")

        # 1️⃣ Try to find slideshow window first
        slide_windows = [t for t in all_titles if any(k.lower() in t.lower() for k in keywords_slideshow)]

        if slide_windows:
            # Focus existing slideshow window
            window_title = slide_windows[-1]
            ppt_window = gw.getWindowsWithTitle(window_title)[0]

            ppt_window.restore()
            ppt_window.activate()
            time.sleep(0.5)
            print(f"✅ Focused PowerPoint slideshow: {window_title}")
            return True

        # 2️⃣ If slideshow window not found, try to find PowerPoint editor and start slideshow
        ppt_windows = [t for t in all_titles if any(k.lower() in t.lower() for k in keywords_editor)]

        if ppt_windows:
            window_title = ppt_windows[-1]
            ppt_window = gw.getWindowsWithTitle(window_title)[0]

            try:
                ppt_window.restore()
                ppt_window.activate()
            except Exception as e:
                print(f"⚠️ Could not restore PowerPoint editor: {e}")

            time.sleep(0.8)
            print("🔁 Slideshow not open, starting presentation now...")
            pyautogui.press("f5")  # Start presentation
            time.sleep(2)

            # Try again to focus slideshow window
            all_titles = gw.getAllTitles()
            slide_windows = [t for t in all_titles if any(k.lower() in t.lower() for k in keywords_slideshow)]
            if slide_windows:
                slide_title = slide_windows[-1]
                slide_window = gw.getWindowsWithTitle(slide_title)[0]
                slide_window.activate()
                print(f"✅ Focused newly started slideshow: {slide_title}")
                return True

        # 3️⃣ Final fallback: use Alt+Tab and try to start slideshow
        print("⚠️ PowerPoint slideshow not found, using fallback...")
        pyautogui.keyDown("alt")
        pyautogui.press("tab")
        pyautogui.keyUp("alt")
        time.sleep(0.5)
        pyautogui.press("f5")
        print("✅ Started slideshow via fallback")
        return True

    except Exception as e:
        print(f"⚠️ Could not refocus PowerPoint slideshow: {e}")
        return False
# -----------------Intent Recognition System----------------------
class IntentRecognizer:
    def __init__(self):
        self.model = None
        self.vectorizer = TfidfVectorizer(ngram_range=(1, 2), max_features=1000)
        self.classifier = MultinomialNB()
        self.pipeline = make_pipeline(self.vectorizer, self.classifier)
        self.intent_labels = []
        self.training_data = self._create_training_data()
        self._train_model()
    
    def _create_training_data(self):
        """Create comprehensive training data for various intents"""
        training_data = {
            'greeting': [
                'hello', 'hi', 'hey', 'good morning', 'good afternoon', 'good evening',
                'how are you', 'whats up', 'hello there', 'hi there', 'hey there',
                'good morning how are you', 'good afternoon how are you'
            ],
            'name_query': [
                'what is your name', 'who are you', 'what should i call you',
                'your name', 'tell me your name', 'what are you called',
                'may i know your name', 'what do people call you'
            ],
            'time_query': [
                'what time is it', 'current time', 'time now', 'tell me the time',
                'what is the time', 'can you tell me the time', 'whats the time',
                'time please', 'current time please'
            ],
            'date_query': [
                'what is the date', 'current date', 'date today', 'todays date',
                'what date is it', 'can you tell me the date', 'whats the date',
                'date please'
            ],
            'search': [
                'search for', 'find', 'look up', 'google', 'search web for',
                'find information about', 'search about', 'search the web for',
                'look for', 'find details about', 'search online for'
            ],
            'location': [
                'location', 'place', 'map', 'locate', 'where is',
                'find location', 'show me location of', 'where can i find',
                'locate place', 'find place', 'show map of'
            ],
            'files': [
                'list files', 'show files', 'display files', 'what files are there',
                'browse files', 'explore directory', 'show folder contents',
                'list directory', 'show me files', 'display folder contents',
                'whats in this folder', 'open folder', 'open file'
            ],
            'open_file_path': [
                'open file', 'open folder', 'open directory', 'open this file', 
                'open this folder', 'go to folder', 'go to file', 'open my downloads',
                'open my desktop', 'open my documents'
            ],
            'file_navigate': [
                'go back', 'back', 'previous folder', 'navigate back',
                'go to previous', 'up one level'
            ],
            'goodbye': [
                'bye', 'goodbye', 'see you', 'farewell', 'quit', 'exit',
                'see you later', 'goodbye for now', 'bye bye', 'take care',
                'see you soon', 'i have to go'
            ],
            'copy': [
                'copy', 'copy this', 'copy text', 'copy that', 'copy selected',
                'copy to clipboard', 'copy the text', 'copy this text'
            ],
            'paste': [
                'paste', 'paste it', 'paste here', 'paste text', 'paste that',
                'paste from clipboard', 'paste the text', 'paste here please'
            ],
            'wake_up': [
                'wake up', 'start', 'activate', 'hello omega', 'are you there',
                'wake up omega', 'start listening', 'activate omega'
            ],
            'presentation_control': [
            'next slide', 'previous slide', 'go back a slide', 'go to next slide',
            'show next', 'show previous', 'start presentation', 'begin slideshow',
            'end presentation', 'stop presentation', 'pause presentation',
            'resume presentation', 'exit slideshow', 'zoom in', 'zoom out',
            'increase zoom', 'decrease zoom', 'make it bigger', 'make it smaller',
            'go full screen', 'exit full screen', 'move to first slide',
            'go to last slide', 'skip to slide', 'show slide number','resume presentation',
            'pause presentation', 'return to presentation', 'come back', 'bring presentation back',
            'continue presentation'
            ],
            'open_app': ['open app', 'launch app', 'start app', 'open application', 'run', 'open program', 'start program'],
            'click_action': ['click', 'click on', 'double click', 'select', 'press', 'open this','tap',
                            'choose', 'scroll and click', 'click the second', 'click the third']

        }
        return training_data
    
    def _train_model(self):
        """Train the intent classification model"""
        texts = []
        labels = []
        
        for intent, examples in self.training_data.items():
            for example in examples:
                texts.append(example)
                labels.append(intent)
        
        if texts:
            self.pipeline.fit(texts, labels)
            self.intent_labels = list(self.training_data.keys())
            print("Intent recognition model trained successfully!")
    
    def predict_intent(self, text):
        """Predict the intent of the given text"""
        try:
            # Preprocess text
            processed_text = self._preprocess_text(text)
            
            if not processed_text.strip():
                return 'unknown'
            
            # Predict using the model
            prediction = self.pipeline.predict([processed_text])[0]
            probability = np.max(self.pipeline.predict_proba([processed_text]))
            
            # Only return prediction if confidence is high enough
            if probability > 0.3:
                return prediction
            else:
                return self._fallback_intent(processed_text)
                
        except Exception as e:
            print(f"Intent prediction error: {e}")
            return self._fallback_intent(text)
    
    def _preprocess_text(self, text):
        """Preprocess text for better matching"""
        if not text:
            return ""
        
        # Convert to lowercase and remove extra spaces
        text = text.lower().strip()
        
        # Remove common filler words
        filler_words = ['the', 'a', 'an', 'please', 'could', 'you', 'would', 'can', 'will', 'should']
        words = text.split()
        cleaned_words = [word for word in words if word not in filler_words]
        
        return ' '.join(cleaned_words)
    
    def _fallback_intent(self, text):
        """Fallback method using keyword matching"""
        text_lower = text.lower()
        
        # Enhanced keyword matching with flexible patterns
        intent_keywords = {
            'greeting': ['hello', 'hi', 'hey', 'good morning', 'good afternoon', 'good evening'],
            'name_query': ['name', 'who are you', 'your name'],
            'time_query': ['time', 'clock', 'what time'],
            'date_query': ['date', 'today', 'what date'],
            'search': ['search', 'find', 'look up', 'google'],
            'location': ['location', 'map', 'where is', 'locate'],
            'files': ['file', 'files', 'directory', 'folder', 'list'],
            'open_file_path': ['open file', 'open folder', 'directory', 'folder', 'file'],
            'file_navigate': ['back', 'previous', 'go back', 'up'],
            'goodbye': ['bye', 'goodbye', 'exit', 'quit'],
            'copy': ['copy', 'copy this'],
            'paste': ['paste', 'paste it'],
            'wake_up': ['wake up', 'start', 'activate'],
            'presentation_control': [
            'next slide', 'previous slide', 'go back a slide', 'go to next slide',
            'show next', 'show previous', 'start presentation', 'begin slideshow',
            'end presentation', 'stop presentation', 'pause presentation',
            'resume presentation', 'exit slideshow', 'zoom in', 'zoom out',
            'increase zoom', 'decrease zoom', 'make it bigger', 'make it smaller',
            'full screen', 'exit full screen', 'first slide', 'last slide',
            'skip to slide', 'show slide number'],
            'open_app': ['open app', 'launch', 'start', 'run'],
            'click_action': ['click', 'click on', 'double click', 'tap', 'press',
            'select', 'choose']

        }
        
        for intent, keywords in intent_keywords.items():
            for keyword in keywords:
                if keyword in text_lower:
                    return intent
        
        return 'unknown'

# Initialize intent recognizer
intent_recognizer = IntentRecognizer()




def is_double_click_command(text):
    return "double" in text.lower()


# ------------------Functions----------------------
def reply(audio):
    if not audio:
        return
        
    app.ChatBot.addAppMsg(audio)
    print(f"omega: {audio}")
    
    if engine:
        try:
            engine.say(audio)
            engine.runAndWait()
        except Exception as e:
            print(f"TTS error: {e}")

def wish():
    hour = int(datetime.datetime.now().hour)

    if hour >= 0 and hour < 12:
        reply("Good Morning!")
    elif hour >= 12 and hour < 18:
        reply("Good Afternoon!")   
    else:
        reply("Good Evening!")  
        
    reply("I am omega, how may I help you?")

def extract_entity(text, intent):
    """Extract entities from text based on intent"""
    text_lower = text.lower()
    
    if intent == 'search':
        # Remove search-related words to get the query
        search_words = ['search', 'find', 'look up', 'google', 'for', 'about']
        query = text_lower
        for word in search_words:
            query = query.replace(word, '').strip()
        return query if query and len(query) > 1 else None
    
    elif intent == 'location':
        # Remove location-related words to get the place
        location_words = ['location', 'place', 'map', 'locate', 'where is', 'of']
        place = text_lower
        for word in location_words:
            place = place.replace(word, '').strip()
        return place if place and len(place) > 1 else None
    
    elif intent == 'open_file_path':
        cleaned = re.sub(
            r'\b(open|file|folder|directory|go to|please|the|my)\b', 
            '', 
            text_lower
        )
        return cleaned.strip() if cleaned.strip() else None

    
    elif intent == 'file_navigate':
        return 'back'
    
    elif intent == 'open_app':
    # remove open keywords
     name = re.sub(r'\b(open|launch|start|run|app|application|program)\b', '', text_lower).strip()
     return name if name else None

    elif intent == 'files':
        cleaned = re.sub(
            r'\b(open|file|files|folder|directory|go to|please|the|my)\b',
            '',
            text_lower
        )
        return cleaned.strip() if cleaned.strip() else None
    
    elif intent == 'click_action':
        text = text_lower

        # Remove click-related words
        remove_words = [
            "double click", "click on", "click", "tap", "press", 
            "select", "choose"
        ]

        for w in remove_words:
            text = text.replace(w, "")

        # Remove meaningless fillers
        filler = ["the", "this", "that", "a", "an", "please"]
        words = [w for w in text.split() if w not in filler]

        # The entity (target) must be the last meaningful words
        return " ".join(words).strip()

    
    elif intent == 'presentation_control':
        # Remove the wake word and common filler words, but keep "slide"
        control_words = ['omega', 'presentation', 'slideshow', 'powerpoint', 'keynote', 'please','to the', 'to presentation', 'back to']
        command = text_lower
        
        # Remove only whole words to avoid partial deletions
        for word in control_words:
            command = re.sub(rf'\b{word}\b', '', command).strip()
        
        # Clean up extra spaces
        command = re.sub(r'\s+', ' ', command).strip()
        
        # Return cleaned command if valid
        return command if command and len(command) > 1 else text_lower

    return None



    



# def open_file_or_folder(full_path):
#     """Open a file or folder using the appropriate method"""
#     try:
#         if exists(full_path):
#             if isfile(full_path):
#                 # It's a file - open it with default application
#                 os.startfile(full_path)
#                 return f"Opened file: {os.path.basename(full_path)}"
#             else:
#                 # It's a folder - navigate into it
#                 global current_directory, files, file_exp_status
#                 current_directory = full_path
#                 files = listdir(current_directory)
                
#                 # Display new directory contents
#                 filestr = ""
#                 counter = 0
#                 for f in files:
#                     counter += 1
#                     item_type = "file" if isfile(join(current_directory, f)) else "folder"
#                     filestr += f"{counter}: {f} ({item_type})<br>"
#                     print(f"{counter}: {f} ({item_type})")
                
#                 file_exp_status = True
#                 app.ChatBot.addAppMsg(filestr)
#                 return f"Opened folder: {os.path.basename(full_path)}"
#         else:
#             return f"Path does not exist: {full_path}"
#     except Exception as e:
#         return f"Error opening: {str(e)}"

def navigate_back():
    """Navigate to parent directory"""
    global current_directory, files, file_exp_status
    
    if current_directory == 'C://':
        return "Already at root directory"
    
    try:
        # Go up one level
        parent_dir = os.path.dirname(current_directory)
        if parent_dir:  # Ensure we don't go above C://
            current_directory = parent_dir
        else:
            current_directory = 'C://'
        
        # List contents of new directory
        files = listdir(current_directory)
        filestr = ""
        counter = 0
        for f in files:
            counter += 1
            item_type = "file" if isfile(join(current_directory, f)) else "folder"
            filestr += f"{counter}: {f} ({item_type})<br>"
            print(f"{counter}: {f} ({item_type})")
        
        app.ChatBot.addAppMsg(filestr)
        return f"Navigated back to: {current_directory}"
    except Exception as e:
        return f"Error navigating back: {str(e)}"

# Audio to String with better error handling
def record_audio():
    try:
        with sr.Microphone() as source:
            r.adjust_for_ambient_noise(source, duration=0.5)
            r.pause_threshold = 1.0
            r.energy_threshold = 300
            print("Listening...")
            audio = r.listen(source, timeout=5, phrase_time_limit=5)
            
        try:
            voice_data = r.recognize_google(audio).lower()
            print(f"Recognized: {voice_data}")
            return voice_data
        except sr.UnknownValueError:
            print("Could not understand audio")
            return ""
        except sr.RequestError as e:
            print(f"Speech recognition error: {e}")
            reply('Speech recognition service error. Check internet connection.')
            return ""
            
    except sr.WaitTimeoutError:
        print("Listening timeout")
        return ""
    except Exception as e:
        print(f"Microphone error: {e}")
        return ""

def handle_greeting():
    """Handle greeting intent"""
    wish()

def handle_name_query():
    """Handle name query intent"""
    reply('My name is omega! I am your personal voice assistant.')

def handle_time_query():
    """Handle time query intent"""
    current_time = datetime.datetime.now().strftime("%I:%M %p")
    reply(f'The current time is {current_time}')

def handle_date_query():
    """Handle date query intent"""
    reply(today.strftime("Today's date is %B %d, %Y"))

def handle_search(query):
    """
    Launch a web search, and remember if we came from a presentation
    so we can return on demand (or automatically if AUTO_RETURN_AFTER_SEARCH=True).
    """
    global last_context
    if query:
        reply(f"Searching for {query}")
        try:
            import pygetwindow as gw
            import time
            import webbrowser

            # Detect current active window (if possible) to remember context
            try:
                active_win = gw.getActiveWindow()
                active_title = active_win.title.lower() if active_win and active_win.title else ""
                if "powerpoint" in active_title or "slide show" in active_title or "presentation" in active_title:
                    last_context = "presentation"
                    # ✅ Minimize PowerPoint temporarily
                    try:
                        active_win.minimize()
                        print("🪄 Minimized PowerPoint before search to keep Chrome visible.")
                    except Exception as e:
                        print(f"Could not minimize PowerPoint: {e}")
                else:
                    last_context = None
            except Exception as _:
                last_context = None

            # Perform the search
            url = "https://google.com/search?q=" + query.replace(" ", "+")
            webbrowser.get().open(url)
            reply("Here are the search results I found")

            # Optional: automatic return if explicitly enabled
            if AUTO_RETURN_AFTER_SEARCH and last_context == "presentation":
                time.sleep(2)
                pyautogui.press("esc")


                success = focus_powerpoint()
                if success:
                    reply("Returned to the presentation.")
                else:
                    reply("I couldn't return to the presentation automatically.")

        except Exception as e:
            print("handle_search error:", e)
            reply("Please check your internet connection")
    else:
        reply("What would you like me to search for?")
        temp_audio = record_audio()
        if temp_audio:
            handle_search(temp_audio)



def handle_location(place):
    """Handle location intent"""
    if place:
        reply(f'Locating {place} on maps')
        url = 'https://google.com/maps/place/' + place.replace(' ', '+')
        try:
            webbrowser.get().open(url)
            reply('Here is the location on maps')
        except:
            reply('Please check your internet connection')
    else:
        reply('Which place are you looking for?')
        temp_audio = record_audio()
        if temp_audio:
            place = temp_audio
            reply(f'Locating {place} on maps')
            url = 'https://google.com/maps/place/' + place.replace(' ', '+')
            try:
                webbrowser.get().open(url)
                reply('This is what I found')
            except:
                reply('Please check your internet connection')

# def handle_files(operation, voice_data):
#     """Handle file listing"""
#     global file_exp_status, files, current_directory
    
#     if operation == 'list' or not operation:
#         counter = 0
#         current_directory = 'C://'
#         try:
#             files = listdir(current_directory)
#             filestr = ""
#             for f in files:
#                 counter += 1
#                 item_type = "file" if isfile(join(current_directory, f)) else "folder"
#                 filestr += f"{counter}: {f} ({item_type})<br>"
#                 print(f"{counter}: {f} ({item_type})")
#             file_exp_status = True
#             reply('These are the files in your root directory')
#             app.ChatBot.addAppMsg(filestr)
#         except Exception as e:
#             reply(f'Error accessing directory: {e}')
            
def handle_open_file_path(name):
    
    if not name:
        reply("Which file or folder should I open?")
        return

    name = name.lower().strip()
    home = os.path.expanduser("~")

    # ---------------------------
    # 1. Friendly folder names
    # ---------------------------
    known = {
        "desktop": os.path.join(home, "Desktop"),
        "documents": os.path.join(home, "Documents"),
        "downloads": os.path.join(home, "Downloads"),
        "pictures": os.path.join(home, "Pictures"),
        "music": os.path.join(home, "Music"),
        "videos": os.path.join(home, "Videos"),
        "c drive": "C:\\",
        "d drive": "D:\\"
    }

    if name in known:
        path = known[name]
        if os.path.exists(path):
            try:
                os.startfile(path)
                reply(f"Opening {name}")
                return
            except Exception as e:
                reply(f"Could not open {name}: {e}")
                return

    # ---------------------------
    # 2. Full path handling
    # ---------------------------
    if " slash " in name:
        name = name.replace(" slash ", os.sep)

    if os.path.exists(name):
        try:
            os.startfile(name)
            reply(f"Opening {name}")
            return
        except Exception as e:
            reply(f"Could not open {name}: {e}")
            return

    # ---------------------------
    # 3. Recent file support
    # ---------------------------
    if name.startswith("recent"):
        parts = name.split()
        pattern = "*"
        if len(parts) > 1:
            q = parts[1]
            pattern = f"*{q}*" if "." not in q else f"*{q}"

        search_dirs = [home, os.path.join(home, "Downloads"), os.path.join(home, "Desktop")]
        latest = (None, 0)

        for d in search_dirs:
            for root, _, files in os.walk(d):
                for f in files:
                    if glob.fnmatch.fnmatch(f.lower(), pattern.lower()):
                        full = os.path.join(root, f)
                        try:
                            mtime = os.path.getmtime(full)
                            if mtime > latest[1]:
                                latest = (full, mtime)
                        except:
                            pass

        if latest[0]:
            try:
                os.startfile(latest[0])
                reply(f"Opening recent file {os.path.basename(latest[0])}")
                return
            except Exception as e:
                reply(f"Could not open recent file: {e}")
                return

        reply("No recent file found.")
        return

    # ---------------------------
    # 4. Search for folders/files (PRIORITIZED)
    # ---------------------------
    search_dirs = [
        os.path.join(home, "Desktop"),
        os.path.join(home, "Documents"),
        os.path.join(home, "Downloads"),
        home
    ]

    best_startswith = None
    best_contains = None

    for d in search_dirs:
        for root, dirs, files in os.walk(d):

            # Avoid system / cache folders
            skip = ["appdata", "cache", "indexeddb", "leveldb", "temp"]
            if any(s in root.lower() for s in skip):
                continue

            # ---------Exact folder match---------
            for fol in dirs:
                if fol.lower() == name:
                    full = os.path.join(root, fol)
                    os.startfile(full)
                    reply(f"Opening folder {fol}")
                    return

            # ---------Starts-with match---------
            for fol in dirs:
                if fol.lower().startswith(name):
                    best_startswith = os.path.join(root, fol)

            # ---------Contains match---------
            for fol in dirs:
                if name in fol.lower():
                    if not best_contains:
                        best_contains = os.path.join(root, fol)

            # ---------Exact file match---------
            for f in files:
                if f.lower() == name:
                    full = os.path.join(root, f)
                    os.startfile(full)
                    reply(f"Opening file {f}")
                    return

            # ---------File starts-with---------
            for f in files:
                if f.lower().startswith(name):
                    best_startswith = os.path.join(root, f)

            # ---------File contains---------
            for f in files:
                if name in f.lower():
                    if not best_contains:
                        best_contains = os.path.join(root, f)

    # ---------------------------
    # 5. Choose the best match
    # ---------------------------
    if best_startswith:
        os.startfile(best_startswith)
        reply(f"Opening {os.path.basename(best_startswith)}")
        return

    if best_contains:
        os.startfile(best_contains)
        reply(f"Opening {os.path.basename(best_contains)}")
        return

    # ---------------------------
    # 6. Nothing found
    # ---------------------------
    reply(f"I couldn't find any file or folder named {name}.")
    return



# def handle_file_open(file_number, voice_data):
#     """Handle file/folder opening by number"""
#     global file_exp_status, files, current_directory
    
#     if not file_exp_status:
#         reply("Please list files first using 'list files' command")
#         return
    
#     try:
#         if file_number and 1 <= file_number <= len(files):
#             file_index = file_number - 1
#             selected_item = files[file_index]
#             full_path = join(current_directory, selected_item)
            
#             result = open_file_or_folder(full_path)
#             reply(result)
#         else:
#             reply(f"Please specify a valid file number between 1 and {len(files)}")
#     except Exception as e:
#         reply(f"Error opening file: {str(e)}")

def handle_file_navigate(operation):
    """Handle file navigation"""
    if operation == 'back':
        result = navigate_back()
        reply(result)

def handle_goodbye():
    """Handle goodbye intent"""
    global is_awake
    reply("Goodbye! Have a nice day!")
    is_awake = False

def handle_copy():
    """Handle copy intent"""
    with keyboard.pressed(Key.ctrl):
        keyboard.press('c')
        keyboard.release('c')
    reply('Copied to clipboard')

def handle_paste():
    """Handle paste intent"""
    with keyboard.pressed(Key.ctrl):
        keyboard.press('v')
        keyboard.release('v')
    reply('Pasted from clipboard')

def handle_wake_up():
    """Handle wake up intent"""
    global is_awake
    is_awake = True
    wish()

def handle_open_app(app_name):
    if not app_name:
        reply("Which application should I open?")
        return
    if open_taskbar_app(app_name):
        return
    result = open_system_app(app_name)
    reply(result)


def handle_presentation_control(command):
    """
    Handle presentation control commands such as:
    next slide, previous slide, start presentation, end presentation,resume presentation etc.
    """
    command = command.lower().strip()

    try:
        # Match against common variations of user commands
        if any(kw in command for kw in ["next", "forward", "advance"]):
            pyautogui.press('right')
            reply("Moved to the next slide.")
        
        elif any(kw in command for kw in ["previous", "back", "go back"]):
            pyautogui.press('left')
            reply("Went back to the previous slide.")
        
        elif any(kw in command for kw in ["start", "begin", "show", "play", "slideshow", "present"]):
            print("🎬 Starting presentation...")
            success = focus_powerpoint()  # make sure PowerPoint is active
            if success:
                time.sleep(0.8)
                pyautogui.press("f5")  # Start presentation from beginning
                reply("Starting your presentation now.")
            else:
                reply("I couldn't find PowerPoint to start the presentation.")
        
        elif any(kw in command for kw in ["end", "exit", "stop", "close"]):
            print("🛑 Exiting full screen or slideshow...")
            minimize_browser_windows()
            time.sleep(0.8)
            success = focus_powerpoint()
            if success:
                time.sleep(0.5)
                pyautogui.press("esc")  # Exit slideshow mode
                reply("Exited the presentation screen.")
            else:
                reply("I couldn't find PowerPoint to exit the presentation.")
        
        elif any(kw in command for kw in ["pause"]):
            pyautogui.press('b')  # black screen pause
            reply("Presentation paused.")
        
        elif any(kw in command for kw in ["resume", "continue"]):
            pyautogui.press('b')  # resumes from black screen
            reply("Resumed the presentation.")
        
        elif any(kw in command for kw in ["zoom in", "increase zoom", "bigger"]):
            pyautogui.hotkey('ctrl', '+')
            reply("Zoomed in.")
        
        elif any(kw in command for kw in ["zoom out", "decrease zoom", "smaller"]):
            pyautogui.hotkey('ctrl', '-')
            reply("Zoomed out.")
        
        elif "full screen" in command:
            pyautogui.press('f11')
            reply("Entered full screen mode.")
        
        elif "exit full screen" in command:
            pyautogui.press('esc')
            reply("Exited full screen mode.")
        
        elif any(kw in command for kw in ["first slide", "beginning"]):
            pyautogui.hotkey('home')
            reply("Moved to the first slide.")
        
        elif any(kw in command for kw in ["last slide", "final slide"]):
            pyautogui.hotkey('end')
            reply("Moved to the last slide.")
        
            # ... existing code ...
        elif any(kw in command for kw in ["return", "come back", "bring back", "continue", "resume"]):
            # Try to refocus PowerPoint
            minimize_browser_windows()
            time.sleep(0.8)

            success = focus_powerpoint()
            if success:
                reply("Returned to the presentation.")
            else:
                reply("I couldn't find the presentation window to return to.")

        
        elif any(kw in command for kw in ["slide", "skip to slide", "go to slide"]):
            # Extract slide number if mentioned
            import re
            match = re.search(r'\d+', command)
            if match:
                slide_num = int(match.group())
                pyautogui.typewrite(str(slide_num))
                pyautogui.press('enter')
                reply(f"Moved to slide number {slide_num}.")
            else:
                reply("Please specify which slide number to go to.")
        
        else:
            reply("I'm not sure what you want to do with the presentation.")
    
    except Exception as e:
        print(f"Presentation control error: {e}")
        reply("Sorry, I couldn’t control the presentation right now.")

def handle_click_action(entity, full_cmd):
    if not entity:
        reply("What should I click on?")
        return
    
    perform_click_action(
        entity=entity,
        is_double=is_double_click_command(full_cmd),
        full_cmd=full_cmd
    )

# Executes Commands (input: string)
def respond(voice_data):
    global file_exp_status, files, is_awake, current_directory
    
    if not voice_data:
        return
        
    print(f"Processing: {voice_data}")
    
    # Store original for display
    original_voice_data = voice_data
    
    # Remove wake word and clean up for processing
    processed_voice = voice_data.replace('omega', '').strip()
    app.ChatBot.addUserMsg(original_voice_data)

    # Check if bot is asleep
    if not is_awake:
        intent = intent_recognizer.predict_intent(processed_voice)
        if intent == 'wake_up':
            handle_wake_up()
        return

    # Use ML-based intent recognition
    intent = intent_recognizer.predict_intent(processed_voice)
    entity = extract_entity(processed_voice, intent)
    
    print(f"Detected Intent: {intent}, Entity: {entity}")
    
    # Handle different intents
    if intent == 'greeting':
        handle_greeting()
    
    elif intent == 'name_query':
        handle_name_query()
    
    elif intent == 'time_query':
        handle_time_query()
    
    elif intent == 'date_query':
        handle_date_query()
    
    elif intent == 'search':
        handle_search(entity)
    
    elif intent == 'location':
        handle_location(entity)
    
    elif intent == 'files':
        if not entity:
            entity = extract_entity(processed_voice, 'open_file_path')
        handle_open_file_path(entity)    
    
    elif intent == 'open_file_path':
        handle_open_file_path(entity)
    
    elif intent == 'file_navigate':
        handle_file_navigate(entity)
    
    elif intent == 'goodbye':
        handle_goodbye()
    
    elif intent == 'copy':
        handle_copy()
    
    elif intent == 'paste':
        handle_paste()
    
    elif intent == 'wake_up':
        handle_wake_up()

    elif intent == 'presentation_control':
        handle_presentation_control(entity)
    
    elif intent == 'open_app':
        handle_open_app(entity)

    elif intent == 'click_action':
        handle_click_action(entity, processed_voice)

    
    else:
        # Fallback for unknown intents
        reply("I'm not sure how to help with that. You can ask me about time, date, search, files, or location.")

# ------------------Driver Code--------------------

if __name__ == "__main__":
    # Start the chatbot GUI in a separate thread
    t1 = Thread(target=app.ChatBot.start, daemon=True)
    t1.start()

    # Wait for chatbot to start
    max_wait = 10  # seconds
    wait_time = 0
    while not app.ChatBot.started and wait_time < max_wait:
        time.sleep(0.5)
        wait_time += 0.5
        print("Waiting for chatbot to start...")

    if not app.ChatBot.started:
        print("Chatbot failed to start within expected time")
    else:
        print("Chatbot started successfully")

    print(get_taskbar_buttons())    

    wish()
    
    while True:
        try:
            if app.ChatBot.isUserInput():
                # Take input from GUI
                voice_data = app.ChatBot.popUserInput()
            else:
                # Take input from Voice
                voice_data = record_audio()

            # Process voice_data if we have input
            if voice_data:
                # Check if omega is mentioned OR if we have GUI input
                if 'omega' in voice_data.lower() or app.ChatBot.isUserInput():
                    result = respond(voice_data)
                    if result == "exit":
                        break
                else:
                    print("Voice activation word 'omega' not detected")
                    
            time.sleep(0.1)  # Small delay to prevent CPU overuse
                    
        except SystemExit:
            reply("Exit Successful")
            break
        except KeyboardInterrupt:
            reply("Interrupted by user")
            break
        except Exception as e:
            print(f"Unexpected error: {e}")
            time.sleep(1)  # Prevent rapid error looping

    print("omega shutdown complete")