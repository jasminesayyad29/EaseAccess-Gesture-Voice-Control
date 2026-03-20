import pyttsx3
import speech_recognition as sr
from datetime import date
import time
import webbrowser
import datetime
from pynput.keyboard import Key, Controller
import pyautogui
import sys
import builtins
import os
import psutil
from os import listdir
from os.path import isfile, join, exists
import wikipedia
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
import pychrome
from queue import Queue, Empty, Full
from threading import Thread, Event
try:
    from voice_indicator import VoiceStateIndicator
except Exception:
    VoiceStateIndicator = None


# Keep track of where we came from so we can return if needed
last_context = None   # possible values: None, 'presentation', 'browser', ...
AUTO_RETURN_AFTER_SEARCH = False # Set True if you want automatic return after search
last_found_boxes = []
last_click_target_text = None

paste_word_mode = None        # None, "before", or "after"
paste_word_boxes = []         # boxes of the matched word(s)
paste_word_target = None      # the word like "stay" / "save"

paste_position_mode = None      
paste_position_boxes = []

pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# -------------Object Initialization---------------
today = date.today()
r = sr.Recognizer()
keyboard = Controller()

# Voice recognition tuning (balanced for noisy rooms)
r.dynamic_energy_threshold = True
r.dynamic_energy_adjustment_damping = 0.15
r.dynamic_energy_ratio = 1.7
r.pause_threshold = 0.8
r.phrase_threshold = 0.25
r.non_speaking_duration = 0.4

engine = None


def _init_tts_engine():
    """Initialize TTS in the same thread that will speak (important on Windows)."""
    try:
        local_engine = pyttsx3.init('sapi5')
    except Exception:
        try:
            local_engine = pyttsx3.init()
        except Exception as e:
            print(f"TTS initialization failed: {e}")
            return None

    try:
        voices = local_engine.getProperty('voices')
        if voices:
            local_engine.setProperty('voice', voices[0].id)
    except Exception:
        pass

    return local_engine

voice_indicator = VoiceStateIndicator() if VoiceStateIndicator is not None else None


def _voice_show_listening():
    if voice_indicator is None:
        return
    try:
        voice_indicator.show_listening()
    except Exception:
        pass


def _voice_show_recognized(recognized_text=""):
    if voice_indicator is None:
        return
    try:
        voice_indicator.show_recognized(recognized_text=recognized_text)
    except Exception:
        pass


def _voice_hide():
    if voice_indicator is None:
        return
    try:
        voice_indicator.hide()
    except Exception:
        pass


def _voice_close():
    if voice_indicator is None:
        return
    try:
        voice_indicator.close()
    except Exception:
        pass

# ----------------Variables------------------------
file_exp_status = False
files = []
path = ''
is_awake = True  # Bot status
current_directory = 'C://'  # Track current directory
import winreg
import glob


# Chrome Tab Memory
last_chrome_tab_id = None
current_chrome_tab_id = None

WAKE_WORD_VARIANTS = (
    "omega", "oh mega", "o mega", "ome ga", "amiga", "omegaa", "omegle", "omagle"
)

COMMON_STT_FIXES = {
    "thankyou": "thank you",
    "thanks you": "thank you",
    "omegle": "omega",
    "omagle": "omega",
    "click on you": "click on two",
    "click on to": "click on two",
    "choose you": "choose two",
    "search four": "search for",
    "switch toe tab": "switch to tab",
    "switch two tab": "switch to tab",
    "go too tab": "go to tab",
    "opun": "open",
    "clik": "click",
}

APP_ALIAS_GROUPS = {
    "chrome": ["chrome", "google chrome", "chrome.exe"],
    "word": ["word", "microsoft word", "winword", "winword.exe"],
    "powerpoint": ["powerpoint", "microsoft powerpoint", "power point", "powerpnt", "powerpnt.exe"],
    "excel": ["excel", "microsoft excel", "excel.exe"],
}

APP_DIRECT_EXECUTABLES = {
    "chrome": ["chrome.exe"],
    "word": ["winword.exe"],
    "powerpoint": ["powerpnt.exe"],
    "excel": ["excel.exe"],
}

AMBIENT_RECALIBRATE_EVERY_SEC = 90
ACTIVE_COMMAND_WINDOW_SEC = 14
_audio_calibrated = False
_last_ambient_calibration_at = 0.0
_last_wake_detected_at = 0.0

APP_DISCOVERY_CACHE_TTL_SEC = 120
_apps_cache = {}
_apps_cache_time = 0.0

COMMAND_QUEUE_MAXSIZE = 32
TTS_QUEUE_MAXSIZE = 256

command_queue = Queue(maxsize=COMMAND_QUEUE_MAXSIZE)
tts_queue = Queue(maxsize=TTS_QUEUE_MAXSIZE)
shutdown_event = Event()

_raw_print = builtins.print


def print(*args, **kwargs):
    """Print to console and mirror the same text to TTS when available."""
    _raw_print(*args, **kwargs)

    sep = kwargs.get("sep", " ")
    try:
        text = sep.join(str(a) for a in args).strip()
    except Exception:
        text = ""

    if not text:
        return

    enqueue_fn = globals().get("_enqueue_tts")
    if callable(enqueue_fn):
        enqueue_fn(text)


def _enqueue_latest(q, item):
    """Keep queue fresh under bursty input by dropping the oldest item if full."""
    try:
        q.put_nowait(item)
    except Full:
        try:
            q.get_nowait()
        except Empty:
            pass
        try:
            q.put_nowait(item)
        except Full:
            pass


def _enqueue_tts(msg):
    if not msg:
        return
    while not shutdown_event.is_set():
        try:
            tts_queue.put(msg, timeout=0.25)
            return
        except Full:
            continue


def _tts_worker():
    global engine
    try:
        pythoncom.CoInitialize()
    except Exception:
        pass

    engine = _init_tts_engine()
    if engine is None:
        print("TTS engine unavailable in worker thread")

    while True:
        try:
            msg = tts_queue.get(timeout=0.25)
        except Empty:
            if shutdown_event.is_set():
                continue
            continue

        if msg is None:
            break

        if engine:
            try:
                engine.say(msg)
                engine.runAndWait()
            except Exception as e:
                print(f"TTS error: {e}")
    try:
        pythoncom.CoUninitialize()
    except Exception:
        pass


def is_thank_you_exit_command(text):
    if not text:
        return False
    t = normalize_voice_text(text)
    return t in {"thank you", "thanks", "thanks omega", "thank you omega", "thankyou"}


def handle_thank_you_exit():
    reply("You're welcome. Shutting down now.")
    return "exit"


def _should_allow_command(voice_data: str) -> bool:
    global _last_wake_detected_at
    normalized_voice = normalize_voice_text(voice_data)
    wake_detected = any(w in normalized_voice for w in WAKE_WORD_VARIANTS)
    if wake_detected:
        _last_wake_detected_at = time.time()
    return wake_detected or ((time.time() - _last_wake_detected_at) <= ACTIVE_COMMAND_WINDOW_SEC)


def _command_worker():
    while not shutdown_event.is_set():
        try:
            voice_data = command_queue.get(timeout=0.25)
        except Empty:
            continue

        if voice_data is None:
            break

        try:
            result = respond(voice_data)
            if result == "exit":
                shutdown_event.set()
                break
        except Exception as e:
            print(f"Command worker error: {e}")


def _normalize_app_key(value: str) -> str:
    if not value:
        return ""
    return re.sub(r"\s+", " ", re.sub(r"[^a-z0-9\s]", " ", value.lower())).strip()


def _add_app_entry(apps, name, target):
    """Store multiple normalized aliases for better matching."""
    if not name or not target:
        return

    raw_name = name.strip().lower()
    norm_name = _normalize_app_key(name)

    for key in (raw_name, norm_name):
        if key:
            apps[key] = target


def _get_start_apps_from_os():
    """Read Windows Start apps (Name + AppID) from this system."""
    ps_cmd = (
        "Get-StartApps | "
        "Select-Object Name,AppID | "
        "Where-Object { $_.Name -and $_.AppID } | "
        "ConvertTo-Json -Depth 3"
    )

    try:
        result = subprocess.run(
            ["powershell", "-NoProfile", "-Command", ps_cmd],
            capture_output=True,
            text=True,
            timeout=8,
        )
        if result.returncode != 0:
            return []

        payload = (result.stdout or "").strip()
        if not payload:
            return []

        data = json.loads(payload)
        if isinstance(data, dict):
            data = [data]
        if not isinstance(data, list):
            return []
        return data
    except Exception:
        return []


def _launch_app_user_model_id(app_id: str) -> bool:
    if not app_id:
        return False
    try:
        subprocess.Popen(["explorer.exe", f"shell:AppsFolder\\{app_id}"])
        return True
    except Exception:
        return False

def normalize_voice_text(text):
    """Normalize transcript to reduce STT noise before intent extraction."""
    if not text:
        return ""
    cleaned = re.sub(r"[^a-zA-Z0-9\s]", " ", text.lower())
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    for wrong, right in COMMON_STT_FIXES.items():
        cleaned = re.sub(rf"\b{re.escape(wrong)}\b", right, cleaned)
    return cleaned

def strip_wake_word(text):
    """Remove wake-word variants from the beginning of a transcript."""
    if not text:
        return ""
    output = text
    for wake_word in WAKE_WORD_VARIANTS:
        output = re.sub(rf"^\s*{re.escape(wake_word)}[\s,.:-]*", "", output, flags=re.IGNORECASE)
    return output.strip()

def normalize_app_query(text):
    """Extract app name from variants like 'open chrome' and 'open app chrome'."""
    if not text:
        return ""
    cleaned = normalize_voice_text(text)
    cleaned = re.sub(
        r"\b(open|launch|start|run|app|application|program|command|please|the)\b",
        " ",
        cleaned
    )
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    return cleaned

def is_open_app_command(text):
    """Detect app-open intent for plain 'open <app>' and 'open app <app>' forms."""
    if not text:
        return False
    t = normalize_voice_text(text)
    if not (t.startswith("open ") or t.startswith("launch ") or t.startswith("start ") or t.startswith("run ")):
        return False

    file_hints = {"file", "folder", "directory", "path", "documents", "downloads", "desktop", "pictures"}
    words = set(t.split())
    if words & file_hints:
        return False
    return True


def normalize_close_app_query(text):
    """Extract target app name from close-app phrasing."""
    if not text:
        return ""
    cleaned = normalize_voice_text(text)
    cleaned = re.sub(
        r"\b(close|quit|exit|stop|end|terminate|kill|force|app|application|program|please|the)\b",
        " ",
        cleaned,
    )
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    return cleaned


def is_close_app_command(text):
    """Detect close-app intent for forms like 'close app chrome' and 'close chrome'."""
    if not text:
        return False

    t = normalize_voice_text(text)
    if not t:
        return False

    if any(k in t for k in ["discard options", "dismiss options", "clear options", "close options", "hide options"]):
        return False

    blocked = {"presentation", "slideshow", "full screen", "tab", "options"}
    if any(k in t for k in blocked):
        return False

    verbs = ("close ", "quit ", "exit ", "stop ", "end ", "terminate ", "kill ")
    if not t.startswith(verbs):
        return False

    candidate = normalize_close_app_query(t)
    return bool(candidate)


def chrome_get_tabs():
    """Return Chromium DevTools connection & list of tabs (any Chromium browser)."""
    try:
        # Debugging ALWAYS listens on localhost, not LAN
        browser = pychrome.Browser(url="http://127.0.0.1:9223")
        tabs = browser.list_tab()
        return browser, tabs
    except:
        return None, []

def ensure_chrome_debugging():
    """
    Upgraded version:
    Ensures ANY Chromium-based browser (Chrome, Edge, Brave, Opera)
    is running with remote debugging enabled at port 9223.
    """
    import psutil
    import subprocess
    import time
    import os

    # Supported Chromium browsers and their possible paths
    browser_paths = [
        # Chrome
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",

        # Edge
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",

        # Brave
        r"C:\Program Files\BraveSoftware\Brave-Browser\Application\brave.exe",

        # Opera
        fr"C:\Users\{os.getlogin()}\AppData\Local\Programs\Opera\launcher.exe",
    ]

    # 1. Check if ANY browser process is already running WITH debugging enabled
    for proc in psutil.process_iter(['name', 'cmdline']):
        try:
            pname = proc.info['name'] or ""
            cmd = " ".join(proc.info['cmdline']) if proc.info['cmdline'] else ""

            if any(b in pname.lower() for b in ["chrome", "msedge", "brave", "opera"]):
                if "--remote-debugging-port=9223" in cmd:
                    return True  # Already enabled
        except:
            pass

    # 2. If debugging is NOT running, launch the FIRST available browser
    for path in browser_paths:
        if os.path.exists(path):
            print(f"Launching browser with debugging enabled: {path}")
            subprocess.Popen([path, "--remote-debugging-port=9223"])
            time.sleep(2)
            return True

    print("No Chromium browsers found to enable debugging.")
    return False

def get_all_installed_apps():
    """
    Returns a dict mapping normalized app names -> launch target.
    Launch target may be:
      - .lnk shortcut path
      - .exe absolute path
      - install folder (we will search .exe inside)
      - built-in executable name (e.g., 'notepad.exe')
    """
    global _apps_cache, _apps_cache_time

    now = time.time()
    if _apps_cache and (now - _apps_cache_time) < APP_DISCOVERY_CACHE_TTL_SEC:
        return dict(_apps_cache)

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
                _add_app_entry(apps, name, shortcut)
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
                                key = display_name
                                # prefer install location if available
                                if install_loc:
                                    _add_app_entry(apps, key, install_loc)
                                else:
                                    # keep what we have (maybe later Start Menu will override)
                                    norm_key = _normalize_app_key(key)
                                    if key.lower() not in apps:
                                        apps[key.lower()] = ""
                                    if norm_key and norm_key not in apps:
                                        apps[norm_key] = ""
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
                                _add_app_entry(apps, name, os.path.join(folder_path, f))
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
            _add_app_entry(apps, os.path.splitext(exe)[0].lower(), exe_path)
        else:
            # fallback to name only (os.startfile will handle these)
            base = os.path.splitext(exe)[0].lower()
            apps.setdefault(base, exe)

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
                    _add_app_entry(apps, name, exe)
            except Exception:
                pass

    # F. Windows user-level app shims (Store/Desktop bridge apps)
    user_windows_apps = os.path.join(
        os.environ.get("LOCALAPPDATA", ""),
        "Microsoft",
        "WindowsApps",
    )
    if os.path.isdir(user_windows_apps):
        try:
            for exe in glob.glob(os.path.join(user_windows_apps, "*.exe")):
                name = os.path.splitext(os.path.basename(exe))[0].lower()
                if name not in ("python", "python3"):
                    _add_app_entry(apps, name, exe)
        except Exception:
            pass

    # G. Start app catalog from OS (Get-StartApps -> AppUserModelID)
    for entry in _get_start_apps_from_os():
        try:
            name = str(entry.get("Name", "")).strip()
            app_id = str(entry.get("AppID", "")).strip()
            if name and app_id:
                _add_app_entry(apps, name, f"appid::{app_id}")
        except Exception:
            continue

    # Normalize: remove empty install locations for registry-only entries if Start Menu found later
    # (Start Menu/WindowsApps entries added earlier will override)
    _apps_cache = dict(apps)
    _apps_cache_time = now
    return apps


def match_app(user_input, app_dict):
    """
    Return the matched key from app_dict (or None).
    Matching order: exact -> startswith -> contains/tokens -> fuzzy
    """
    if not user_input:
        return None

    name = normalize_app_query(user_input)
    keys = list(app_dict.keys())

    # Exact
    if name in app_dict:
        return name

    # startswith
    for k in keys:
        if k.startswith(name):
            return k

    # contains all query tokens (helps "chrome" -> "google chrome")
    name_tokens = [t for t in name.split() if t]
    if name_tokens:
        for k in keys:
            if all(tok in k for tok in name_tokens):
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
    name = normalize_app_query(app_name)
    if not name:
        return False

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

def _app_alias_terms(query):
    terms = [query]
    for canonical, aliases in APP_ALIAS_GROUPS.items():
        if query == canonical or query in aliases:
            terms.extend(aliases)
            terms.append(canonical)
            break
    unique_terms = []
    for t in terms:
        t = t.strip().lower()
        if t and t not in unique_terms:
            unique_terms.append(t)
    return unique_terms

def _bring_app_window_to_front(query_terms):
    """Best-effort focus/maximize of newly opened app window."""
    try:
        titles = gw.getAllTitles()
        for title in titles:
            t = (title or "").lower()
            if any(term in t for term in query_terms if len(term) > 2):
                wins = gw.getWindowsWithTitle(title)
                if wins:
                    w = wins[0]
                    if w.isMinimized:
                        w.restore()
                    w.activate()
                    try:
                        w.maximize()
                    except Exception:
                        pass
                    return True
    except Exception:
        pass
    return False

def _launch_target(target):
    if not target:
        return False

    if isinstance(target, str) and target.startswith("appid::"):
        app_id = target.split("appid::", 1)[1].strip()
        return _launch_app_user_model_id(app_id)

    if target.lower().endswith(".lnk"):
        os.startfile(target)
        return True

    if target.lower().endswith(".exe") and os.path.exists(target):
        os.startfile(target)
        return True

    if os.path.isdir(target):
        for f in os.listdir(target):
            if f.lower().endswith(".exe"):
                os.startfile(os.path.join(target, f))
                return True
    return False

def open_system_app(app_name):
    """
    Universal app launcher:
    - handles plain 'open <app>' and 'open app <app>' forms
    - resolves aliases (Chrome/Word/PowerPoint, etc.)
    - starts app and attempts to bring the app window to front
    """
    query = normalize_app_query(app_name)
    if not query:
        return "No application name provided."

    try:
        query_terms = _app_alias_terms(query)

        # 1) Direct executable hints for common apps
        for canonical, executables in APP_DIRECT_EXECUTABLES.items():
            if canonical in query_terms:
                for exe in executables:
                    try:
                        os.startfile(exe)
                        time.sleep(1.0)
                        _bring_app_window_to_front(query_terms + [canonical])
                        return f"Opening {canonical}"
                    except Exception:
                        continue

        # 2) Installed app lookup (Start Menu/registry/UWP/system tools)
        apps = get_all_installed_apps()

        match = None
        for term in query_terms:
            match = match_app(term, apps)
            if match:
                break

        if not match:
            for key in apps.keys():
                k = key.lower()
                if any(term in k for term in query_terms):
                    match = key
                    break

        if not match:
            return f"I could not find an application named '{query}'."

        target = apps.get(match)
        if _launch_target(target):
            time.sleep(1.0)
            _bring_app_window_to_front(query_terms + [match.lower()])
            return f"Opening {match}"

        # 3) Generic fallback starts
        for term in [match] + query_terms:
            try:
                os.startfile(term)
                time.sleep(1.0)
                _bring_app_window_to_front(query_terms + [match.lower()])
                return f"Opening {match}"
            except Exception:
                continue

        for term in [match] + query_terms:
            try:
                subprocess.Popen([term], shell=True)
                time.sleep(1.0)
                _bring_app_window_to_front(query_terms + [match.lower()])
                return f"Opening {match}"
            except Exception:
                continue

        return f"I found {match}, but could not start it."

    except Exception as e:
        return f"Error launching {query}: {e}"


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

import pytesseract
from pytesseract import Output
from PIL import ImageGrab
import pyautogui
import time
from difflib import SequenceMatcher

last_tab_boxes = []
last_tab_target = None
# ----------------- RANGE SELECTION STATE (Independent) -----------------
range_mode = None              # "select" or "copy" or None
range_phase = None             # "start" or "end" or None
range_start_phrase = ""
range_end_phrase = ""

range_start_occurrences = []   # OCR boxes for the start phrase
range_end_occurrences = []     # OCR boxes for the end phrase

range_selected_start = None    # selected start OCR box (x, y, w, h)


def chrome_switch_tab_by_name(name):
    global last_chrome_tab_id, current_chrome_tab_id
    global last_tab_boxes, last_tab_target

    name = name.lower().strip()

    # Capture only top area where tabs live
    screen = ImageGrab.grab()
    w, h = screen.size
    tab_region = screen.crop((0, 0, w, 120))

    data = pytesseract.image_to_data(tab_region, output_type=Output.DICT)

    matches = []  # list of (x, y, w, h)

    for i, txt in enumerate(data["text"]):
        if not txt.strip():
            continue

        tab_text = txt.lower().strip()
        score = SequenceMatcher(None, name, tab_text).ratio()

        # fuzzy or direct match
        if name in tab_text or score > 0.55:
            x = data["left"][i]
            y = data["top"][i]
            w_box = data["width"][i]
            h_box = data["height"][i]
            matches.append((x, y, w_box, h_box))

    # No matches
    if not matches:
        reply(f"I couldn't find a tab named {name}.")
        return False

    # MULTIPLE matches → show options
    if len(matches) > 1:
        last_tab_boxes = matches
        last_tab_target = name
        show_numbered_boxes(matches)
        reply("I found multiple matching tabs. Say 'choose 1', 'choose 2', etc.")
        return True

    # Exactly one match → click it immediately
    x, y, w_box, h_box = matches[0]
    cx = x + w_box // 2
    cy = y + h_box // 2

    pyautogui.moveTo(cx, cy, duration=0.2)
    pyautogui.click()

    last_chrome_tab_id = current_chrome_tab_id
    current_chrome_tab_id = name

    reply(f"Switched to tab: {name}")
    return True

def chrome_return_to_previous_tab():
    """
    Return to previously clicked tab (OCR-based).
    """
    global last_chrome_tab_id

    if not last_chrome_tab_id:
        reply("No previous tab recorded.")
        return False

    return chrome_switch_tab_by_name(last_chrome_tab_id)
overlay_window = None
overlay_labels = []
import tkinter as tk

def show_on_screen_boxes(boxes):
    global overlay_window, overlay_labels

    # Close old overlay if exists
    if overlay_window:
        overlay_window.destroy()

    overlay_window = tk.Tk()
    overlay_window.attributes("-topmost", True)
    overlay_window.attributes("-transparentcolor", "white")
    overlay_window.overrideredirect(True)

    screen_w = overlay_window.winfo_screenwidth()
    screen_h = overlay_window.winfo_screenheight()

    overlay_window.geometry(f"{screen_w}x{screen_h}+0+0")
    overlay_window.config(bg="white")

    overlay_labels = []

    for i, (x, y, w, h) in enumerate(boxes):
        label = tk.Label(
            overlay_window,
            text=str(i+1),
            bg="yellow",
            fg="black",
            font=("Arial", 14, "bold")
        )
        label.place(x=x + w//2, y=y - 20)  # draw above the box
        overlay_labels.append(label)

    overlay_window.update()

def clear_on_screen_boxes():
    global overlay_window
    if overlay_window:
        overlay_window.destroy()
        overlay_window = None


def show_numbered_boxes(boxes):
    """
    Provides visual feedback for multiple OCR matches.
    Shows list in console and voice.
    """
    for i, (x, y, w, h) in enumerate(boxes):
        print(f"Option {i+1}: ({x}, {y}, {w}, {h})")

    show_on_screen_boxes(boxes)
    reply("I found multiple matches. Say 'choose 1', 'choose 2', etc.")


def find_text_boxes(target, fuzzy_threshold=0.65):
    """
    Full-screen OCR with improved accuracy and Chrome-specific fallback.
    Keeps your original structure and parameters.
    Behavior:
      - Normalizes spoken numbers
      - Word-level matching with OCR-error corrections
      - Cluster splitting for horizontal groups
      - Chrome-only high-contrast / sharpen fallback when first-pass fails
      - Returns list of boxes as (x, y, w, h) in screen coordinates
    """
    from difflib import SequenceMatcher
    import re
    import ctypes
    import psutil
    from PIL import ImageEnhance, ImageOps

    # ----------------------------
    # Normalize spoken numbers
    # ----------------------------
    num_words = {
        "zero":"0","one":"1","two":"2","three":"3","four":"4","five":"5",
        "six":"6","seven":"7","eight":"8","nine":"9","ten":"10",
    }
    if not isinstance(target, str):
        return []
    parts = target.lower().strip().split()
    target = " ".join([num_words.get(w, w) for w in parts]).strip()
    if not target:
        return []

    # Short UI labels stricter
    if len(target) <= 5:
        fuzzy_threshold = max(fuzzy_threshold, 0.85)

    # ----------------------------
    # Detect active window process (to enable Chrome mode)
    # ----------------------------
    chrome_mode = False
    try:
        # Get foreground window handle
        user32 = ctypes.windll.user32
        hwnd = user32.GetForegroundWindow()
        pid = ctypes.c_ulong()
        ctypes.windll.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
        proc = psutil.Process(pid.value)
        proc_name = proc.name().lower()
        # enable chrome mode only for chrome/edge/brave
        if any(x in proc_name for x in ("chrome", "msedge", "brave", "opera")):
            chrome_mode = True
        # ensure PowerPoint unaffected
        if "powerpnt" in proc_name or "powerpoint" in proc_name:
            chrome_mode = False
    except Exception:
        chrome_mode = False

    # ----------------------------
    # helper: OCR normalization corrections
    # ----------------------------
    def fix_ocr_token(w):
        """Common char confusions and some chrome-specific patterns."""
        s = w.replace("|", "i").replace("l", "i").replace("1", "i")
        s = s.replace("0", "o").replace("5", "s").replace("$", "s")
        # chrome-specific common misreads
        s = s.replace("rn", "m").replace("vv", "w").replace("cl", "d")
        s = s.replace("¬", "t")
        return s

    # ----------------------------
    # 1) Capture screen (first pass)
    # ----------------------------
    img = ImageGrab.grab()
    data = pytesseract.image_to_data(img, output_type=pytesseract.Output.DICT)

    # quick guard
    n = len(data.get("text", []))
    if n == 0:
        return []

    # ----------------------------
    # 2) Word-level matching (high priority)
    # ----------------------------
    word_hits = []
    for i, w in enumerate(data["text"]):
        if not w.strip():
            continue
        w_raw = w.lower().strip()
        w_fixed = fix_ocr_token(w_raw)

        x = data["left"][i]
        y = data["top"][i]
        wid = data["width"][i]
        h = data["height"][i]

        # direct exact / substring
        if target == w_raw or target in w_raw or w_raw in target:
            word_hits.append((x, y, wid, h)); continue

        # fuzzy on raw
        if SequenceMatcher(None, target, w_raw).ratio() >= fuzzy_threshold:
            word_hits.append((x, y, wid, h)); continue

        # fuzzy on corrected token
        if SequenceMatcher(None, target, w_fixed).ratio() >= fuzzy_threshold:
            word_hits.append((x, y, wid, h)); continue

    # If multi-token target and many token hits, return them to let click logic choose
    target_tokens = target.split()
    if len(target_tokens) > 1 and len(word_hits) >= max(1, len(target_tokens)//2):
        return word_hits

    # For single-word target, prefer word_hits immediately
    if len(target_tokens) == 1 and word_hits:
        return word_hits

    # ----------------------------
    # 3) Line grouping (preserve original structure)
    # ----------------------------
    lines = {}
    for i, w in enumerate(data["text"]):
        if not w.strip():
            continue
        block = data["block_num"][i]
        line = data["line_num"][i]
        key = (block, line)
        x = data["left"][i]
        y = data["top"][i]
        wid = data["width"][i]
        h = data["height"][i]

        if key not in lines:
            lines[key] = {"words": [], "boxes": []}
        lines[key]["words"].append(w.lower())
        lines[key]["boxes"].append((x, y, wid, h))

    matches = []

    # ----------------------------
    # 4) Cluster splitting + matching
    # ----------------------------
    for (block, line), info in lines.items():
        entries = list(zip(info["words"], info["boxes"]))
        entries.sort(key=lambda e: e[1][0])  # left->right

        # split clusters by horizontal gap
        clusters = []
        if entries:
            cur = [entries[0]]
            for i in range(1, len(entries)):
                prev_x, prev_w = entries[i-1][1][0], entries[i-1][1][2]
                curr_x = entries[i][1][0]
                gap = curr_x - (prev_x + prev_w)
                if gap < 50:
                    cur.append(entries[i])
                else:
                    clusters.append(cur)
                    cur = [entries[i]]
            clusters.append(cur)

        for cluster in clusters:
            cl_text = " ".join([c[0] for c in cluster])
            xs = [b[0] for _, b in cluster]
            ys = [b[1] for _, b in cluster]
            ws = [b[2] for _, b in cluster]
            hs = [b[3] for _, b in cluster]
            x1 = min(xs); y1 = min(ys)
            x2 = max(xs[i] + ws[i] for i in range(len(ws)))
            y2 = max(ys[i] + hs[i] for i in range(len(hs)))
            box = (x1, y1, x2 - x1, y2 - y1)

            # direct phrase substring
            if target in cl_text:
                matches.append(box); continue

            # fuzzy phrase match
            if SequenceMatcher(None, target, cl_text).ratio() >= fuzzy_threshold:
                matches.append(box); continue

            # token-level majority
            token_hits = sum(1 for t in target_tokens if t in cl_text)
            if token_hits >= max(1, len(target_tokens)//2):
                matches.append(box); continue

    # If matches found on first pass, return
    if matches:
        return matches

    # ----------------------------
    # 5) CHROME-SPECIFIC SECOND PASS (only when chrome_mode)
    # ----------------------------
    if chrome_mode:
        try:
            # produce enhanced image: grayscale -> autocontrast -> enhance contrast/sharpness
            enh = img.convert("L")
            enh = ImageOps.autocontrast(enh)
            enh = ImageEnhance.Contrast(enh).enhance(1.8)
            enh = ImageEnhance.Sharpness(enh).enhance(2.0)
            # re-run OCR on enhanced image
            data2 = pytesseract.image_to_data(enh, output_type=pytesseract.Output.DICT)
            n2 = len(data2.get("text", []))
            if n2 == 0:
                return []
            # repeat word-level and cluster matching but with extra token corrections
            word_hits2 = []
            for i in range(n2):
                raw = (data2["text"][i] or "").strip()
                if not raw:
                    continue
                wr = raw.lower()
                # extra chrome fixes
                wf = fix_ocr_token(wr)
                # also a cleaned variant removing spaces inside tokens (case: "govern ance")
                wf2 = re.sub(r"\s+", "", wr)

                x = data2["left"][i]
                y = data2["top"][i]
                wid = data2["width"][i]
                h = data2["height"][i]

                if target == wr or target in wr or wr in target:
                    word_hits2.append((x, y, wid, h)); continue
                if SequenceMatcher(None, target, wr).ratio() >= fuzzy_threshold:
                    word_hits2.append((x, y, wid, h)); continue
                if SequenceMatcher(None, target, wf).ratio() >= fuzzy_threshold:
                    word_hits2.append((x, y, wid, h)); continue
                if SequenceMatcher(None, target, wf2).ratio() >= fuzzy_threshold:
                    word_hits2.append((x, y, wid, h)); continue

            if word_hits2:
                return word_hits2

            # cluster pass for enhanced OCR data
            lines2 = {}
            for i in range(n2):
                raw = (data2["text"][i] or "").strip()
                if not raw:
                    continue
                blk = data2["block_num"][i]
                ln = data2["line_num"][i]
                key = (blk, ln)
                x = data2["left"][i]
                y = data2["top"][i]
                wid = data2["width"][i]
                h = data2["height"][i]
                if key not in lines2:
                    lines2[key] = {"words": [], "boxes": []}
                lines2[key]["words"].append(raw.lower())
                lines2[key]["boxes"].append((x, y, wid, h))

            matches2 = []
            for key, info in lines2.items():
                entries = list(zip(info["words"], info["boxes"]))
                entries.sort(key=lambda e: e[1][0])
                if not entries:
                    continue
                # cluster split
                clusters2 = []
                cur = [entries[0]]
                for i in range(1, len(entries)):
                    prev_x, prev_w = entries[i-1][1][0], entries[i-1][1][2]
                    curr_x = entries[i][1][0]
                    gap = curr_x - (prev_x + prev_w)
                    if gap < 50:
                        cur.append(entries[i])
                    else:
                        clusters2.append(cur)
                        cur = [entries[i]]
                clusters2.append(cur)

                for cluster in clusters2:
                    text = " ".join([c[0] for c in cluster])
                    xs = [b[0] for _, b in cluster]
                    ys = [b[1] for _, b in cluster]
                    ws = [b[2] for _, b in cluster]
                    hs = [b[3] for _, b in cluster]
                    x1 = min(xs); y1 = min(ys)
                    x2 = max(xs[i] + ws[i] for i in range(len(ws)))
                    y2 = max(ys[i] + hs[i] for i in range(len(hs)))
                    box = (x1, y1, x2 - x1, y2 - y1)

                    if target in text:
                        matches2.append(box); continue
                    if SequenceMatcher(None, target, text).ratio() >= fuzzy_threshold:
                        matches2.append(box); continue
                    # try corrected token comparison
                    tclean = re.sub(r"\s+", "", text)
                    if SequenceMatcher(None, target.replace(" ", ""), tclean).ratio() >= 0.85:
                        matches2.append(box); continue

            if matches2:
                return matches2

        except Exception:
            # if anything breaks in chrome fallback, silently ignore and return []
            pass

    # end chrome fallbackhan
    return matches

def move_caret_after_phrase(phrase: str):
    """
    Move text cursor (caret) to just AFTER the given phrase using OCR.
    """
    phrase = phrase.strip().lower()
    if not phrase:
        reply("After which word should I paste?")
        return False

    boxes = find_text_boxes(phrase)
    if not boxes:
        reply(f"I couldn't find {phrase} on the screen.")
        return False

    x, y, w, h = boxes[0]

    # click AFTER the word
    click_x = x + w + 5
    click_y = y + h // 2

    try:
        pyautogui.click(click_x, click_y)
        time.sleep(0.1)
        return True
    except Exception as e:
        print(f"move_caret_after_phrase error: {e}")
        reply("I couldn't move the cursor after that word.")
        return False


def move_caret_before_phrase(phrase: str):
    """
    Move caret to just BEFORE the given phrase using OCR.
    """
    phrase = phrase.strip().lower()
    if not phrase:
        reply("Before which word should I paste?")
        return False

    boxes = find_text_boxes(phrase)
    if not boxes:
        reply(f"I couldn't find {phrase} on the screen.")
        return False

    x, y, w, h = boxes[0]

    # Click slightly to the LEFT of the word so caret is before it
    click_x = max(x - 5, 0)
    click_y = y + h // 2

    try:
        pyautogui.click(click_x, click_y)
        time.sleep(0.1)
        return True
    except Exception as e:
        print(f"move_caret_before_phrase error: {e}")
        reply("I couldn't move the cursor before that word.")
        return False


def move_caret_to_line_start_for_phrase(phrase: str):
    """
    Move caret to the START of the line containing the phrase.
    """
    phrase = phrase.strip().lower()
    if not phrase:
        reply("Which line should I paste at the start of?")
        return False

    boxes = find_text_boxes(phrase)
    if not boxes:
        reply(f"I couldn't find {phrase} on the screen.")
        return False

    x, y, w, h = boxes[0]

    # Far left on the same line
    click_x = max(x - 200, 0)
    click_y = y + h // 2

    try:
        pyautogui.click(click_x, click_y)
        time.sleep(0.1)
        return True
    except Exception as e:
        print(f"move_caret_to_line_start_for_phrase error: {e}")
        reply("I couldn't move the cursor to the start of that line.")
        return False


def move_caret_to_line_end_for_phrase(phrase: str):
    """
    Move caret to the END of the line / paragraph containing the phrase.
    """
    phrase = phrase.strip().lower()
    if not phrase:
        reply("Which line or paragraph end should I paste at?")
        return False

    boxes = find_text_boxes(phrase)
    if not boxes:
        reply(f"I couldn't find {phrase} on the screen.")
        return False

    x, y, w, h = boxes[0]

    # Far right after the phrase on the same line
    click_x = x + w + 200
    click_y = y + h // 2

    try:
        pyautogui.click(click_x, click_y)
        time.sleep(0.1)
        return True
    except Exception as e:
        print(f"move_caret_to_line_end_for_phrase error: {e}")
        reply("I couldn't move the cursor to the end of that line.")
        return False


def move_to_excel_cell(cell_ref: str):
    """
    Move selection to a specific Excel cell (e.g., A5, B10) using Go To (Ctrl+G).
    Assumes Excel is focused.
    """
    cell_ref = cell_ref.strip().upper()
    if not cell_ref:
        reply("Which cell should I paste into?")
        return False

    try:
        # Open 'Go To' dialog
        with keyboard.pressed(Key.ctrl):
            keyboard.press('g')
            keyboard.release('g')

        time.sleep(0.2)
        pyautogui.typewrite(cell_ref)
        pyautogui.press('enter')
        time.sleep(0.2)
        return True
    except Exception as e:
        print(f"move_to_excel_cell error: {e}")
        reply("I couldn't jump to that cell. Make sure Excel is active.")
        return False

def apply_range_selection(start_box, end_box, mode, start_phrase, end_phrase):
    """
    Given two OCR boxes and a mode:
      - 'select' → just select range
      - 'copy'   → select range + Ctrl+C
    """
    global range_mode, range_phase
    global range_start_occurrences, range_end_occurrences
    global range_start_phrase, range_end_phrase, range_selected_start

    try:
        sx, sy, sw, sh = start_box
        ex, ey, ew, eh = end_box

        # 0) Small safety pause
        time.sleep(0.1)

        # 1) Click at start (this also focuses the underlying window)
        # Click slightly to the *left* of the OCR box so caret goes to the start of the word
        start_x = max(sx - 5, 0)
        start_y = sy + sh // 2

        pyautogui.click(start_x, start_y)
        time.sleep(0.15)

        # 2) Shift+click at end to extend selection
        end_x = ex + ew - 3
        end_y = ey + eh // 2
        pyautogui.keyDown('shift')
        pyautogui.click(end_x, end_y)
        pyautogui.keyUp('shift')
        time.sleep(0.15)

        # 3) Optional copy
        if mode == 'copy':
            with keyboard.pressed(Key.ctrl):
                keyboard.press('c')
                keyboard.release('c')
            reply(f"Copied text from {start_phrase} to {end_phrase}.")
        else:
            reply(f"Selected text from {start_phrase} to {end_phrase}.")

    except Exception as e:
        print(f"apply_range_selection error: {e}")
        reply("I couldn't select that text range.")

    # Reset range state
    range_mode = None
    range_phase = None
    range_start_occurrences = []
    range_end_occurrences = []
    range_start_phrase = ""
    range_end_phrase = ""
    range_selected_start = None
    clear_on_screen_boxes()

def start_range_selection(full_cmd: str, mode: str):
    """
    Common logic for:
      - 'select from X to Y'  (mode='select')
      - 'copy from X to Y'    (mode='copy')

    If multiple matches for start or end, it will:
      - show numbered boxes,
      - then wait for 'choose N' (handled by handle_range_choice).
    """
    global range_mode, range_phase
    global range_start_phrase, range_end_phrase
    global range_start_occurrences, range_end_occurrences
    global range_selected_start

    text = full_cmd.lower()
    m = re.search(r'from (.+?) to (.+)', text)
    if not m:
        if mode == 'copy':
            reply("Please say something like: copy from <start> to <end>.")
        else:
            reply("Please say something like: select from <start> to <end>.")
        return

    start_phrase = m.group(1).strip()
    end_phrase = m.group(2).strip()

    if not start_phrase or not end_phrase:
        reply("I need both a start and an end phrase.")
        return

    # OCR search
    start_boxes = find_text_boxes(start_phrase)
    if not start_boxes:
        reply(f"I couldn't find the phrase {start_phrase} on the screen.")
        return

    end_boxes = find_text_boxes(end_phrase)
    if not end_boxes:
        reply(f"I couldn't find the phrase {end_phrase} on the screen.")
        return

    # Reset/initialize state
    range_mode = mode
    range_phase = None
    range_start_phrase = start_phrase
    range_end_phrase = end_phrase
    range_start_occurrences = start_boxes
    range_end_occurrences = end_boxes
    range_selected_start = None

    # Case 1: unique start & unique end → execute immediately
    if len(start_boxes) == 1 and len(end_boxes) == 1:
        apply_range_selection(start_boxes[0], end_boxes[0], mode, start_phrase, end_phrase)
        return

    # Case 2: ambiguous start (multiple occurrences)
    if len(start_boxes) > 1:
        range_phase = "start"
        show_numbered_boxes(start_boxes)
        reply("I found multiple matches for the start phrase. Say 'choose 1', 'choose 2', etc.")
        return

    # Case 3: unique start, ambiguous end
    range_selected_start = start_boxes[0]
    range_phase = "end"
    show_numbered_boxes(end_boxes)
    reply("I found multiple matches for the end phrase. Say 'choose 1', 'choose 2', etc.")

def select_text_range_by_phrases(start_phrase: str, end_phrase: str):
    """
    Use OCR to select text on screen from start_phrase up to end_phrase.
    Works in editors/browsers where text is visually rendered.

    Returns True if selection likely succeeded, False otherwise.
    """
    if not start_phrase or not end_phrase:
        reply("I need both a start and an end phrase.")
        return False

    start_phrase = start_phrase.strip().lower()
    end_phrase = end_phrase.strip().lower()

    # 1) Find start phrase boxes
    start_boxes = find_text_boxes(start_phrase)
    if not start_boxes:
        reply(f"I couldn't find the phrase {start_phrase} on the screen.")
        return False

    # 2) Find end phrase boxes
    end_boxes = find_text_boxes(end_phrase)
    if not end_boxes:
        reply(f"I couldn't find the phrase {end_phrase} on the screen.")
        return False

    # For now: use the first occurrence of each
    sx, sy, sw, sh = start_boxes[0]
    ex, ey, ew, eh = end_boxes[0]

    # 3) Click at the start phrase (left side, vertically centered)
    start_x = sx + 3
    start_y = sy + sh // 2

    # 4) Click end phrase while holding Shift, so we select everything
    #    between them (direction doesn't really matter; selection will be from anchor to this click)
    end_x = ex + ew - 3
    end_y = ey + eh // 2

    try:
        # First click: set caret at start
        pyautogui.click(start_x, start_y)
        time.sleep(0.1)

        # Second click with Shift held: extend selection to end
        pyautogui.keyDown('shift')
        pyautogui.click(end_x, end_y)
        pyautogui.keyUp('shift')

        return True
    except Exception as e:
        print(f"select_text_range_by_phrases error: {e}")
        reply("I couldn't select that text range.")
        return False


def capture_screen_region(x, y, width, height):
    """Capture a screen region and return a PIL image."""
    left = max(int(x), 0)
    top = max(int(y), 0)
    right = max(left + int(width), left + 1)
    bottom = max(top + int(height), top + 1)
    return ImageGrab.grab(bbox=(left, top, right, bottom))


def preprocess_image(img):
    """Light OCR-oriented preprocessing that preserves text edges."""
    if img is None:
        return img

    np_img = np.array(img)
    if len(np_img.shape) == 3:
        gray = cv2.cvtColor(np_img, cv2.COLOR_BGR2GRAY)
    else:
        gray = np_img

    denoised = cv2.bilateralFilter(gray, 7, 60, 60)
    enhanced = cv2.convertScaleAbs(denoised, alpha=1.2, beta=6)
    return enhanced

def get_all_ocr_boxes():
    """
    Use the SAME OCR pipeline we already use for find_text_boxes(),
    but return ALL detected word boxes (no text matching).
    """
    try:
        screen_width, screen_height = pyautogui.size()

        # 1) Capture full screen using your accurate capturer
        img = capture_screen_region(0, 0, screen_width, screen_height)

        # 2) Preprocess using your pipeline for accuracy
        processed = preprocess_image(img)

        # 3) Run OCR with your existing settings
        data = pytesseract.image_to_data(processed, output_type=Output.DICT)

        boxes = []
        n = len(data.get("text", []))

        for i in range(n):
            txt = data["text"][i]
            conf = int(data["conf"][i]) if data["conf"][i].isdigit() else -1

            # skip empty or low-confidence words
            if not txt or not txt.strip():
                continue
            if conf < 30:   # your normal threshold
                continue

            x = data["left"][i]
            y = data["top"][i]
            w = data["width"][i]
            h = data["height"][i]

            # skip tiny words/noise
            if w < 10 or h < 10:
                continue

            boxes.append((x, y, w, h))

        return boxes

    except Exception as e:
        print("get_all_ocr_boxes error:", e)
        return []

def start_paste_position_selection():
    """
    OCR the screen, show numbered boxes as possible paste positions,
    and wait for the user to say 'choose 1', 'choose 2', etc.
    """
    global paste_position_mode, paste_position_boxes

    boxes = get_all_ocr_boxes()
    if not boxes:
        reply("I couldn't find any paste positions on the screen.")
        return

    paste_position_boxes = boxes
    paste_position_mode = "active"

    # Reuse existing overlay system
    show_numbered_boxes(paste_position_boxes)
    reply("I found multiple paste positions. Say 'choose 1', 'choose 2', etc.")


def looks_like_file_name(text):
    """Detect if target appears to be a file or folder."""
    file_exts = [
        ".pdf", ".docx", ".xlsx", ".xls", ".pptx",
        ".png", ".jpg", ".jpeg", ".zip", ".rar",
        ".txt", ".py", ".csv", ".mp4", ".mp3"
    ]
    t = text.lower()
    if any(ext in t for ext in file_exts):
        return True
    
    # folder-like (no spaces, no dots)
    if " " not in t and "." not in t and len(t) <= 20:
        return True
    
    return False

def parse_choice_number(text):
    text = text.lower().strip()

    number_words = {
        "one":1, "first":1, "1":1, "1st":1,
        "two":2, "second":2, "2":2, "2nd":2,
        "three":3, "third":3, "3":3, "3rd":3,
        "four":4, "fourth":4, "4":4, "4th":4,
        "five":5, "fifth":5, "5":5, "5th":5
    }

    for word, num in number_words.items():
        if word in text:
            return num

    # fallback: digit detection
    digits = ''.join([ch for ch in text if ch.isdigit()])
    if digits.isdigit():
        return int(digits)

    return None


def perform_click_action(entity, is_double=False, full_cmd=""):
    """
    FINAL version:
     - Multi-match disambiguation
     - File/folder auto double-click
     - nth occurrence support
     - URL/Domain fallback
     - Region targeting (top left, bottom right)
     - Avoid browser tabs unless explicitly allowed
     - Uses your full find_text_boxes() engine
    """
    import re
    global last_found_boxes

    target = entity.lower().strip()

    # normalize spoken numbers
    num_map = {
        "zero":"0","one":"1","two":"2","three":"3","four":"4","five":"5",
        "six":"6","seven":"7","eight":"8","nine":"9","ten":"10"
    }
    target = " ".join([num_map.get(w, w) for w in target.split()]).strip()

    reply(f"Searching for {target}...")

    # 1. OCR for initial matches
    boxes = find_text_boxes(target)

    # nth occurrence detection
    index = 0
    mapping = {
        "first":0,"1st":0,"second":1,"2nd":1,"third":2,"3rd":2,
        "fourth":3,"4th":3,"fifth":4,"5th":4
    }
    for w in full_cmd.lower().split():
        if w in mapping:
            index = mapping[w]

    # URL-like handling
    is_url_like = False
    url_candidate = None
    if re.match(r"(https?://)|www\.", entity.lower()) or (("." in entity) and (" " not in entity.strip())):
        is_url_like = True
        url_candidate = entity.lower().strip()
        dom = re.sub(r"^https?://", "", url_candidate)
        dom = dom.split("/")[0]
        dom_token = dom.replace("www.", "").split(".")[0]
    else:
        dom_token = None

    # 2. scroll + retry
    if not boxes:
        boxes = scroll_and_find(target)

    # domain-token fallback
    if not boxes and is_url_like and dom_token:
        boxes = find_text_boxes(dom_token)

    # link fallback
    if not boxes and ("link" in full_cmd.lower() or "open link" in full_cmd.lower()) and dom_token:
        boxes = find_text_boxes(dom_token)

    if not boxes:
        reply(f"I could not find {entity} on the screen.")
        return


    # --------------------------------------------------
    # MULTIPLE MATCHES → list them & wait for user choice
    # --------------------------------------------------
    if len(boxes) > 1:
        global last_click_target_text
        last_click_target_text = entity   # remember what the user wanted to click
        last_found_boxes = boxes
        show_numbered_boxes(boxes)
        return



    # --------------------------------------------------
    # 3. REGION TARGETING (top left, top right, etc.)
    # --------------------------------------------------
    cmd = full_cmd.lower()

    # “top right”
    if "top right" in cmd:
        boxes = sorted(boxes, key=lambda b: (b[1], -b[0]))

    elif "top left" in cmd:
        boxes = sorted(boxes, key=lambda b: (b[1], b[0]))

    elif "bottom right" in cmd:
        boxes = sorted(boxes, key=lambda b: (-b[1], -b[0]))

    elif "bottom left" in cmd:
        boxes = sorted(boxes, key=lambda b: (-b[1], b[0]))


    # --------------------------------------------------
    # AVOID CLICKING TOP BROWSER BAR (TABS)
    # unless user said “top”
    # --------------------------------------------------
    if "top" not in cmd:
        filtered = [b for b in boxes if b[1] > 80]   # ignore any box within top 80px
        if filtered:
            boxes = filtered


    # --------------------------------------------------
    # ICON MODE (unchanged)
    # --------------------------------------------------
    if "icon" in cmd:
        icon_name = entity + ".png"
        icon_path = os.path.join("icons", icon_name)
        box = find_icon_on_screen(icon_path)
        if box:
            x, y, w, h = box
            pyautogui.moveTo(x + w//2, y + h//2, duration=0.2)
            pyautogui.doubleClick() if is_double else pyautogui.click()
            reply("Icon clicked.")
            return
        reply("Icon not found.")
        return


    # --------------------------------------------------
    # If STILL no boxes (edge-case)
    # --------------------------------------------------
    if not boxes:
        reply(f"I could not find {entity} on the screen.")
        return


    # --------------------------------------------------
    # 4. FILE/FOLDER → AUTO DOUBLE CLICK
    # --------------------------------------------------
    x, y, w, h = boxes[index]

    if looks_like_file_name(entity):
        pyautogui.moveTo(x + w//2, y + h//2, duration=0.25)
        pyautogui.doubleClick()
        reply("Opened.")
        return


    # --------------------------------------------------
    # 5. NORMAL CLICK (UI elements)
    # --------------------------------------------------
    pyautogui.moveTo(x + w//2, y + h//2, duration=0.25)
    pyautogui.doubleClick() if is_double else pyautogui.click()
    reply("Done.")


def scroll_and_find(target, max_scrolls=20):
    """Scroll the screen and search for text."""
    for _ in range(max_scrolls):
        boxes = find_text_boxes(target)
        if boxes:
            return boxes
        pyautogui.scroll(-800)  # scroll down
        time.sleep(0.3)
    return []

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
        self.exact_phrase_to_intent = {}
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
                    'copy from',
                    'copy text from',
                    'copy everything from',
                    'copy data from',
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
                'continue presentation','highlight','scroll', 'find', 'search on page'
            ],
            'open_app': ['open app', 'launch app', 'start app', 'open application', 'run', 'open program', 'start program'],
            'click_action': ['click', 'click on', 'double click', 'select', 'press', 'open this','tap',
                'choose', 'scroll and click', 'click the second', 'click the third'
            ],
            'tab_control': [
                'switch tab', 'switch to tab', 'go to tab', 'open tab',
                'change tab', 'previous tab', 'next tab', 'back to tab'
            ],
            'range_select': [
                'select from', 
                'select text from',
                'select data from',
                'highlight from',
                'highlight text from',
                'select everything from'
            ],
            'paste_after': [
                'paste after',
                'paste after this',
                'paste after the word',
                'insert after',
                'insert text after',
                'paste it after',
                'paste immediately after',
                'paste right after'
            ],

            'paste_before': [
                'paste before',
                'paste before this',
                'paste before the word',
                'insert before',
                'place text before',
                'paste it before',
                'paste right before'
            ],

            'paste_line_start': [
                'paste at start of line',
                'paste at the start of the line',
                'insert at line start',
                'place text at start of line',
                'paste before the line starts'
            ],

            'paste_line_end': [
                'paste at end of line',
                'paste at the end of the line',
                'insert at line end',
                'place text at end of line',
                'paste at end of paragraph',
                'insert at end of paragraph'
            ],

            'paste_cell': [
                'paste in cell',
                'paste into cell',
                'paste to cell',
                'insert in cell',
                'insert into cell',
                'paste inside cell'
            ],



        }
        return training_data
    
    def _train_model(self):
        """Train the intent classification model"""
        texts = []
        labels = []
        
        for intent, examples in self.training_data.items():
            for example in examples:
                normalized_example = normalize_voice_text(example)
                texts.append(normalized_example)
                labels.append(intent)
                self.exact_phrase_to_intent[normalized_example] = intent
        
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

            # Prefer exact command phrase match when available
            if processed_text in self.exact_phrase_to_intent:
                return self.exact_phrase_to_intent[processed_text]
            
            # Predict using the model
            prediction = self.pipeline.predict([processed_text])[0]
            probability = np.max(self.pipeline.predict_proba([processed_text]))
            
            # Only return prediction if confidence is high enough
            if probability > 0.45:
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
        text = normalize_voice_text(text)
        
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
            'file_navigate': ['back', 'previous', 'go back'],
            'goodbye': ['bye', 'goodbye', 'exit', 'quit'],
            'copy': ['copy', 'copy this', 'copy text', 'copy from',],
            'paste': ['paste', 'paste it'],
            'wake_up': ['wake up', 'start', 'activate'],
            'presentation_control': [
            'next slide', 'previous slide', 'go back a slide', 'go to next slide',
            'show next', 'show previous', 'start presentation', 'begin slideshow',
            'end presentation', 'stop presentation', 'pause presentation',
            'resume presentation', 'exit slideshow', 'zoom in', 'zoom out',
            'increase zoom', 'decrease zoom', 'make it bigger', 'make it smaller',
            'full screen', 'exit full screen', 'first slide', 'last slide',
            'skip to slide', 'show slide number','highlight','scroll', 'find', 'search on page'],
            'open_app': ['open app', 'launch', 'start', 'run'],
            'click_action': ['click', 'click on', 'double click', 'tap', 'press',
            'select', 'choose'],
            'tab_control': ['switch tab', 'switch to', 'open tab', 'go to tab','change tab', 'next tab', 'previous tab', 'back to tab'],
            'choice_select': ['choose', 'option', 'select number', 'number'],
            'range_select': ['select from', 'select text from', 'highlight from'],
            'paste_after': ['paste after', 'insert after', 'paste right after', 'paste immediately after'],
            'paste_before': ['paste before', 'insert before', 'paste right before'],
            'paste_line_start': [ 'start of line', 'paste at start', 'line start'],
            'paste_line_end': ['end of line', 'end of paragraph', 'paste at end', 'line end' ],
            'paste_cell': [ 'cell', 'paste in cell', 'paste into cell', 'insert in cell'],

        }
        
        import re

        for intent, keywords in intent_keywords.items():
            for keyword in keywords:
                # Match whole words only
                pattern = rf'\b{re.escape(keyword)}\b'
                if re.search(pattern, text_lower):
                    return intent

        return 'unknown'

# Initialize intent recognizer
intent_recognizer = IntentRecognizer()

COMMAND_LIBRARY = [
    normalize_voice_text(example)
    for examples in intent_recognizer.training_data.values()
    for example in examples
]

def _score_command_candidate(candidate):
    """
    Score transcript candidates so command-like phrases rank higher.
    Combines STT confidence and fuzzy similarity against known commands.
    """
    transcript = normalize_voice_text(candidate.get("transcript", ""))
    if not transcript:
        return -1.0, ""

    stt_confidence = float(candidate.get("confidence", 0.0))
    similarity = 0.0
    if COMMAND_LIBRARY:
        similarity = max(SequenceMatcher(None, transcript, known).ratio() for known in COMMAND_LIBRARY)

    score = (0.65 * similarity) + (0.35 * stt_confidence)
    return score, transcript

def _best_transcript_from_google(audio):
    """
    Request multiple STT alternatives and choose the candidate that best matches
    known assistant command phrases.
    """
    alt_data = r.recognize_google(audio, show_all=True)
    if isinstance(alt_data, dict):
        alternatives = alt_data.get("alternative", [])
        if alternatives:
            ranked = [_score_command_candidate(candidate) for candidate in alternatives]
            ranked = [item for item in ranked if item[1]]
            if ranked:
                ranked.sort(key=lambda item: item[0], reverse=True)
                return ranked[0][1]
    return normalize_voice_text(r.recognize_google(audio))




def is_double_click_command(text):
    return "double" in text.lower()


# ------------------Functions----------------------
def reply(audio):
    if not audio:
        return

    print(f"omega: {audio}")


def handle_scroll_command(full_cmd: str):
    """Perform a smooth wheel-scroll at the current cursor position."""
    cmd = normalize_voice_text(full_cmd)
    is_up = "up" in cmd
    wheel_delta = 50 if is_up else -50

    try:
        # Wheel event is delivered to the control under current cursor.
        win32api.mouse_event(win32con.MOUSEEVENTF_WHEEL, 0, 0, wheel_delta, 0)
    except Exception:
        # Fallback for environments where low-level wheel event is blocked.
        pyautogui.scroll(50 if is_up else -50)

    reply("Scrolled up." if is_up else "Scrolled down.")

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
        # remove open keywords and command fillers
        name = normalize_app_query(text_lower)
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
    
    elif intent == 'tab_control':
        # Remove trigger words
        text = text_lower
        remove = ["switch", "to", "tab", "open", "go", "back", "previous", "next"]
        words = [w for w in text.split() if w not in remove]
        return " ".join(words).strip()

    elif intent == 'choice_select':
    # return exactly what user said, like "option 4"
        return text_lower
    
    elif intent == 'range_select':
        # whole command needed to extract "from X to Y"
        return text_lower

    elif intent == 'copy' and "from" in text_lower and "to" in text_lower:
        # treat copy-from-to as range copy
        return text_lower
    
    elif intent == 'paste_after':
        text = text_lower
        remove_words = ['paste', 'after', 'the', 'word', 'insert', 'text', 'right', 'immediately']
        phrase = " ".join([w for w in text.split() if w not in remove_words]).strip()
        return phrase if phrase else None

    elif intent == 'paste_before':
        text = text_lower
        remove_words = ['paste', 'before', 'the', 'word', 'insert', 'text', 'right']
        phrase = " ".join([w for w in text.split() if w not in remove_words]).strip()
        return phrase if phrase else None

    elif intent == 'paste_line_start':
        text = text_lower
        # extract phrase after "line"
        m = re.search(r'line (.+)', text)
        return m.group(1).strip() if m else None

    elif intent == 'paste_line_end':
        text = text_lower
        # extract phrase after line/paragraph
        m = re.search(r'(?:line|paragraph) (.+)', text)
        return m.group(1).strip() if m else None

    elif intent == 'paste_cell':
        text = text_lower
        m = re.search(r'cell\s+([a-z]+[0-9]+)', text)
        return m.group(1).strip().upper() if m else None


        
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

        return f"Navigated back to: {current_directory}"
    except Exception as e:
        return f"Error navigating back: {str(e)}"

# Audio to String with better error handling
def record_audio():
    global _audio_calibrated, _last_ambient_calibration_at
    try:
        _voice_show_listening()
        with sr.Microphone() as source:
            now = time.time()
            should_recalibrate = (
                (not _audio_calibrated)
                or (now - _last_ambient_calibration_at) > AMBIENT_RECALIBRATE_EVERY_SEC
            )
            if should_recalibrate:
                r.adjust_for_ambient_noise(source, duration=0.35)
                _audio_calibrated = True
                _last_ambient_calibration_at = now

            r.energy_threshold = max(180, int(r.energy_threshold))
            audio = r.listen(source, timeout=None, phrase_time_limit=5)
            
        try:
            voice_data = _best_transcript_from_google(audio)
            voice_data = normalize_voice_text(voice_data)
            print(f"Recognized: {voice_data}")
            _voice_show_recognized(voice_data)
            return voice_data
        except sr.UnknownValueError:
            return ""
        except sr.RequestError as e:
            print(f"Speech recognition error: {e}")
            reply('Speech recognition service error. Check internet connection.')
            _voice_hide()
            return ""
            
    except sr.WaitTimeoutError:
        print("Listening timeout")
        _voice_hide()
        return ""
    except Exception as e:
        print(f"Microphone error: {e}")
        _voice_hide()
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

def handle_range_choice(num: int):
    """
    Handles 'choose N' specifically for range selection (from X to Y).
    Does NOT interfere with existing handle_choice_selection.
    """
    global range_mode, range_phase
    global range_start_occurrences, range_end_occurrences
    global range_selected_start
    global range_start_phrase, range_end_phrase

    # If not in a range-selection workflow, do nothing
    if range_mode is None or range_phase is None:
        return

    index = num - 1

    # -------------------- Choosing the start phrase --------------------
    if range_phase == "start":
        if index < 0 or index >= len(range_start_occurrences):
            reply("Invalid option number for the start phrase.")
            return

        # Remove overlay so clicks hit the real document
        clear_on_screen_boxes()
        range_selected_start = range_start_occurrences[index]

        # If multiple end matches → ask for them
        if len(range_end_occurrences) > 1:
            range_phase = "end"
            show_numbered_boxes(range_end_occurrences)
            reply("Now choose which ending phrase you want. Say 'choose 1', 'choose 2', etc.")
            return

        # Otherwise: apply immediately (unique end)
        apply_range_selection(
            range_selected_start,
            range_end_occurrences[0],
            range_mode,
            range_start_phrase,
            range_end_phrase
        )
        return

    # -------------------- Choosing the end phrase --------------------
    if range_phase == "end":
        if index < 0 or index >= len(range_end_occurrences):
            reply("Invalid option number for the end phrase.")
            return

        clear_on_screen_boxes()
        end_box = range_end_occurrences[index]

        # Edge case: if start wasn't set for some reason
        if range_selected_start is None and range_start_occurrences:
            range_selected_start = range_start_occurrences[0]

        apply_range_selection(
            range_selected_start,
            end_box,
            range_mode,
            range_start_phrase,
            range_end_phrase
        )

def handle_choice_selection(num):
    global last_found_boxes, last_click_target_text
    global last_tab_boxes, last_tab_target

    # ----------------- CASE 1: TAB SELECTION -----------------
    if last_tab_boxes:
        index = num - 1
        if index < 0 or index >= len(last_tab_boxes):
            reply("Invalid option number.")
            return

        clear_on_screen_boxes()

        x, y, w, h = last_tab_boxes[index]
        cx = x + w // 2
        cy = y + h // 2

        pyautogui.moveTo(cx, cy, duration=0.2)
        pyautogui.click()

        reply(f"Switched to tab: option {num}.")

        # update memory for "previous tab"
        global last_chrome_tab_id, current_chrome_tab_id
        last_chrome_tab_id = current_chrome_tab_id
        current_chrome_tab_id = last_tab_target

        last_tab_boxes = []
        last_tab_target = None
        return

    # ----------------- CASE 2: NORMAL CLICK SELECTION -----------------
    if not last_found_boxes:
        reply("No options available to choose from.")
        return

    index = num - 1
    if index < 0 or index >= len(last_found_boxes):
        reply("Invalid option number.")
        return

    clear_on_screen_boxes()

    x, y, w, h = last_found_boxes[index]

    if looks_like_file_name(last_click_target_text):
        pyautogui.moveTo(x + w//2, y + h//2, duration=0.25)
        pyautogui.doubleClick()
        reply(f"Opened file (option {num}).")
    else:
        pyautogui.moveTo(x + w//2, y + h//2, duration=0.25)
        pyautogui.click()
        reply(f"Selected option {num}.")

    last_found_boxes = []
    last_click_target_text = None

def handle_tab_control(entity):
    if not entity:
        reply("Which tab should I switch to?")
        return

    entity = entity.lower()

    if "previous" in entity or "back" in entity:
        chrome_return_to_previous_tab()
    else:
        chrome_switch_tab_by_name(entity)

def handle_paste_after(full_cmd: str):
    """
    Advanced paste handler.

    Supports:
      - 'paste before <word>'
      - 'paste after <word>'
      (You can extend later for start/end of line / cell logic.)

    Behavior:
      - Use OCR to find <word> on screen
      - If multiple matches -> show numbered options -> user says "choose N"
      - Then move caret just before/after that word and paste, without selecting a range.
    """
    global paste_word_mode, paste_word_boxes, paste_word_target

    text = full_cmd.lower()

    # Detect "before" vs "after"
    mode = None
    if "paste before" in text or "before" in text:
        mode = "before"
    elif "paste after" in text or "after" in text:
        mode = "after"

    if mode not in ("before", "after"):
        # For now, just fallback to simple paste if we don't recognize pattern
        handle_paste()
        return

    # Extract the word/phrase after "before" or "after"
    m = re.search(r'(before|after)\s+(.+)', text)
    if not m:
        reply("Please say 'paste before <word>' or 'paste after <word>'.")
        return

    target = m.group(2).strip()
    if not target:
        reply("I didn't catch the word after before/after.")
        return

    paste_word_target = target
    paste_word_mode = mode

    # Use your OCR engine to find this word on screen
    boxes = find_text_boxes(target)
    if not boxes:
        paste_word_mode = None
        paste_word_boxes = []
        paste_word_target = None
        reply(f"I couldn't find {target} on the screen.")
        return

    # If only one match → paste immediately
    if len(boxes) == 1:
        x, y, w, h = boxes[0]

        if paste_word_mode == "before":
            click_x = x + 2
        else:  # "after"
            click_x = x + w - 2

        click_y = y + h // 2

        pyautogui.click(click_x, click_y)
        time.sleep(0.1)

        with keyboard.pressed(Key.ctrl):
            keyboard.press('v')
            keyboard.release('v')

        reply(f"Pasted {paste_word_mode} {target}.")
        paste_word_mode = None
        paste_word_boxes = []
        paste_word_target = None
        return

    # Multiple matches → show options and wait for 'choose N'
    paste_word_boxes = boxes
    show_numbered_boxes(paste_word_boxes)
    reply(f"I found multiple '{target}' positions. Say 'choose 1', 'choose 2', etc.")

def handle_paste_position_choice(num: int):
    """
    When in paste_position_mode, user says 'choose N'.
    Move caret to that box and paste from clipboard.
    """
    global paste_position_mode, paste_position_boxes

    if paste_position_mode != "active":
        return  # not in paste-position mode; let normal handler manage

    index = num - 1
    if index < 0 or index >= len(paste_position_boxes):
        reply("Invalid option number.")
        return

    clear_on_screen_boxes()

    x, y, w, h = paste_position_boxes[index]
    click_x = x + w // 2
    click_y = y + h // 2

    # Click to move caret
    pyautogui.click(click_x, click_y)
    time.sleep(0.1)

    # Paste from clipboard
    with keyboard.pressed(Key.ctrl):
        keyboard.press('v')
        keyboard.release('v')

    reply("Pasted at the selected position.")

    # Reset state
    paste_position_mode = None
    paste_position_boxes = []

def handle_paste_word_choice(num: int):
    """
    User said 'choose N' while we are in the paste_word_mode workflow.
    Paste before/after the chosen occurrence of the word.
    """
    global paste_word_mode, paste_word_boxes, paste_word_target

    if paste_word_mode not in ("before", "after") or not paste_word_boxes:
        return  # not in word-paste mode, ignore

    index = num - 1
    if index < 0 or index >= len(paste_word_boxes):
        reply("Invalid option number.")
        return

    clear_on_screen_boxes()

    x, y, w, h = paste_word_boxes[index]

    if paste_word_mode == "before":
        click_x = x + 2
    else:  # "after"
        click_x = x + w - 2

    click_y = y + h // 2

    pyautogui.click(click_x, click_y)
    time.sleep(0.1)

    with keyboard.pressed(Key.ctrl):
        keyboard.press('v')
        keyboard.release('v')

    reply(f"Pasted {paste_word_mode} {paste_word_target}.")

    # Reset state
    paste_word_mode = None
    paste_word_boxes = []
    paste_word_target = None
   

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

def handle_copy_range(full_cmd: str):
    # Use shared range-selection engine (with disambiguation)
    start_range_selection(full_cmd, mode="copy")


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


def handle_discard_options():
    """Dismiss any active on-screen option overlays and reset option state."""
    global last_found_boxes, last_click_target_text
    global last_tab_boxes, last_tab_target
    global paste_position_mode, paste_position_boxes
    global paste_word_mode, paste_word_boxes, paste_word_target
    global range_mode, range_phase
    global range_start_occurrences, range_end_occurrences
    global range_start_phrase, range_end_phrase, range_selected_start

    clear_on_screen_boxes()

    last_found_boxes = []
    last_click_target_text = None
    last_tab_boxes = []
    last_tab_target = None

    paste_position_mode = None
    paste_position_boxes = []
    paste_word_mode = None
    paste_word_boxes = []
    paste_word_target = None

    range_mode = None
    range_phase = None
    range_start_occurrences = []
    range_end_occurrences = []
    range_start_phrase = ""
    range_end_phrase = ""
    range_selected_start = None

    reply("Okay, discarded the options.")


def _has_pending_options():
    return (
        bool(last_tab_boxes)
        or bool(last_found_boxes)
        or (range_mode is not None)
        or (paste_position_mode == "active")
        or (paste_word_mode is not None)
    )


def _dispatch_option_number(num: int):
    if not num:
        reply("Please say a valid option number like choose one or choose two.")
        return

    if range_mode is not None:
        handle_range_choice(num)
        return

    if paste_position_mode == "active":
        handle_paste_position_choice(num)
        return

    if paste_word_mode is not None:
        handle_paste_word_choice(num)
        return

    handle_choice_selection(num)

def handle_open_app(app_name):
    if not app_name:
        reply("Which application should I open?")
        return
    target_name = normalize_app_query(app_name)
    if open_taskbar_app(target_name):
        return


def handle_close_app(app_name):
    """Close app windows first, then terminate remaining matching processes."""
    target = normalize_close_app_query(app_name)
    if not target:
        reply("Which application should I close?")
        return

    query_terms = _app_alias_terms(target)
    for canonical, aliases in APP_ALIAS_GROUPS.items():
        if target == canonical or target in aliases:
            query_terms.extend([canonical] + aliases)

    query_terms = list({t for t in query_terms if t and len(t) > 1})
    safe_block = {"omega", "voice", "assistant", "python"}

    windows_closed = 0
    try:
        for title in gw.getAllTitles():
            t = (title or "").lower().strip()
            if not t:
                continue
            if any(term in t for term in query_terms if len(term) > 2):
                wins = gw.getWindowsWithTitle(title)
                for w in wins:
                    try:
                        w.close()
                        windows_closed += 1
                    except Exception:
                        pass
    except Exception:
        pass

    procs_stopped = 0
    proc_terms = set(query_terms)
    for canonical, executables in APP_DIRECT_EXECUTABLES.items():
        if canonical in proc_terms:
            proc_terms.update(os.path.splitext(e)[0].lower() for e in executables)

    for proc in psutil.process_iter(["pid", "name", "exe", "cmdline"]):
        try:
            pname = (proc.info.get("name") or "").lower()
            pexe = os.path.basename(proc.info.get("exe") or "").lower()
            pcmd = " ".join(proc.info.get("cmdline") or []).lower()
            joined = f"{pname} {pexe} {pcmd}"

            if any(b in joined for b in safe_block):
                continue

            if any(term in joined for term in proc_terms if len(term) > 2):
                proc.terminate()
                procs_stopped += 1
        except Exception:
            continue

    if procs_stopped:
        time.sleep(0.4)
        for proc in psutil.process_iter(["pid", "name"]):
            try:
                pname = (proc.info.get("name") or "").lower()
                if any(term in pname for term in proc_terms if len(term) > 2):
                    proc.kill()
            except Exception:
                continue

    if windows_closed or procs_stopped:
        reply(f"Closed {target}.")
    else:
        reply(f"I could not find a running app named {target} to close.")

def handle_select_range(full_cmd: str):
    # Use shared range-selection engine (with disambiguation)
    start_range_selection(full_cmd, mode="select")


def handle_presentation_control(command):
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

        elif any(kw in command for kw in ["highlight", "find", "search on page"]):
            phrase = None

            # Extract text after the first keyword that appears
            for kw in ["highlight", "find", "search for"]:
                if kw in command:
                    phrase = command.split(kw, 1)[1].strip()
                    break

            if phrase:
                pyautogui.hotkey("ctrl", "f")
                time.sleep(0.2)
                pyautogui.typewrite(phrase)
                pyautogui.press("enter")
                reply(f"Highlighted {phrase}")
            else:
                reply("What should I highlight?")
            return True


        # Scroll down
        elif any(kw in command for kw in ["scroll down", "go down", "move down"]):
            pyautogui.scroll(-700)
            reply("Scrolling down.")
            return True

        # Scroll up
        elif any(kw in command for kw in ["scroll up", "go up", "move up"]):
            pyautogui.scroll(700)
            reply("Scrolling up.")
            return True

        
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
    if not entity or entity.strip() in {"", "here", "cursor", "current", "current position", "position"}:
        pyautogui.click()
        reply("Clicked.")
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
       
    normalized_input = normalize_voice_text(voice_data)
    print(f"Processing: {normalized_input}")
    
    # Store original for display
    # Remove wake word and clean up for processing
    processed_voice = strip_wake_word(normalized_input)
    if not processed_voice:
        return

    pv = processed_voice

    # Deterministic fast-path so dismiss works even if intent classification misses it.
    if any(cmd in pv for cmd in (
        "discard options",
        "dismiss options",
        "clear options",
        "cancel options",
        "close options",
        "hide options",
        "discard",
        "dismiss",
        "clear",
        "cancel",
    )):
        handle_discard_options()
        return

    if is_close_app_command(pv):
        close_candidate = normalize_close_app_query(pv)
        if close_candidate:
            handle_close_app(close_candidate)
            return

    if re.match(r"^scroll(?:\s|$)", pv):
        handle_scroll_command(pv)
        return

    if is_thank_you_exit_command(pv):
        return handle_thank_you_exit()

    # Fast path: when options are visible, allow phrases like 'click on two' immediately.
    if _has_pending_options():
        num = parse_choice_number(pv)
        if num:
            _dispatch_option_number(num)
            return

    # Deterministic app-open handling for:
    # - "open chrome"
    # - "open app chrome"
    # - "launch powerpoint"
    if is_open_app_command(pv):
        app_candidate = normalize_app_query(pv)
        if app_candidate:
            handle_open_app(app_candidate)
            return

        # Force Tab Control
    if "switch" in pv and "tab" in pv:
        entity = extract_entity(processed_voice, "tab_control")
        handle_tab_control(entity)
        return
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
    
        if "from" in processed_voice.lower() and "to" in processed_voice.lower():
            handle_copy_range(processed_voice)
        else:
            handle_copy()

    elif intent == 'range_select':
        handle_select_range(processed_voice)

    
    elif intent == 'paste':
        text = processed_voice.lower()

        # 1️⃣ Optional: generic OCR-options paste (if you implemented start_paste_position_selection)
        if "paste using options" in text or "paste at position" in text or "paste with options" in text:
            start_paste_position_selection()

        # 2️⃣ Word-based before/after/start/end/cell paste
        elif any(kw in text for kw in ["after", "before", "start of line", "end of line", "end of paragraph", "cell "]):
            handle_paste_after(processed_voice)

        # 3️⃣ Simple paste at current caret
        else:
            handle_paste()

    elif intent in ['paste_after', 'paste_before', 'paste_line_start', 'paste_line_end', 'paste_cell']:
        handle_paste_after(processed_voice)




    elif intent == 'wake_up':
        handle_wake_up()

    elif intent == 'choice_select':
            _dispatch_option_number(parse_choice_number(entity))


    elif intent == 'presentation_control':
        handle_presentation_control(entity)
    
    elif intent == 'open_app':
        handle_open_app(entity)

    elif intent == 'tab_control':
        handle_tab_control(entity)

    elif intent == 'click_action':
        handle_click_action(entity, processed_voice)

    
    else:
        # Fallback for unknown intents
        reply("I'm not sure how to help with that. You can ask me about time, date, search, files, or location.")

# ------------------Driver Code--------------------

if __name__ == "__main__":
    print(get_taskbar_buttons())    

    tts_thread = Thread(target=_tts_worker, daemon=True)
    cmd_thread = Thread(target=_command_worker, daemon=True)
    tts_thread.start()
    cmd_thread.start()

    wish()

    while not shutdown_event.is_set():
        try:
            # Take input from voice only
            voice_data = record_audio()

            # Process voice_data if we have input
            if voice_data:
                reply(f"I heard {voice_data}")
                if _should_allow_command(voice_data):
                    _enqueue_latest(command_queue, voice_data)
                    
            time.sleep(0.02)
                    
        except SystemExit:
            shutdown_event.set()
        except KeyboardInterrupt:
            reply("Interrupted by user")
            shutdown_event.set()
        except Exception as e:
            print(f"Unexpected error: {e}")
            time.sleep(1)  # Prevent rapid error looping

    _enqueue_latest(command_queue, None)
    _enqueue_latest(tts_queue, None)
    cmd_thread.join(timeout=1.2)
    tts_thread.join(timeout=1.2)
    _voice_close()
    print("omega shutdown complete")
