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
import winreg
from difflib import get_close_matches
# Keep track of where we came from so we can return if needed
last_context = None   # possible values: None, 'presentation', 'browser', ...
AUTO_RETURN_AFTER_SEARCH = False # Set True if you want automatic return after search


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


# --------------------------
# MATCHER
# --------------------------
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


# --------------------------
# UNIVERSAL LAUNCHER
# --------------------------
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
            'open_app': ['open app', 'launch app', 'start app', 'open application', 'run', 'open program', 'start program']

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
            'open_app': ['open app', 'launch', 'start', 'run']

        }
        
        for intent, keywords in intent_keywords.items():
            for keyword in keywords:
                if keyword in text_lower:
                    return intent
        
        return 'unknown'

# Initialize intent recognizer
intent_recognizer = IntentRecognizer()

def minimize_browser_windows():
    """Minimize Chrome, Edge, or Firefox windows if open."""
    import pygetwindow as gw

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
    """
    Open files or folders by natural names:
    - 'downloads', 'desktop', 'documents', 'c drive', 'd drive'
    - filename (searches common user folders)
    - full path (if provided)
    - 'recent <type>' opens recent files
    """
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