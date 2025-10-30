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
warnings.filterwarnings('ignore')

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
            'file_open': [
                'open file', 'open number', 'open folder', 'launch file',
                'open item', 'start file', 'open this file'
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
            'go to last slide', 'skip to slide', 'show slide number'
            ]
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
            'file_open': ['open', 'launch', 'start'],
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
            'skip to slide', 'show slide number'
            ]
        }
        
        for intent, keywords in intent_keywords.items():
            for keyword in keywords:
                if keyword in text_lower:
                    return intent
        
        return 'unknown'

# Initialize intent recognizer
intent_recognizer = IntentRecognizer()

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
    
    elif intent in ['files', 'file_open']:
        # Extract file number or name
        numbers = re.findall(r'\d+', text_lower)
        if numbers:
            return int(numbers[0])  # Return the first number found
        return None
    
    elif intent == 'file_navigate':
        return 'back'
    
    elif intent == 'presentation_control':
        # Remove the wake word and common filler words to get the presentation command
        control_words = ['omega', 'presentation', 'slideshow', 'powerpoint', 'keynote', 'slide', 'please']
        command = text_lower
        for word in control_words:
            command = command.replace(word, '').strip()
        
        # Remove extra spaces and return the clean command
        command = re.sub(r'\s+', ' ', command).strip()
        return command if command and len(command) > 1 else text_lower
    return None

def open_file_or_folder(full_path):
    """Open a file or folder using the appropriate method"""
    try:
        if exists(full_path):
            if isfile(full_path):
                # It's a file - open it with default application
                os.startfile(full_path)
                return f"Opened file: {os.path.basename(full_path)}"
            else:
                # It's a folder - navigate into it
                global current_directory, files, file_exp_status
                current_directory = full_path
                files = listdir(current_directory)
                
                # Display new directory contents
                filestr = ""
                counter = 0
                for f in files:
                    counter += 1
                    item_type = "file" if isfile(join(current_directory, f)) else "folder"
                    filestr += f"{counter}: {f} ({item_type})<br>"
                    print(f"{counter}: {f} ({item_type})")
                
                file_exp_status = True
                app.ChatBot.addAppMsg(filestr)
                return f"Opened folder: {os.path.basename(full_path)}"
        else:
            return f"Path does not exist: {full_path}"
    except Exception as e:
        return f"Error opening: {str(e)}"

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
    """Handle search intent"""
    if query:
        reply(f'Searching for {query}')
        url = 'https://google.com/search?q=' + query.replace(' ', '+')
        try:
            webbrowser.get().open(url)
            reply('Here are the search results I found')
        except:
            reply('Please check your internet connection')
    else:
        reply('What would you like me to search for?')
        temp_audio = record_audio()
        if temp_audio:
            query = temp_audio
            reply(f'Searching for {query}')
            url = 'https://google.com/search?q=' + query.replace(' ', '+')
            try:
                webbrowser.get().open(url)
                reply('Here are the search results')
            except:
                reply('Please check your internet connection')

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

def handle_files(operation, voice_data):
    """Handle file listing"""
    global file_exp_status, files, current_directory
    
    if operation == 'list' or not operation:
        counter = 0
        current_directory = 'C://'
        try:
            files = listdir(current_directory)
            filestr = ""
            for f in files:
                counter += 1
                item_type = "file" if isfile(join(current_directory, f)) else "folder"
                filestr += f"{counter}: {f} ({item_type})<br>"
                print(f"{counter}: {f} ({item_type})")
            file_exp_status = True
            reply('These are the files in your root directory')
            app.ChatBot.addAppMsg(filestr)
        except Exception as e:
            reply(f'Error accessing directory: {e}')

def handle_file_open(file_number, voice_data):
    """Handle file/folder opening by number"""
    global file_exp_status, files, current_directory
    
    if not file_exp_status:
        reply("Please list files first using 'list files' command")
        return
    
    try:
        if file_number and 1 <= file_number <= len(files):
            file_index = file_number - 1
            selected_item = files[file_index]
            full_path = join(current_directory, selected_item)
            
            result = open_file_or_folder(full_path)
            reply(result)
        else:
            reply(f"Please specify a valid file number between 1 and {len(files)}")
    except Exception as e:
        reply(f"Error opening file: {str(e)}")

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

def handle_presentation_control(command):
    """
    Handle presentation control commands such as:
    next slide, previous slide, start presentation, end presentation, etc.
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
        
        elif any(kw in command for kw in ["start", "begin", "slideshow", "present"]):
            pyautogui.press('f5')
            reply("Starting the presentation.")
        
        elif any(kw in command for kw in ["end", "exit", "stop", "close"]):
            pyautogui.press('esc')
            reply("Presentation closed.")
        
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
        handle_files(entity, processed_voice)
    
    elif intent == 'file_open':
        handle_file_open(entity, processed_voice)
    
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