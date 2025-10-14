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
from os.path import isfile, join
import wikipedia
import app
from threading import Thread

# # Try to import gesture controller with error handling
# try:
#     import Gesture_Controller
#     GESTURE_AVAILABLE = True
# except ImportError as e:
#     print(f"Gesture controller not available: {e}")
#     GESTURE_AVAILABLE = False

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
files =[]
path = ''
is_awake = True  # Bot status

# ------------------Functions----------------------
def reply(audio):
    if not audio:
        return
        
    app.ChatBot.addAppMsg(audio)
    print(f"Aura: {audio}")
    
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
        
    reply("I am Aura, how may I help you?")

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

# Executes Commands (input: string)
def respond(voice_data):
    global file_exp_status, files, is_awake, path
    
    if not voice_data:
        return
        
    print(f"Processing: {voice_data}")
    voice_data = voice_data.replace('Aura', '')
    app.ChatBot.addUserMsg(voice_data)

    if is_awake == False:
        if 'wake up' in voice_data:
            is_awake = True
            wish()
        return

    # STATIC CONTROLS
    elif 'hello' in voice_data:
        wish()

    elif 'what is your name' in voice_data:
        reply('My name is Aura!')

    elif 'date' in voice_data:
        reply(today.strftime("%B %d, %Y"))

    elif 'time' in voice_data:
        reply(str(datetime.datetime.now()).split(" ")[1].split('.')[0])

    elif 'search' in voice_data:
        query = voice_data.split('search')[-1].strip()
        if query:
            reply('Searching for ' + query)
            url = 'https://google.com/search?q=' + query
            try:
                webbrowser.get().open(url)
                reply('This is what I found Sir')
            except:
                reply('PlAura check your Internet')
        else:
            reply('What would you like me to search for?')

    elif 'location' in voice_data:
        reply('Which place are you looking for?')
        temp_audio = record_audio()
        if temp_audio:
            app.ChatBot.addUserMsg(temp_audio)
            reply('Locating...')
            url = 'https://google.com/maps/place/' + temp_audio.replace(' ', '+')
            try:
                webbrowser.get().open(url)
                reply('This is what I found Sir')
            except:
                reply('PlAura check your Internet')

    elif ('bye' in voice_data) or ('by' in voice_data):
        reply("Good bye Sir! Have a nice day.")
        is_awake = False

    # elif ('exit' in voice_data) or ('terminate' in voice_data):
    #     if GESTURE_AVAILABLE and Gesture_Controller.GestureController.gc_mode:
    #         Gesture_Controller.GestureController.gc_mode = 0
    #     app.ChatBot.close()
    #     reply("Shutting down...")
    #     return "exit"
    
    # # DYNAMIC CONTROLS
    # elif 'launch gesture recognition' in voice_data:
    #     if not GESTURE_AVAILABLE:
    #         reply('Gesture recognition is not available on this system')
    #     elif Gesture_Controller.GestureController.gc_mode:
    #         reply('Gesture recognition is already active')
    #     else:
    #         try:
    #             gc = Gesture_Controller.GestureController()
    #             t = Thread(target=gc.start, daemon=True)
    #             t.start()
    #             reply('Launched Successfully')
    #         except Exception as e:
    #             reply(f'Failed to launch gesture recognition: {e}')

    # elif 'stop gesture recognition' in voice_data:
    #     if not GESTURE_AVAILABLE:
    #         reply('Gesture recognition is not available')
    #     elif Gesture_Controller.GestureController.gc_mode:
    #         Gesture_Controller.GestureController.gc_mode = 0
    #         reply('Gesture recognition stopped')
    #     else:
    #         reply('Gesture recognition is already inactive')
        
    elif 'copy' in voice_data:
        with keyboard.pressed(Key.ctrl):
            keyboard.press('c')
            keyboard.relAura('c')
        reply('Copied')
          
    elif 'paste' in voice_data:
        with keyboard.pressed(Key.ctrl):
            keyboard.press('v')
            keyboard.relAura('v')
        reply('Pasted')
        
    # File Navigation (Default Folder set to C://)
    elif 'list files' in voice_data:
        counter = 0
        path = 'C://'
        try:
            files = listdir(path)
            filestr = ""
            for f in files:
                counter += 1
                print(str(counter) + ':  ' + f)
                filestr += str(counter) + ':  ' + f + '<br>'
            file_exp_status = True
            reply('These are the files in your root directory')
            app.ChatBot.addAppMsg(filestr)
        except Exception as e:
            reply(f'Error accessing directory: {e}')
        
    elif file_exp_status == True:
        counter = 0   
        if 'open' in voice_data:
            try:
                file_index = int(voice_data.split(' ')[-1]) - 1
                if 0 <= file_index < len(files):
                    full_path = join(path, files[file_index])
                    if isfile(full_path):
                        os.startfile(full_path)
                        file_exp_status = False
                        reply('File opened')
                    else:
                        path = full_path + '//'
                        files = listdir(path)
                        filestr = ""
                        for f in files:
                            counter += 1
                            filestr += str(counter) + ':  ' + f + '<br>'
                            print(str(counter) + ':  ' + f)
                        reply('Opened Successfully')
                        app.ChatBot.addAppMsg(filestr)
                else:
                    reply('Invalid file number')
            except (ValueError, IndexError):
                reply('PlAura specify a valid file number')
            except Exception as e:
                reply(f'Error opening file: {e}')
                                    
        elif 'back' in voice_data:
            filestr = ""
            if path == 'C://':
                reply('Sorry, this is the root directory')
            else:
                try:
                    a = path.split('//')[:-2]
                    path = '//'.join(a)
                    path += '//'
                    files = listdir(path)
                    for f in files:
                        counter += 1
                        filestr += str(counter) + ':  ' + f + '<br>'
                        print(str(counter) + ':  ' + f)
                    reply('Navigated back')
                    app.ChatBot.addAppMsg(filestr)
                except Exception as e:
                    reply(f'Error navigating back: {e}')
                   
    else: 
        reply('I am not functioned to do this!')

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
            if voice_data and 'Aura' in voice_data:
                result = respond(voice_data)
                if result == "exit":
                    break
                    
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

    print("Aura shutdown complete")