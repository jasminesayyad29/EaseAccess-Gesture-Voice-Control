import eel
import os
from queue import Queue

class ChatBot:

    started = False
    userinputQueue = Queue()

    def isUserInput():
        return not ChatBot.userinputQueue.empty()

    def popUserInput():
        return ChatBot.userinputQueue.get()

    def close_callback(route, websockets):
        exit()

    @eel.expose
    def getUserInput(msg):
        ChatBot.userinputQueue.put(msg)
        print(f"User input: {msg}")
    
    def close():
        ChatBot.started = False
    
    def addUserMsg(msg):
        eel.addUserMsg(msg)
    
    def addAppMsg(msg):
        eel.addAppMsg(msg)

    def start():
        path = os.path.dirname(os.path.abspath(__file__))
        eel.init(path + r'\web', allowed_extensions=['.js', '.html'])
        try:
            eel.start('index.html', mode='chrome',
                                    host='localhost',
                                    port=27005,
                                    block=False,
                                    size=(350, 480),
                                    position=(10,100),
                                    disable_cache=True,
                                    close_callback=ChatBot.close_callback)
            ChatBot.started = True
            while ChatBot.started:
                try:
                    eel.sleep(1.0)  # Reduced sleep time for better responsiveness
                except:
                    break
        
        except Exception as e:
            print(f"Eel start error: {e}")

# Initialize on import
if __name__ == "__main__":
    ChatBot.start()