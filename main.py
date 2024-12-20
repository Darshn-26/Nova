import speech_recognition as sr
import webbrowser
import os
import subprocess
from datetime import datetime
import pyautogui
import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pyttsx3
import json
import logging
from dotenv import load_dotenv
import groq
import keyboard
import win32gui
import win32con
import win32com.client
from PIL import Image
import pytesseract
import cv2
import numpy

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    filename='assistant.log'
)

class VoiceAssistant:
    def __init__(self):
        # Load environment variables
        load_dotenv()
        
        # Initialize speech recognition
        self.recognizer = sr.Recognizer()

        # Adjust recognition settings
        self.recognizer.energy_threshold = 4000
        self.recognizer.dynamic_energy_threshold = True
        self.recognizer.pause_threshold = 0.5

        # Initialize text-to-speech
        self.engine = pyttsx3.init()
        self.setup_voice()
        
        # Initialize Llama model with Groq
        self.setup_ai()
        
        # Initialize Selenium WebDriver
        self.driver = None
        
        # Initialize context management
        self.context = {
            "user_name": os.getenv('USER_NAME', 'Sir'),
            "last_command": None,
            "conversation_history": [],
        }

        # Initialize application paths and website shortcuts
        self.app_paths = {
            'notepad': 'notepad.exe',
            'calculator': 'calc.exe',
            'chrome': r'C:\Program Files\Google\Chrome\Application\chrome.exe',
            'firefox': r'C:\Program Files\Mozilla Firefox\firefox.exe',
            'word': 'WINWORD.EXE',
            'excel': 'EXCEL.EXE',
            'powerpoint': 'POWERPNT.EXE',
            'vlc': r'C:\Program Files\VideoLAN\VLC\vlc.exe',
        }

        self.website_shortcuts = {
            'google': 'google.com',
            'youtube': 'youtube.com',
            'facebook': 'facebook.com',
            'twitter': 'twitter.com',
            'linkedin': 'linkedin.com',
            'github': 'github.com',
        }

    def setup_voice(self):
        """Configure text-to-speech settings"""
        voices = self.engine.getProperty('voices')
        self.engine.setProperty('voice', voices[0].id)  # Change from voices[0] to voices[1]
        # Slower rate for more natural speech
        self.engine.setProperty('rate', 180)  # Reduced from 180
        # Add pitch variation
        self.engine.setProperty('pitch', 1.2)  # Default is 1.0
        self.engine.setProperty('volume', 0.9)


    def setup_ai(self):
        """Initialize Llama model with Groq"""
        try:
            api_key = os.getenv('GROQ_API_KEY')
            if not api_key:
                raise ValueError("GROQ_API_KEY not found in environment variables")
            
            self.groq_client = groq.Groq(api_key=api_key)
            logging.info("AI model initialized successfully")
        except Exception as e:
            logging.error(f"Error initializing AI model: {e}")
            self.speak("Warning: AI capabilities are currently limited")

    def get_ai_response(self, user_input):
        """Get response from Llama model via Groq"""
        try:
            prompt = f"User: {user_input}\nAssistant: "
            
            chat_completion = self.groq_client.chat.completions.create(
                messages=[
                    {
                        "role": "system",
                        "content": "You are a helpful assistant named Pax. You have to assist users with basic tasks such as opening applications, searching the web, setting alarms and remainders etc. Keep your answers sweet and short. You occasionally use casual expressions and friendly metaphors. You show empathy and understanding. You're knowledgeable but humble. You sometimes express emotions through tone (e.g., 'I'm excited to help!). Keep responses consise but warm. Have persistent Memory of the previous conversation."
                    },
                    {
                        "role": "user",
                        "content": user_input
                    }
                ],
                model="llama3-8b-8192",
                temperature=0.7,
                max_tokens=100
            )
            
            return chat_completion.choices[0].message.content
        except Exception as e:
            logging.error(f"AI model error: {e}")
            return "I'm sorry, I didn't understand that command."

    def speak(self, text):
        """Convert text to speech"""
        try:
            print("Assistant:", text)
            self.engine.say(text)
            self.engine.runAndWait()
        except Exception as e:
            logging.error(f"Error in text-to-speech: {e}")
            print(f"Speech Error: {text}")

    def listen(self):
        """Continuously listen for commands"""
        try:
            with sr.Microphone() as source:
                print("Listening...")
                self.recognizer.adjust_for_ambient_noise(source, duration=0.5)
                audio = self.recognizer.listen(source, timeout=5, phrase_time_limit=5)
                command = self.recognizer.recognize_google(audio).lower()
                print(f"You said: {command}")
                self.process_command(command)
        except sr.WaitTimeoutError:
            pass
        except sr.UnknownValueError:
            pass
        except Exception as e:
            logging.error(f"Error in listen: {e}")

    def process_command(self, command):
        """Process and execute commands"""
        try:
            command = command.lower().strip()

                      
            # Website handling
            if any(x in command for x in ["go to", "visit", "open website"]):
                for trigger in ["go to", "visit", "open website"]:
                    if trigger in command:
                        url = command.split(trigger)[-1].strip()
                        if url:
                            self.visit_website(url)
                            return

            # Typing handling
            if any(x in command for x in ["type in", "write in", "input in"]):
                for trigger in ["type in", "write in", "input in"]:
                    if trigger in command:
                        rest = command.split(trigger)[-1].strip()
                        parts = rest.split(" ", 1)
                        if len(parts) == 2:
                            app_name, text = parts
                            self.type_in_application(app_name, text)
                            return

            # Application opening
            if any(x in command for x in ["open", "launch", "start", "run"]):
                for trigger in ["open", "launch", "start", "run"]:
                    if trigger in command:
                        app_name = command.split(trigger)[-1].strip()
                        if app_name:
                            self.open_application(app_name)
                            return

            # Search handling
            if any(x in command for x in ["search for", "look up", "google"]):
                for trigger in ["search for", "look up", "google"]:
                    if trigger in command:
                        search_term = command.split(trigger)[-1].strip()
                        if search_term:
                            self.web_search(search_term)
                            return

             # File opening commands with more variations
            file_triggers = [
                "open file", "find file", "show file", 
                "open the file", "find the file", "show the file",
                "open", "find", "show"
            ]

            if any(trigger in command for trigger in file_triggers):
                for trigger in file_triggers:
                    if trigger in command:
                        file_name = command.split(trigger)[-1].strip()
                    if file_name:
                        self.open_file(file_name)
                        return

            # Screenshot handling
            if "screenshot" in command or "capture screen" in command:
                self.take_screenshot()
                return

            # Close window handling
            if "close" in command and ("window" in command or "application" in command):
                self.close_active_window()
                return

            # Exit handling
            if command in ["exit", "quit", "goodbye", "bye"]:
                self.shutdown()
                return

            # If no direct command match, use AI to interpret
            response = self.get_ai_response(command)
            self.speak(response)

        except Exception as e:
            logging.error(f"Error in process_command: {e}")
            self.speak("Sorry, I encountered an error processing that command")

    def open_file(self, file_name):
        """Open any file in the system by name with voice selection"""
        try:
            # Input validation
            if not file_name:
                self.speak("Please specify a file name to open")
                return False
    
            self.speak(f"Searching for files named {file_name}")
            
            common_extensions = [
                '', '.txt', '.pdf', '.doc', '.docx', '.xls', '.xlsx', 
                '.ppt', '.pptx', '.jpg', '.jpeg', '.png', '.mp3', 
                '.mp4', '.wav', '.zip', '.rar'
            ]
            
            search_paths = [
                os.path.expanduser('~'),  # Home directory
                os.path.join(os.path.expanduser('~'), 'Desktop'),
                os.path.join(os.path.expanduser('~'), 'Documents'),
                os.path.join(os.path.expanduser('~'), 'Downloads'),
                os.path.join(os.path.expanduser('~'), 'Pictures'),
                os.path.join(os.path.expanduser('~'), 'Music'),
                os.path.join(os.path.expanduser('~'), 'Videos')
            ]
    
            def find_file(name, paths):
                """Search for file in given paths"""
                found_files = []
                try:
                    for path in paths:
                        if os.path.exists(path):
                            for root, dirs, files in os.walk(path):
                                for ext in common_extensions:
                                    search_name = f"{name}{ext}".lower()
                                    for file in files:
                                        if file.lower() == search_name:
                                            full_path = os.path.join(root, file)
                                            found_files.append(full_path)
                except Exception as e:
                    logging.error(f"Error in find_file: {e}")
                return found_files
    
            def select_file_by_voice(files):
                """Handle voice selection of files"""
                max_attempts = 3
                attempts = 0
                
                while attempts < max_attempts:
                    try:
                        self.speak("Please say the number of the file you want to open")
                        
                        with sr.Microphone() as source:
                            self.recognizer.adjust_for_ambient_noise(source, duration=0.5)
                            print("Listening for file selection...")
                            audio = self.recognizer.listen(source, timeout=5, phrase_time_limit=3)
                            selection = self.recognizer.recognize_google(audio).lower()
                            print(f"Heard selection: {selection}")  # Debug print
                            
                            # Convert word numbers to digits
                            number_mapping = {
                                'one': '1', 'two': '2', 'three': '3', 'four': '4', 
                                'five': '5', 'six': '6', 'seven': '7', 'eight': '8',
                                'nine': '9', 'ten': '10'
                            }
                            
                            for word, digit in number_mapping.items():
                                selection = selection.replace(word, digit)
                            
                            # Extract numbers from the selection
                            numbers = [int(s) for s in selection.split() if s.isdigit()]
                            
                            if numbers:
                                index = numbers[0] - 1
                                if 0 <= index < len(files):
                                    return files[index]
                                else:
                                    self.speak(f"Please select a number between 1 and {len(files)}")
                            else:
                                self.speak("I didn't hear a number. Please try again.")
                            
                    except sr.WaitTimeoutError:
                        self.speak("I didn't hear your selection. Please try again.")
                    except sr.UnknownValueError:
                        self.speak("I couldn't understand your selection. Please try again.")
                    except Exception as e:
                        logging.error(f"Error in voice selection: {e}")
                        self.speak("There was an error with the selection. Please try again.")
                    
                    attempts += 1
                
                self.speak("Maximum attempts reached. Opening the first file by default.")
                return files[0] if files else None
    
            # Search for the file
            found_files = find_file(file_name, search_paths)
    
            if found_files:
                if len(found_files) > 1:
                    # List found files
                    self.speak(f"I found {len(found_files)} files with that name.")
                    for i, file in enumerate(found_files, 1):
                        file_info = f"{i}. {os.path.basename(file)} in {os.path.dirname(file)}"
                        print(file_info)
                        self.speak(file_info)
                    
                    # Get user's selection
                    selected_file = select_file_by_voice(found_files)
                    if selected_file:
                        try:
                            os.startfile(selected_file)
                            self.speak(f"Opening {os.path.basename(selected_file)}")
                            return True
                        except Exception as e:
                            logging.error(f"Error opening selected file: {e}")
                            self.speak("Sorry, I couldn't open the selected file")
                            return False
                else:
                    # Open the single file found
                    try:
                        os.startfile(found_files[0])
                        self.speak(f"Opening {os.path.basename(found_files[0])}")
                        return True
                    except Exception as e:
                        logging.error(f"Error opening single file: {e}")
                        self.speak("Sorry, I couldn't open the file")
                        return False
            else:
                self.speak(f"Sorry, I couldn't find any file named {file_name}")
                return False
    
        except Exception as e:
            logging.error(f"Error in open_file: {e}")
            self.speak(f"Sorry, I encountered an error while trying to open {file_name}")
            return False
    
    

    def visit_website(self, url):
        """Enhanced website navigation"""
        try:
            url = url.lower().strip()
            
            if url in self.website_shortcuts:
                url = self.website_shortcuts[url]

            if not url.startswith(('http://', 'https://')):
                url = 'https://' + url

            if not url.startswith('https://www.') and '.' in url.split('/', 3)[2]:
                url = url.replace('https://', 'https://www.')

            webbrowser.open(url)
            self.speak(f"Opening {url}")
        except Exception as e:
            logging.error(f"Error visiting website: {e}")
            self.speak("I couldn't open that website")

    def type_in_application(self, app_name, text):
        """Enhanced typing functionality"""
        try:
            app_name = app_name.lower().strip()
            
            if not self.focus_window(app_name):
                self.open_application(app_name)
                time.sleep(2)
                if not self.focus_window(app_name):
                    raise Exception(f"Could not focus on {app_name}")

            time.sleep(0.5)
            
            keyboard.press_and_release('ctrl+a')
            keyboard.press_and_release('delete')
            
            for char in text:
                keyboard.write(char)
                time.sleep(0.02)
            
            self.speak(f"Text typed in {app_name}")
        except Exception as e:
            logging.error(f"Error typing in {app_name}: {e}")
            self.speak(f"I couldn't type in {app_name}")

    def focus_window(self, window_title):
        """Improved window focusing"""
        try:
            def callback(hwnd, windows):
                if win32gui.IsWindowVisible(hwnd):
                    title = win32gui.GetWindowText(hwnd).lower()
                    if window_title.lower() in title:
                        windows.append((hwnd, title))
                return True

            windows = []
            win32gui.EnumWindows(callback, windows)
            
            if windows:
                windows.sort(key=lambda x: len(x[1]))
                hwnd = windows[0][0]
                
                if win32gui.IsIconic(hwnd):
                    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                
                win32gui.SetForegroundWindow(hwnd)
                
                try:
                    shell = win32com.client.Dispatch("WScript.Shell")
                    shell.SendKeys('%')
                    win32gui.SetForegroundWindow(hwnd)
                except:
                    pass
                
                time.sleep(0.5)
                return True
                
            return False
        except Exception as e:
            logging.error(f"Error focusing window: {e}")
            return False

    def open_application(self, app_name):
        """Open application with improved error handling"""
        try:
            app_name = app_name.lower().strip()
            
            if app_name in self.app_paths:
                subprocess.Popen(self.app_paths[app_name])
                self.speak(f"Opening {app_name}")
                return True

            try:
                subprocess.Popen(f"{app_name}.exe")
                self.speak(f"Opening {app_name}")
                return True
            except:
                pass

            common_paths = [
                f"C:\\Program Files\\{app_name}\\{app_name}.exe",
                f"C:\\Program Files (x86)\\{app_name}\\{app_name}.exe",
                f"C:\\Users\\{os.getenv('USERNAME')}\\AppData\\Local\\{app_name}\\{app_name}.exe"
            ]

            for path in common_paths:
                if os.path.exists(path):
                    subprocess.Popen(path)
                    self.speak(f"Opening {app_name}")
                    return True

            self.speak(f"Sorry, I couldn't find {app_name}")
            return False
            
        except Exception as e:
            logging.error(f"Error opening application {app_name}: {e}")
            self.speak(f"Sorry, I couldn't open {app_name}")
            return False

    def web_search(self, query):
        """Perform a web search"""
        try:
            search_url = f"https://www.google.com/search?q={query.replace(' ', '+')}"
            webbrowser.open(search_url)
            self.speak(f"Searching for {query}")
        except Exception as e:
            logging.error(f"Error performing web search: {e}")
            self.speak("I couldn't perform the search")

    def take_screenshot(self):
        """Take and save a screenshot"""
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"screenshot_{timestamp}.png"
            screenshot = pyautogui.screenshot()
            screenshot.save(filename)
            self.speak(f"Screenshot saved as {filename}")
        except Exception as e:
            logging.error(f"Error taking screenshot: {e}")
            self.speak("I couldn't take the screenshot")

    def close_active_window(self):
        """Close the currently active window"""
        try:
            pyautogui.hotkey('alt', 'f4')
            self.speak("Closing active window")
        except Exception as e:
            logging.error(f"Error closing window: {e}")
            self.speak("I couldn't close the window")

    def shutdown(self):
        """Shutdown the assistant"""
        self.speak("Goodbye! Have a great day!")
        if self.driver:
            self.driver.quit()
        exit()

def main():
    """Main execution function"""
    try:
        assistant = VoiceAssistant()
        print("Voice Assistant Started!")
        print("Listening for commands...")
        
        while True:
            try:
                assistant.listen()
            except KeyboardInterrupt:
                print("\nExiting Voice Assistant...")
                break
            except Exception as e:
                print(f"An error occurred: {e}")
                continue
    except Exception as e:
        print(f"Critical error: {e}")
        logging.error(f"Critical error: {e}")

if __name__ == "__main__":
    main()
