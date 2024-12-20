import os
import re
import dotenv
import subprocess
import speech_recognition as sr
import pyttsx3
import google.generativeai as genai
import cv2
import mediapipe as mp
import pyautogui
import webbrowser
import mss
import numpy as np
from datetime import datetime
import gesture
import spacy
from spacy.cli import download
import analyse

# Ensure spaCy model is available
try:
    nlp = spacy.load("en_core_web_sm")
except OSError:
    print("Downloading 'en_core_web_sm' model...")
    download("en_core_web_sm")
    nlp = spacy.load("en_core_web_sm")

# Load environment variables
dotenv.load_dotenv()
genai.configure(api_key=os.getenv("GEMINI_API_KEY"))

# Create the model
generation_config = {
    "temperature": 1,
    "top_p": 0.95,
    "top_k": 40,
    "max_output_tokens": 8192,
    "response_mime_type": "text/plain",
}

model = genai.GenerativeModel(
    model_name="gemini-1.5-pro",
    generation_config=generation_config,
    system_instruction="You are Nova, an AI assistant. Respond to commands with a personal touch and assist with various tasks. Give short and sweet responses to questions. Don't use emojis (abbreviations are ok).with the ability to find direct links to websites or research papers."
)

class VoiceAssistant:
    def __init__(self):
        self.recognizer = sr.Recognizer()
        self.tts_engine = pyttsx3.init()
        self.driver = None
        self.chat_session = None
        self.stop_tracking_flag = False

        self.commands = {
            "open browser": self.open_browser,
            "launch browser": self.open_browser,
            "start browser": self.open_browser,
            "search for": self.web_search,
            "google": self.web_search,
            "find about": self.web_search,
            "take screenshot": self.take_screenshot,
            "capture screen": self.take_screenshot,
            "screenshot": self.take_screenshot,
            "launch application": self.launch_application,
            "start app": self.launch_application,
            "open app": self.launch_application,
            "close app": self.close_app,
            "terminate app": self.close_app,
            "end app": self.close_app,
            "read file": self.read_file_content,
            "open file": self.read_file_content,
            "view file": self.read_file_content,
            "what time is it": self.tell_time,
            "current time": self.tell_time,
            "tell me the time": self.tell_time,
            "who are you": self.who_are_you,
            "introduce yourself": self.who_are_you,
            "your identity": self.who_are_you,
            "nova": self.hey_nova,
            "hey nova": self.hey_nova,
            "call nova": self.hey_nova,
            "close window": self.close_window,
            "exit window": self.close_window,
            "minimize window": self.minimize_window,
            "shrink window": self.minimize_window,
            "hide window": self.minimize_window,
            "track gesture": self.track_gesture,
            "gesture tracking": self.track_gesture,
            "start tracking": self.track_gesture,
            "analyse window":self.analyze_window,
        }

    def speak(self, text):
        """Convert text to speech."""
        self.tts_engine.say(text)
        self.tts_engine.runAndWait()

    def listen(self):
        """Listen to user commands."""
        with sr.Microphone() as source:
            print("Listening...")
            self.recognizer.adjust_for_ambient_noise(source)
            audio = self.recognizer.listen(source)

        try:
            command = self.recognizer.recognize_google(audio).lower()
            print(f"You said: {command}")
            self.process_command(command)
        except sr.UnknownValueError:
            self.speak("Go ahead. I'm listening. ")
        except sr.RequestError:
            self.speak("There was an error with the speech recognition service.")

    def process_command(self, command):
        """Process user commands using NLP for better understanding."""
        if "stop tracking" in command:
            self.stop_tracking_flag = True
            self.speak("Stopping gesture tracking.")
            return
        if "analyze window" in command or "analyse window" in command:
            self.analyze_window()
            return

        # Use NLP to match commands
        doc = nlp(command)
        for key in self.commands:
            if all(token.lemma_ in command for token in nlp(key)):
                if key in ["search for", "read file", "launch application"]:
                    argument = command.split(key)[-1].strip()
                    self.commands[key](argument)
                else:
                    self.commands[key]()
                return

        self.handle_undefined_command(command)

    def handle_undefined_command(self, command):
        """Handle undefined commands by responding conversationally."""
        self.speak("Let me think about that for a moment.")
        print(f"User command not recognized: {command}")

        if not self.chat_session:
            self.chat_session = model.start_chat(
                history=[
                    {
                        "role": "user",
                        "parts": [command],
                    },
                ]
            )

        response = self.chat_session.send_message(command)
        reply = response.text.strip()
        print(f"Nova's response: {reply}")
        self.speak(reply)

    def open_browser(self):
        """Open the browser."""
        webbrowser.open("http://www.google.com")
        self.speak("Browser opened.")

    def web_search(self, query):
        """Search for direct links using Gemini or fallback to Google search."""
        try:
            # Initialize chat session if not already done
            if not self.chat_session:
                self.chat_session = model.start_chat(
                    history=[
                        {
                            "role": "user",
                            "parts": ["You are Nova, a helpful assistant. Please help find direct links to resources."],
                        }
                    ]
                )

            # Send a query to Gemini to find a direct link
            response = self.chat_session.send_message(
                {"parts": [f"Can you find a direct link for {query}? Provide only the most relevant URL."]}
            )
            reply = response.text.strip()
            print(f"Gemini's response: {reply}")

            # Extract URL using a regex
            url_match = re.search(r'https?://[^\s\]]+', reply)  # Matches URLs starting with http or https
            if url_match:
                url = url_match.group(0)
                self.speak(f"Opening the link I found for {query}.")
                webbrowser.open(url)
            else:
                # Fallback to a Google search
                self.speak("I couldn't find a direct link. Performing a Google search instead.")
                url = f"https://www.google.com/search?q={query.replace(' ', '+')}"
                webbrowser.open(url)
        except Exception as e:
            self.speak("There was an error fetching the link. I'll perform a Google search instead.")
            print(f"Error: {e}")
            url = f"https://www.google.com/search?q={query.replace(' ', '+')}"
            webbrowser.open(url)

    def close_window(self):
        """Close the active window."""
        try:
            subprocess.run(["osascript", "-e", 'tell application "System Events" to keystroke "w" using {command down}'])
            self.speak("Closing the active window.")
        except Exception as e:
            self.speak("There was an error closing the window.")
            print(f"Error: {e}")

    def hey_nova(self):
        """Handle the special command for calling Nova."""
        self.speak("Hello there! What can I do for you?")
        print("Hello there! What can I do for you?")

    def minimize_window(self):
        """Minimize the active window."""
        try:
            subprocess.run(["osascript", "-e", 'tell application "System Events" to keystroke "m" using {command down}'])
            self.speak("Minimizing the active window.")
        except Exception as e:
            self.speak("There was an error minimizing the window.")
            print(f"Error: {e}")

    def take_screenshot(self):
        """Take a screenshot."""
        with mss.mss() as sct:
            filename = f"screenshot_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
            sct.shot(output=filename)
            self.speak(f"Screenshot saved as {filename}.")

    def launch_application(self, app_name):
        """Launch a specific application."""
        try:
            self.speak(f"Launching {app_name}.")
            subprocess.Popen(["open", "-a", app_name])
        except Exception as e:
            self.speak(f"Unable to launch {app_name}. Please try again.")
            print(f"Error: {e}")

    def close_app(self, app_name):
        """Close a specific application."""
        try:
            self.speak(f"Closing {app_name}.")
            subprocess.Popen(["osascript", "-e", f'tell application "{app_name}" to quit'])
        except Exception as e:
            self.speak(f"Unable to close {app_name}. Please try again.")
            print(f"Error: {e}")

    def read_file_content(self, file_path):
        """Read the content of a file."""
        try:
            with open(file_path, 'r') as file:
                content = file.read()
                self.speak(content)
        except FileNotFoundError:
            self.speak("File not found.")
        except Exception as e:
            self.speak(f"Error reading the file: {e}")

    def tell_time(self):
        """Tell the current time."""
        current_time = datetime.now().strftime("%H:%M:%S")
        self.speak(f"The current time is {current_time}.")

    def who_are_you(self):
        """Respond to the question of identity."""
        self.speak("I am Nova, your personal AI assistant. How can I assist you?")

    def track_gesture(self):
        gesture.track_gesture(self.speak)
    
    def analyze_window(self):
        """Invoke the screenshot and analyze functionality."""
        self.speak("Capturing and analyzing the window. Please wait")
        try:
            response_text = analyse.screenshot_and_analyze()
            self.speak(response_text)
            self.speak("Analysis completed !.")
        except Exception as e:
            self.speak("An error occurred while analyzing the window.")
            print(f"Error: {e}")

# Main program loop
def main():
    assistant = VoiceAssistant()
    print("Nova Voice Assistant Started!")
    assistant.speak("Hello, I am Nova, your personal assistant.")
    print("Available commands:")
    for command in assistant.commands:
        print(f"- {command}")

    while True:
        try:
            assistant.listen()
        except KeyboardInterrupt:
            assistant.speak("Goodbye!")
            print("\nExiting Nova Voice Assistant...")
            break
        except Exception as e:
            assistant.speak("An unexpected error occurred.")
            print(f"Error: {e}")

if __name__ == "__main__":
    main()
