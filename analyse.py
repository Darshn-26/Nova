import os
import mss
import google.generativeai as genai
import pyttsx3


def speak(text):
    """Uses pyttsx3 to convert text to speech."""
    engine = pyttsx3.init()
    engine.say(text)
    engine.runAndWait()


def configure_gemini():
    """Configures Gemini with the API key from environment variables."""
    genai.configure(api_key='AIzaSyBZr3iIG2PQ12UPhg5h6NTm0GXu3fh-7G4')
    generation_config = {
        "temperature": 1,
        "top_p": 0.95,
        "top_k": 40,
        "max_output_tokens": 8192,
        "response_mime_type": "text/plain",
    }
    return genai.GenerativeModel(
        model_name="gemini-1.5-pro",
        generation_config=generation_config,
        system_instruction="Ignore terminal window image .Analyze the uploaded screenshot imageand provide the analysis (make sure use professional words and phrases) If u see educational documents summeraise with longer seponse but keep it simple .",
    )


def upload_to_gemini(path, mime_type=None):
    """Uploads the given file to Gemini."""
    file = genai.upload_file(path, mime_type=mime_type)
    print("Analysing ur active screen !")
    speak("Analysing ur active screen !")
    return file


def screenshot_and_analyze():
    """Takes a screenshot, uploads it to Gemini, and returns the analysis."""
    model = configure_gemini()

    # Take a screenshot and save it as a file
    with mss.mss() as sct:
        screenshot_folder = "Analyse"
        os.makedirs(screenshot_folder, exist_ok=True)
        screenshot_path = os.path.join(screenshot_folder, "snapshot.png")
        sct.shot(output=screenshot_path)
        print("Captured ur active screen starting analysis !")
        speak("Captured ur active screen !")

    # Upload the screenshot to Gemini
    uploaded_file = upload_to_gemini(screenshot_path, mime_type="image/png")

    # Start a chat session and send the image
    chat_session = model.start_chat(
        history=[
            {
                "role": "user",
                "parts": [uploaded_file],
            }
        ]
    )

    response = chat_session.send_message("Please analyze the content of this image.Ignore the terminal window here")
    response_text = response.text.strip()
    print("AI Response:", response_text)


    return response_text