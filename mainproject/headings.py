import speech_recognition as sr
import pyttsx3
from docx import Document
from docx.shared import RGBColor
from webcolors import name_to_rgb, IntegerRGB
import os

def speak(audio):
    engine = pyttsx3.init('sapi5')
    voices = engine.getProperty('voices')
    engine.setProperty('voices', voices[0].id)
    engine.setProperty('rate', 160)
    engine.say(audio)
    engine.runAndWait()

def change_heading_color(doc, color):
    speak("changing the color...")
    for paragraph in doc.paragraphs:
        if paragraph.style.name.startswith('Title'):
            for run in paragraph.runs:
                if isinstance(color, IntegerRGB):
                    color = RGBColor(color.red, color.green, color.blue)
                run.font.color.rgb = color

# def get_color_from_voice():
#     recognizer = sr.Recognizer()
#     recognizer.energy_threshold = 4000
#     recognizer.pause_threshold = 0.8
#     while True:  # Continue listening until a color is detected
#         with sr.Microphone() as source:
#             speak("Please tell me the color in which you want the headings")
#             speak("Listening...")
        
#             audio = recognizer.listen(source,phrase_time_limit=3)
            
#             try:
#                 speak("Recognizing...")
#                 color_name = recognizer.recognize_google(audio).lower()
#                 print(f"Color detected: {color_name}")
#                 return name_to_rgb(color_name)
#             except sr.UnknownValueError:
#                 print("Could not understand audio.")
#             except sr.RequestError as e:
#                 print("Could not request results; {0}".format(e))



def take_info(docname,color):
    doc = Document(docname)

    change_heading_color(doc, color)

    doc.save(docname)
    speak("Color changed Successfully")
    speak("I will open the document for you")
    print(f"Modified document saved to {docname}")
    os.startfile(f"C:\\Users\\Dell\\Desktop\\final year project\\mainproject\\{docname}")


