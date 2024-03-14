import re
from g4f.client import Client
import pyttsx3

def speak(audio):
    engine = pyttsx3.init('sapi5')
    voices = engine.getProperty('voices')
    engine.setProperty('voice', voices[0].id)
    engine.setProperty('rate', 150)

    # Remove URLs and '#' symbol using regular expressions
    audio = re.sub(r'http[s]?://(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\\(\\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+', '', audio)
    audio = audio.replace('#', '')

    engine.say(audio)
    engine.runAndWait()

def generate_info(topic):
    client = Client()
    response = client.chat.completions.create(
        model="gpt-3.5-turbo",
        messages=[{"role": "user", "content": "Generate the information about " + topic + " in just 50 words"}],
    )
    res = response.choices[0].message.content
    print(res)
    speak(res)

if __name__ == "__main__":
    generate_info("Tiger")
