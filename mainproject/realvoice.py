import speech_recognition as sr
import pyautogui
import time
import pyttsx3
import pygetwindow as gw
import borders
import headings

def speak(audio):
    engine = pyttsx3.init('sapi5')
    voices = engine.getProperty('voices')
    engine.setProperty('voices', voices[0].id)
    engine.setProperty('rate',180)
    engine.say(audio)
    engine.runAndWait()

def text_editor():
    # Initialize the speech recognizer
    recognizer = sr.Recognizer()
    recognizer.pause_threshold = 0.5

    speak("Opening Microsoft Word...")

    # Open a new Word document
    pyautogui.press('win')
    time.sleep(1)  # Add a delay to ensure the start menu is open
    pyautogui.write('word')
    time.sleep(5)
    pyautogui.press('enter')
    time.sleep(5)  # Add a delay to allow Word to open

    # Initialize the text variable
    text = ""

    # Continuously listen for speech input and type it into Word
    running = True
    while running:
        with sr.Microphone() as source:
            speak("Listening...")
            audio = recognizer.listen(source,phrase_time_limit=5)

        try:
            speak("Recognizing...")
            # Recognize speech using Google Speech Recognition
            new_text = recognizer.recognize_google(audio)
            print("You said:", new_text)

            # Check if the command is "stop" to exit the loop
       # Check if the command is "stop" to exit the loop
            if new_text.lower() == "stop":
                print("Stopping the program...")
                running = False  # Set running to False to exit the loop
                break
            elif new_text.lower() in ["enter", "inter"]:
                pyautogui.press('enter')  # Press enter key to move to the next line
            elif new_text.lower() == "space":
                pyautogui.press('space')
            elif new_text.lower() == "save":
                speak("Sure Sir,Saving the document")
                pyautogui.hotkey('ctrl', 's')
                speak("Please enter the filename")
                time.sleep(3)
                pyautogui.press('enter') 
                speak("Document Saved successfully ")
            elif new_text.lower() == "create new file":
                pyautogui.hotkey('ctrl','n')
                speak("New file created successfully")
            elif new_text.lower() == "add border":
                borders.add_page_border_to_word_document()
             
            elif new_text.lower() == "justify the document":
                pyautogui.hotkey('ctrl','a')
                pyautogui.hotkey('ctrl','j')
                speak("Document Justified SUccessfully")
            elif new_text.lower() == "close":
                pyautogui.hotkey('alt','f4')
                speak("Thanks Sir for using our built in text editor")
                running = False  # Set running to False to exit the loop
                break

            else:
                # Update text only if new speech input is successfully recognized and it's not "enter" or "space"
                text = new_text

                # Type the recognized text into Word
                pyautogui.write(text)


        except sr.UnknownValueError:
            print("Google Speech Recognition could not understand audio")
        except sr.RequestError as e:
            print("Could not request results from Google Speech Recognition service; {0}".format(e))

if __name__ == "__main__":
    text_editor()
