import pyautogui
import time
import pyttsx3

def speak(audio):
    engine = pyttsx3.init('sapi5')
    voices = engine.getProperty('voices')
    engine.setProperty('voices', voices[0].id)
    engine.setProperty('rate',180)
    engine.say(audio)
    engine.runAndWait()

def add_page_border_to_word_document():
    # Open Word and create a new document
    # Navigate to the Page Layout tab
    speak("Adding page border to your file")
    pyautogui.hotkey('alt', 'p')

    # Click on the Page Borders button
    pyautogui.click(565, 83)  # Adjust the coordinates as per your screen resolution
    
    # Apply the border style (e.g., Box border)
    time.sleep(1)
    pyautogui.press('right')  # Select the Box option
    time.sleep(1)
    
    # Apply the border to all sides of the page
    pyautogui.press('tab', presses=3, interval=0.5)  # Move to the Apply to dropdown
    pyautogui.press('down', presses=3, interval=0.5)  # Select the All option
    
    # Close the Page Borders dialog
    pyautogui.click(827,532)
    speak("Border added successfully...")    
    
if __name__ == "__main__":
    add_page_border_to_word_document()


