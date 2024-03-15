import datetime
import sys
import pyttsx3
import speech_recognition as sr
import ppt
import img
import realvoice
from PyQt5 import QtWidgets, QtCore, QtGui
from PyQt5.QtCore import QTimer, QDate, QTime, Qt, QThread
from PyQt5.QtGui import QMovie
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from PyQt5.uic import loadUiType
from final import Ui_MainWindow
import time
import invitation
import table2
import cgpt

def speak(audio):
    engine = pyttsx3.init('sapi5')
    voices = engine.getProperty('voices')
    engine.setProperty('voices', voices[0].id)
    engine.setProperty('rate', 180)
    engine.say(audio)
    engine.runAndWait()

def wishme():
    hour = datetime.datetime.now().hour
    if 0 <= hour < 12:
        speak("Good morning")
    elif 12 <= hour < 18:
        speak("Good afternoon")
    else:
        speak("Good evening")
    speak("Sir, I am at your assistance")

class MainThread(QThread):
    def __init__(self, docname):  # Receive docname as a parameter
        super(MainThread, self).__init__()
        self.presentation_in_progress = False
        self.image_generation_in_progress = False
        self.docname = docname  # Store docname

    def run(self):
        self.taskexecution()

    def takecommand(self):
        r = sr.Recognizer()
        with sr.Microphone() as source:
            speak("Listening")
            r.pause_threshold = 0.8
            audio = r.listen(source, phrase_time_limit=5)
        try:
            speak("Recognizing...")
            self.query = r.recognize_google(audio, language='en-in')
            print(self.query)
        except Exception as e:
            print(e)
            speak("Say that again")
            return "none"
        return self.query

    def taskexecution(self):
        wishme()
        while True:
            if not self.presentation_in_progress and not self.image_generation_in_progress:
                self.query = self.takecommand().lower()
                if "hello" in self.query:
                    speak("Hello Sir, how can I help you?")
                    

                elif "introduction" in self.query or "introduce" in self.query:
                    speak("Hi there I am Your word assistant AI generated for making the task of creating word documenst easier and for guding you throught microsoft word.You can use me to generate images,create essays,invitations and also to create some powerpoint presentations")
                    speak("Also feel free to ask me anything as I am integrated with Google bard too")
                    speak("Thank You!")

                elif "presentation" in self.query:
                    if "generate a powerpoint presentation on" in self.query:
                        topic = self.query.replace("generate a powerpoint presentation on", "")

                    elif "generate a presentation on" in self.query:
                        topic = self.query.replace("generate a presentation on", "")
                    speak(f"Generating presentation on {topic}")
                    speak("Sir, this might take a few seconds, I will open the presentation when it is ready!")
                    self.presentation_in_progress = True
                    try:
                        ppt.get_bot_response(topic)
                    except Exception as e:
                        print(e)
                    self.presentation_in_progress = False

                elif "images" in self.query or "image" in self.query:
                    topic = self.query.replace("generate images on"," ")
                    speak(f"Generating images on {topic}")
                    img.generate_images(topic)
                    self.image_generation_in_progress = True

                elif "invitation" in self.query:
                    speak("Sir, I am generating your invitations and will notify you when completed")
                    template_path = 'template2.docx'
                    invitation.fill_automatically('read.csv', template_path)

                elif "table" in self.query:
                    speak("Make Sure That you have provided the file name to me")
                    print("Document Name:", self.docname)
                    speak("Sure, sir.")
                    rows = None
                    while rows is None:
                        speak("Please tell me the number of rows.")
                        rows_text = self.takecommand()
                        try:
                            rows = int(rows_text)
                        except ValueError:
                            speak("Sorry, I didn't get that. Please provide a valid number.")

                    cols = None
                    while cols is None:
                        speak("Please tell me the number of columns.")
                        cols_text = self.takecommand()
                        try:
                            cols = int(cols_text)
                        except ValueError:
                            speak("Sorry, I didn't get that. Please provide a valid number.")

                    # Now that we have valid row and column inputs, call the function
                    table2.create_empty_table(rows, cols, self.docname)
                    speak("I am generating the table.")
           
                elif "give me access to google bard" in self.query or "give me access to google bird" in self.query  or "give me access to google word" in self.query:
                    speak("Sure sir,Now you have the acces to the Google bard")
                    speak("Please tell me what can I do for you")
                    topic = self.takecommand()
                    topic  = topic.replace("generate information on",' ')
                    speak(f"Generating information about{topic}")
                    cgpt.generate_info(topic)
                elif "text editor" in self.query:
                      speak("Sure Sir, moving you to the Microsoft Word text Editor")
                      realvoice.text_editor()
                elif "stop" in self.query:
                    speak("Thank You sir for using me")
                    self.close()
            time.sleep(1)  # Add a small delay to prevent excessive CPU usage
            # After image generation is complete, allow listening again
            if self.image_generation_in_progress:
                self.image_generation_in_progress = False

startExecution = None

class Main(QMainWindow):
    
    def __init__(self):
        super().__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.ui.pushButton_3.clicked.connect(self.startTask)
        self.ui.lineEdit.setText("")
        self.ui.pushButton_4.clicked.connect(self.get_text)  
        self.docname = ""  # Initialize docname variable

    def startTask(self):
        global startExecution
        self.ui.movie = QtGui.QMovie("XDZT.gif")
        self.ui.label_2.setMovie(self.ui.movie)
        self.ui.movie.start()
        docname = self.ui.lineEdit.text()  # Retrieve text from QLineEdit
        startExecution = MainThread(docname)  # Pass docname to MainThread
        startExecution.start()  # Start the thread

    def get_text(self):
        self.docname = self.ui.lineEdit.text()  # Retrieve text from QLineEdit

app = QApplication(sys.argv)
word = Main()
word.show()
exit(app.exec_())