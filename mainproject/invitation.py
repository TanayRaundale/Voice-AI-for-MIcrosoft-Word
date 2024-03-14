from docx import Document
import pandas as pd
import pyttsx3
import os
def speak(audio):
    engine = pyttsx3.init('sapi5')
    voices = engine.getProperty('voices')
    engine.setProperty('voices', voices[0].id)
    engine.say(audio)
    engine.runAndWait()

def fill_invitation(template_path, output_path, data):
    doc = Document(template_path)
    for paragraph in doc.paragraphs:
        for key, value in data.items():
            if key in paragraph.text:
                paragraph.text = paragraph.text.replace(key, value)
    doc.save(output_path)



def fill_automatically(csv_path, template_path):
    df = pd.read_csv(csv_path) #df=dataframe
    for index, row in df.iterrows():
        data = {
            '[sender name]': row['Sender_Name'],
            '[email]': row['Email'],
            '[r name]': row['Receiver_Name'],
            '[designation]': row['Designation'],
            '[company name]': row['Company_Name'],
            '[address]': row['Address']
        }
        output_path = f'invitations/invitation{index}.docx'
        fill_invitation(template_path, output_path, data)
    speak("I Have created the invitations for you,Sir")
    speak("I will open the folder for you")
    os.startfile(r"C:\Users\Dell\Desktop\final year project\mainproject\invitations")


