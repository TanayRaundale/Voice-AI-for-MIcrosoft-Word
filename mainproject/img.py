import requests
import os
import pyttsx3

def speak(audio):
    engine = pyttsx3.init('sapi5')
    voices = engine.getProperty('voices')
    engine.setProperty('voices', voices[0].id)
    engine.say(audio)
    engine.runAndWait()

def generate_images(topic):
    
    PIXABAY_API_KEY = '41925790-ffb825eb27f460e8ccf41f878'
    url = f'https://pixabay.com/api/?key={PIXABAY_API_KEY}&q={topic}&image_type=photo'
    response = requests.get(url)
    data = response.json()
    urls = []


    if 'hits' in data:
        for hit in data['hits']:
            urls.append(hit['largeImageURL'])
    else:
        print('Error in API response:', data['message'])

    for url in urls:
        print(url)

    os.mkdir(f'{topic}')

    i = 1
    for index, img_link in enumerate(urls):
        if i <= 5:
            img_data = requests.get(img_link).content
            with open(f"{topic}/" + str(index + 1) + '.jpg', 'wb+') as f:
                f.write(img_data)
            i = i + 1
        else:               
            f.close()
            speak(f"sir,I have generated the images on{topic}")
            speak(f"I will open the {topic} folder for you")
            os.startfile(fr"C:\Users\Dell\Desktop\final year project\mainproject\{topic}")
            break
