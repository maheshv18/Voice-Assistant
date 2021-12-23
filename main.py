import pyttsx3
import datetime
import speech_recognition as sr
import webbrowser
import os
import pywhatkit
import cv2
import requests, json
import keyboard
import math
from geopy.geocoders import Nominatim
from pprint import pprint
import pyjokes
import wikipedia
import smtplib
from bs4 import BeautifulSoup as soup
from urllib.request import urlopen
import time
from glob import glob
import re
import win32com.client as win32
from win32com.client import constants
import easygui
import wolframalpha
from textblob import TextBlob


Mobile = {'Mahesh': 6384803545 }
app = Nominatim(user_agent="tutorial")

BASE_URL = "https://api.openweathermap.org/data/2.5/weather?"
CITY = "Coimbatore"
API_KEY = "API key"
URL = BASE_URL + "q=" + CITY + "&appid=" + API_KEY
response = requests.get(URL)

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)




def sendEmail(to, content):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
    server.login('your@email.com', 'pass')
    server.sendmail('to@mail.com', to, content)
    server.close()


def speak(audio):
    engine.say(audio)
    engine.runAndWait()


def greet():
    hour = int(datetime.datetime.now().hour)
    if hour >= 0 and hour < 12:
        speak("Good Morning!")
    elif hour >= 12 and hour < 18:
        speak("Good Afternoon")
    else:
        speak("Good Evening")
    speak(" How may i help?  ")

def chc():
    i=int(input("Enter ip choice\n0->write\n1->speak\n"))
    if(i==0):
        a='write'
    else:
        a='speak'
    return a


op=chc()
def takeUserAudio():

    if op=='write':
        query = input("Enter command-->")

    else:
        r = sr.Recognizer()
        with sr.Microphone() as source:
            print("Listening...")
            r.pause_threshold = 1.75
            audio = r.listen(source,timeout=15)

        try:
            print("Recognising..")
            query = r.recognize_google(audio, language='en-us')
            #query2 = input("Enter command-->")

            print("User said-> " + query)
        except Exception as e:
            print("Please repeat....Say that again")
            return "None"
    return query


if __name__ == "__main__":
    greet()
    while True:
        query = takeUserAudio().lower()
        if 'open youtube' in query:
            webbrowser.open("youtube.com")

        elif 'open google' in query:
            webbrowser.open('google.com')

        elif 'search' in query or 'find' in query:
            speak("What should I search?")
            srchstr = input()
            pywhatkit.search(srchstr)

        elif 'calendar' in query:
            webbrowser.open('https://calendar.google.com/calendar/u/0/r?tab=rc&pli=1')

        elif 'play' in query:
            speak("what should i play?")
            pl = input()
            pywhatkit.playonyt(pl)

        elif ('take my photo' in query or 'take picture' in query or 'open camera' in query):
            speak("Enter space to save picture and Escape to exit")
            cam = cv2.VideoCapture(0)
            cv2.namedWindow("test")
            img_counter = 0
            while True:
                ret, frame = cam.read()
                if not ret:
                    print("failed to grab frame")
                    break
                cv2.imshow("test", frame)

                k = cv2.waitKey(1)
                if k % 256 == 27:
                    print("Escape hit, closing...")
                    break
                elif k % 256 == 32:
                    img_name = "Pic_{}.png".format(img_counter)
                    cv2.imwrite(img_name, frame)
                    print("{} written!".format(img_name))
                    speak("Picture saved!! You may click space to save another,  escape to save and exit")
                    img_counter += 1

            cam.release()
            cv2.destroyAllWindows()

        elif ('whatsapp message' in query or 'send message' in query):
            flag = 0
            num = 0
            msg = ""
            speak("Enter name")
            name = input()
            num = Mobile[name]
            if num != 0:
                speak("Enter message")
                msg = input()
                if len(msg) != 0:
                    speak("Enter hour, minute")
                    h, m = map(int, input().split())
                    flag = 1
            try:
                if flag == 1:
                    pywhatkit.sendwhatmsg(f"+91{num}", msg, h, m, 10)
                    print("Successfully Sent!")
            except:
                print("An Unexpected Error!")
        elif "where is" in query:
            loc = query.replace("where is ", "")
            pywhatkit.search("Where is  " + loc)
            time.sleep(4)

        elif 'near me' in query:
            loc = query.replace("near me ", "")
            pywhatkit.search(loc + " near me")
            time.sleep(4)

        elif 'play music' in query:
            music = "C:\\Users\\UserName\\AppData\\Local\\Microsoft\\WindowsApps\\spotify.exe"
            os.startfile(music)

        elif 'time' in query :
            strTime = datetime.datetime.now().strftime("%H:%M:%S")
            speak("The time is " + strTime)
            print(strTime)

        elif ('to word' in query or 'convert to' in query):
            # os.system('cmd /c "pdf2docx gui"')
            paths = glob(easygui.fileopenbox(), recursive=True)


            def save_as_docx(path):
                # Opening MS Word
                word = win32.gencache.EnsureDispatch('Word.Application')
                doc = word.Documents.Open(path)
                doc.Activate()

                new_file_abs = os.path.abspath(path)
                new_file_abs = re.sub(r'\.\w+$', '.docx', new_file_abs)

                word.ActiveDocument.SaveAs(
                    new_file_abs, FileFormat=constants.wdFormatXMLDocument
                )
                doc.Close(False)
                codePath = new_file_abs
                os.startfile(codePath)


            for path in paths:
                save_as_docx(path)
            speak("pdf document has been converted to word")


        elif ('internet speed' in query or 'speedtest' in query):
            os.system('cmd /c "speedtest"') # pip install speedtest-cli

        elif 'open vs code' in query:
            codePath = "C:\\Users\\UserName\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe"
            os.startfile(codePath)

        elif 'open whatsapp' in query:
            codePath = "C:\\Users\\UserName\\AppData\\Local\\WhatsApp\\app-2.2123.8\\WhatsApp.exe"
            os.startfile(codePath)

        elif 'who are you' in query:
            speak(
                "I am your personal Voice Assistant I can do simple things like opening websites, desktop apps,tell you the time,play music, give weather info etc")

        elif ('weather' in query or 'temperature' in query):
            if response.status_code == 200:
                data = response.json()
                main = data['main']
                temperature = main['temp']
                humidity = main['humidity']
                pressure = main['pressure']
                report = data['weather']
                Celsius = math.trunc(temperature - 273.15)
                speak("Temperature is " + str(Celsius) + "Celsius, scroll down for more details")
                print(f"{CITY:-^30}")
                print(f"Temperature: {Celsius}")
                print(f"Humidity: {humidity}")
                print(f"Pressure: {pressure}")
                print(f"Weather Report: {report[0]['description']}")
                print("------------------------------")
            else:
                # showing the error message
                print("Error in the HTTP request")
        elif 'location info' in query:
            def get_location_by_address(address):

                try:
                    return app.geocode(address).raw
                except:
                    return get_location_by_address(address)


            speak("enter address")
            address = input()
            location = get_location_by_address(address)
            latitude = location["lat"]
            longitude = location["lon"]
            print(f"{latitude}, {longitude}")
            pprint(location)
            speak("Location information printed")

        elif 'joke' in query:
            a = pyjokes.get_joke()
            print(a)
            speak(a)

        elif 'am i correct' in query or 'correct me' in query or 'grammar check' in query:
            a = TextBlob("idk what doin now")
            a = a.correct()
            print(a)

        elif 'voice to text' in query:
            r = sr.Recognizer()
            with sr.Microphone() as source:
                audio = r.listen(source)
                print("Your input text-->")
                r.adjust_for_ambient_noise(source, duration=0.2)
                text = r.recognize_google(audio, language='en-us')
                print(text)

        elif 'wikipedia' in query:
            speak('Searching Wikipedia...')
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences=2)
            speak("According to Wikipedia")
            print(results)
            speak(results)

        elif 'art' in query:
            pywhatkit.image_to_ascii_art(easygui.fileopenbox())



        elif 'email' in query:
            try:
                speak("Enter 1 for simple, 2 with attachment")
                ch = int(input())

                if ch == 1:
                    speak("Enter the email address")
                    to = input()
                    speak("Enter the content")
                    content = input()
                    sendEmail(to, content)
                    speak("Email has been sent!")
                else:
                    mail = "C:\\Users\\UserName\\OneDrive\\Documents\\NetBeansProjects\\TimeTableManagement\\src\\SendingEmail.form"
                    os.startfile(mail)
                    time.sleep(17)
                    keyboard.send("Shift+F6")
                    time.sleep(4)

            except Exception as e:
                print(e)
                speak("Sorry . I am not able to send this email")
            time.sleep(4)

        elif "don't listen" in query or "stop listening" in query:
            speak("for how much time you want to stop jarvis from listening commands")
            a = int(input())
            time.sleep(a)
            print(a)

        elif 'news' in query:
            news_url = "https://news.google.com/news/rss"
            Client = urlopen(news_url)
            xml_page = Client.read()
            Client.close()
            soup_page = soup(xml_page, "xml")
            news_list = soup_page.findAll("item")
            speak("Enter number to be displayed!")
            i = int(input())
            for news in news_list:
                if i > 0:
                    print(news.title.text)
                    print(news.link.text)
                    print(news.pubDate.text)
                    print("-" * 60)
                    i = i - 1
            speak("have a great Day!)")

        elif "calculate" in query:
            app_id = "API key"
            client = wolframalpha.Client(app_id)
            indx = query.lower().split().index('calculate')
            query = query.split()[indx + 1:]
            res = client.query(' '.join(query))
            answer = next(res.results).text
            print("The answer is " + answer)
            speak("The answer is " + answer)

        elif "what is" in query or "who is" in query:
            app_id = "Api Key"
            client = wolframalpha.Client(app_id)
            res = client.query(query)
            try:
                print(next(res.results).text)
                speak(next(res.results).text)
            except StopIteration:
                print("No results")

        elif "don't listen" in query or "stop listening" in query or "sleep" in query:
            speak("for how much time you want me to stop listening commands")
            print("Enter-> ")
            a = int(input())
            print("sleeping....")
            speak("I am going to hibernate now")
            time.sleep(a)
            speak("I am awake")



        elif ('quit' in query or 'exit' in query):
            speak("Thankyou")
            exit()
        else:
            print('I am lost!')
