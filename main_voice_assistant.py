import subprocess
import pygame
import pyttsx3
import datetime
import speech_recognition as sr
import wikipedia
import webbrowser
import os
from email.message import EmailMessage
import ssl
import smtplib
import random
import requests
from bs4 import BeautifulSoup
import time
import playsound
import pyautogui
import pyjokes
import pywhatkit
from win32com.client import GetObject

pygame.init()

white = (255, 255, 255)
red = (255, 0, 0)
black = (0, 0, 0)

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')

engine.setProperty('voice', voices[2].id)
# WakeUp = False
global query

screen_width = 400
screen_height = 400

gameWindow = pygame.display.set_mode((screen_width, screen_height))

bgimg = pygame.image.load(
    "C:/Users/aryan/Desktop/python/Project/img/robo_sleep.jpg")
bgimg = pygame.transform.scale(
    bgimg, (screen_width, screen_height)).convert_alpha()
# Game Title
pygame.display.set_caption("Pushpa")
pygame.display.update()


clock = pygame.time.Clock()
font = pygame.font.SysFont(None, 20)  # select font from system
welcome_font = pygame.font.SysFont(None, 40)  # select font from system


def speak(audio):
    engine.say(audio)
    engine.runAndWait()


def wishMe():
    speak("Hello Boss!")
    hour = int(datetime.datetime.now().hour)
    if (hour >= 0 and hour < 12):
        speak("Good Morning")
    elif (hour >= 12 and hour < 18):
        speak("Good Afternoon")
    else:
        speak("Good Evening")

    # speak("I am pushpa sir, How may i help you?")
    speak("I am pushpa, How may I help you?")


def firstCommand():
    try:
        r = sr.Recognizer()
        with sr.Microphone() as source:
            r.pause_threshold = 1
            audio = r.listen(source)
            query = r.recognize_google(audio)
            print(f"User said: {query}\n")
    except:
        # firstCommand()
        return 'None'
    return query


def takeCommand():
    try:
        r = sr.Recognizer()
        with sr.Microphone() as source:
            print("listening...")
            r.pause_threshold = 1
            audio = r.listen(source)
            print("Recognizing...")
            query = r.recognize_google(audio)
            print(f"User said: {query}\n")

    except:
        speak("Sorry I can't recognize your command. Please, say the command again")
        # takeCommand()
        return 'None'
    return query


def welcome():
    exit_assistant = False
    while not exit_assistant:
        gameWindow.fill((45, 43, 85))
        bgimg = pygame.image.load(
            "C:/Users/aryan/Desktop/python/Project/img/robo_sleep.jpg")
        bgimg = pygame.transform.scale(
            bgimg, (screen_width, screen_height)).convert_alpha()
        gameWindow.blit(bgimg, (0, 0))
        pygame.display.update()

        query = firstCommand().lower()
        if 'pushpa wake up' in query:
            # exit_assistant = True
            gameloop()

        for event in pygame.event.get():
            if event.type == pygame.QUIT:
                exit_assistant = True

        clock.tick(30)


# Game loop
def gameloop():
    exit_assistant = False
    quit_assistant = False
    fps = 30
    wished = False

    while not exit_assistant:
        if not quit_assistant:
            gameWindow.fill(white)
            gameoverimg = pygame.image.load(
                "C:/Users/aryan/Desktop/python/Project/img/robo_active.jpg")
            gameoverimg = pygame.transform.scale(
                gameoverimg, (screen_width, screen_height)).convert_alpha()
            gameWindow.blit(gameoverimg, (0, 0))
            pygame.display.update()
            for event in pygame.event.get():
                if event.type == pygame.QUIT:
                    exit_assistant = True

            if not wished:
                # time.sleep(3)
                wishMe()
                wished = True

        query = takeCommand().lower()

        if ('wikipedia' in query):
            speak("Searching wikipedia...")
            query = query.replace('wikipedia', '')
            result = wikipedia.summary(query, sentences=4)
            print(result)
            speak("According to wikipedia")
            speak(result)

        elif ('open youtube' in query):
            webbrowser.open('http://www.youtube.com/')

        elif ('open google' in query):
            webbrowser.open('http://www.google.com/')

        elif ('classroom' in query):
            webbrowser.open('https://classroom.google.com/u/0/h')

        elif ('ai class' in query):
            webbrowser.open(
                'https://classroom.google.com/u/0/c/NTA5MjEzMTE3MjU3')
        elif ('ai lab class' in query):
            webbrowser.open(
                'https://classroom.google.com/u/0/c/NTgxMDM2Mjc1Njc3')

        elif ('iot class' in query):
            webbrowser.open(
                'https://classroom.google.com/u/0/c/NTgxNzIxMTUxNjI5')

        elif ('graphics class' in query):
            webbrowser.open(
                'https://classroom.google.com/u/0/c/NTA5MjIxNjA0MDUy')

        elif ('play music' in query):
            speak("Playing music")
            music_dir = 'D:\\Music'
            songs = os.listdir(music_dir)
            num = random.randint(0, len(songs)-1)
            os.startfile(os.path.join(music_dir, songs[num]))

        elif ('time' in query):
            strtime = datetime.datetime.now().strftime("%I:%M %p")
            speak(f"The time is {strtime}")

        elif 'date' in query:
            today = datetime.date.today()
            d2 = str(today.strftime("%d %B, %Y"))
            speak(d2)

        elif 'day' in query:
            weeklist = ['monday', 'tuesday', 'wednesday',
                        'thursday', 'friday', 'saturday', 'sunday']
            dt = datetime.datetime.now()
            # get day of week as an integer from 0 to 6
            x = dt.weekday()
            day = weeklist[x]
            speak("Today is "+day)

        elif ('are you single' in query):
            speak("No I am not single. I have a lot of data")
            playsound.playsound(
                'C:\\Users\\aryan\\Desktop\\python\\Project\\music\\hahaha.mp3')
            time.sleep(1)

        elif 'tell me a joke' in query:
            My_joke = pyjokes.get_joke(language="en", category="all")
            speak(My_joke)
            playsound.playsound(
                'C:\\Users\\aryan\\Desktop\\python\\Project\\music\\hahaha.mp3')

        elif 'what kind of joke' in query:
            speak("Isn't it funny...! ... sorry")
            playsound.playsound(
                'C:\\Users\\aryan\\Desktop\\python\\Project\\music\\gameover.wav')

        elif ("are you married") in query:
            speak("As an AI model, I don't have the ability to get married or have a personal life. My purpose is to assist you with information and answer your questions to the best of my knowledge.")

        elif 'what is your name' in query:
            # speak("My name is Pushpa")
            playsound.playsound(
                'C:\\Users\\aryan\\Desktop\\python\\Project\\music\\Name.mp3')

        elif ('who are you') in query:
            # speak("I am Pushpa. an ai model developed by Python. My purpose is response to your queries and provide information.")
            playsound.playsound(
                'C:\\Users\\aryan\\Desktop\\python\\Project\\music\\Pushpa.mp3')

        elif 'search in google' in query:
            speak("What do you want to search")
            search = takeCommand().lower()
            # search = query.replace('search', '')
            url = 'https://google.com/search?q=' + search
            webbrowser.get().open(url)
            speak('Here is what i found about ' + search)

        elif "search in youtube" in query:
            speak("What do you want to search")
            search = takeCommand().lower()
            url = 'https://www.youtube.com/results?search_query=' + search
            webbrowser.get().open(url)
            speak("Here is what i found about " + search)

        elif "temperature" in query:
            base_url = "http://api.openweathermap.org/data/2.5/weather?q=dhaka,BAN&units=metric&appid=ea45752424c9cad83b4f5c836ced6b1a"
            data = requests.get(base_url).json()
            speak("According to Weather Report of DHAKA City")
            speak("Temperature:   " +
                  str(int(data['main']['temp']))+' degree Celsius')

        elif ('open vs code') in query:
            path = 'C:\\Users\\aryan\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe'
            os.startfile(path)

        elif "graphics project" in query:
            path = "C:\\Users\\aryan\\Desktop\\Computer Graphics\\Home\\build\\project.exe"
            os.startfile(path)

        elif ('open pycharm') in query:
            path = "C:\\Program Files\\JetBrains\\PyCharm Community Edition 2022.3.1\\bin\\pycharm64.exe"
            os.startfile(path)

        elif ('open browser') in query:
            path = "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"
            os.startfile(path)

        elif ('google chrome') in query:
            path = "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"
            os.startfile(path)

        elif ('edge') in query:
            path = "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe"
            os.startfile(path)

        elif ('open word') in query:
            path = "C:\\Program Files\\Microsoft Office\\root\\Office16\\WINWORD.EXE"
            os.startfile(path)

        elif ('open excel') in query:
            path = "C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE"
            os.startfile(path)

        elif ('open powerpoint') in query:
            path = "C:\\Program Files\\Microsoft Office\\root\Office16\\POWERPNT.EXE"
            os.startfile(path)

        elif ('open access') in query:
            path = "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe"
            os.startfile(path)

        elif ('open gmail') in query:
            webbrowser.open('https://mail.google.com/mail/u/2/#inbox')

        elif ('open facebook') in query:
            webbrowser.open('https://www.facebook.com/suvo242/')

        elif ('open github') in query:
            webbrowser.open('https://github.com/alemam242/')

        elif ('whatsapp') in query:
            speak("Tell me the phone number")
            while True:
                phone = takeCommand()
                phone = "+88" + phone.replace(" ", "")
                if len(phone) == 14:
                    break
                else:
                    speak(
                        "The phone number is not valid. Tell me the valid phone number")
            speak("What do you want to say?")
            while True:
                msg = takeCommand()
                if msg != "None":
                    break

            pywhatkit.sendwhatmsg_instantly(
                phone_no=phone,
                message=msg,
                wait_time=10,
                tab_close=True,
                close_time=5
            )
            speak("I have send the message")
            print(phone)
            print(msg)

        elif ('send mail to') in query:
            # name = query.replace('send mail to ', '')
            # print(name)
            try:
                speak("What should I say?")
                while True:
                    content = takeCommand()
                    if content != 'None':
                        break

                name = query.replace('send mail to ', '')
                print(name)
                if (name in ('imam', 'emam', 'suvo', 'shuvo', 'al imam', 'al emam')):
                    to = "alemam242@gmail.com"
                elif (name == 'rifat'):
                    to = "srrifat781@gmail.com"
                elif (name == 'naima' or name == 'nayema'):
                    to = "nayemasddika@gmail.com"

                email_sender = 'ecommerce.redhut.store@gmail.com'
                email_password = 'deqhpvlbubfuwjkq'

                subject = """Test email from Pushpa (Voice Assistant)"""
                print(to)
                # body = """This is a test email"""
                em = EmailMessage()
                em['From'] = email_sender
                em['To'] = to
                em['subject'] = subject
                em.set_content(content)

                context = ssl.create_default_context()

                with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=context) as smtp:
                    smtp.login(email_sender, email_password)
                    smtp.sendmail(email_sender, to, em.as_string())
                    speak("Email has been sent!")

            except:
                speak("Sorry, I am not able to send this email")

        elif ('great job') in query:
            speak("Thank you sir")

        elif ("toss" or "toss a coin") in query:
            moves = ["head", "tails"]
            cmove = random.choice(moves)
            playsound.playsound(
                'C:\\Users\\aryan\\Desktop\\python\\Project\\music\\toss.mp3')
            time.sleep(1)
            speak("It's " + cmove)

        elif ('note') in query:
            speak("What would you like to write down?")
            data = takeCommand().lower()
            date = datetime.datetime.now()
            filename = str(date).replace(':', '-')+'-note.txt'
            a = os.getcwd()
            if not os.path.exists('Notes'):
                os.mkdir('Notes')
            os.chdir(a+r'\Notes')
            with open(filename, 'w') as f:
                f.write(data)
            subprocess.Popen(['notepad.exe', filename])
            os.chdir(a)
            speak("I have made a note of that.")
            WMI = GetObject('winmgmts:')
            for process in WMI.ExecQuery('select * from Win32_Process where Name="Notepad.exe"'):
                os.system("taskkill /pid  "+str(process.ProcessId)+" /f")

        elif 'screenshot' in query:
            speak("Taking screenshot")
            img_captured = pyautogui.screenshot()
            a = os.getcwd()
            if not os.path.exists("Screenshots"):
                os.mkdir("Screenshots")
            os.chdir(a+'\Screenshots')
            ImageName = 'screenshot-' + \
                str(datetime.datetime.now()).replace(':', '-')+'.png'
            img_captured.save(ImageName)
            os.startfile(ImageName)
            os.chdir(a)
            speak('Captured screenshot is saved in Screenshots folder.')

        # elif ('quit' or 'exit' or 'bye' in query):
        elif "bye" in query:
            speak("Okay Sir, GoodBye. See you soon. Take care")
            # exit(0)
            exit_assistant = True

        elif "None" in query:
            speak("Sorry i can't recognize your command. Please say the command again")
        else:
            continue

        pygame.display.update()
        clock.tick(fps)

    pygame.quit()
    quit()


welcome()
