import ctypes
from itertools import count
import json
import os
import pyodbc
import pyautogui
import shutil
import smtplib
import subprocess
import time

from typing import MutableSet
import webbrowser
from datetime import datetime
from urllib.request import urlopen
from winreg import QueryValue
import ecapture as ec
import winshell
import difflib
import pyjokes
import pyttsx3
import requests
import speech_recognition as sr
import wikipedia
import twilio
import pywhatkit
import pathlib
import AppOpener
from tkinter import Label, Button, Entry, StringVar, END, messagebox,filedialog
import tkinter as Tk

from email.mime.multipart import MIMEMultipart

from tkinter import filedialog

def window():
        window=Tk.Tk()
        lbl=Label(window, text="voice recognition system", fg='red', font=("Helvetica", 16))
        lbl=Label(window, text="Listening.............", fg='red', font=("Helvetica", 16))
        frame_a = Tk.Frame(master=window, width=100, height=100, bg="red")
        frame_a.pack(fill=Tk.BOTH ,side=Tk.LEFT, expand=True)

        label_a = Tk.Label(master=frame_a, text="speak....", fg='red', font=("Helvetica", 16))
        label_a.pack()
        lbl.place(x=50, y=200)
        window.title('voice recognition system')
        window.geometry("500x500+15+15")
        window.mainloop()

engine = pyttsx3.init("sapi5")
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)
assname = ("NARUTO 1 point o")

global query

def speak(audio):
    engine.say(audio)
    engine.runAndWait()

def lower(str):
    
    # Converting the given string to lower case
    # and return the string
    
    return str.lower()

def wishMe():
    hour = int(datetime.now().hour)
    if hour >= 0 and hour < 12:
        speak("Good Morning Sir !")

    elif hour >= 12 and hour < 18:
        speak("Good Afternoon Sir !")

    else:
        speak("Good Evening Sir !")

    
    speak("I am your Assistant")
    speak(assname)

def username():
   
    uname = "Shashank"
    speak("Welcome Mister " + uname)
    
    
    columns = shutil.get_terminal_size().columns

    print("#####################".center(columns))
    print("Welcome Mr.")
    print(uname)
    print("#####################".center(columns))

    speak("How can i Help you,")

def takeCommand():
    try:
        r = sr.Recognizer()
        with sr.Microphone() as source:
            r.adjust_for_ambient_noise(source)
            r.energy_threshold = 1000
            r.pause_threshold = 0.5
            r.dynamic_energy_threshold = True
            r.dynamic_energy_adjustment_damping = 0.15
            
            print("Listening...")
            
            audio = r.listen(source)
            print("Recognizing...")
            query= r.recognize_google(audio, language='en-in and kn-in')
            query=lower(query)
            print(query)
    except Exception as ex:
        
        print("exception")
        print(ex)
        print("Unable to Recognize your voice.")
        return "None"
    if len(query) > 0:
     try:
        file1 = open("MyFile.txt", "w+")
        file1.write("\n")
        file1.write(query)

        
        print("file written")
        
        b=file1.read()
        print(b)
        file1.close()
     except Exception as e:
        print("exception in file write")
        print(e)
        return " file error"

    return query
    
def sendEmail(to, content):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()

    # Enable low security in gmail
    server.login('your email id', 'your email password')
    server.sendmail('your email id', to, content)
    server.close()

def validate_ip(ipa):  
       
    # check number of periods  
    if ipa.count('.') != 3:  
        return 'Invalid Ip address'  
   
    ip_list = list(map(str, ipa.split('.')))  
   
    # check range of each number between periods  
    for element in ip_list:  
        if int(element) < 0 or int(element) > 255 or (element[0]=='0' and len(element)!=1):  
            return 'Invalid IP address'  
   
    return 'Valid IP address'   
      
if __name__ == '__main__' :

    speak("Hello ")
    def clear():
        return os.system('cls')

    # This Function will clean any
    # command before execution of this python file
    clear()
    wishMe()
    username()

    while True:

        query= takeCommand()
        query=lower(query)
        print(query)
        speak(query)
        

        
        if 'open browser' in query:
            webbrowser.open("https://www.google.com")

        elif 'open youtube' in query:
            speak("Here you go to Youtube\n")
            webbrowser.open("youtube.com")

        elif 'open' in query or 'open app' in query :
            appName = query.replace("open ", "")
            appName = appName.replace("app", "")    
            speak(f"opening{appName}\n")
            AppOpener.open(appName,match_closest=True)
            
        elif '''open  file manager''' in query or  'file manager' in query:
            speak("Here you go to file\n")
            
            difflib.get_close_matches(query, os.listdir())
            OpenFileDialog = filedialog.askopenfilename(initialdir = "/",title = "Select file",filetypes = (("Text files","*.txt"),("all files","*.*")))
            
        elif 'open google' in query:
            speak("Here you go to Google\n")
            webbrowser.open("google.com")

        elif "open settings" in query:
            os.system("C:\\Windows\\System32\\control.exe")
            
        elif 'open whatsapp' in query:
            webbrowser.open("web.whatsapp.com") 

        elif 'search' in query:
                    query = query.replace('search', '')
                    webbrowser.open(query)
                    
        elif 'system restart' in query:
                    speak('bye sir! hope you enjoyed the day')
                    os.system(("shutdown /r /t 1"))

        elif 'validate ip' in query:
            speak("Enter the ip address")
            query=takeCommand() or input()
            valid=validate_ip(query)
            print(valid)

        elif 'open stackoverflow' in query or 'open stack overflow' in query:
            speak("Here you go to Stack Over flow.Happy coding")
            webbrowser.open("stackoverflow.com")

        elif 'play music' in query or "play song" in query:
            speak("Here you go with music")
            # music_dir = "H:\\Song"
            music_dir = "H:\\music"
            songs = os.listdir(music_dir)
            print(songs)
            os.startfile(os.path.join(music_dir, songs[0]))

        elif 'the time' in query:
            strTime = datetime.now().strftime("%H:%M:%S")
            speak(f"Sir, the time is {strTime}")

        elif 'play video' in query:
            speak("Here you go with video")
            video_dir = "F:"
            videos = os.listdir(video_dir)
            print(videos)
            os.startfile(os.path.join(video_dir, videos[0]))

        elif "play" in query:
            song = query.replace('play', '')
            
            speak('playing ' + song)
            song= query.replace(' ', '+')
            webbrowser.open("https://www.youtube.com/results?search_query=" + song)

        elif 'email to shashank' in query:
            try:
                speak("What should I say?")
                content = takeCommand()
                to = "shashankvs2@outlook.com"
                sendEmail(to, content)
                speak("Email has been sent !")
            except Exception as e:
                print(e)
                speak("I am not able to send this email")

        elif 'send a mail' in query or 'send email' in query:
            try:
                speak("What should I say?")
                content = takeCommand()
                speak("whome should i send")
                print("enter the email id to send:")
                to = input()
                AppOpener.open("gmail.com",match_closest=True)
                speak("Email has been sent !")
            except Exception as e:
                print(e)
                speak("I am not able to send this email")

        elif 'how are you' in query:
            print("I am fine, Thank you")
            speak("I am fine, Thank you")
            speak("How are you, Sir")

        elif 'take a screenshot' in query:
            myScreenshot = pyautogui.screenshot()
            val=count()
            myScreenshot.save(r'C:\Users\User\Desktop\New folder\Screenshot'+str(next(val))+'.png')

        elif 'fine' in query or "good" in query:
            speak("It's good to know that your fine")
            print("It's good to know that your fine")
            
        elif "change my name to" in query:
            query = query.replace("change my name to", "")
            assname = query
            print("name chaged to", assname)
            speak("name chaged to "+assname)

        elif "change your name" in query:
            speak("What would you like to call me, Sir ")
            assname = takeCommand().lower()
            speak("Thanks for naming me")

        elif "what's your name" in query or "What is your name" in query:
            speak("My friends call me")
            speak(assname)
            print("My friends call me", assname)

        elif 'exit' in query or 'close' in query:
            speak("Thanks for giving me your time")
            exit()

        elif "who made you" in query or "who created you" in query:
            speak("I have been created by shashank")

        elif 'joke' in query:
            speak(pyjokes.get_joke())

        
        elif 'search' in query or 'play' in query:

            query = query.replace("search", "")
            query = query.replace("play", "")
            webbrowser.open(query)

        elif "why you came to world" in query:
            speak("Thanks to SHASHANK. further It's a secret")

        elif 'powerpoint presentation' in query or 'powerpoint' in query:
            speak("opening Power Point presentation")
            power = r"C:\Users\User\Desktop\Speech recognition.pptx"
            os.startfile(power)

        
        elif "who are you" in query:
            speak("I am your virtual assistant created by shashank")

        elif 'reason for you' in query:
            speak("I was created as a Minor project by Mister shashank")

        elif 'change background' in query:
            ctypes.windll.user32.SystemParametersInfoW(20,
                                                        0,
                                                        "Location of wallpaper",
                                                        0)
            speak("Background changed successfully")

        

        elif 'news' in query:

            try:
                jsonObj = urlopen(
                    '''https://newsapi.org / v1 / articles?source = the-times-of-india&sortBy = top&apiKey =\\times of India Api key\\''')
                data = json.load(jsonObj)
                i = 1

                speak('here are some top news from the times of india')
                print('''=============== TIMES OF INDIA ============''' + '\n')

                for item in data['articles']:

                    print(str(i) + '. ' + item['title'] + '\n')
                    print(item['description'] + '\n')
                    speak(str(i) + '. ' + item['title'] + '\n')
                    i += 1
            except Exception as e:

                print(str(e))

        elif 'lock window' in query:
            speak("locking the device")
            ctypes.windll.user32.LockWorkStation()

        elif 'shutdown system' in query:
            speak("Hold On a Sec ! Your system is on its way to shut down")
            subprocess.call('shutdown / p /f')

        elif 'empty recycle bin' in query:
            winshell.recycle_bin().empty(confirm=False, show_progress=False, sound=True)
            speak("Recycle Bin Recycled")

        

        elif "where is" in query:
            query = query.replace("where is", "")
            location = query
            speak("User asked to Locate")
            speak(location)
            webbrowser.open(
                "https://www.google. / maps / place/" + location + "")

        elif "camera" in query or "take a photo" in query:
            ec.capture(0, "Jarvis Camera ", "img.jpg")

        elif "restart" in query:
            subprocess.call(["shutdown", "/r"])

        elif "hibernate" in query or "sleep" in query:
            speak("Hibernating")
            subprocess.call("shutdown / h")

        elif "log off" in query or "sign out" in query:
            speak("Make sure all the application are closed before sign-out")
            time.sleep(5)
            subprocess.call(["shutdown", "/l"])

       
        elif 'naruto' in query:
            speak("Assistant naruto at your service sir")
            
            speak(assname)
        
     #   elif 'open ' in query:
      #      query = query.replace("open ", "")
         #   speak("Opening " + query)
         #   AppOpener.open(query,match_closest=True)

        elif 'weather' in query:

            # Google Open weather website
            # to get API of Open weather
            api_key = "Api key"
            base_url = "https://www.google.com/search?q=weather&rlz=1C1SQJL_enIN867IN867&oq=weather+&aqs=chrome..69i57j35i39i650l2j0i131i433i512l2j0i131i433i457i512j0i402i650l2j0i131i433i512l2.8804j1j4&sourceid=chrome&ie=UTF-8"
            response = requests.get(base_url)
            x = response.json()

            if x["code"] != "404":
                y = x["main"]
                current_temperature = y["temp"]
                current_pressure = y["pressure"]
                current_humidiy = y["humidity"]
                z = x["weather"]
                weather_description = z[0]["description"]
                print(" Temperature (in kelvin unit) = " + str(current_temperature)+"\n atmospheric pressure (in hPa unit) ="+str(
                    current_pressure) + "\n humidity (in percentage) = " + str(current_humidiy) + "\n description = " + str(weather_description))

            else:
                speak(" weather information Not Found ")

        elif "wikipedia" in query:
            webbrowser.open("wikipedia.com")

        elif "good morning" in query or 'good afternoon' in query or 'good evening' in query:
            speak("A warm" + query)
            speak("How are you Mister")
            speak(assname)

        # most asked question from google Assistant
        

        elif "how are you" in query:
            speak("I'm fine, glad you me that")

        elif "open notepad" in query:
            speak("opening notepad")
            os.system("notepad")

        elif 'show last command' in query or 'last command' in query or "what did i say" in query:
            speak('last command is')
            lc1=open("MyFile.txt","r")
            lc=lc1.read()
            speak(lc)
            lc1.close()            
        
        elif "" in query:
            speak ("say something")
    

