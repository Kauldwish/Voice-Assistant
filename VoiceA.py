import os
import re 
import webbrowser
import speech_recognition as sr
import requests
from gtts import gTTS
import win32com.client as wincl
import datetime
import random

chrome_path = 'C:/Program Files (x86)/Google/Chrome/Application/chrome.exe %s'
speak = wincl.Dispatch("SAPI.SpVoice")
currentDT = datetime.datetime.now()


def reply(audio):
    print(audio)
    for line in audio.splitlines():
        os.system("say" + audio)


def words():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print ('Hi i am The Rock your personal voice assistant Say Something!')
        audio = r.listen(source)
        print ('I heard what you said!')
        speak.Speak('I heard what you said!')
    try:
        command = r.recognize_google(audio)
        print('You said: ' + command + '\n')
    except sr.UnknownValueError:
        print('Your last command couldn\'t be heard, try again')
        speak.Speak('Your last command couldn\'t be heard, try again')
        command = words();
    return command
def assistant(command):
    if 'open website' in command:
        reg_ex = re.search('open website (.+)', command)
        if reg_ex:
            domain = reg_ex.group(1)
            url = 'https://www.' + domain + '.com'
            reply("i am opening it")
            speak.Speak("I am opening it")
            webbrowser.open(url)
            print('Done!')
        else:
            pass
    
    elif 'hi' in command:
        reply("hello there ")
        speak.Speak("hello there")
        
    elif 'how are you' in command:
        reply("I am fine wanna hear a joke?")
        speak.Speak("I am fine wanna hear a joke? ")
        
    elif 'bye' in command:
        reply('Bye user')
        speak.Speak("Bye user")
        sys.exit()
     
    elif 'who are you' in command:
        reply('I am Rock your personal Voice Assistant')
        speak.Speak('I am Rock our personal Voice Assistant')
        
    elif 'what is your name' in command:
        reply('My name is Rock')
        speak.Speak('My name is rock') 

    elif 'joke' in command:
        res = requests.get(
                'https://icanhazdadjoke.com/',
                headers={"Accept":"application/json"}
                )
        if res.status_code == requests.codes.ok:
            reply(str(res.json()['joke']))
            speak.Speak(str(res.json()['joke']))
        else:
            reply('oops!I ran out of jokes')
            
    elif 'date' in command:
        reply(str(currentDT))
        speak.Speak(currentDT.strftime("%Y-%m-%d "))
    elif 'time' in command:
        reply(currentDT.strftime("%I:%M:%S"))
        speak.Speak(currentDT.strftime("%I:%M:%S"))
        
        speak.Speak(currentDT.strftime("%I:%M:%S"))
    elif 'date and time' in command:
        reply(str(currentDT))
        speak.Speak(currentDT.strftime("%Y-%m-%d %H:%M:%S:"))
        
        
    elif 'day' in command:
        reply(currentDT.strftime("%a, %b %d, %Y"))
        speak.Speak(currentDT.strftime("%a, %b %d, %Y"))

    elif 'Google search' in command:
        reply('I am searching')
        speak.Speak('I am searching')
        command = 'https://www.google.co.in/search?q=' + command
        webbrowser.get(chrome_path).open(command)
        
    
    elif 'weather forecast' in command:
        reply('I am looking at it')
        speak.Speak('I am looking at it')
        command = 'https://www.google.co.in/search?q=' + command
        webbrowser.get(chrome_path).open(command)
        
reply('I am Here')

while True:
    assistant(words())
