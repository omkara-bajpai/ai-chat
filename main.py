import speech_recognition as sr
from speech_recognition.exceptions import *
import win32com.client
import webbrowser
import wikipedia
import datetime
import requests
import json
import os

city = 'pune'
speaker = win32com.client.Dispatch('SAPI.SpVoice')


def getweather():
    data = requests.get(f'http://api.openweathermap.org/data/2.5/weather?q={city}&units=imperial&APPID=143454aa39bbe3442a890cdbf3f9db36').text
    data = json.loads(data)
    return data['main']['temp']


def say(text):
    speaker.Speak(text)


def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        audio = r.listen(source)
        final = ''
        try:
            print('Recognizing ......')
            final = r.recognize_google(audio, language='en-in')
            print(f'User Said : {final}')

            return final
        except RequestError:

            # Net issue
            print('Network issue')
            return final
        except:
            # issue
            print('Some error occurred')
            return final


def main():
    user_command = takeCommand()
    if user_command[:7].lower() == 'Jarvis '.lower():
        user_command = user_command[7:]
        sites = [['youtube', 'https://www.youtube.com'], ['wikipedia', 'https://www.wikipedia.org'],
                 ['bootstrap documentation', 'https://www.getbootstrap.com/docs']]
        for site in sites:
            if f'Open {site[0]}'.lower() in user_command.lower():
                print(f'Jarvis : Opening {site[0]}')
                say(f'Opening {site[0]}')
                webbrowser.open(site[1])

        apps = [['vs code', 'code'], ['pycharm', 'pycharm', 'C:\\Users\\Omkar\\PycharmProjects'], [
            'visual studio code', 'code'], ['hyper', 'hyper'], ['notepad', 'notepad'], ["fleet", 'fleet']]
        for app in apps:

            if f'Open {app[0]}'.lower() == user_command.lower():
                print(f'Jarvis : Opening {app[0]}')
                say(f'Opening {app[0]}')
                os.system(f'{app[1]}')
                if app[0] == 'pycharm':
                    os.system(f'explorer {app[2]}')
        if 'Search Wikipedia'.lower() == user_command.lower()[:16]:
            say('Searching Wikipedia')
            user_command = user_command.lower()
            user_command = user_command.replace('Search Wikipedia'.lower(), '')
            try:
                results = wikipedia.summary(user_command, sentences=4)

                print('Jarvis : According to Wikipedia')
                print(results)
                say('According to Wikipedia' + results)
            except wikipedia.exceptions.WikipediaException:
                print('Please give a value to search')

        elif 'tell the time' == user_command.lower():
            hour = datetime.datetime.now().strftime('%I')
            minute = datetime.datetime.now().strftime('%M')
            if str(hour)[0] == '0':
                if len(str(hour)) > 1:
                    hour = str(hour)[1:]
                else:
                    hour = ''
            if str(minute)[0] == '0':
                if len(str(minute)) > 1:
                    minute = str(minute)[1:]
                else:
                    minute = ''

            am_or_pm = datetime.datetime.now().strftime('%p')
            time = hour + ' ' + minute + ' ' + am_or_pm
            say('The current time is' + time)

            print('Jarvis : The current time is', end=' ')
            print(datetime.datetime.now().strftime('%I:%M %p'))
        elif user_command == '':
            print('Jarvis : I could not understand please try again')
            say('I could not understand please try again')
        elif user_command.lower() == 'Quit'.lower():
            exit()
        elif user_command.lower() == 'tell the weather':
            temp = getweather()
            print(f'The current weather is {temp}°F')
            say(f'The current weather is {temp} degree fahrenheit')

        else:
            # todo : Make openai working
            pass


if __name__ == '__main__':
    say('Hello I am Jarvis AI')
    while True:
        main()
