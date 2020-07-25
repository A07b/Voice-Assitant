#import pyttsx3
import datetime
import speech_recognition as sr
import wikipedia
import webbrowser
import os
import random
import pyyoutube
import youtube_search
import smtplib
import pywhatkit
from googletrans import  Translator
#import pyyoutube
import wolframalpha
from win32com.client import Dispatch
#import pyaudio



client = wolframalpha.Client("U2K2XP-XH36V3268G")
def speak(voice):
    
    pro = Dispatch('SAPI.SpVoice')
    pro.speak(voice)


def send_wh_msg(phoneno,msg):
    hr = int(datetime.datetime.now().hour)
    mi = (int(datetime.datetime.now().minute) + 2)
    try:
        pywhatkit.sendwhatmsg(phoneno,msg,hr,mi)
    except Exception as e:
        print(e)


def wish_me():
    hour = int(datetime.datetime.now().hour)
    if hour >= 0 and hour < 12:
        speak("Good Morning!")
    elif hour >= 12 and hour < 18:
        speak("Good Afternoon")
    else:
        speak("Good Evening")
    speak("I am Sunny, Please tell me how can i help you")


def take_command():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening....")
        r.pause_threshold = 1
        audio = r.listen(source,timeout=2)

    try:
        print("Recognizing....")
        query = r.recognize_google(audio, language='en-in')
        print(f"User Said: {query}\n")

    except Exception as e:
        #print(e)
        speak("Sorry.....Say that again please...")
        return "None"
    return query

def send_mail(sub,to,content):
    message ='Subject: {}\n\n{}'.format(sub, content)
    server = smtplib.SMTP('smtp.gmail.com',587)
    server.ehlo()
    server.starttls()
    server.login('abjain7218@gmail.com','Abhi@Sanju123')

    server.sendmail('abjain7218@gmail.com',to,message)
    server.close()

if __name__ == "__main__":
    
    speak("Hello Abhishek")
    wish_me()
    while True:
        query = take_command().lower()
        #all external activity on internet like open website  
        #wikipedia-----------------------------------------------------------------
        if 'wikipedia' in query:
            speak('Searching Wikipedia.....')
            query = query.replace('wikipedia', '')
            results = wikipedia.summary(query, sentences=3)
            speak("According To Wikipedia")
            print(results)
            speak(results)
        
        #youtube-----------------------------------------------------------------
        elif 'open youtube' in query:
            speak("opening youtube")
            webbrowser.open("https://www.youtube.com/")

        elif 'start youtube' in query:
            speak("opening youtube")
            webbrowser.open("https://www.youtube.com/")

        elif 'youtube' in query:
            speak("opening youtube")
            webbrowser.open("https://www.youtube.com/")

        #instagram---------------------------------------------------------------

        elif 'open instagram' in query:
            speak("opening instagram")
            webbrowser.open("https://www.instagram.com/")

        elif 'open insta' in query:
            speak("opening instagram")
            webbrowser.open("https://www.instagram.com/")

        elif 'instagram' in query:
            speak("opening instagram")
            webbrowser.open("https://www.instagram.com/")

        #facebook----------------------------------------------------------------
        elif 'open facebook' in query:
            speak("opening facebook")
            webbrowser.open("https://www.facebook.com/")

        elif 'open fb' in query:
            speak("opening facebook")
            webbrowser.open("https://www.facebook.com/")

        elif 'facebook' in query:
            speak("opening facebook")
            webbrowser.open("https://www.facebook.com/")


        elif 'open whatsapp web' in query:
            speak("opening whatsapp")
            webbrowser.open("https://web.whatsapp.com/")

        elif 'open whatsapp' in query:
            speak("opening whatsapp")
            webbrowser.open("https://web.whatsapp.com/")

        elif 'whatsapp' in query:
            speak("opening whatsapp")
            webbrowser.open("https://web.whatsapp.com/")
        #news----------------------------------------------------------------------
        elif 'english news' in query:
            speak("opening hindustan times")
            webbrowser.open("https://www.hindustantimes.com/")
        
        elif 'hindi news' in query:
            speak("opening aaj tak")
            webbrowser.open("https://aajtak.intoday.in/")

        elif 'marathi news' in query:
            speak("opening A.B.P Maza")
            webbrowser.open("https://marathi.abplive.com/")

        elif 'todays news' in query:
            speak("which news you want to see English Hindi Or Marathi")
            news_qry = take_command().lower()
            if 'english' in news_qry:
                speak("opening hindustan times")
                webbrowser.open("https://www.hindustantimes.com/")
            elif 'hindi' in news_qry:
                speak("opening aaj tak")
                webbrowser.open("https://aajtak.intoday.in/")
            elif 'marathi' in news_qry:
                speak("opening A B P Maza")
                webbrowser.open("https://marathi.abplive.com/")
            else:
                speak("No News")
        
        elif 'latest news' in query:
            speak("which news you want to see English Hindi Or Marathi")
            news_qry = take_command().lower()
            if 'english' in news_qry:
                speak("opening hindustan times")
                webbrowser.open("https://www.hindustantimes.com/")
            elif 'hindi' in news_qry:
                speak("opening aaj tak")
                webbrowser.open("https://aajtak.intoday.in/")
            elif 'marathi' in news_qry:
                speak("opening A B P Maza")
                webbrowser.open("https://marathi.abplive.com/")
            else:
                speak("No News")

        elif 'show me news' in query:
            speak("which news you want to see English Hindi Or Marathi")
            news_qry = take_command().lower()
            if 'english' in news_qry:
                speak("opening hindustan times")
                webbrowser.open("https://www.hindustantimes.com/")
            elif 'hindi' in news_qry:
                speak("opening aaj tak")
                webbrowser.open("https://aajtak.intoday.in/")
            elif 'marathi' in news_qry:
                speak("opening A B P Maza")
                webbrowser.open("https://marathi.abplive.com/")
            else:
                speak("No News")

        elif 'news' in query:
            speak("which news you want to see English Hindi Or Marathi")
            news_qry = take_command().lower()
            if 'english' in news_qry:
                speak("opening hindustan times")
                webbrowser.open("https://www.hindustantimes.com/")
            elif 'hindi' in news_qry:
                speak("opening aaj tak")
                webbrowser.open("https://aajtak.intoday.in/")
            elif 'marathi' in news_qry:
                speak("opening A B P Maza")
                webbrowser.open("https://marathi.abplive.com/")
            else:
                speak("No News")
                
        #all internal activity like play music, open application
        #music---------------------------------------------------------------------------------------
        elif 'play music' in query:
            music_dir = 'E:\\Songs'
            songs = os.listdir(music_dir)
            print(songs)
            os.startfile(os.path.join(music_dir,songs[random.randint(0,15)]))

        elif 'play songs' in query:
            music_dir = 'E:\\Songs'
            songs = os.listdir(music_dir)
            print(songs)
            os.startfile(os.path.join(music_dir,songs[random.randint(0,15)]))

        elif 'play song' in query:
            music_dir = 'E:\\Songs'
            songs = os.listdir(music_dir)
            print(songs)
            os.startfile(os.path.join(music_dir,songs[random.randint(0,15)]))

        elif 'play audio' in query:
            music_dir = 'E:\\Songs'
            songs = os.listdir(music_dir)
            print(songs)
            os.startfile(os.path.join(music_dir,songs[random.randint(0,15)]))

        elif 'songs' in query:
            music_dir = 'E:\\Songs'
            songs = os.listdir(music_dir)
            print(songs)
            os.startfile(os.path.join(music_dir,songs[random.randint(0,15)]))

        elif 'music' in query:
            music_dir = 'E:\\Songs'
            songs = os.listdir(music_dir)
            print(songs)
            os.startfile(os.path.join(music_dir,songs[random.randint(0,15)]))

        #photos---------------------------------------------------------------------------------------

        elif 'show photos' in query:
            photos_dir = "F:\\Photoshop Express"
            photos = os.listdir(photos_dir)
            os.startfile(os.path.join(photos_dir,photos[random.randint(0,1500)]))
        
        elif 'open photos' in query:
            photos_dir = "F:\\Photoshop Express"
            photos = os.listdir(photos_dir)
            os.startfile(os.path.join(photos_dir,photos[random.randint(0,1500)]))

        elif 'open pictures' in query:
            photos_dir = "F:\\Photoshop Express"
            photos = os.listdir(photos_dir)
            os.startfile(os.path.join(photos_dir,photos[random.randint(0,1500)]))
        
        elif 'show pictures' in query:
            photos_dir = "F:\\Photoshop Express"
            photos = os.listdir(photos_dir)
            os.startfile(os.path.join(photos_dir,photos[random.randint(0,1500)]))

        elif 'pictures' in query:
            photos_dir = "F:\\Photoshop Express"
            photos = os.listdir(photos_dir)
            os.startfile(os.path.join(photos_dir,photos[random.randint(0,1500)]))

        elif 'photos' in query:
            photos_dir = "F:\\Photoshop Express"
            photos = os.listdir(photos_dir)
            os.startfile(os.path.join(photos_dir,photos[random.randint(0,1500)]))

       

        elif 'show photo' in query:
            photos_dir = "F:\\Photoshop Express"
            photos = os.listdir(photos_dir)
            os.startfile(os.path.join(photos_dir,photos[random.randint(0,1500)]))
        
        elif 'open photo' in query:
            photos_dir = "F:\\Photoshop Express"
            photos = os.listdir(photos_dir)
            os.startfile(os.path.join(photos_dir,photos[random.randint(0,1500)]))

        elif 'open picture' in query:
            photos_dir = "F:\\Photoshop Express"
            photos = os.listdir(photos_dir)
            os.startfile(os.path.join(photos_dir,photos[random.randint(0,1500)]))
        
        elif 'show picture' in query:
            photos_dir = "F:\\Photoshop Express"
            photos = os.listdir(photos_dir)
            os.startfile(os.path.join(photos_dir,photos[random.randint(0,1500)]))

        elif 'picture' in query:
            photos_dir = "F:\\Photoshop Express"
            photos = os.listdir(photos_dir)
            os.startfile(os.path.join(photos_dir,photos[random.randint(0,1500)]))

        elif 'photo' in query:
            photos_dir = "F:\\Photoshop Express"
            photos = os.listdir(photos_dir)
            os.startfile(os.path.join(photos_dir,photos[random.randint(0,1500)]))

        
        #date and time------------------------------------------------------------------------

        elif 'the time' in query:
            strtime = datetime.datetime.now().strftime("%H:%M:%S")
            speak(strtime)
        
        elif 'current time' in query:
            strtime = datetime.datetime.now().strftime("%H:%M:%S")
            speak(strtime)

        #applications openning-----------------------------------------------------------------
        #Word Doc
        elif 'open word document' in query:
            os.startfile(r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office\Microsoft Word 2010.lnk")    
        
        elif 'open word file' in query:
            os.startfile(r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office\Microsoft Word 2010.lnk") 

        elif 'open new word file' in query:
            os.startfile(r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office\Microsoft Word 2010.lnk")

        elif 'open new word document' in query:
            os.startfile(r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office\Microsoft Word 2010.lnk") 

        elif 'open word' in query:
            os.startfile(r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office\Microsoft Word 2010.lnk") 
        
        #Powerpoint
        elif 'open powerpoint slide' in query:
            os.startfile(r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office\Microsoft PowerPoint 2010.lnk")
        
        elif 'open new powerpoint slide' in query:
            os.startfile(r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office\Microsoft PowerPoint 2010.lnk")

        elif 'open powerpoint document' in query:
            os.startfile(r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office\Microsoft PowerPoint 2010.lnk")

        elif 'open new powerpoint document' in query:
            os.startfile(r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office\Microsoft PowerPoint 2010.lnk")

        elif 'open powerpoint' in query:
            os.startfile(r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office\Microsoft PowerPoint 2010.lnk")

        #Excel Sheet
        elif 'open new excel sheet' in query:
            os.startfile(r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office\Microsoft Excel 2010.lnk")
        
        elif 'open new excel document' in query:
            os.startfile(r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office\Microsoft Excel 2010.lnk")

        elif 'open new excel file' in query:
            os.startfile(r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office\Microsoft Excel 2010.lnk")
        
        elif 'open excel sheet' in query:
            os.startfile(r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office\Microsoft Excel 2010.lnk")
        
        elif 'open excel document' in query:
            os.startfile(r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office\Microsoft Excel 2010.lnk")
        
        elif 'open excel file' in query:
            os.startfile(r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office\Microsoft Excel 2010.lnk")
        
        elif 'open excel' in query:
            os.startfile(r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office\Microsoft Excel 2010.lnk")

        #obs recorder
        elif 'open screen recorder' in query:
            os.startfile(r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\OBS Studio\OBS Studio (64bit).lnk")

        elif 'start screen recorder' in query:
            os.startfile(r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\OBS Studio\OBS Studio (64bit).lnk")
        
        elif 'screen recorder' in query:
            os.startfile(r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs\OBS Studio\OBS Studio (64bit).lnk")
        

        #translation----------------------------------------------------------------------------
        elif 'translate' in query:
            sentence = str(input("Type Something\n"))
            trans_sent = Translator().translate(sentence, dest = 'en')
            print(trans_sent.text)
            speak(str(trans_sent.text))

        elif 'translate to english' in query:
            sentence = str(input("Type Something\n"))
            trans_sent = Translator().translate(sentence, dest = 'en')
            print(trans_sent.text)
            speak(str(trans_sent.text))

        elif 'send mail' in query:
            try:
                sub = str(input('please enter subject name\n'))
                to = str(input('please enter email id\n'))
                msg = str(input('please enter message you want to send\n'))
                send_mail(sub,to,msg)

            except Exception as e:
                print(e)
                speak('sorry.... im not able to send mail.... please try again')

        elif 'message' in query:
            phoneno = input("Enter Phone number\n")
            msg = input("Enter Message\n")
            send_wh_msg('+91'+phoneno,msg)

        else:
            newq = query
            speak('Searching...')
            try:
                try:
                    res = client.query(newq)
                    results = next(res.results).text
                    speak('Got it.')
                    speak('WOLFRAM-ALPHA says - ')
                    print(results)
                    speak(results)
                    
                except:
                    results = wikipedia.summary(query, sentences=2)
                    speak('Got it.')
                    speak('WIKIPEDIA says - ')
                    print(results)
                    speak(results)
        
            except:
                webbrowser.open('www.google.com')
        
