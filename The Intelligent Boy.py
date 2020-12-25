from win32com.client import Dispatch
import pyttsx3
import datetime
import speech_recognition as sr
import wikipedia
import webbrowser

engine =pyttsx3.init('sapi5')
voices=engine.getProperty('voices')
engine.setProperty('voice',voices[1].id)
 
def speak(audio):
    engine.say(audio)
    engine.runAndWait()

def takeCommand():
    # It take input from user and return string outputs...
    r=sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold=1.0
        audio=r.listen(source)

    try:
        print("Recognizing...")
        query=r.recognize_google(audio, language='en-in')
        print("You Said: ",query)
    except Exception as e:
        # print(e)
        speak("I was not able to understand it properly, will you say that again?")    
        return "None"
    return query    

def wishMe():
    speak("Hey There!!!")
    hour=int(datetime.datetime.now().hour)
    if hour>=0 and hour<12:
        speak("Good Morning!!!")
    elif hour>=12 and hour<=16:
        speak("Good Afternoon!!!")
    else:
        speak("Good Evening!!!") 

    speak("I am an Intelligent Voice, always ready for your Service!!!. What can I do for you?")                

def preface():
    speak=("Here are the thing that I can do for you...")
    print(1.,"Search something in Wikipedia?, or")
    print(2.,"anything else...")
    return takeCommand().lower()    

if __name__ == "__main__":
    wishMe()
    while True:
        query=takeCommand().lower()
        #This will be for wikipedia
        if "wiki" in query :
            query=query.replace("wiki","")
            speak("searching for {} in wikipedia...".format(query))
            results=wikipedia.summary(query,sentences=5)
            speak("According to results from wikipedia")
            speak(results)
            speak("This is about {}".format(query))
        elif "wikipedia" in query:
            query=query.replace("wikipedia","")
            speak("searching for {} in wikipedia...".format(query))
            results=wikipedia.summary(query,sentences=5)
            speak("According to results from wikipedia")
            print(results)
            speak(results)
            speak("This is about {}".format(query))         
        elif 'tell'in query and 'me'in query and 'about'in query:      
            query=query.replace("tell","")
            query=query.replace("me","")
            query=query.replace("about","")
            speak("searching for {} in wikipedia...".format(query))
            results=wikipedia.summary(query,sentences=5)
            speak("According to results from wikipedia")
            print(results)
            speak(results)
            speak("This is about {}".format(query))
        elif 'open youtube' in query:
            speak("Opening Youtube for you!!")
            webbrowser.open("youtube.com")
        elif 'open codeforces' in query:
            speak("Opening codeForces for you!!")
            webbrowser.open("code forces.com")    
        elif 'open hackerrank' in query:
            speak("Opening hackerrank for you!!")
            webbrowser.open("hackerrank.com/dashboard")
        elif 'open Code Chef' in query:
            speak("Opening codechef for you!!")
            webbrowser.open("codechef.com")
        elif 'open github' in query:
            speak("Opening github for you!!")
            webbrowser.open("github.com/Priyank2912")
        elif 'open gmail' in query:
            speak("Opening gmail for you!!")
            webbrowser.open("https://mail.google.com/mail/u/0/#inbox")
        elif 'open mail' in query:
            speak("Opening university mail for you!!")
            webbrowser.open("https://mail.google.com/mail/u/1/#inbox")                