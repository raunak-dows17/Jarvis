import pyttsx3
import speech_recognition as sr
import datetime
import wikipedia
import webbrowser
import os
import smtplib


# emails = {
#     "bhushanji" : "ritikbhushanmain@gmail.com",
#     "jay" : "jaytyagi08@gmail.com",
#     "mallika ma'am" : "mallika@samatrix.io",
#     "subh" : "subhkab143@gmail.com",
#     "abhinav" : "Abhinavsenapati2000@gmail.com",
#     "yash" : "yashverma20@gmail.com"
# }


engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)

def speak(audio):
    engine.say(audio)
    engine.runAndWait()

def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour >=0 and hour <=12:
        speak("Good Morning Sir!")
    
    elif hour >=12 and hour < 4:
        speak("Good Afternoon Sir!")
    
    else:
        speak("Good Evening Sir!")
    
    speak("I am Friday, How may I help you!")

def takeCommand():
    #It takes microphone input from the user and returns strin output
    
    r= sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 0.8
        audio = r.listen(source)
        
    try:
        print("Recognizing...")
        query = r.recognize_google(audio, language = 'en-In')
        print(f"{query}\n")
        
    except Exception as e:
        # print(e)
        
        print("Say That again Please....")
        return "None"
    return query

def sendEmail(to, content):
    server = smtplib.SMTP('smtp.gmail.com', 537)
    server.ehlo()
    server.starttls()
    server.login('raunak17pandey@gmail.com', '******')
    server.sendmail('raunak17pandey@gmail.com', to, content)
    server.close()

if __name__ == '__main__':
    wishMe()
    while True:
        query = takeCommand().lower()
        
        #logic for excequting tasks based on query
        if 'wikipedia' in query:
            speak('Searching Wikipedia....')
            query = query.replace('wikipedia', "")
            results = wikipedia.summary(query, sentences = 10)
            speak("According to Wikipedia")
            print(results)
            speak(results)

        elif 'how are you' in query:
            speak("I'm fine sir!...")

        elif 'who are you' in query:
            speak("I'm Friday as I told you earlier!...")
            
        elif 'thank you' in query:
            speak('My pleasure sir!.....')
            
        elif 'open youtube' in query:
            webbrowser.open("https://www.youtube.com")
            speak('Opening Youtube....')
            
        elif 'open google' in query:
            webbrowser.open("https://www.google.co.in")
            speak('Opening Google....')
            
        elif 'open bing' in query:
            webbrowser.open("https://www.bing.com/")
            speak('Opening Bing...')
        
        elif 'open stackoverflow' in query:
            webbrowser.open("https://stackoverflow.com/")
            speak('Opening Stackoverflow....')
            
        elif 'open classroom' in query:
            webbrowser.open("https://classroom.google.com/u/0/h")
            speak('Opening classroom....')
        
        elif 'open instagram' in query:
            webbrowser.open("https://www.instagram.com/")
            speak('Opening Instagram....')
        
        elif 'open facebook' in query:
            webbrowser.open("https://www.facebook.com")
            speak('Opening Facebook....')
        
        elif 'open twitter' in query:
            webbrowser.open("https://twitter.com")
            speak('Opening Twitter....')
        
        elif 'open amazon' in query:
            webbrowser.open("https://www.amazon.in")
            speak('Opening amazon....')
        
        elif 'open github' in query:
            webbrowser.open("https://github.com")
            speak('Opening GitHub....')
        
        elif 'open samatrix' in query:
            webbrowser.open("https://learning.samatrix.io/dashboard")
            speak('Opening Samatrix....')

        
        
        # elif 'play music' in query:
        #     music dir = 
        #     song = os.listdir(music_dir)
        #     print(songs)
        #     os.startfile(os.path.join(music_dir, songs[0]))
        
        elif 'the time' in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")
            speak(f"Sir, The time is {strTime}")
        
        elif 'open code' in query:
            codePath = "C:\\Users\\Raunak\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe"
            os.startfile(codePath)
            speak("Opening Code...")
        
        elif 'open chrome' in query:
            chromePath = "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"
            os.startfile(chromePath)
            speak("Opening Chrome...")
        
        elif 'open excel' in query:
            excelPath = "C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE"
            os.startfile(excelPath)
            speak("Opening Excel....")
        
        elif 'open edge' in query:
            edgePath = "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe"
            os.startfile(edgePath)
            speak("Opening Edge...")
        
        elif 'open onedrive' in query:
            drivePath = "C:\\Program Files\\Microsoft OneDrive\\OneDrive.exe"
            os.startfile(drivePath)
            speak("Opening Onedrive...")
        
        elif 'open onenote' in query:
            notePath = "C:\\Program Files\\Microsoft Office\\root\\Office16\\ONENOTE.EXE"
            os.startfile(notePath)
            speak("Opening Onenote...")
        
        elif 'open powerpoint' in query:
            powerpointPath = "C:\\Program Files\\Microsoft Office\\root\\Office16\\POWERPNT.EXE"
            os.startfile(powerpointPath)
            speak("Opening Powerpoint...")
        
        elif 'open word' in query:
            wordPath = "C:\\Program Files\\Microsoft Office\\root\\Office16\\WINWORD.EXE"
            os.startfile(wordPath)
            speak("Opening Word...")
        
        elif 'open image to pdf' in query:
            imgtopdfpath = "F:\\Imgtopdf\\dist\\imgconverter.exe"
            os.startfile(imgtopdfpath)
            speak("OPening image to pdf...")
        
        elif 'open Dr. Sanke' in query:
            snakepath = "F:\\Dr.Snake\\dist\\snake.exe"
            os.startfile(snakepath)
            speak("Opening Dr.Snake....")
            
        elif 'send email to me' in query:
            try:
                speak("What should I say?...")
                content = takeCommand()
                to = "raunak17pandey@gmail.com"
                sendEmail(to, content)
                speak("Email has been sent!")
            except Exception as e:
                print(e)
                speak("Sorry Sir, I'm not able to send this email!")
        
        elif 'good bye' in query:
            speak('bye sir!....')
            exit()