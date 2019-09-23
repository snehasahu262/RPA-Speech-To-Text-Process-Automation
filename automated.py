import pyttsx3 #pip install pyttsx3
import speech_recognition as sr #pip install speechRecognition
import datetime
import wikipedia #pip install wikipedia
import webbrowser
import os
import smtplib
import time
import imaplib
import email

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
# print(voices[1].id)
engine.setProperty('voice', voices[0].id)


def speak(audio):
    engine.say(audio)
    engine.runAndWait()
    
speak("hello tell me how may I help you")


def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour>=0 and hour<12:
        speak("Good Morning!")

    elif hour>=12 and hour<18:
        speak("Good Afternoon!")   

    else:
        speak("Good Evening!")  

         

def takeCommand():
    #It takes microphone input from the user and returns string output

    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 1
        audio = r.listen(source)

    try:
        print("Recognizing...")    
        query = r.recognize_google(audio, language='en-in')
        print(f"User said: {query}\n")

    except Exception as e:
        # print(e)    
        print("Say that again please...")  
        return "None"
    return query

def sendEmail(to, content):

    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
    server.login('from_mail ie yourmail', 'yourp_password')
    server.sendmail('to_mail', to, content)
    server.close()
    speak('Email sent.')

if __name__ == "__main__":
    wishMe()
    while True:
    # if 1:
        query = takeCommand().lower()

        # Logic for executing tasks based on query
        if 'wikipedia' in query:
            speak('Searching Wikipedia...')
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences=2)
            speak("According to Wikipedia")
            print(results)
            speak(results)

        elif 'open youtube' in query:
            webbrowser.open("youtube.com")

        elif 'open google' in query:
            webbrowser.open("google.com")

        elif 'open stackoverflow' in query:
            webbrowser.open("stackoverflow.com")   


        elif 'play music' in query:
            music_dir = 'D:\\Non Critical\\songs\\Favorite Songs2'
            songs = os.listdir(music_dir)
            print(songs)    
            os.startfile(os.path.join(music_dir, songs[0]))

        elif 'the time' in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")    
            speak(f"Sir, the time is {strTime}")

        
            
        elif 'open word' in query:
            wordpath="C:/Program Files/Microsoft Office/root/Office16/WINWORD.exe"
            os.startfile(wordpath)
            
        elif 'open powerpoint' in query:
            powerpointpath="C:/Program Files/Microsoft Office/root/Office16/POWERPNT.exe"
            os.startfile(powerpointpath)

        elif 'email' in query:
            try:
            
                speak('What is the subject?')
                time.sleep(3)
                subject = takeCommand()
                speak('What should I say?')
                message = takeCommand()
                content = 'Subject: {}\n\n{}'.format(subject, message)
    
                #init gmail SMTP
                mail = smtplib.SMTP('smtp.gmail.com', 587)
    
                #identify to server
                mail.ehlo()
    
                #encrypt session
                mail.starttls()
    
                #login
                mail.login('snehasahu1920@gmail.com', 'welcome!2#')
    
                #send message
                mail.sendmail('snehasahu1920@gmail.com', 'ssahu@gemini-us.com', content)
    
                #end mail connection
                mail.close()
    
                speak('Email sent.')
            
            except Exception as e:
                print(e)
                speak("Sorry my friend. I am not able to send this email")
        
        elif 'read' in query:
            try:
                mail = imaplib.IMAP4_SSL('imap.gmail.com')
                mail.login("snehasahu1920@gmail.com","welcome!2#")
                mail.list()
                mail.select('inbox')
                result, data = mail.uid('search', None, "UNSEEN") # (ALL/UNSEEN)
                i = len(data[0].split())

                for x in range(i):
                    latest_email_uid = data[0].split()[x]
                    result, email_data = mail.uid('fetch', latest_email_uid, '(RFC822)')
                    # result, email_data = conn.store(num,'-FLAGS','\\Seen') 
                    # this might work to set flag to seen, if it doesn't already
                    raw_email = email_data[0][1]
                    raw_email_string = raw_email.decode('utf-8')
                    email_message = email.message_from_string(raw_email_string)
                
                    # Header Details
                    date_tuple = email.utils.parsedate_tz(email_message['Date'])
                    if date_tuple:
                        local_date = datetime.datetime.fromtimestamp(email.utils.mktime_tz(date_tuple))
                        local_message_date = "%s" %(str(local_date.strftime("%a, %d %b %Y %H:%M:%S")))
                    email_from = str(email.header.make_header(email.header.decode_header(email_message['From'])))
                    email_to = str(email.header.make_header(email.header.decode_header(email_message['To'])))
                    subject = str(email.header.make_header(email.header.decode_header(email_message['Subject'])))
                
                    # Body details
                    for part in email_message.walk():
                        if part.get_content_type() == "text/plain":
                            body = part.get_payload(decode=True)
                            file_name = "email_" + str(x) + ".txt"
                            output_file = open(file_name, 'w')
                            var=output_file.write("From: %s\nTo: %s\nDate: %s\nSubject: %s\n\nBody: \n\n%s" %(email_from, email_to,local_message_date, subject, body.decode('utf-8')))
                            print("#################")
                            print(body.decode('utf-8'))
                            speak(body.decode('utf-8'))
                            output_file.close()
                            #speak("read email")
                        else:
                            
                            continue
                
            except Exception as e:
                print(e)
                speak("Sorry not able to read mail")


















