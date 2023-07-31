import speech_recognition as sr
import os
import win32com.client
import webbrowser
import datetime
import random
import numpy as np
import subprocess
import smtplib
def wish_me():
    hour = datetime.datetime.now().hour
    greeting = ""
    if 0 <= hour < 12:
        greeting = "Good morning!"
    elif 12 <= hour < 18:
        greeting = "Good afternoon!"
    else:
        greeting = "Good evening!"

    say(greeting + " I am Jarvis, your virtual assistant. How can I assist you today?")
def open_application(application_name):
    try:
        subprocess.run(application_name)
    except FileNotFoundError:
        print(f"Error: {application_name} not found. Make sure it's installed or provide the correct path.")

def say(text):
        speaker = win32com.client.Dispatch("SAPI.SpVoice")
        speaker.speak(text)

def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        # r.pause_threshold =  0.6
        audio = r.listen(source)
        try:
            print("Recognizing...")
            query = r.recognize_google(audio, language="en-in")
            print(f"User said: {query}")
            return query
        except Exception as e:
            return "Some Error Occurred. Sorry from Jarvis"

if __name__ == '__main__':
    print('Welcome to Jarvis A.I')
    wish_me()
    while True:
        print("Listening...")
        query = takeCommand()
        # todo: Add more sites
        sites = [["youtube", "https://www.youtube.com"], ["wikipedia", "https://www.wikipedia.com"], ["google", "https://www.google.com"],["GPT","https://chat.openai.com/"],["bd","https://bard.google.com/"],["git","https://github.com/"],["linkedin","https://www.linkedin.com/in/"],["ttm","ttm.sh"]]
        for site in sites:
            if f"Open {site[0]}".lower() in query.lower():
                say(f"Opening {site[0]} sir...")
                webbrowser.open(site[1])
        # todo: Add a feature to play a specific son

        if "the time" in query:
            musicPath = "/Users/harry/Downloads/downfall-21371.mp3"
            hour = datetime.datetime.now().strftime("%H")
            min = datetime.datetime.now().strftime("%M")
            say(f"Sir time is {hour} bajke {min} minutes")

        elif "Jarvis Quit".lower() in query.lower():
            exit()
        elif "open notepad".lower() in query.lower():
            appli = r"C:\Users\Dell\OneDrive\Desktop\index.html"
            os.startfile(appli)
        elif "what's your name" .lower() in query.lower() or "What is your name".lower() in query.lower():
            say("My Master call me")
            say("Jarvis")
            print("My Master call me Jarvis")
        elif "restart".lower() in query.lower():
            subprocess.call(["shutdown", "/r"])
        elif "lock window". lower() in query.lower():
            say("locking the device")
            ctypes.windll.user32.LockWorkStation()

        elif "shutdown system".lower() in query.lower():
            say("Hold On a Sec ! Your system is on its way to shut down")
            subprocess.call('shutdown / p /f')

        elif "empty recycle bin".lower() in query.lower():
            winshell.recycle_bin().empty(confirm=False, show_progress=False, sound=True)
            say("Recycle Bin Recycled")
        elif "open notepad".lower() in query.lower():
            open_application("Notepad")
        elif "open file explorer".lower() in query.lower():
            open_application("File Explorer")
        elif "joke".lower() in query.lower():
            jokes = [
                "Why don't scientists trust atoms? Because they make up everything!",
                "Did you hear about the mathematician who’s afraid of negative numbers? He’ll stop at nothing to avoid them.",
                "Parallel lines have so much in common. It’s a shame they’ll never meet.",
                "What do you call a fish with no eyes? Fsh.",
                "I told my wife she was drawing her eyebrows too high. She looked surprised.",
            ]
            say(random.choice(jokes))
            exit()
        elif "exit".lower() in query.lower():
            say("Goodbye!")
            exit()
        elif "send email".lower() in query.lower():
            try:
                server = smtplib.SMTP("smtp.gmail.com", 587)
                server.starttls()
                say("Please enter your email:")
                email = input("Enter your email: ")
                say("Please enter your password:")
                password = input("Enter your password: ")
                server.login(email, password)

                say("What should I say in the email?")
                email_message = take_command()

                say("Whom should I send the email to?")
                recipient = input("Enter the recipient's email: ")

                server.sendmail(email, recipient, email_message)
                say("Email sent successfully!")
            except Exception as e:
                say("Sorry, I couldn't send the email. Please try again later.")
                print(e)
            exit()
        elif "open application".lower() in query.lower():
            say("Sure, which application would you like me to open?")
            app_name = takeCommand().lower()
            if "chrome" in app_name:
                webbrowser.open("chrome")
            elif "notepad".lower() in app_name:
                subprocess.Popen(["notepad.exe"])
            elif "calculator" in app_name:
                subprocess.Popen(["calc.exe"])
            elif "word" in app_name:
                subprocess.Popen(["winword"])
            elif "excel" in app_name:
                subprocess.Popen(["excel"])
            exit()
        else:
            print("Retry...")
            say("I'm sorry, I don't understand that command")
            exit()