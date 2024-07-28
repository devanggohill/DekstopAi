import datetime
import webbrowser
import openai
import speech_recognition as sr
import win32com.client as wincom

from dotenv import load_dotenv
import os

load_dotenv()
openai.api_key = os.getenv("sk-proj-DOmbQmWh3DJrHbDByIk2T3BlbkFJDgnOxZrIRx36SAdrv2oD")

speak = wincom.Dispatch("SAPI.SpVoice")

def ai(prompt):
    try:
        response = openai.ChatCompletion.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "system", "content": "You are a helpful assistant."},
                {"role": "user", "content": prompt}
            ],
            temperature=1,
            max_tokens=256,
            top_p=1,
            frequency_penalty=0,
            presence_penalty=0
        )
        say(response['choices'][0]['message']['content'].strip())
    except Exception as e:
        print(f"Error: {e}")


def say(text):
    speak.Speak(text)

def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Adjusting for ambient noise...")
        r.adjust_for_ambient_noise(source)
        print("Listening...")
        audio = r.listen(source)
        try:
            print("Recognizing...")
            query = r.recognize_google(audio, language="en-in")
            print(f"User said: {query}\n")
        except sr.UnknownValueError:
            print("Sorry, I did not understand that.")
            query = "Sorry, I did not understand that."
        except sr.RequestError as e:
            print(f"Could not request results; {e}")
            query = "Sorry, the service is down."
        return query

if __name__ == "__main__":
    say("Hello, I am Jarvis AI")
    while True:
        print("Listening..........")
        query = takeCommand()
        #say(query)
        sites = [
            ["youtube", "https://youtube.com"],
            ["wikipedia", "https://www.wikipedia.com"],
            ["google", "https://www.google.com"],
            ["facebook", "https://www.facebook.com"],
            ["twitter", "https://www.twitter.com"],
            ["instagram", "https://www.instagram.com"],
            ["reddit", "https://www.reddit.com"],
            ["linkedin", "https://www.linkedin.com"],
            ["amazon", "https://www.amazon.com"],
            ["netflix", "https://www.netflix.com"]
        ]

        for site in sites:
            if f"open {site[0]}".lower() in query.lower():
                say(f"Opening {site[0]} sirr")
                webbrowser.open(site[1])

        if "the time".lower() in query.lower():
            strfTime = datetime.datetime.now().strftime("%H:%M:%S")
            say(f"the time is {strfTime}")

        if "using ai ".lower() in query.lower():
            ai(prompt=query)