import speech_recognition as sr
import win32com.client as wincl
from time import ctime
import webbrowser as web
import pywhatkit
import datetime

# from pywikihow import WikiHow, search_wikihow
# from geopy.distance import great_circle
# from geopy.geocoders import Nominatim
# import geocoder
# import os
# import time
# import wikipedia
# from gtts import gTTS
# import requests, json
# from googlesearch import search


speaker_number = 0
speaker = wincl.Dispatch("SAPI.SpVoice")
vcs = speaker.Getvoices()
# SVSFlag = 11
# print(vcs.Item (speaker_number) .GetAttribute ("Name"))
speaker.Voice
speaker.SetVoice(vcs.Item(speaker_number)) # set voice (see Windows Text-to-Speech settings)
# speaker.Speak("Hello, it works!")

def takecaommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.pause_threshold = 0.8
        audio = r.listen(source)
        try:
            print("Recognizing...")
            querry = r.recognize_google(audio, language="en-in")
            print(f"user said: {querry}")
            return querry
        except Exception as e:
            return "Sorry I can't understand. Can you say again."

if __name__ == '__main__':
    # print("Enter the word you want to speak")
    # s = input()
    speaker.Speak("Hello I am jarvis A.I. How can I help you")

    while True:
        print("Listening...")
        querry = takecaommand()
        sites = [["google", "https://www.google.com"],["flipkart", "https://www.flipkart.com"],["youtube", "https://www.youtube.com"],
                ["wikipedia", "https://www.wikipedia.com"],["amazon", "https://www.amazon.in"],["nasa", "https://www.nasa.gov"],
                ["instagram", "https://www.instagram.com"]]
        for site in sites:
            if f"Open {site[0]}".lower() in querry.lower():
                speaker.Speak(f"Opening {site[0]}")
                web.open(site[1])

        if 'play'.lower() in querry.lower():
            song = querry.replace('play', '')
            speaker.Speak(f"Playing {song}")
            pywhatkit.playonyt(song)

        folders = [["open music", "C:/Users/rawat/Music/"], ["open images", "C:/Users/rawat/Pictures/"],["C drive", "C:/"],
                ["open documents", "C:/Users/rawat/Documents/"],["open video", "C:/Users/rawat/Videos/"],
                ["vs code", "C:/Users/rawat/Desktop/Visual Studio Code.lnk"]]
        for folder in folders:
            if f"Open {folder[0]}".lower() in querry.lower():
                speaker.Speak(f"Opening {folder[0]}")
                web.open(folder[1])

        if "time" in querry:
            strf = datetime.datetime.now().strftime("%H:%M:%S")
            speaker.speak(f"The time is {strf}")
            

        # try:
        #     query = querry
 
        #     for j in search(query, tld="co.in", num=10, stop=10, pause=2):
        #         print(j)
        # except ImportError:
        #     print("No module named 'google' found")

        # to search

        # def googlesearch(term):
        #     queery = term.replace("what is","")
        #     queery = term.replace("where is","")
        #     queery = term.replace("who is","")
        #     queery = term.replace("how to","")
        #     queery = term.replace("why ","")
        #     queery = term.replace("when ","")
        #     # queery = term.replace(" ","")
        #     queery = term.replace("what do you mean by","")
        #     writeab = str(queery)

        #     da = open("C:\\LANGUAGES\\Python Course\\project\\Jarivis\\jarvisdata.txt", 'a')
        #     da.write(writeab)
        #     da.close()

        #     Queery = str(term)
        #     pywhatkit.search(Queery)

        #     if "how to" in Queery:
        #         max_result = 1
        #         how_to_fun = search_wikihow(query=Queery,max_results=max_result)
        #         assert len(how_to_fun) == 1
        #         how_to_fun[0].print()
        #         speaker.Speak(how_to_fun[0].summary)

        #     else:
        #         search = wikipedia.summary(Queery,2)
        #         speaker.Speak(f": According to your search: {search}")
        # googlesearch(querry)
        

        def digital_assistant(data):
            if "how are you" in data:
                listening = True
                speaker.Speak("I am well")

            if "what time is it" in data:
                listening = True
                speaker.Speak(ctime())

        if "where is" in querry:
            listening = True
            data = querry.split(" ")
            location_url = "https://www.google.com/maps/place/" + str(querry[9:len(querry)+1])
            speaker.Speak("Hold on, I will show you where " + querry[9:len(querry)+1] + " is.")
            web.open(location_url)
        digital_assistant(querry)
        
        # if "stop listening" in querry:
        #     listening = False
        #     print('Listening stopped')
        #     return listening
        # return listening
        
        # def googlemap(place):
            
        
        # speaker.Speak(querry)
        


