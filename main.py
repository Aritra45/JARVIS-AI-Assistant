import speech_recognition as sr
import win32com.client
import os
import pyttsx3
import webbrowser
import pywhatkit
import pyjokes
import wikipedia

import datetime
import time
from time import sleep
import requests
from requests import get
from bs4 import BeautifulSoup
import pyautogui
import operator
from JarvisUi import Ui_JarvisUi
from PyQt5 import QtCore, QtGui, QtWidgets, sip
from PyQt5.QtCore import *
from PyQt5.QtGui import QMovie
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from PyQt5.QtCore import Qt, QTimer, QTime, QDate
from PyQt5.uic import loadUiType
import sys


speaker = win32com.client.Dispatch("SAPI.SpVoice")


def wish():
    try:
        hour = int(datetime.datetime.now().hour)
        if hour >= 0 and hour <= 12:
            print("Good Morning")
            speaker.Speak("Good Morning")
        elif hour > 12 and hour < 18:
            print("Good Afternoon")
            speaker.Speak("Good Afternoon")
        else:
            print("Good Evening")
            speaker.Speak("Good Evening")
        print("Hi I Am Jarvis On Your Service")
        speaker.Speak("Hi, I Am Jarvis, on your service")
    except Exception as e:
        print(f"An error occurred while speaking: {str(e)}")


class MainThread(QThread):
    def __init__(self):
        super(MainThread, self).__init__()

    def sipPyTypeDictRef(self):
        return super().sipPyTypeDictRef()

    def run(self):
        self.Task_Gui()

    def takeCommand(self):
        r = sr.Recognizer()
        with sr.Microphone() as source:
            r.pouse_threshold = 0.6
            audio = r.listen(source)
            try:
                self.query = r.recognize_google(audio, language="en-in")
                print(f"User Said: {self.query}")
                return self.query
            except Exception as e:
                return "Some Error Occurred. Sorry from jarvis"

    def Task_Gui(self):
        while 1:
            print("Listing...")
            self.query = self.takeCommand()

            if "open".lower() in self.query.lower():
                print("Launching...")
                speaker.Speak("Launching...")
                self.query = self.query.replace("open", "")
                self.query = self.query.replace("jarvis", "")

                pyautogui.press("super")
                pyautogui.typewrite(self.query)
                pyautogui.sleep(2)
                pyautogui.press("enter")

            elif "open".lower() in self.query.lower():
                from Dictapp import openappweb

                openappweb(self.query)

            elif "close".lower() in self.query.lower():
                from Dictapp import closeappweb

                closeappweb(self.query)

            elif (
                "hello".lower() in self.query.lower()
                or "hi jarvis".lower() in self.query.lower()
                or "what's up".lower() in self.query.lower()
                or "hey".lower() in self.query.lower()
            ):
                print("Hello! How can I assist you today?")
                speaker.Speak("Hello! How can I assist you today?")

            elif (
                "how are you".lower() in self.query.lower()
                or "how r u".lower() in self.query.lower()
            ):
                print(
                    "I'm just a computer program, so I don't have feelings, but I'm here and ready to help you with any questions or tasks you have. How can I assist you today? By The Way how are you?"
                )
                speaker.Speak(
                    "I'm just a computer program, so I don't have feelings, but I'm here and ready to help you with any questions or tasks you have. How can I assist you today? By The Way how are you?"
                )
            elif (
                "Thank You".lower() in self.query.lower()
                or "Thank u".lower() in self.query.lower()
                or "Thanks".lower() in self.query.lower()
                or "well done".lower() in self.query.lower()
            ):
                print(
                    "You're welcome! If you ever have more questions or need assistance in the future, don't hesitate to reach out. Have a great day!"
                )
                speaker.Speak(
                    "You're welcome! If you ever have more questions or need assistance in the future, don't hesitate to reach out. Have a great day!"
                )

            elif (
                "fine".lower() in self.query.lower()
                or "excellent".lower() in self.query.lower()
                or "great".lower() in self.query.lower()
            ):
                print(
                    "I'm glad to hear that you're doing fine! If you have any questions or if there's anything specific you'd like to know or discuss, feel free to let me know. I'm here to help with any information or assistance you may need."
                )
                speaker.Speak(
                    "I'm glad to hear that you're doing fine! If you have any questions or if there's anything specific you'd like to know or discuss, feel free to let me know. I'm here to help with any information or assistance you may need."
                )

            elif (
                "tell me about you".lower() in self.query.lower()
                or "tell me about yourself".lower() in self.query.lower()
                or "who are you".lower() in self.query.lower()
            ):
                print(
                    "J.A.R.V.I.S. (Just A Rather Very Intelligent System) is a fictional character voiced by Paul Bettany in the Marvel Cinematic Universe (MCU) film franchise, based on the Marvel Comics characters Edwin Jarvis and H.O.M.E.R., respectively the household butler of the Stark family and another AI designed by Stark. J.A.R.V.I.S. is an artificial intelligence created by Tony Stark, who later controls his Iron Man and Hulkbuster armor for him. In Avengers: Age of Ultron, after being partially destroyed by Ultron, J.A.R.V.I.S. is given physical form as Vision, physically portrayed by Bettany. Different versions of the character also appear in comics published by Marvel Comics, depicted as AI designed by Iron Man and Nadia van Dyne."
                )
                speaker.Speak(
                    "J.A.R.V.I.S. (Just A Rather Very Intelligent System) is a fictional character voiced by Paul Bettany in the Marvel Cinematic Universe (MCU) film franchise, based on the Marvel Comics characters Edwin Jarvis and H.O.M.E.R., respectively the household butler of the Stark family and another AI designed by Stark. J.A.R.V.I.S. is an artificial intelligence created by Tony Stark, who later controls his Iron Man and Hulkbuster armor for him. In Avengers: Age of Ultron, after being partially destroyed by Ultron, J.A.R.V.I.S. is given physical form as Vision, physically portrayed by Bettany. Different versions of the character also appear in comics published by Marvel Comics, depicted as AI designed by Iron Man and Nadia van Dyne."
                )

            elif (
                "gf".lower() in self.query.lower()
                or "bf".lower() in self.query.lower()
                or "friend".lower() in self.query.lower()
            ):
                print("Sorry!!! I Can't Provide You")
                speaker.Speak("Sorry!!! I Can't Provide You")

            elif "The Time".lower() in self.query.lower():
                strfTime = datetime.datetime.now().strftime("%H:%M%p")
                print(strfTime)
                speaker.Speak(f"It's, {strfTime}")

            elif "The Year".lower() in self.query.lower():
                strfTime = datetime.datetime.now().strftime("%Y")
                print(strfTime)
                speaker.Speak(f"It's, {strfTime}")

            elif "The Date".lower() in self.query.lower():
                strfTime = datetime.datetime.now().strftime("%d %B")
                print(strfTime)
                speaker.Speak(f"It's, {strfTime}")

            elif "The Month".lower() in self.query.lower():
                strfTime = datetime.datetime.now().strftime("%B")
                print(strfTime)
                speaker.Speak(f"It's, {strfTime}")

            elif "temperature".lower() in self.query.lower():
                search = "Temperature in Bally"
                url = f"https://www.google.com/search?q={search}"
                r = requests.get(url)
                data = BeautifulSoup(r.text, "html.parser")
                temp = data.find("div", class_="BNeawe").text
                print(f"Current {search} is {temp}")
                speaker.Speak(f"current{search} is {temp}")

            elif "ip address".lower() in self.query.lower():
                ip = get("https://api.ipify.org").text
                print(f"Your IP Address is {ip}")
                speaker.Speak(f"Your IP Address is {ip}")

            elif (
                "where i am".lower() in self.query.lower()
                or "my location".lower() in self.query.lower()
                or "the location".lower() in self.query.lower()
            ):
                print("wait, let me check...")
                speaker.Speak("wait sir, let me check")
                try:
                    ipAdd = requests.get("https://api.ipify.org").text
                    url = "https://get.geojs.io/v1/ip/geo/" + ipAdd + ".json"
                    geo_requests = requests.get(url)
                    geo_data = geo_requests.json()
                    city = geo_data["city"]
                    country = geo_data["country"]
                    print(
                        f"Sir I am Not Sure, But I Think We are in {city} city of {country} Country"
                    )
                    speaker.Speak(
                        f"Sir I am Not Sure, But I Think We are in {city} city of {country} Country"
                    )
                except Exception as e:
                    speaker.Speak(
                        "Sorry, Due To Network Issue I am Not Able To Find The Location"
                    )
                    pass

            elif "screenshot".lower() in self.query.lower():
                print("Please Tell Me The Name For This Screenshot File")
                speaker.Speak("Please tell me the name for this screenshot file")
                print("Listing...")
                name = self.takeCommand().lower()
                print("Please Hold The Screen For Few Seconds, I Am Taking Screenshot")
                speaker.Speak(
                    "please hold the screen for few seconds, i am taking screenshot"
                )
                time.sleep(3)
                img = pyautogui.screenshot()
                img.save(f"{name}.png")
                speaker.Speak("I am done , the screenshot is saved in out main folder")

            elif (
                "calculate" in self.query.lower() or "calculation" in self.query.lower()
            ):
                from calculator import Calculator

                Calculator(self.query)

            elif "turn on bluetooth".lower() in self.query.lower():
                print("Turning on bluetooh")
                speaker.Speak("Turning on bluetooh")
                pyautogui.click(x=1883, y=1070)
                sleep(2)
                pyautogui.click(x=1735, y=641)
                sleep(1)
                pyautogui.hotkey("win", "k")
                pyautogui.hotkey("tab")
                pyautogui.hotkey("enter")

            elif "turn off bluetooth".lower() in self.query.lower():
                print("Turning off bluetooh")
                speaker.Speak("Turning off bluetooh")
                pyautogui.click(x=1883, y=1070)
                sleep(2)
                pyautogui.click(x=1735, y=641)
                sleep(1)

            elif (
                "turn off wifi".lower() in self.query.lower()
                or "turn off wi-fi".lower() in self.query.lower()
            ):
                print("Turning off wifi")
                speaker.Speak("Turning off wifi")
                pyautogui.click(x=1666, y=1052)
                sleep(2)
                pyautogui.click(x=1537, y=989)
                sleep(1)

            elif (
                "search".lower() in self.query.lower()
                or "what is the".lower() in self.query.lower()
                or "what is".lower() in self.query.lower()
                or "how".lower() in self.query.lower()
            ):
                self.query = self.query.replace("Search", "")
                self.query = self.query.replace("search", "")
                self.query = self.query.replace("jarvis", "")
                self.query = self.query.replace("Jarvis", "")
                self.query = self.query.replace(" ", "")
                webbrowser.open(f"{self.query}")
                print(f"Searching...")
                speaker.Speak(f"searching... {self.query}")

            elif "play".lower() in self.query.lower():
                self.query = self.query.replace("jarvis", "")
                self.query = self.query.replace("Jarvis", "")
                song = self.query.replace("play", "")
                speaker.Speak(f"playing  {song}")
                pywhatkit.playonyt(song)

            elif "pause".lower() in self.query.lower():
                pyautogui.press("k")
                speaker.Speak("Video Paused")

            elif "Resume".lower() in self.query.lower():
                pyautogui.press("k")
                speaker.Speak("Video Resumed")

            elif "mute".lower() in self.query.lower():
                pyautogui.press("m")
                speaker.Speak("Video Muted")

            elif "unmute".lower() in self.query.lower():
                pyautogui.press("m")
                speaker.Speak("Video Un-muted")

            elif (
                "volume up".lower() in self.query.lower()
                or "increase the volume".lower() in self.query.lower()
                or "increase volume".lower() in self.query.lower()
                or "volume increase".lower() in self.query.lower()
            ):
                from keyboard import volumeup

                speaker.Speak("Increasing Volume")
                volumeup()

            elif (
                "volume down".lower() in self.query.lower()
                or "decrease the volume".lower() in self.query.lower()
                or "idecrease volume".lower() in self.query.lower()
                or "volume decrease".lower() in self.query.lower()
            ):
                from keyboard import volumedown

                speaker.Speak("Decreasing Volume")
                volumedown()

            elif "joke".lower() in self.query.lower():
                speaker.Speak(pyjokes.get_joke())

            elif "who the heck is".lower() in self.query.lower():
                person = self.query.replace("who the heck is", "")
                try:
                    info = wikipedia.summary(person, 1)
                    print(info)
                    speaker.Speak(info)
                except Exception as e:
                    speaker.Speak("I can't find...sorry from jarvis")

            elif (
                "stop".lower() in self.query.lower()
                or "bye".lower() in self.query.lower()
                or "exit".lower() in self.query.lower()
                or "quit".lower() in self.query.lower()
            ):
                print("By By...Have A Good Day")
                speaker.Speak("Bye Bye...have a good day")
                exit()


startExe = MainThread()


class Gui_Start(QMainWindow):
    def __init__(self):
        super().__init__()
        self.gui = Ui_JarvisUi()
        self.gui.setupUi(self)
        self.gui.Btn_1.clicked.connect(self.startTask)
        self.gui.Btn_2.clicked.connect(self.close)

    def startTask(self):
        wish()
        self.gui.label1 = QtGui.QMovie("..//G.U.I Material//B.G//Iron_Template_1.gif")
        self.gui.Gif_1.setMovie(self.gui.label1)
        self.gui.label1.start()

        self.gui.label2 = QtGui.QMovie("..//G.U.I Material//ExtraGui//Earth.gif")
        self.gui.Gif_2.setMovie(self.gui.label2)
        self.gui.label2.start()

        self.gui.label3 = QtGui.QMovie(
            "..//G.U.I Material//ExtraGui//B.G_Template_1.gif"
        )
        self.gui.Gif_3.setMovie(self.gui.label3)
        self.gui.label3.start()

        self.gui.label4 = QtGui.QMovie("..//G.U.I Material//VoiceReg//jarvis_jj.gif")
        self.gui.Gif_4.setMovie(self.gui.label4)
        self.gui.label4.start()

        self.gui.label5 = QtGui.QMovie("..//G.U.I Material//VoiceReg//__1.gif")
        self.gui.Gif_5.setMovie(self.gui.label5)
        self.gui.label5.start()

        self.gui.label6 = QtGui.QMovie("..//G.U.I Material//ExtraGui//initial.gif")
        self.gui.Gif_6.setMovie(self.gui.label6)
        self.gui.label6.start()

        self.gui.label7 = QtGui.QMovie(
            "..//G.U.I Material//ExtraGui//Health_Template.gif"
        )
        self.gui.Gif_7.setMovie(self.gui.label7)
        self.gui.label7.start()

        t_ime = QTime.currentTime()
        time = t_ime.toString()
        d_ate = QDate.currentDate()
        date = d_ate.toString()
        label_time = "Time : " + time
        label_date = "Date : " + date

        self.gui.Text_Time.setText(label_time)
        self.gui.Text_Date.setText(label_date)

        startExe.start()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    jarvis_gui = Gui_Start()
    jarvis_gui.show()
    sys.exit(app.exec_())
