import os
import pyautogui
import webbrowser
import pyttsx3
import win32com.client
from time import sleep

speaker = win32com.client.Dispatch("SAPI.SpVoice")

dictapp = {
    "commandprompt": "cmd",
    "paint": "paint",
    "word": "winword",
    "excel": "excel",
    "chrome": "chrome",
    "vscode": "code",
    "powerpoint": "powerpnt",
}


def openappweb(query):
    if ".com" in query or ".co.in" in query or ".org" in query:
        query = query.replace("open", "")
        query = query.replace("jarvis", "")
        query = query.replace("Jarvis", "")
        query = query.replace("launch", "")
        query = query.replace(" ", "")
        webbrowser.open(f"https://www.{query}")

    else:
        keys = list(dictapp.keys())
        for app in keys:
            if app in query:
                os.system(f"start {dictapp[app]}")


def closeappweb(query):
    speaker.Speak("Closeing...")
    if "one tab" in query or "1 tab".lower() in query.lower():
        pyautogui.hotkey("ctrl", "w")
        speaker.Speak("One tab closed")
    elif "2 tabs" in query or "two tabs".lower() in query.lower():
        pyautogui.hotkey("ctrl", "w")
        sleep(0.5)
        pyautogui.hotkey("ctrl", "w")
        speaker.Speak("Two tabs closed")
    elif "3 tabs" in query or "three tabs".lower() in query.lower():
        pyautogui.hotkey("ctrl", "w")
        sleep(0.5)
        pyautogui.hotkey("ctrl", "w")
        sleep(0.5)
        pyautogui.hotkey("ctrl", "w")
        speaker.Speak("Three tabs closed")

    elif "4 tabs" in query or "four tabs".lower() in query.lower():
        pyautogui.hotkey("ctrl", "w")
        sleep(0.5)
        pyautogui.hotkey("ctrl", "w")
        sleep(0.5)
        pyautogui.hotkey("ctrl", "w")
        sleep(0.5)
        pyautogui.hotkey("ctrl", "w")
        speaker.Speak("Four tabs closed")
    elif "5 tabs" in query or "five tabs".lower() in query.lower():
        pyautogui.hotkey("ctrl", "w")
        sleep(0.5)
        pyautogui.hotkey("ctrl", "w")
        sleep(0.5)
        pyautogui.hotkey("ctrl", "w")
        sleep(0.5)
        pyautogui.hotkey("ctrl", "w")
        sleep(0.5)
        pyautogui.hotkey("ctrl", "w")
        speaker.Speak("Five tabs closed")

    else:
        keys = list(dictapp.keys())
        for app in keys:
            if app in query:
                os.system(f"taskkill /f /im {dictapp[app]}.exe")
