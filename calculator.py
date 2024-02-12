import wolframalpha
import requests
import win32com.client

speaker = win32com.client.Dispatch("SAPI.SpVoice")


def WolfRam(query):
    api_key = "KPV36R-4W74JHJRR5"
    requester = wolframalpha.Client(api_key)
    requested = requester.query(query)

    try:
        Answer = next(requested.results).text
        return Answer
    except:
        speaker.Speak("An Sring Value is not answerable")


def Calculator(query):
    Term = str(query)

    Term = Term.replace("JARVIS", "")
    Term = Term.replace("Jarvis", "")
    Term = Term.replace("jarvis", "")
    Term = Term.replace("multiply by", "*")
    Term = Term.replace("Multiply By", "*")
    Term = Term.replace("plus", "+")
    Term = Term.replace("Plus", "+")
    Term = Term.replace("minus", "-")
    Term = Term.replace("Minus", "-")
    Term = Term.replace("into", "*")
    Term = Term.replace("Into", "*")

    Final = str(Term)

    try:
        result = WolfRam(Final)
        print(f"Your Result is : {result}")
        speaker.Speak(f"Your Result is : {result}")
    except:
        print("An Sring Value is not answerable")
        speaker.Speak("An Sring Value is not answerable")
