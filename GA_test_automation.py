from openpyxl import load_workbook
from openpyxl import Workbook
import pyttsx3
import time
import speech_recognition as sr
from gtts import gTTS 
import os


#setup the speech_recognition
r = sr.Recognizer()
mic = sr.Microphone(device_index=1)

class Bot():
    def __init__(self, input_file, output_file):
        self.input_file = input_file
        self.output_file = output_file
        self.sheet = load_workbook(self.input_file).active

#defining the speech to tex function and adjust the sensitivity of ambient noise
#and record audio from mic
#the converted text will store in result variable
def stt(recognizer, microphone):
    print("collecting audio feedback...")

    with microphone as source:
        recognizer.adjust_for_ambient_noise(source, duration = 1)
        audio = recognizer.listen(source, timeout= 10)
    
    response = {'success': True, "error": None, "transcription": None}

    try:
        recog = recognizer.recognize_google(audio, language='en-US')
        response["transcription"] = recog

    except sr.RequestError:
        response['success'] = False
        response['error'] = 'API unavailable'

    except sr.UnknownValueError:
        response['error'] = "Unable to recognitized the speech"

    return response

#funtion that determine the respond
#If the GA's response opened with sorry, return test fail, otherwise return pass?
#this need some further examination
def determine_result(response):
    if response['transcription'] == None:
        return "Fail, no respond from GA"

    elif 'sorry' in response['transcription'][:5]:
        return "Fail"
    
    else:
        return "Pass"

#Use google's voice
def google_speak(text):
    language = 'en'
    speech = gTTS(text = text, lang = language, slow = False)
    speech.save('text.mp3')
    os.system("start text.mp3")

#use pyttsx3 to speak
def pyttsx3_speak(text):
    engine.say(text)
    engine.runAndWait()

#text-to-voice engine setup
engine = pyttsx3.init()
rate = engine.getProperty('rate')
engine.setProperty('rate', 120)

def tts(file):
    #say "hey google" first and than say the command with 0.5 second delay
    engine.say(steps[0][6:])
    engine.runAndWait()
    time.sleep(0.5)
    engine.say(steps[1][6:])
    engine.runAndWait()
            
            response = stt(r, mic)
            print(response)
            print(determine_result(response))

       