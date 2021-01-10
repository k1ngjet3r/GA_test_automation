import csv
import pyttsx3
import time
import speech_recognition as sr
from gtts import gTTS 
import os

engine = pyttsx3.init()
# rate = engine.getProperty('rate')
engine.setProperty('rate', 120)
engine.say("hello world")
time.sleep(1)
engine.runAndWait()
engine.say('test')
engine.runAndWait()