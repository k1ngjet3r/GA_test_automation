import csv
import pyttsx3
import time
import speech_recognition as sr
from gtts import gTTS 
import os

engine = pyttsx3.init()
rate = engine.getProperty('rate')
engine.setProperty('rate', 120)
engine.say("ok google")
engine.runAndWait()