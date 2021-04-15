import os, pyttsx3, time

def activate_ga():
    os.system('adb shell am start -n com.google.android.carassistant/com.google.android.apps.gsa.binaries.auto.app.voiceplate.VoicePlateActivity')

def tts(step):
    engine = pyttsx3.init()
    engine.setProperty('rate', 105)
    # say "hey google" first and than say the command with 0.5 second delay
    activate_ga()
    time.sleep(1.5)
    # Giving the commend
    engine.say(step)
    engine.runAndWait()
