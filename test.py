from gtts import gTTS 
import os

text = "Hey google, Call Frank"
language = 'en'
speech = gTTS(text = text, lang = language, slow = True)

speech.save('text.mp3')

os.system("start text.mp3")