import csv
import pyttsx3
import time
import speech_recognition as sr


#text-to-voice engine setup
engine = pyttsx3.init()
rate = engine.getProperty('rate')
engine.setProperty('rate', 160)

#convert the text cases in to speech
with open('simple_test_cases.csv', 'r') as file:
    reader = csv.reader(file)
    for row in reader:
        tts = []
        name, steps = row
        print("running test case: {0}".format(name))

        #for test case involved with 'hey google' activation
        if '\n' in steps:
            steps = steps.split('\n')

            #say "hey google" first and than say the command with 0.5 second delay
            engine.say(steps[0][6:])
            engine.runAndWait()
            time.sleep(0.5)
            engine.say(steps[1][6:])
            engine.runAndWait()
            print("collecting audio feedback...")
            
            
            # for i in steps:
            #     engine.say(i[6:])
            #     engine.runAndWait()
            #     time.sleep(0.5)
            # print("collecting audio feedback...")
            # time.sleep(2)
            
        


            


            