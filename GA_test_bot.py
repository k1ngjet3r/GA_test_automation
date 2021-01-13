from openpyxl import load_workbook
from openpyxl import Workbook
from datetime import datetime
from pushbullet import Pushbullet
from gtts import gTTS
import speech_recognition as sr
import cv2
import pyttsx3
import time
import os

r = sr.Recognizer()
mic = sr.Microphone(device_index=1)

def ambient_noise(recog, mic):
    with mic as source:
        print('Adjusting amnient noise...')
        # Collect and adjust the ambient nosie threadhole
        recog.adjust_for_ambient_noise(source, duration = 5)


def stt(recognizer, microphone):
    with microphone as source:
        # print('Adjusting ambient noise')
        # recognizer.adjust_for_ambient_noise(source, duration = 1)
        
        # Collecting the respond with 150 seconds of waiting time
        print('Collecting Respond...')
        audio = recognizer.listen(source, timeout= 150)
    
    response = {'success': True, "error": None, "transcription": None}
    
    try:
        # set the recognize language to English and convert the speech to text
        recog = recognizer.recognize_google(audio, language='en-US')
        response["transcription"] = recog

    except sr.RequestError:
        response['success'] = False
        response['error'] = 'API unavailable'

    except sr.UnknownValueError:
        response['success'] = False
        response['error'] = "Unable to recognitized the speech"

    except:
        response['success'] = False
        response['error']  = 'No respond from Google Assistant'
    return response

def tts(step):
    engine = pyttsx3.init()
    engine.setProperty('rate', 105)
    #say "hey google" first and than say the command with 0.5 second delay
    engine.say('Hey Google')
    engine.runAndWait()
    time.sleep(0.6)
    # Giving the commend
    engine.say(step)
    engine.runAndWait()

def filename_formater(date):
    y = date[:4]
    m = date[5:7]
    d = date[8:10]
    h = date[11:13]
    minute = date[14:16]
    s = date[17:19]
    return y + '_' + m + '_' + d + '_' + h + '_' + minute + '_' + s

def capturing(tcid):
    cam = cv2.VideoCapture(0)
    cv2.namedWindow("test")
    img_counter = 0

    while True:
        ret, frame = cam.read()

        if not ret: 
            print("failed to grab frame")
            break
        cv2.imshow("test", frame)
        cv2.waitKey(1)
        
        if img_counter < 1:
            img_name = "{}.png".format(tcid)
            cv2.imwrite(img_name, frame)
            print("{} written!".format(img_name))
            img_counter += 1
        else:
            break

    cam.release()
    cv2.destroyAllWindows()

def push_noti(message):
    # load the key from the pushbullet_api_key.txt
    key = open('pushbullet_api_key.txt', 'r').read()
    pb = Pushbullet(key)
    dev = pb.get_device('Google Pixel 4a (5G)')
    dev.push_note('Automation Notification', message)

class Automation():
    def __init__(self, input_file, output_file):
        self.input_file = str(input_file)
        self.output_file = str(output_file)
        # loading the test cases file
        self.sheet = load_workbook(self.input_file).active
        # Create an empty workbook for output
        self.wb = Workbook()
        self.wb.active
        # Create the sheetname
        self.wb.create_sheet('AutoResult')
        # Append title of the sheet
        self.wb['AutoResult'].append(['TCID', 'Test Step', 'Time of Execution', 'GA_respond'])

    def execute(self):
        # Generate the dictionary for the case
        cases = {str(tcid):str(step) for tcid, step in self.sheet.iter_rows(max_col=2, values_only=True) if tcid is not None}
        
        # Iterate the case and feed it to the main loop
        for tcid in cases:
            ToEx = datetime.now()
            print('{} Execute Case {}'.format(ToEx, tcid))
            result = [tcid, cases[tcid], ToEx]
            print('Commend: {}'.format(cases[tcid]))

            # Adjusting the ambient noise threadhole for 5 seconds
            ambient_noise(r, mic)

            # Generate the speech
            tts(cases[tcid])
            # time.sleep(0.5)

            # Reciving Respond
            respond = stt(r, mic)
            print("Respond: {}".format(str(respond['transcription'])))

            # Capturing the image if the computer captured the respond
            if respond['transcription'] is not None:
                capturing(tcid)
            
            # If the computer cannot get the respond, it will execute the case again
            else:
                # give it 5 sec to clear the previous condition
                time.sleep(5)
                # Try to perform the test case again
                print('Try to perform the case again')
                tts(cases[tcid])
                # time.sleep(0.5)
                # Reciving Respond
                # print('Collecting Respond...')
                respond = stt(r, mic)
                print("Respond: {}".format(str(respond['transcription'])))

                if respond['transcription'] is not None:
                    capturing(tcid)
                
                # Won't capture photo if the respond is still none
                else:
                    print('Fail to perform the test case {}'.format(tcid))

            # Append the result to the output excel file
            result.append(str(respond['transcription']))
            self.wb['AutoResult'].append(result)
            print(respond)
            print('====================================================')
            time.sleep(3)
        # Push the complete notification to the phone using PushBullet
        push_noti('All test cases executed.')
        # Export the result
        print('Saving the file {}'.format(self.output_file))
        self.wb.save(self.output_file)

test = Automation('ac_online_signin.xlsx', 'ac_online_signin_Result.xlsx')
test.execute()