from openpyxl import load_workbook
from openpyxl import Workbook
from datetime import datetime
from datetime import date
from pushbullet import Pushbullet
from gtts import gTTS
import speech_recognition as sr
import subprocess, sys, cv2, pyttsx3, time, os

r = sr.Recognizer()
mic = sr.Microphone(device_index=1)

today = str(date.today())
exe_date = today[:4] + today[5:7] + today[8:]

output_name = 'ac_auto_result_{}.xlsx'.format(exe_date)
sheet_titles = ['Online_In', 'Offline_In', 'Online_Out', 'Offline_Out']
out = Workbook().active
for name in sheet_titles:
    out.create_sheet(name)
    out[name].append(['TCID', 'Test Step', 'Time of Execution', 'GA_respond'])

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

def wifi_controller(online=True):
    if online:
        p = subprocess.Popen(["powershell.exe", '\\connection\\enableWIFI.ps1'], stdout=sys.stdout)
    elif online is False:
        p = subprocess.Popen(["powershell.exe", '\\connection\\disableWIFI.ps1'], stdout=sys.stdout)
    p.communicate()

def sign_out():
    p = subprocess.Popen(['Powershell.exe', '\\sign_status\\SignOut.ps1'])
    p.communicate()

class Automation():
    def __init__(self, input_file):
        self.input_file = str(input_file)
        # loading the test cases file
        self.sheet = load_workbook(self.input_file).active

    def execute(self, sheet_name):
        # Generate the dictionary for the case
        cases = {str(tcid):str(step) for tcid, step in self.sheet.iter_rows(max_col=2, values_only=True) if tcid is not None}
        case_amount = len(cases)
        print('executing file {}'.format(self.input_file))
        
        # Iterate the case and feed it to the main loop
        for tcid in cases:
            ToEx = datetime.now()
            print('{} Execute Case {}/{}'.format(ToEx, tcid, case_amount))

            # Adjusting the ambient noise threadhole for 5 seconds
            ambient_noise(r, mic)

            result = [tcid, cases[tcid], ToEx]
            print('Commend: {}'.format(cases[tcid]))

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
                # Try to perform the test case again
                print('===> Try to perform the case again')
                 
                # give it 5 sec to clear the previous condition and recalibrate the ambient noise threadhole
                ambient_noise(r, mic)

                tts(cases[tcid])
        
                # Reciving Respond
                respond = stt(r, mic)
                print("Respond: {}".format(str(respond['transcription'])))

                if respond['transcription'] is not None:
                    capturing(tcid)
                
                # Won't capture photo if the respond is still none
                else:
                    print('Fail to perform the test case {}'.format(tcid))

            # Append the result to the output excel file
            result.append(str(respond['transcription']))
            out[sheet_name].append(result)
            print(respond)
            print('====================================================================================')
            time.sleep(3)
        # Push the complete notification to the phone using PushBullet

        

# online_signin
test_1 = Automation('ac_online_signIn.xlsx')
test_1.execute(sheet_titles[0])

# disconnect WiFi
wifi_controller(False)

# offline_signin
test_2 = Automation('ac_offline_signIn.xlsx')
test_2.execute(sheet_titles[1])

# connect WiFi
wifi_controller(True)
# Sign out google account
sign_out()

# online_signout
test_3 = Automation('ac_online_signOut.xlsx')
test_3.execute(sheet_titles[2])

# disconnect WiFi
wifi_controller(False)

# offline_signout
test_4 = Automation('ac_offline_signOut.xlsx')
test_4.execute(sheet_titles[3])

push_noti('All test cases executed.')
# Export the result
print('Saving the file {}'.format(output_name))
out.save(output_name)