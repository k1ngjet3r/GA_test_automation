import speech_recognition as sr
from time import sleep

r = sr.Recognizer()
mic = sr.Microphone(device_index=1)

def ambient_noise(recog, mic):
    with mic as source:
        print('Adjusting amnient noise...')
        # Collect and adjust the ambient nosie threadhole
        recog.adjust_for_ambient_noise(source, duration=5)
        
def stt(recognizer, microphone):
    with microphone as source:
        # print('Adjusting ambient noise')
        # recognizer.adjust_for_ambient_noise(source, duration = 1)

        # Collecting the respond with 150 seconds of waiting time
        print('    Collecting Respond...')
        audio = recognizer.listen(source, timeout=2000)

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
        response['error'] = 'No respond from Google Assistant'
    return response

if __name__ == '__main__':
    print('Adjusting the ambient noise...')
    ambient_noise(r, mic)
    print('Recording in 5..')
    sleep(1)
    print('Recording in 4..')
    sleep(1)
    print('Recording in 3..')
    sleep(1)
    print('Recording in 2..')
    sleep(1)
    print('Recording in 1..')
    sleep(1)
    s = stt(r, mic)
    print(s)