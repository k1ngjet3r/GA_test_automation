import os, json, time
from func.img_search import find_and_tap, checking_info_screen

def sign_in_google_account():

    '''
        Steps for signing in the Google Account via Google Maps
        1. click maps icon
        2. Tap user icon on the top right conor
        3. Tap "sign in on the car screen
        4. Enter username
        5. Tap Next btn
        6. Enter password
        7. Tap Next btn
        8. Tap Done
    '''

    with open('json\\google_account.json') as ac:
        account = json.load(ac)

    steps = [
        'google_maps_icon.png',
        'sign_out_user_icon.png',
        'sign_in_to_google_text.png',
        'sign_in_on_car_screen.png',
        'username_entry_field.png',
    ]

    i = 0

    while i < len(steps):
        print('[Sign-In] {}'.format(steps[i][:-4]))
        progress = find_and_tap(steps[i])
        
        if progress == False:
            print('[Error] Fail on step {}'.format(steps[i][:-4]))
            break
        else:
            time.sleep(2)

            if steps[i][-9:-4] == 'field' and steps[i][:8] == 'username':
                print('[DEBUG] entering username')
                os.system('adb shell input text "{}"'.format(account['username']))
                time.sleep(1)
                find_and_tap('next_btn.png')
                time.sleep(5)
                print('[DEBUG] entering password')
                os.system('adb shell input text "{}"'.format(account['password']))
                time.sleep(3)
                find_and_tap('next_btn.png')
            i += 1
    else:
        print('[DEBUG] Signin Successful!')


def sign_out_google_account():
    steps = ['google_maps_icon.png', 
            'sign_in_user_icon.png',
            'sign_out_btn.png',
            ]
    
    i = 0

    
    while i < len(steps):
        print('[Sign-In] {}'.format(steps[i][:-4]))
        progress = find_and_tap(steps[i])
        if progress == False:
            print('[Error] Fail on step {}'.format(steps[i][:-4]))
            break
        else:
            i += 1
            time.sleep(2)
    else:
        print('[DEBUG] Sign Out Successful!')