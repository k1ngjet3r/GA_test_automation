import os, subprocess, time

path = 'C:\\control\\'

def activate_ga():
    os.system('adb shell am start -n com.google.android.carassistant/com.google.android.apps.gsa.binaries.auto.app.voiceplate.VoicePlateActivity')

def screenshot(status):
    os.system('$ex = Test-Path C:\Screenshot -PathType Container')
    os.system('if($ex -ne 1) {mkdir C:\Screenshot}')
    os.system('adb shell screencap -p /sdcard/{}.png'.format(status))
    os.system('adb pull /sdcard/{}.png C:/Users/GM-PC-03/Documents/Python/GA_test_automation/screenshot'.format(status))
    os.system('adb shell rm /sdcard/{}.png'.format(status))

def turn_on_wifi():
    os.system('adb root')
    os.system('adb shell "svc wifi enable"')

def turn_off_wifi():
    os.system('adb root')
    os.system('adb shell "svc wifi disable"')


# def sign_out():
#     p = subprocess.Popen(
#         ['powershell.exe', path+'SignOut.ps1'])
#     p.communicate()

# def sign_in():
#     p = subprocess.Popen(
#         ['powershell.exe', path+'SignIn.ps1'])
#     p.communicate()

def reset():
    os.system('adb shell input keyevent 3')
    time.sleep(3)


if __name__ == '__main__':
    reset()