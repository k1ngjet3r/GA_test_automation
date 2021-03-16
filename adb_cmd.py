import os
import time
import subprocess

path = 'C:\\control\\'


def activate_ga():
    os.system('adb shell am start -n com.google.android.carassistant/com.google.android.apps.gsa.binaries.auto.app.voiceplate.VoicePlateActivity')


def screenshot(status):
    os.system('$ex = Test-Path C:\Screenshot -PathType Container')
    os.system('if($ex -ne 1) {mkdir C:\Screenshot}')
    os.system('adb shell screencap -p /sdcard/{}.png'.format(status))
    os.system('adb pull /sdcard/{}.png /Screenshot'.format(status))
    os.system('adb shell rm /sdcard/{}.png'.format(status))


def turn_on_wifi():
    print('Turning wifi on')
    os.system('adb root')
    os.system('adb shell "svc wifi enable"')


def turn_off_wifi():
    print('Turning wifi off')
    os.system('adb root')
    os.system('adb shell "svc wifi disable"')


def sign_out():
    print('Signing out')
    p = subprocess.Popen(
        ['powershell.exe', path+'SignOut.ps1'])
    p.communicate()


def sign_in():
    p = subprocess.Popen(
        ['powershell.exe', path+'SignIn.ps1'])
    p.communicate()


def reset():
    os.system('adb shell input keyevent 3')
    time.sleep(3)
