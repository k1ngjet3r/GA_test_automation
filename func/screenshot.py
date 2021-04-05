import os, datetime

def screenshot(name):
    print('getting the screenshot')
    path = os.getcwd().replace('\\', '/')+'/screenshot'
    now = datetime.now()
    # name = '{}_{}_{}_{}'.format(now.day, now.hour, now.minute, now.second)
    os.system('adb shell screencap -p /sdcard/{}.png'.format(name))
    os.system('adb pull /sdcard/{}.png {}'.format(name, path))
    os.system('adb shell rm /sdcard/{}.png'.format(name))