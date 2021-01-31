import subprocess, sys, time

path = 'C:\\control\\'

def screenshot():
    p = subprocess.Popen(['powershell.exe', path+'screenshot.ps1'], stdout=sys.stdout)
    p.communicate()

def wifi_controller(online=True):
    if online:
        p = subprocess.Popen(["powershell.exe", path+'enableWIFI.ps1'], stdout=sys.stdout)
    elif online is False:
        p = subprocess.Popen(["powershell.exe", path+'disableWIFI.ps1'], stdout=sys.stdout)
    p.communicate()

def sign_out():
    p = subprocess.Popen(['powershell.exe', path+'SignOut.ps1'])
    p.communicate()

def sign_in():
    p = subprocess.Popen(['powershell.exe', path+'SignIn.ps1'])
    p.communicate()


print('disconnecting wifi')
wifi_controller(False)
time.sleep(5)

print('connecting wifi')
wifi_controller(True)
time.sleep(15)

print('sign out')
sign_out()

print('sign in')
sign_in()

print('taking screenshot')
screenshot()

print('syetem check complete')