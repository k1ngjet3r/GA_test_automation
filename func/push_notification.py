from pushbullet import pushbullet

def push_noti(message):
    # load the key from the pushbullet_api_key.txt
    key = open('pushbullet_api_key.txt', 'r').read()
    pb = Pushbullet(key)
    dev = pb.get_device('Google Pixel 4a (5G)')
    dev.push_note('Automation Notification', message)