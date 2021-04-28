import os, time, subprocess
from cv2 import cv2

'''
    Search for a pattern from an image

    <<based on https://github.com/drov0/python-imagesearch/blob/master/python_imagesearch/imagesearch.py>>

    input:
    target_img: target image that you want to search from
    pattern: desired pattern you wish to find in the target_img
    precision : the higher, the lesser tolerant and fewer false positives are found. Default is 0.8

    Return:
    the coordinate of the center of the matched pattern
'''

def image_search(target_img, pattern, precision=0.8):
    # preprocess image
    target = cv2.imread(target_img, 0)
    template = cv2.imread(pattern, 0)

    # if target_img is None:
    #     raise FileNotFoundError('Image name {} cannot be found'.format(target_img))
    # if template is None:
    #     raise FileNotFoundError('Image name {} cannot be found'.format(template))

    height, width = template.shape

    x_offset = width/2
    y_offset = height/2

    try:
        result = cv2.matchTemplate(target, template, cv2.TM_CCOEFF_NORMED)
        min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)

        print('         Image matching rate: {}'.format(max_val))

        if max_val < precision:
            return False

        x, y = (max_loc[0]+x_offset, max_loc[1]+y_offset)
        # print(x, y)

        return x, y

    except:
        print("[ImgNotFound] OpenCV couldn't find the image file in the given directory")

def tap_xy(x, y):
    print('         Tap coordinates ({}, {})'.format(x, y))
    os.system('adb shell input tap {} {}'.format(x, y))

def get_cur_screenshot():
    #get device's current screen shot and place it in img\temp folder
    current_dir = os.getcwd() + '/img/screenshot'
    current_dir.replace('\\', '/')
    # capture = subprocess.check_output(['adb', 'shell', 'screencap', '-p', '/sdcard/current.png']).splitlines()
    # move_file = subprocess.check_output(['adb', 'pill', '/sdcard/current.png', current_dir]).splitlines()

    # print('         Screenshot Captured')
    os.system('adb shell screencap -p /sdcard/current.png')
    os.system('adb pull /sdcard/current.png {}'.format(current_dir))

def find_and_tap(pattern):
    get_cur_screenshot()
    target_img = 'img\\temp\\current.png'
    img_dir = 'img\\ui_icon\\'
    try:
        if image_search(target_img, img_dir+pattern) ==  False:
            return False
        else:
            x, y = image_search(target_img, img_dir+pattern)
            tap_xy(x, y)


    except TypeError:
        print('[ImgNotFound] Please check your setting or device connection')

def checking_screen(img):
    checking_info_img = 'img\\ui_icon\\' + img
    while True:
        get_cur_screenshot()
        target_img = 'img\\temp\\current.png'
        if image_search(target_img, checking_info_img) != False:
            time.sleep(1)
            continue
        else:
            break

# if __name__ == '__main__':
#     get_cur_screenshot()
#     find_and_tap('sign_in_user_icon.png')
