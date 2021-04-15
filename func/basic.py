import cv2, re

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

def match_slice(sentence, keywords):
    sen = sentence.lower()
    for key in keywords:
        if re.search(key, sen):
            return True
    return False

def analysis(respond):
    fail_keyword = ["offline", "can't do that", 'sorry',
                    'trouble', 'wrong', 'try again', "don't", "none"]
    if match_slice(respond.lower(), fail_keyword):
        return False
    return True