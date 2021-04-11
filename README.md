# GA_test_automation
Using the text-to-speech and speech-to-text to test the Google Assistant function automatially

# Enviornment Setup
Since pyttx3 was a pretty old module, it can only support up to python 3.6
Also, there are few module that need to install:
1. openeyxl
2. pyttx3
3. pushbullet
4. opencv
5. speech_recognition

Those modules can be install by copy the following text into cmd:
```shell
    pip install <package>
```

# Summary
This project consissts of five major parts: 
1. give command from the test case (excel file)
2. receive the Google assistant respond
3. capture the system test result for the conformation using webcam
4. output the results and export to an excel file
5. push notification to the Android phone

## 1. Giving the command
The command need to be convert from text to speech from the test case, since the test cases are a google sheet format, it need to be download into a .xlsx file format so that we can read it using openpyxl. Also the file need to be filtered, we only need the test case's name and the steps.

## 2. Recieve the Google assistant's response
In this part, the response need to be convert into the text. Here we are using the pyttx3 to convert the speech to text.

## 3. Determine the results
After we collect the respond from the Google Assisatant, we need to determine the results is vaild or not. This part is the most difficult part since the google assistant respond are different for different input. The only way I come up with is the fail cases, if the respond start with "sorry", this is the sign for Google assistant state that it cannot execute the command. But the hardest part is to determine the right respond and wether the command is block or not. This is the part that need to be explored furthermore.

This project is paused since other team was already working on this part accroding to Rick. I might come back in the future for integrading this with other project.
