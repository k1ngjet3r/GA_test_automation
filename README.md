# GA_test_automation
Using the text-to-speech and speech-to-text to automate the Google Assistant function

This project consissts of four major parts: giving command from the test case, receive the Google assistant respond, and output the results and determine the results.

## 1. Giving the command
The command need to be convert from text to speech from the test case, since the test cases are a google sheet format, it need to be download into a CSV file format so that we can read it using Python. Also the file need to be filtered, we only need the test case's name and the steps.
Once we have the CSV. file, the steps are then converted into speech by using pyttsx3 library.

## 2. Recieve the Google assistant's response
In this part, the response need to be convert into the text. Here we are using the speech_recognition library and gtts (Google text-to-speech) engine to convert the speech to text.

## 3. Determine the results
After we collect the respond from the Google Assisatant, we need to determine the results is vaild or not. This part is the most difficult part since the google assistant respond are different for different input. The only way I come up with is the fail cases, if the respond start with "sorry", this is the sign for Google assistant state that it cannot execute the command. But the hardest part is to determine the right respond and wether the command is block or not. This is the part that need to be explored furthermore.

This project is paused since other team was already working on this part accroding to Rick. I might come back in the future for integrading this with other project.
