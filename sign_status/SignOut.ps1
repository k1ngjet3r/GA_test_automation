Write-Output "adb root"
adb root
Write-Output "tap Google map on thethird of left hand side"
adb shell input tap 0 600
Write-Output "tap account icon"
adb shell input tap 1850 200
Write-Output "tap sign out"
adb shell input tap 600 700
Write-Output "finish"