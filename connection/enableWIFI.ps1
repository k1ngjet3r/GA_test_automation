adb devices
Write-Output "adb root"
adb root
Write-Output "Input WIFI enable command"
adb shell "svc wifi enable"
Write-Output "Finished"