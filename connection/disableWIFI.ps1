adb devices
Write-Output "adb root"
adb root
Write-Output "Input WIFI disable command"
adb shell "svc wifi disable"
Write-Output "Finished"