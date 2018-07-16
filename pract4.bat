set /p File="Enter Log File Name (Saved in D:\VGS\Micelleneous Directory): "
set "extention=.txt"
echo %extention%
Set "LogFile=%File%%extention%"
echo %LogFile%
echo Check Log File: %logFile% for detail