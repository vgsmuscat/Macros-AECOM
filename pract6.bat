@echo on
set /p File="Enter Log File Name: "
set "Prefix=SyncDrive_"
set "extention=.txt"
Set  "LogFile=%Prefix%%File%%extention%"


robocopy "D:\VGS\Miscelleneous\Anushka Files" "D:\VGS\SkyDrive Pro\Miscelleneous\Anushka Files" /e /purge /LOG+:"D:\VGS\Miscelleneous\Log Files\%LogFile%"
echo Check Log File: D:\VGS\Miscelleneous\Log Files\%LogFile% for detail