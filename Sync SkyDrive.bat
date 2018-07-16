@echo on
set /p File="Enter Log File Name: "
set "Prefix=SyncDrive_"
set "extention=.txt"
Set  "LogFile=%Prefix%%File%%extention%"


robocopy "D:\VGS\Miscelleneous\Anushka Files" "D:\VGS\SkyDrive Pro\Miscelleneous\Anushka Files" /e /purge /LOG+:"D:\VGS\Miscelleneous\Log Files\%LogFile%"
robocopy "D:\VGS\Miscelleneous\General" "D:\VGS\SkyDrive Pro\Miscelleneous\general" /e /purge /LOG+:"D:\VGS\Miscelleneous\Log Files\%LogFile%"
robocopy "D:\VGS\Miscelleneous\Keshavi Files" "D:\VGS\SkyDrive Pro\Miscelleneous\Keshavi Files" /e /purge /LOG+:"D:\VGS\Miscelleneous\Log Files\%LogFile%"
robocopy "D:\VGS\Miscelleneous\LANOSREP" "D:\VGS\SkyDrive Pro\Miscelleneous\LANOSREP" /e /purge /LOG+:"D:\VGS\Miscelleneous\Log Files\%LogFile%"
robocopy "D:\VGS\Miscelleneous\Vandana Files" "D:\VGS\SkyDrive Pro\Miscelleneous\Vandana Files" /e /purge /LOG+:"D:\VGS\Miscelleneous\Log Files\%LogFile%"
robocopy "D:\VGS\Miscelleneous\Papa Files" "D:\VGS\SkyDrive Pro\Miscelleneous\Papa Files" /e /purge /LOG+:"D:\VGS\Miscelleneous\Log Files\%LogFile%"
robocopy "D:\VGS\Miscelleneous\Mummy Files" "D:\VGS\SkyDrive Pro\Miscelleneous\Mummy Files" /e /purge /LOG+:"D:\VGS\Miscelleneous\Log Files\%LogFile%"

robocopy "D:\VGS\Projects\ECC Dwarka\Deliveries by Contractors" "D:\VGS\SkyDrive Pro\Projects\ECC Dwarka\Deliveries by Contractors" /e /purge /LOG+:"D:\VGS\Miscelleneous\Log Files\%LogFile%"
robocopy "D:\VGS\Projects\ECC Dwarka\Planning\Formats" "D:\VGS\SkyDrive Pro\Projects\ECC Dwarka\Planning\Formats" /e /purge /LOG+:"D:\VGS\Miscelleneous\Log Files\%LogFile%"
robocopy "D:\VGS\Projects\ECC Dwarka\Planning\Miscelleneous" "D:\VGS\SkyDrive Pro\Projects\ECC Dwarka\Planning\Miscelleneous" /e /purge /LOG+:"D:\VGS\Miscelleneous\Log Files\%LogFile%"
robocopy "D:\VGS\Projects\ECC Dwarka\Planning\MPR" "D:\VGS\SkyDrive Pro\Projects\ECC Dwarka\Planning\MPR" /e /purge /LOG+:"D:\VGS\Miscelleneous\Log Files\%LogFile%"
robocopy "D:\VGS\Projects\ECC Dwarka\Planning\Programme" "D:\VGS\SkyDrive Pro\Projects\ECC Dwarka\Planning\Programme" /e /purge /LOG+:"D:\VGS\Miscelleneous\Log Files\%LogFile%"
robocopy "D:\VGS\Projects\ECC Dwarka\Planning\Updates" "D:\VGS\SkyDrive Pro\Projects\ECC Dwarka\Planning\Updates" /e /purge /LOG+:"D:\VGS\Miscelleneous\Log Files\%LogFile%"
robocopy "D:\VGS\Projects\ECC Dwarka\Planning\Deliveries by SubConsultant\PEAC\costing" "D:\VGS\SkyDrive Pro\Projects\ECC Dwarka\Planning\Deliveries by SubConsultant\PEAC\costing" /e /purge /LOG+:"D:\VGS\Miscelleneous\Log Files\%LogFile%"
robocopy "D:\VGS\Projects\ECC Dwarka\Statutory Approval" "D:\VGS\SkyDrive Pro\Projects\ECC Dwarka\Statutory Approval" /e /purge /LOG+:"D:\VGS\Miscelleneous\Log Files\%LogFile%"
robocopy "D:\VGS\Projects\ECC Dwarka\Work Order" "D:\VGS\SkyDrive Pro\Projects\ECC Dwarka\Work Order" /xd "D:\VGS\Projects\ECC Dwarka\Work Order\EPC\Corrigendum - Not yet uploaded, 28.10.2017" "D:\VGS\Projects\ECC Dwarka\Work Order\EPC\IICC Tender Documents as uploaded on _14-10-2017" /e /purge /LOG+:"D:\VGS\Miscelleneous\Log Files\%LogFile%"


robocopy "C:\Users\guptav1\AppData\Roaming\Microsoft\Excel\XLSTART" "D:\VGS\SkyDrive Pro\Miscelleneous\XLSTART Direct" /e /purge /LOG+:"D:\VGS\Miscelleneous\Log Files\%LogFile%"
robocopy "D:\VGS\Projects\ECC Dwarka\Milestone Submission" "D:\VGS\SkyDrive Pro\Projects\ECC Dwarka\Milestone Submission" /e /purge /LOG+:"D:\VGS\Miscelleneous\Log Files\%LogFile%"





echo Check Log File: "D:\VGS\Miscelleneous\Log Files\%LogFile%" for detail
notepad "D:\VGS\Miscelleneous\Log Files\%LogFile%"
