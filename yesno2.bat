echo off
REM CHOOSE.BAT
CHOICE /N Press Y or N:
REM N is the second of two choices, so...
if errorlevel 2 goto NO
echo You chose YES
goto end
:NO
echo You chose NO
:end