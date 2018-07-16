ECHO OFF
CLS
:MENU
ECHO.
ECHO ...............................................
ECHO PRESS 1, 2 OR 3 to select your task, or 4 to EXIT.
ECHO ...............................................
ECHO.
ECHO 1 - Open Notepad
ECHO 2 - Open Calculator
ECHO 3 - Open Notepad AND Calculator
ECHO 4 - EXIT
ECHO.
CHOICE /C:1234
IF ERRORLEVEL 1 SET M=1
IF ERRORLEVEL 2 SET M=2
IF ERRORLEVEL 3 SET M=3
IF ERRORLEVEL 4 SET M=4
IF %M%==1 GOTO NOTE
IF %M%==2 GOTO CALC
IF %M%==3 GOTO BOTH
IF %M%==4 GOTO EOF
:NOTE
cd %windir%\system32\notepad.exe
start notepad.exe
GOTO MENU
:CALC
cd %windir%\system32\calc.exe
start calc.exe
GOTO MENU
:BOTH
cd %windir%\system32\notepad.exe
start notepad.exe
cd %windir%\system32\calc.exe
start calc.exe
GOTO MENU