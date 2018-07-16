ECHO OFF
CLS
REM YesNo.BAT
REM Delete any file named YN
IF EXIST YN DEL YN
REM Prompt the user
ECHO If you want the Yes branch,
ECHO press Any Key followed by
ECHO Ctrl-Z and Enter. If you
ECHO want the No branch, press
ECHO only Ctrl-Z and Enter.
REM Get the response. If only Ctrl-Z REM is pressed,
REM then no YN file will be created.
COPY CON yn > nul
REM If the YN file exists, branch
REM to Yes otherwise branch to No
IF EXIST yn GOTO yes
GOTO no
:yes
ECHO You chose Yes.
GOTO end
:no
ECHO You chose No.
:end
REM Erase YN file, if it exists.
IF EXIST YN del YN 