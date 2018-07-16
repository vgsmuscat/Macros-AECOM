set "var1=A"
set "var2=B"
set "AB=hi"
set "newvar=%var1%%var2%"
echo %newvar% is the concat of var1 and var2
echo !%newvar%! is the indirect content of newvar