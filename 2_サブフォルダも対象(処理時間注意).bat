@echo off
setlocal

if "%~dpnx1" == "" (
  echo log�t�@�C��/�t�H���_���w�肵�Ă��������B
  goto ERROR
)


echo;
echo;
echo;
echo;
echo;
echo;
echo;
echo;
echo;
echo;

Powershell -file "%~dp0tools\invoke.ps1" -Recurse true "%~dpnx1" "%~dpnx2" "%~dpnx3" "%~dpnx4" "%~dpnx5" "%~dpnx6" "%~dpnx7" "%~dpnx8" "%~dpnx9"

echo �������܂���

pause
endlocal
exit /B


:ERROR
endlocal
pause
exit /B
