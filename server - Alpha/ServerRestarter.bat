@echo off
color 0a
title "Fail-Safe Server Restarter"

echo Welcome to the Fail safe server restarter...
echo Opening server...
echo Server was Opened at...
time /t
echo.
ServerBigScreen.exe
goto loop

:loop
echo The server has crashed...
echo Reopening server...
echo Server was Reopened at...
time /t
echo.
ServerBigScreen.exe
goto loop