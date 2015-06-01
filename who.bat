@echo off
:CHECK
echo Checking Connection ...
ping %1 -n 1  | findstr "TTL=" >nul && GOTO MAIN || GOTO ERROR
GOTO END
:ERROR
echo Device Unreachable
GOTO END
:MAIN
wmic /node: "%1" computersystem get username
GOTO END
:END
echo.
