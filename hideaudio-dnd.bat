@echo off
setlocal enabledelayedexpansion

rem --------------------------------------------------------------------
rem Copyright (c) 2020 Hiroaki Kawashima
rem This software is released under the MIT License, see LICENSE.txt.
rem 
rem Preparation:
rem     Download hideaudio.ps1 and move to this folder
rem        
rem Usage:
rem     1. Select pptx files from file explorer
rem     2. Drag and drop to this batch file
rem     3. Converted pptx/ppsx files can be found in "ppt-out" folder
rem --------------------------------------------------------------------

echo %0 begin
cd /d %~dp0
for %%f in (%*) do (
    powershell -NoProfile -ExecutionPolicy Unrestricted .\hideaudio.ps1 %%f
    if %ERRORLEVEL% neq 0 (
        echo Error occurred while processing [%%f]
        pause
        exit /b 1
    )  
)
echo done

