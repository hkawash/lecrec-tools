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
rem     1. Put pptx files to "ppt-in" folder
rem     2. Run this batch file (double-click or run from command prompt)
rem     3. Converted pptx files can be found in "ppt-out" folder
rem --------------------------------------------------------------------

echo %0 begin
set INPUT_DIR=ppt-in

if not exist %INPUT_DIR% (
    echo Error: Folder [%INPUT_DIR%] does not exist
    pause
    exit /b 1
)

dir %INPUT_DIR%\*pptx %INPUT_DIR%\*ppsx
if %ERRORLEVEL% neq 0 (
    echo Error: No pptx/ppsx files in [%INPUT_DIR%] folder
    pause
    exit /b 1
)

for %%f in (%INPUT_DIR%\*.pptx %INPUT_DIR%\*.ppsx) do (
    powershell -NoProfile -ExecutionPolicy Unrestricted .\hideaudio.ps1 %%f
    if %ERRORLEVEL% neq 0 (
        echo Error occurred while processing [%%f]
        pause
        exit /b 1
    )  
)
echo done!

