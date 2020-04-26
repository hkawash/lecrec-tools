@echo off
setlocal enabledelayedexpansion

rem --------------------------------------------------------------------
rem Copyright (c) 2020 Hiroaki Kawashima
rem This software is released under the MIT License, see LICENSE.txt.
rem 
rem Preparation:
rem     1. Download ffmpeg (zip archived) and extract the zip file
rem     2. Move ffmpeg.exe in the "bin" folder to this folder (which 
rem        contains this batch file). 
rem           Instead of moving the exe file, you can also set the PATH
rem        by setting system environment variable or editting the
rem        following code (the line starts with "set PATH").
rem        
rem Usage:
rem     1. Put pptx/ppsx files to "ppt-in" folder
rem     2. Run this batch file (double-click or run from command prompt)
rem     3. Compressed pptx/ppsx files can be found in "ppt-out" folder
rem --------------------------------------------------------------------

set PATH=%PATH%;%SYSTEMROOT%\System32;"C:\opt\ffmpeg-4.2.2-win64-static\bin"

rem Check ffmpeg
ffmpeg -version
if not %ERRORLEVEL% == 0 (
    echo Error: Cannot find ffmpeg
    pause
    exit /b 1
)


rem ===========================================================
rem Set bitrate here. (Do not insert space around '=')
set BITRATE=64k
rem ===========================================================

set INPUT_DIR=ppt-in
set OUTPUT_DIR=ppt-out
set WORK_DIR=work

if not exist %INPUT_DIR% (
    echo Error: Folder [%INPUT_DIR%] does not exist
    pause
    exit /b 1
)

dir %INPUT_DIR%\*pptx %INPUT_DIR%\*ppsx
if not %ERRORLEVEL% == 0 (
    echo Error: No pptx/ppsx files in [%INPUT_DIR%] folder
    pause
    exit /b 1
)

if not exist %WORK_DIR% mkdir %WORK_DIR%
if not exist %OUTPUT_DIR% mkdir %OUTPUT_DIR%

for %%f in (%INPUT_DIR%\*.pptx %INPUT_DIR%\*.ppsx) do (
    rem Setup filenames
    echo %%f
    set pptfname=%%~nxf
    set pptfbase=%%~nf
    set zipfname=%WORK_DIR%\!pptfbase!.zip
    echo !zipfname!
    copy "%%f" "!zipfname!"

    rem Create folder for pptx content
    if exist "!pptfbase!" rd /s /q "!pptfbase!"
    mkdir "!pptfbase!"

    rem Expand zip file
    powershell -Command Expand-Archive -Path '!zipfname!' -DestinationPath '!pptfbase!'
    if not exist "!pptfbase!"\ppt\media (
        echo Error: No media folder in pptx/ppsx
        pause
        exit /b 1
    )

    rem Compress audio files
    for %%a in ("!pptfbase!"\ppt\media\*.m4a) do (
        set m4afname=%%~nxa
        ffmpeg -i "%%a" -ab %BITRATE% "%WORK_DIR%\!m4afname!"
        move /y "%WORK_DIR%\!m4afname!" "%%a"
    )

    rem Archive again
    powershell -Command Compress-Archive -Path '!pptfbase!\*' -DestinationPath '%OUTPUT_DIR%\!pptfbase!.zip'
    move /y "%OUTPUT_DIR%\!pptfbase!.zip" "%OUTPUT_DIR%\!pptfname!"

    rem Remove temoporary files/folders
    del "!zipfname!"
    rd /s /q "!pptfbase!"
)
