@echo off
set SOURCE_PATH=%~dp0
echo.
echo		Installing Clipboard-Link
echo.
if not exist "%APPDATA%\clipboard-link\" mkdir "%APPDATA%\clipboard-link\"
copy %SOURCE_PATH%\assets\clipboard-link.vbs "%APPDATA%\clipboard-link\" /y
cscript %SOURCE_PATH%\create_link_file.vbs

for %%X in (clip.exe) do (set FOUND=%%~$PATH:X)
if defined FOUND GOTO CLIP_FOUND 
copy %SOURCE_PATH%\assets\clip.exe "%APPDATA%\clipboard-link\" /y
GOTO FINISHED
:CLIP_FOUND
echo.
echo   clip.exe utility found. 
:FINISHED
echo.
echo		Clipboard-Link install complete.
echo.
pause
