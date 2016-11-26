@echo off
set SOURCE_PATH=\\omegafs01\software$\OpenSource\ClipLink
echo.
echo		Installing Clipboard-Link
echo.
if not exist "%APPDATA%\clipboard-link\" mkdir "%APPDATA%\clipboard-link\"
copy %SOURCE_PATH%\assets\clip.exe "%APPDATA%\clipboard-link\" /y
copy %SOURCE_PATH%\assets\clipboard-link.vbs "%APPDATA%\clipboard-link\" /y
copy %SOURCE_PATH%\assets\Clipboard-link.lnk "%APPDATA%\Microsoft\Windows\SendTo" /y
echo.
echo		Clipboard-Link install complete.
echo.
pause
