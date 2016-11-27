Set oWS = WScript.CreateObject("WScript.Shell")
sLinkFile = "%APPDATA%\Microsoft\Windows\SendTo\Clipboard-link.lnk"
Set oLink = oWS.CreateShortcut(sLinkFile)
  oLink.TargetPath = "%APPDATA%\clipboard-link\clipboard-link.vbs"
 '  oLink.Arguments = ""
  oLink.Description = "Clipboard-link"   
 '  oLink.HotKey = ""
  oLink.IconLocation = "%SystemRoot%\system32\SHELL32.dll, 27"
  oLink.WindowStyle = "1"   
 '  oLink.WorkingDirectory = ""
oLink.Save