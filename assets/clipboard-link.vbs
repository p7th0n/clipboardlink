'EMAILASLINK.COM
Option Explicit
Dim bpath, cnt, ws, fso, strAttach, bfile, regEx
Dim WshShell, bshortpath, bfolder, sTitle, iLinks, sCountMsg, sRestrictedFileTypes, sName, CLIP_EXE 

Set WshShell = WScript.CreateObject("WScript.Shell")
Set ws = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

CLIP_EXE = "%APPDATA%\clipboard-link\clip.exe"


If WScript.Arguments.Count = 0 Then
' **************** Install shortcut to Send To

	MsgBox "Clipboard-Link is not installed -- " & vbCrLf & "run setup from S:\OpenSource\ClipLink\ to install", vbOkOnly, "Clipboard-Link"

	WScript.Quit
' **************** Install shortcut to Send To
End If
	
sTitle = "Clipboard Link"
sRestrictedFileTypes = wshShell.ExpandEnvironmentStrings( "%PATHEXT%" )

wshShell.popup "Creating shortened link...please wait... ", 1, sTitle, 0

iLinks = WScript.Arguments.Count

For cnt = 0 To (WScript.Arguments.Count -1)

	bpath = WScript.Arguments.Item(cnt)

    if instr(sRestrictedFileTypes, ucase(right(bpath, 4))) > 0 then
      MsgBox "Restricted file type cannot create link." & vbCrLf & vbCrLf & bpath, vbOkOnly, "RESTRICTED FILE TYPE"
    else
      if (fso.FileExists(bpath)) then
        Set bfile = fso.GetFile(bpath)
        bshortpath = bfile.ShortPath
        sName = fso.GetFileName(bpath)
      elseif (fso.FolderExists(bpath)) then
        Set bfolder = fso.GetFolder(bpath)
        bshortpath = bfolder.ShortPath
        sName = bfolder.Name
      end if

      strAttach = "[" & sName & "] file://" & bshortpath  

      Set regEx = New RegExp
      regEx.Pattern = "&"
      regEx.IgnoreCase = True
      'strAttach = """" & regEx.Replace(strAttach, "^&") & """"
      strAttach = """" & strAttach & """"


      WshShell.Run "cmd.exe /c echo " & strAttach & " | " & CLIP_EXE, 0, TRUE

      if cnt < iLinks - 1 then
        sCountMsg = "  Paste this link first and press Ok for the next link."
      else
        sCountMsg = ""
      end if

      MsgBox "Source:  " & bpath & vbCrLf & vbCrLf & "Link:  " & strAttach & vbCrLf & vbCrLf & "Link " & cnt + 1 & " of " & iLinks & " was copied to the clipboard." & sCountMsg, vbOkOnly, sTitle
        
    end if
Next

WScript.Quit

