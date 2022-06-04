Set objShell = CreateObject("WScript.Shell")
Set FSO = CreateObject("Scripting.FileSystemObject")
strPath = FSO.GetParentFolderName (WScript.ScriptFullName)
If FSO.FileExists(strPath & "\MAIN.VBS") Then
     objShell.Run "wscript.exe", _
        Chr(34) & strPath & "\MAIN.VBS" & Chr(34), "", "runas", 1
Else
     objShell.Run("C:\SystemRun\Run.py "&chr(34)+WScript.Arguments(0)&chr(34))
End If