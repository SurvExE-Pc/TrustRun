Dim Shell
Set Shell = CreateObject( "WScript.Shell" )
Shell.Run("C:\SystemRun\AdvancedRun.exe /EXEFilename "&chr(34)+WScript.Arguments(0)+chr(34)&" /RunAs 8 /Run")