dim source, destination, WshShell, ObjShell, fso, OverWriteExisting, current

current= Left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName, "\"))
source = current & "timeDate synchronizer.vbs"
destination = "C:\timeDate synchronizer.vbs"

Set WshShell = WScript.CreateObject("WScript.Shell")

If WScript.Arguments.length = 0 Then
	Set ObjShell = CreateObject("Shell.Application")
	ObjShell.ShellExecute "wscript.exe", """" & WScript.ScriptFullName & """" & " RunAsAdministrator", , "runas", 1
Else
	OverWriteExisting = True
	Set fso = CreateObject("Scripting.FileSystemObject")
	fso.CopyFile source, destination, OverWriteExisting
end If

CreateObject("WScript.Shell").Run(current & "install.reg")
