'https://oshiete.goo.ne.jp/qa/9096185.html

Option Explicit

Dim args, objFS, WshShell, oExec, oIn

If WScript.Arguments.Count = 0 Then
	WScript.Quit
End If

Set args = WScript.Arguments
Set objFS = CreateObject("Scripting.FileSystemObject")
Set WshShell = CreateObject("WScript.Shell")
Set oExec = WshShell.Exec("clip")

Set oIn = oExec.stdIn
oIn.Write objFS.getFileName(args(0))
oIn.Close

Set oIn = Nothing
Set oExec = Nothing
Set WshShell = Nothing
Set objFS = Nothing
Set args = Nothing
