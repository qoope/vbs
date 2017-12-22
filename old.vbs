'右クリックメニューから呼ばれ，ファイルのバックアップを取る
'バックアップには日時が入る
 
Option Explicit

Dim args, i, objFS, strNow, strMF,strMT, RE, strHMS

strNow = Now
strHMS = Right("0" & Trim(Hour(Now)) , 2) & Right("0" & Minute(Now) , 2) & Right("0" & Second(Now) , 2)

If WScript.Arguments.Count = 0 Then
	WScript.Quit
End If

Set args = WScript.Arguments
Set objFS = CreateObject("Scripting.FileSystemObject")
Set RE = CreateObject("VBScript.RegExp")
With RE
	.Pattern = "^.+\-\d{8}\-\d{6}\..+$"
	.IgnoreCase = True
	.Global = True
End With

With objFS
	For i = 0 To args.Count - 1
		If RE.Test(args(i)) = True Then
			If MsgBox("元のファイル名" & vbCrLf & Left(.getBaseName(args(i)), Len(.getBaseName(args(i))) - 16) & "." & .GetExtensionName(args(i)) & vbCrLf & "に戻します．よろしいですか？", vbOKCancel + vbCritical, "復帰の確認") = vbOk then
				strMT = .getParentFolderName(args(i)) & "\" & Left(.getBaseName(args(i)), Len(.getBaseName(args(i))) - 16) & "." & .GetExtensionName(args(i))
				If .FileExists(strMT) Then
					If MsgBox("元のファイルが同じ場所にあります．上書きしてよろしいですか？", vbOKCancel + vbCritical, "上書き確認") = vbOk then
						strMF = args(i)
						objFS.CopyFile strMF,strMT
					Else
						WScript.Quit
					End If
				Else
					strMF = args(i)
					objFS.CopyFile strMF,strMT
				End If
			Else
				WScript.Quit
			End If
		Else
			strMF = args(i)
			strMT = .getParentFolderName(args(i)) & "\" & .getBaseName(args(i)) & "-" & Replace(Left(strNow, 10), "/", "") & "-" & strHMS & "." & .GetExtensionName(args(i))
			objFS.CopyFile strMF,strMT
		End If
	Next
End With

Set RE = Nothing
Set objFS = Nothing
Set args = Nothing

WScript.Quit
