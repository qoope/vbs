'�E�N���b�N���j���[����Ă΂�C�t�@�C���̃o�b�N�A�b�v�����
'�o�b�N�A�b�v�ɂ͓���������
 
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
			If MsgBox("���̃t�@�C����" & vbCrLf & Left(.getBaseName(args(i)), Len(.getBaseName(args(i))) - 16) & "." & .GetExtensionName(args(i)) & vbCrLf & "�ɖ߂��܂��D��낵���ł����H", vbOKCancel + vbCritical, "���A�̊m�F") = vbOk then
				strMT = .getParentFolderName(args(i)) & "\" & Left(.getBaseName(args(i)), Len(.getBaseName(args(i))) - 16) & "." & .GetExtensionName(args(i))
				If .FileExists(strMT) Then
					If MsgBox("���̃t�@�C���������ꏊ�ɂ���܂��D�㏑�����Ă�낵���ł����H", vbOKCancel + vbCritical, "�㏑���m�F") = vbOk then
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
