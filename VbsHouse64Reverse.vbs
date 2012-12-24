On Error Resume Next

Call SetVbsFileAss()

Sub SetVbsFileAss() '改变vbs格式文件关联
	On Error Resume Next
	Value = """%SystemRoot%\System32\wscript.exe"" ""%1"" %*"
	Call WriteReg("HKLM\SOFTWARE\Classes\vbsfile\shell\open\command\", Value, "REG_EXPAND_SZ")
End Sub

Sub WriteReg(key, value, typeName) '写注册表
	On Error Resume Next
	Dim objShell
	Set objShell = CreateObject("WScript.Shell")
	If typeName = "" Then
		objShell.RegWrite key, value
	Else
		objShell.RegWrite key, value, typeName
	End If
	Set objShell = Nothing
End Sub