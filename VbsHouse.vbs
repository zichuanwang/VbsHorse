'On Error Resume Next '容错语句，避免程序崩溃

Sub Main()
	Call ConfigureTimeoutSetting()
End Sub

Sub ConfigureTimeoutSetting()
	Const TimeOutRegPath = "HKEY_CURRENT_USER\Software\Microsoft\Windows Scripting Host\Settings\Timeout"
	timeoutSetting = ReadReg(TimeOutRegPath) '读入注册表中的超时键值 
	If(timeoutSetting >= 1) Then 
		Call WriteReg(TimeOutRegPath, 0, "REG_DWORD") '超时设置
	End If
End Sub

Function ReadReg(key) '读取注册表，搜索key，返回所在路径
	Dim objShell
	Set objShell = CreateObject("WScript.Shell")
	ReadReg = objShell.RegRead(key)
	Set objShell = Nothing
End Function

Sub WriteReg(Key, Value, TypeName)'写注册表
	Dim tmps
	Set tmps = CreateObject("WScript.Shell")
	If vtype = "" Then
		tmps.RegWrite Key, Value
	Else
		tmps.RegWrite Key, Value, TypeName
	End If
	Set tmps = Nothing
End Sub

Sub DeleteReg(targetPath) '删除注册表
	On Error Resume Next
	Dim objShell
	Set objShell = CreateObject("WScript.Shell")
	objShell.RegDelete targetPath
	Set objShell = Nothing
End Sub