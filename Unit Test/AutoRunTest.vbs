On Error Resume Next '容错语句，避免程序崩溃
Call ConfigureAutoRun

Sub ConfigureAutoRun
	On Error Resume Next
	Const REG_PATH = "HKLM\Software\Microsoft\Windows\CurrentVersion\Run\"	'开机启动的注册表地址
	Const KEY_VALUE = "C:\VbsHouse.vbs" 									'需要开机启动的程序的具体路径
	Const KEY_NAME = "VbsHouse"												'需要开机启动的程序的注册表主键
	Call WriteReg(REG_PATH & KEY_NAME, KEY_VALUE, "") '修改注册表启动项
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