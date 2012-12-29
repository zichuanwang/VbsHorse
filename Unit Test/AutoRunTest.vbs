Call ConfigureAutoRun()

WScript.Echo("Auto Run")

Sub ConfigureAutoRun()
	Dim objFileSystem
	Set objFileSystem = WScript.CreateObject("scripting.filesystemobject")
	Const REG_PATH = "HKLM\Software\Microsoft\Windows\CurrentVersion\Run\"	'开机启动的注册表地址
	Const KEY_NAME = "VbsHouse"												'需要开机启动的程序的注册表主键
	horsePath = objFileSystem.GetFile(Wscript.ScriptFullName)				'需要开机启动的程序的具体路径
	Call WriteReg(REG_PATH & KEY_NAME, horsePath, "") '修改注册表启动项
	Set objFileSystem = Nothing
End Sub

Sub WriteReg(key, value, typeName) '写注册表
	Dim objShell
	Set objShell = CreateObject("WScript.Shell")
	If typeName = "" Then
		objShell.RegWrite key, value
	Else
		objShell.RegWrite key, value, typeName
	End If
	Set objShell = Nothing
End Sub