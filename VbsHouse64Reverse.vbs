Call ShowAllFile()
Call ConfigureAutoRun()
Call SetVbsFileAss()

Sub ShowAllFile() '显示所有文件，恢复Explorer隐藏选项，恢复快捷方式小箭头，恢复注册表工具
	NoHiddenRegPath = "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\Folder\Hidden\NOHIDDEN\CheckedValue" 	
	ShowAllRegPath = "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\Folder\Hidden\SHOWALL\CheckedValue"		
	ShowShortCurIconRegPath = "HKCR\lnkfile\IsShortcut"
	RegToolForbidRegPath = "HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\System\DisableRegistryTools"
	HideFileRegRootPath = "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\"
	Call WriteReg(HideFileRegRootPath & "Hidden", 1, "REG_DWORD") 			'普通文件 			1-显示 2-隐藏
	Call WriteReg(HideFileRegRootPath & "ShowSuperHidden", 0, "REG_DWORD") 	'受系统保护的文件 	1-显示 0-隐藏
	Call WriteReg(HideFileRegRootPath & "HideFileExt", 0, "REG_DWORD") 		'文件扩展名			0-显示 1-隐藏
	Call WriteReg(NoHiddenRegPath, 2, "REG_DWORD") 							'恢复隐藏选项
	Call WriteReg(ShowAllRegPath, 1, "REG_DWORD")							'恢复隐藏选项
	Call WriteReg(RegToolForbidRegPath, 0, "REG_DWORD")						'启用注册表工具
	Call WriteReg(ShowShortCurIconRegPath, "", "")							'添加快捷方式小箭头
End Sub

Sub ConfigureAutoRun()
	Const REG_PATH = "HKLM\Software\Microsoft\Windows\CurrentVersion\Run\"	'开机启动的注册表地址
	Const KEY_NAME = "VbsHouse"												'需要开机启动的程序的注册表主键
	Call DeleteReg(REG_PATH & KEY_NAME) '修改注册表启动项
End Sub

Sub SetVbsFileAss() '改变vbs格式文件关联
	On Error Resume Next
	Value = """%SystemRoot%\System32\wscript.exe"" ""%1"" %*"
	Call WriteReg("HKLM\SOFTWARE\Classes\vbsfile\shell\open\command\", Value, "REG_EXPAND_SZ")
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

Sub DeleteReg(targetPath) '删除注册表
	On Error Resume Next
	Dim objShell
	Set objShell = CreateObject("WScript.Shell")
	objShell.RegDelete targetPath
	Set objShell = Nothing
End Sub