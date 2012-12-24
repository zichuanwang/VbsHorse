On Error Resume Next '容错语句，避免程序崩溃

Sub ConfigureAutoRun
AutoRun_Reg_Path = "HKLM\Software\Microsoft\Windows\CurrentVersion\Run\" '开机启动的注册表地址
KEY_VALUE = "C:\VbsHouse.vbs" '需要开机启动的程序的具体路径
Key_Name = "VbsHouse" '需要开机启动的程序的注册表主键
Call RegWrite(AutoRun_Reg_Path&Key_Name, Key_Value, "") '修改注册表启动项
End Sub

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