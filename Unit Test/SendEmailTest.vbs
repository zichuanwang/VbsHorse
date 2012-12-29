On Error Resume Next '容错语句，避免程序崩溃 
Const SENDER_MAIL_ADDRESS	= "2637252534@qq.com" '用于发送邮件的邮箱地址
Const SENDER_MAIL_PASSWORD 	= "gy920711"          '用于发送邮件的邮箱密码
Const SENDEE_MAIL_ADDRESS  	= "610955867@qq.com"  '用于接收邮件的邮箱地址

Dim objFileSystem
Set objFileSystem = CreateObject("Scripting.FileSystemObject")
selfPath = objFileSystem.GetFile(WScript.ScriptFullName)
success = SendEmail(SENDER_MAIL_ADDRESS, SENDER_MAIL_PASSWORD, SENDEE_MAIL_ADDRESS, "", "来自VbsHouse的邮件", "感谢您测试VbsHouse。", selfPath)
If success Then
WScript.Echo "发送邮件成功。"
Else
WScript.Echo "发送邮件失败。"
End If

Set objFileSystem = Nothing

Function SendEmail(senderAddress, senderPassword, sendeeAddress, backupAddress, mailTitle, mailContent, mailAttachment)
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
    
    objEmail.Send '发送邮件
	
    Set objEmail = Nothing
    SendEmail = True
    
    If Err Then
        Err.Clear
        SendEmail = False
    End If
End Function