Dim StdOut : Set StdOut = WScript

StdOut.Echo cp & vbCrLf & " --- === Call up the script Check Bios ver. 1.0 === --- " & vbCrLf
StdOut.Echo " [Date call out the script ] :" & Date () & "" & Time ()
Set WshShell = CreateObject (" Wscript.Shell " )
strValue = " HKLM \ SYSTEM \ CurrentControlSet \ Services \ Tcpip \ Parameters \ Hostname "
Hostname = wshShell.RegRead ( strValue )
StdOut.Echo " [ Host on which the script is executed ] :" & hostname & vbCrLf & cp
StdOut.Echo " ============================================== ====== "

'Checking parameters
Select Case StdOut.Arguments.Count
 Case 0
 '' Designation what the computer should be checked by default is our pc
 Set objWMIService = GetObject ( " winmgmts :/ / ./root/cimv2 " )
 Set colItems = objWMIService.ExecQuery ( " Select * from Win32_ComputerSystem ", , 48 )
 For Each objItem in colItems
 strComputer = objItem.Name
 Next
End Select

On Error Resume Next
' Use WMI service
Set objWMIService = GetObject ( " winmgmts :/ / " & strComputer & " / root/cimv2 " )
' displaying the error number , if it is
If Err.Number Then ShowError ()

'Info on the bios -u
Set colItems = objWMIService.ExecQuery ( " Select * from Win32_BIOS where PrimaryBIOS = true" , 48 )
' displaying the error number , if it is
If Err.Number Then ShowError ()

strMsg = vbCrLf & " BIOS summary for " & strComputer & " :" & vbCrLf & vbCrLf

' Details BIOS
For Each objItem in colItems
 strMsg = strMsg _
 & " BIOS Name :" & vbCrLf _ & objItem.Name
 & "Version :" & vbCrLf _ & objItem.Version
 & " Manufacturer :" & vbCrLf _ & objItem.Manufacturer
 & " SMBIOS Version : " & vbCrLf & objItem.SMBIOSBIOSVersion
Next

' Displaying the result
StdOut.Echo strMsg

' ending
StdOut.Quit (0 )

sub ShowError
 strMsg = vbCrLf & " Error # " & Err.Number & vbCrLf & _
 Err.Description & vbCrLf & vbCrLf
 Syntax
End Sub