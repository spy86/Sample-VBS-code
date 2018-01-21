'Checking parameters
Select Case WScript.Arguments.Count
Case 0
 '' Designation what the computer should be checked by default is our pc
 Set objWMIService = GetObject ( " winmgmts :/ / ./root/cimv2 " )
 Set colItems = objWMIService.ExecQuery ( " Select * from Win32_ComputerSystem ", , 48 )
 For Each objItem in colItems
 strComputer = objItem.Name
 Next
End Select

' Setting permanent
Const wbemFlagReturnImmediately = & h10
Const wbemFlagForwardOnly = & h20

strMsg = vbCrLf & "CPU load percentage for " & strComputer & " :" & vbCrLf & vbCrLf

'
On Error Resume Next

' Use WMI service
Set objWMIService = GetObject ( " winmgmts :/ / " & strComputer & " / root/cimv2 " )
' displaying the error number , if it is
If Err Then ShowError

' Info. processor
Set colItems = objWMIService.ExecQuery ( "SELECT * FROM Win32_Processor ", " WQL ", _
 wbemFlagReturnImmediately + wbemFlagForwardOnly )
' displaying the error number , if it is
If Err Then ShowError
' Detail regarding the processor of Use
For Each objItem In colItems
 strMsg = strMsg _
 & "Device ID :" & vbCrLf _ & objItem.DeviceID
 & "Load Percentage :" & objItem.LoadPercentage & vbCrLf & vbCrLf
Next

' result
WScript.Echo strMsg

WScript.Quit (0 )

Sub ShowError ()
 strMsg = vbCrLf & " Error # " & Err.Number & vbCrLf & _
 Err.Description & vbCrLf & vbCrLf
 Syntax
End Sub