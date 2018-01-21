On Error Resume Next

Dim objNetwork, strRemotePath1, strRemotePath2, strRemotePath3
Dim strDriveLetter1, strDriveLetter2, strDriveLetter3

Set objNetwork = CreateObject("WScript.Network")

'Drive letters and paths
strDriveLetter1 = "Z:"
strDriveLetter2 = "S:"
strRemotePath1 =  "\\storage\admins"
strRemotePath2 = "https:\\sharepoint.tbscg.com"

' Remove all network drives
objNetwork.RemoveNetworkDrive strDriveLetter1, True, True
objNetwork.RemoveNetworkDrive strDriveLetter2, True, True
WScript.Echo "The following drives have been removed " & strDriveLetter1 & " & " & strDriveLetter2 

' Map network drives
objNetwork.MapNetworkDrive strDriveLetter1, strRemotePath1
objNetwork.MapNetworkDrive strDriveLetter2, strRemotePath2
WScript.Echo "New drives mapped " & strDriveLetter1 & " & " & strDriveLetter2 

' Quit wscript process
Wscript.Quit