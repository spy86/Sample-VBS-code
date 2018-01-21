'#### Install .NET 3.5 for Windows 8.1 ##############
Sub main() 
	Dim sExecutable,SystemSet,objShell,objExecObject,System,caption
	sExecutable = LCase(Mid(Wscript.FullName, InstrRev(Wscript.FullName,"\")+1))
	If sExecutable <> "cscript.exe" Then 
	  WScript.Echo "Please run this script with cscript.exe"
	  Wscript.Quit
	End If
	Set SystemSet = GetObject("winmgmts:").InstancesOf ("Win32_OperatingSystem") 
	Set objShell = WScript.CreateObject("WScript.Shell")
	for each System in SystemSet 
		caption = System.Caption 
	next 
	If InStr(caption,"Microsoft Windows 8") Then 
		If GetError = True Then 
			WScript.StdOut.WriteLine "Please run this script as Administrator"
		Else 
			If GetNetFramewordstatus = True  Then
				wscript.stdout.WriteLine ".Net Framework 3.5 has been installed and enabled."
			Else	
				Set objExecObject = objShell.Exec("Dism /online /Enable-feature /featurename:NetFx3 /All")
				WScript.StdOut.WriteLine "Installing .Net Framework 3.5 online,please wait.... "
				While objExecObject.Status = 0
					WScript.Sleep 1
				Wend
				If GetNetFramewordstatus = True Then 
					WScript.StdOut.WriteLine "Install .Net Framework 3.5 successfully."
				Else 
					WScript.StdOut.WriteLine "Failed to install .Net Framework 3.5 online. You can use local source to install it."
					WScript.StdOut.Write "Local source :"
					Source  = WScript.StdIn.ReadLine
					WScript.StdOut.WriteLine "Installing .Net Framework 3.5 in local,please wait.... "
					Set objExecObject = objShell.Exec("DISM /Online /Enable-Feature /FeatureName:NetFx3 /All /LimitAccess /Source:" & Source )	
					While objExecObject.Status = 0
						WScript.Sleep 1
					Wend	
					If GetNetFramewordstatus = True Then 
						WScript.StdOut.WriteLine "Install .Net Framework 3.5 successfully."
					Else 
						WScript.StdOut.WriteLine "Failed to install .Net Framework 3.5,please make sure the local source is correct."
					End If 
				End If 
			End If 
		End If 
	Else 
		WScript.StdOut.WriteLine "Please run this script in Windows 8"
	End If 

End Sub 

Function GetNetFramewordstatus
	Dim objShell,objExecObject,Flag,strText,returnValue,Result
	Set objShell = WScript.CreateObject("WScript.Shell")
	Set objExecObject = objShell.Exec("Dism /online /Get-FeatureInfo /FeatureName:NetFx3")
	Flag = 0
	Do While Not objExecObject.StdOut.AtEndOfStream
	    strText = objExecObject.StdOut.ReadLine()
	    returnValue = InStr(strText,"Enabled")
		Flag = Flag + returnValue
	Loop
	If Flag > 0 Then 
		Result = True 
	Else 
		Result = False 
	End If 
	GetNetFramewordstatus = Result 
End Function 

Function GetError 
	Dim objShell,objExecObject,Flag,strText,returnValue,Result
	Set objShell = WScript.CreateObject("WScript.Shell")
	Set objExecObject = objShell.Exec("Dism /online /Get-FeatureInfo /FeatureName:NetFx3")
	Flag = 0
	Do While Not objExecObject.StdOut.AtEndOfStream
	    strText = objExecObject.StdOut.ReadLine()
	    returnValue = InStr(strText,"Error")
		Flag = Flag + returnValue
	Loop
	If Flag > 0 Then 
		Result = True 
	Else 
		Result = False 
	End If 
	GetError = Result 
End Function 

Call main 