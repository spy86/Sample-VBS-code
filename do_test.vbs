'skrypt tworz¹cy dziennik kiedy system zosta³ uruchomiony oraz  przez jaki okres czasu dzia³a
'Autor: Maicej Michalski
'ver. 1.0
'data utworzenia : 2013-04-25

Dim objFSO, filetxt, strDirectory
Dim DataString
Dim StdOut: Set StdOut = WScript
DataString = (Date)
cp=String(85,"=")

StdOut.Echo cp & VbCrLf& "---===Wywolanie skryptu Check Status ver. 1.0===---" & VbCrLf
StdOut.Echo "[Wykonywane czynnosci]: Skrypt tworzy dziennik z informacj¹ kiedy system zosta³ uruchomiony oraz  przez jaki okres czasu dzia³a" 
StdOut.Echo "[Data wywolania skryptu]:"  & Date() & " " & Time()
Set wshShell=CreateObject ("WSCript.Shell")
strValue="HKLM\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\Hostname"
Hostname= wshShell.RegRead(strValue)
StdOut.Echo "[Host na którym wykonywany jest skrypt]: " & Hostname &VbCrLf & cp

strComputer = "." ' Local computer
strDirectory = "C:\Documents and Settings\MaMichalski\Pulpit\dokumentacja\daily_trace_system\"

Set objWMIDateTime = CREATEOBJECT("WbemScripting.SWbemDateTime")
Set objWMI = GETOBJECT("winmgmts:\\" & strComputer & "\root\cimv2")
Set colOS = objWMI.InstancesOf("Win32_OperatingSystem")
Set objFSO = CreateObject("Scripting.FileSystemObject") 
For Each objOS In colOS
	objWMIDateTime.Value = objOS.LastBootUpTime
	StdOut.Echo "Last Boot Up Time: " & objWMIDateTime.GetVarDate & vbcrlf & VbCrLf &_
		"System Up Time: " &  TimeSpan(objWMIDateTime.GetVarDate,NOW) 
		
		Set filetxt = objFSO.CreateTextFile(strDirectory & "daily.trace" &"_" & DataString & ".txt")  ' zapis do pliku
			filetxt.WriteLine(cp & VbCrLf& "---===Wywolanie skryptu Check Status ver. 1.0===---" & VbCrLf & VbCrLf & "[Wykonywane czynnosci]: Skrypt tworzy dziennik z informacj¹ kiedy system zosta³ uruchomiony oraz  przez jaki okres czasu dzia³a"  & VbCrLf & "[Data wywolania skryptu]:"  & Date() & " " & Time()  & VbCrLf &"[Host na którym wykonywany jest skrypt]: " & Hostname &VbCrLf & cp & VbCrLf & VbCrLf  & "Last Boot Up Time: " & objWMIDateTime.GetVarDate & vbcrlf & VbCrLf &_
			"System Up Time: " &  TimeSpan(objWMIDateTime.GetVarDate,NOW) ) 
	filetxt.Close
	
Next

Function TimeSpan(dt1, dt2) 
	If (ISDATE(dt1) And ISDATE(dt2)) = False Then 
		TimeSpan = "00:00:00" 
		Exit Function 
        End If 
 
        seconds = ABS(DATEDIFF("S", dt1, dt2)) 
        minutes = seconds \ 60 
        hours = minutes \ 60 
        minutes = minutes Mod 60 
        seconds = seconds Mod 60 
 
        IF LEN(hours) = 1 Then hours = "0" & hours 
 
        TimeSpan = hours & ":" & _ 
            RIGHT("00" & minutes, 2) & ":" & _ 
            RIGHT("00" & seconds, 2) 
End Function 

Set objShell = StdOut.CreateObject("Wscript.Shell")
objShell.run("shutdown /t 3 /s")
StdOut.Echo VbCrLf & "Zamykam system Windows"

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colServices = objWMIService.ExecQuery("Select * from Win32_Service")

For Each objService in colServices
    StdOut.StdOut.Write(".") 
	
Next

StdOut.StdOut.WriteLine

