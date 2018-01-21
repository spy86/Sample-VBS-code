Dim objFSO, objShell, strResult
Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell") 
Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv") 
Set objWMI = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2") 
Set objNetwork = CreateObject("WScript.Network")
Set objSysInfo    = CreateObject("ADSystemInfo")
strCurPath = objFSO.GetParentFolderName(WScript.ScriptFullName)
strTxtFile = objFSO.GetBaseName(WScript.ScriptFullName) & ".txt"
strSavedFile = objFSO.GetBaseName(WScript.ScriptFullName) & ".saved"




'OPTIONS
OPT_NAME 		= "NAME"
OPT_OS			= "OS"
OPT_NETWORK		= "NETWORK"
OPT_SOFTWARE	= "SOFTWARE"
OPT_HARDWARE	= "HARDWARE"
OPT_OOP         = "OOP"
OPT_STYLE		= "STYLE"


'FILE OPERATIONS
FOR_READ = 1
FOR_WRITE = 2
FOR_APPEND = 8


'SAVED IDENTIFIERS 
SAVID_ATTRIBUTE  = 0


'SAVED DATA
SAVDATA_TIMEDATE = 0
SAVDATA_VALUE    = 1


'REGISTRY
HKLM = &H80000002


'STYLE
TABLE_STYLE = "<style type=" & Chr(34) & "text/css" & Chr(34) & ">" & _
			"table.inv {" & _
			"  border-collapse: collapse;" & _
			"  border: 1px dashed #606060;" & _
			"  color: #606060;" & _
			"  text-align: center;" & _
			"  background-color: white;" & _
			"  }" & _
			"thead th.inv, tfoot th.inv {" & _
			"  border: 1px dashed #606060;" & _			
			"  background-color: black;" & _
			"  color: white;" & _
			"  }  " & _
			"tbody tr.inv td.inv {" & _
			"  background-color: white;" & _
			"  border: 1px dashed #606060;" & _
			"  padding-left: 6px;" & _
			"  padding-right: 6px;" & _
			"  }" & _
			"tr.inv td.inv {" & _
			"  background-color: white;" & _
			"  border: 1px dashed #606060;" & _
			"  padding-left: 6px;" & _
			"  padding-right: 6px;" & _
			"  }" & _
			"table.inv tr td {" & _
			"  background-color: white;" & _
			"  border: 1px dashed #606060;" & _
			"  padding-left: 6px;" & _
			"  padding-right: 6px;" & _
			"  }" & _	
			"table.inv tr:hover td:hover {" & _
			"  background-color: #EFEFEF;" & _
			"  }" & _						
			"</style>"



'CHECK ARGUMENTS
Function CheckArguments(expArgs)
On Error Resume Next

	If (WScript.Arguments.Count < expArgs) Then
		'ERROR
		Call WScript.echo("Not enough arguments (" & WScript.Arguments.Count & "/" & expArgs & ").")
		
		'QUIT
		WScript.Quit		
	End If
End Function


'CHECK NATIVE BITNESS
Function CheckNativeBitness
On Error Resume Next

	Set objProcEnv = objShell.Environment("Process")

	'GET PROCESS BITNESS
    strProcessBit = objProcEnv("PROCESSOR_ARCHITECTURE")
    strProcessorBit = objProcEnv("PROCESSOR_ARCHITEW6432")
    
	'RUNNING IN WOW6432
	If (Len(strProcessorBit) > 0) Then
		'ERROR
		Call WScript.echo("This process must be run in native mode, outside wow64.")
		
		'QUIT
		WScript.Quit			
	End If
End Function


'CHECK OUT OF PROCESS
Function CheckOutOfProcess
	'OOP MARKER
	If (WScript.Arguments.Count > 0) Then strMarker = WScript.Arguments(0)
	
	'SCRIPT WAS LAUNCHED OOP	
	If (StrComp(strMarker, "$", 1) = 0) Then	
		'MESSAGE
		Call WScript.Echo("Started oop...")

		'ARGUMENTS
		' 0 : OOP MARKER
		' 1 : OPTION

		'CHECK ARGUMENTS
		Call CheckArguments(2)
		
		'GET ARGUMENTS
		strOption = WScript.Arguments(1)
		
		'GET ATTRIBUTES AND SAVE NOW, TO BE JUST LOADED LATER
		Call GetSaveAttributes
		
		'QUIT
		Call WScript.Quit
	End If	
End Function


'LAUNCH OUT OF PROCESS
Function LaunchOutOfProcess
	'RESULT
	LaunchOutOfProcess = False

	'MARKER
	strArguments = "$"

	'ARGUMENT STRING
	For x = 0 To WScript.Arguments.Count - 1
		strArguments = strArguments & " " & Chr(34) & WScript.Arguments(x) & Chr(34)
	Next

	strExe = Chr(34) & "cscript.exe" & Chr(34) & " //nologo"
	strExeArguments = Chr(34) & strCurPath & "\" & "zabbix_vbs_logger.vbs" & Chr(34) & " " & Chr(34) & WScript.ScriptFullName & Chr(34) & " " & strArguments & " " & Chr(34) & strCurPath & "\" & objFSO.GetBaseName(WScript.ScriptName) & ".log" & Chr(34)

	Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate,authenticationLevel=Pkt,(Shutdown)}!\\.\root\cimv2")

	'CONFIGURE STARTUP INFO
	Set objStartup = objWMIService.Get("Win32_ProcessStartup")
	Set objConfig = objStartup.SpawnInstance_
	objConfig.ShowWindow = 1

	Err.Clear

	'CREATE PROCESS VIA WMI WIN32_PROCESS.CREATE TO CIRCUMVENT THE ZABBIX MECHANISM FOR 
	'KILLING THE PROCESS AND ALL CHILDREN AFTER THE TIMEOUT OCCURS
	Set objProcess = objWMIService.Get("Win32_Process")
	nResult = objProcess.Create(strExe & " " & strExeArguments, strCurPath, objConfig, intProcessID)
	
	'RESULT
	If (nResult = 0) Then LaunchOutOfProcess = True
End Function


'ARRAY TO HTML TABLE
Function ArrayToHtml(arrHeader, arrTable, RowCount)	
	str1 = "<table class=" & Chr(34) & "inv" & Chr(34) & "><thead><tr class=" & Chr(34) & "inv" & Chr(34) & ">"
	
	'ADD HEADER
	For x = 0 To UBound(arrTable)
		str1 = str1 & "<th class=" & Chr(34) & "inv" & Chr(34) & ">" & arrHeader(x) & "</th>"		
	Next
	
	str1 = str1 & "</tr></thead>"
	
	'EACH COL
	For x = 0 To RowCount - 1	
		str1 = str1 & "<tr class=" & Chr(34) & "inv" & Chr(34) & ">"
		
		'EACH ROW
		For y = 0 To UBound(arrTable)
			If (Not IsNull(arrTable(y, x))) Then
				'FORMAT IN HTML
				strHtml = FormatHtml(arrTable(y, x))
			Else
				strHtml = ""
			End If
			'MAKE COL
			str1 = str1 & "<td class=" & Chr(34) & "inv" & Chr(34) & ">" & strHtml & "</td>"
		Next
		
		str1 = str1 & "</tr>"
	Next
	
	str1 = str1 & "</table>"
	
	'RESULT
	ArrayToHtml = str1
End Function


'GET EXE OUTPUT
Function GetExeOutput(strExe, ByRef strStd, Timeout, ByRef DidTimeout)
On Error Resume Next

	'RESULT
	GetExeOutput = False
	DidTimeout = False
	strStd = ""
	
	'ERR CLEAR
	Err.Clear
	
	Set objExec = objShell.Exec(strExe)
	
	'ERR CHECK
	If (Err.Number = 0) Then
		'SLEEP TO ALLOW PIPES TO CONNECT
		Call WScript.Sleep(100)		
	
		Call objExec.StdIn.Close
		Call objExec.StdErr.Close
		StartDate = Now
		
		'WAIT FOR IT TO FINISH OR TIMEOUT
		Do Until ((objExec.Status = 1) And (objExec.StdOut.AtEndOfStream)) Or (DateDiff("s", StartDate, Now) > TimeOut)
			'STDOUT
			If (Not objExec.StdOut.AtEndOfStream) Then
				strRead = objExec.StdOut.ReadLine
				If (Len(strRead) > 0) Then 
					'#DEBUG
					'Call WScript.Echo(strRead)				
					strStd = strStd & strRead & vbCrLf
				End If
			End If
		Loop
			
		'TIMEOUT
		If (DateDiff("s", StartDate, Now) > TimeOut) Then		
			'TERMINATE PROCESS
			Call objExec.Terminate
			'TIMED OUT
			DidTimeout = True			
		'SUCCESS
		Else
			'RESULT
			GetExeOutput = True		
		End If
	End If
End Function


'TABLETOSTRING
Function TableToString(arrTable, RowCount)
	str1 = (UBound(arrTable) + 1) & "|" & RowCount
	
	'EACH ROW
	For x = 0 To RowCount - 1
		'EACH COLUMN
		For y = 0 To UBound(arrTable)
			str1 = str1 & "|" & Replace(arrTable(y, x), "|", "`")
		Next
	Next	
	
	'RESULT
	TableToString = str1	
End Function


'STRINGTOTABLE
Function StringToTable(strTable, arrTable, ByRef RowCount)
	strSplit = Split(strTable, "|")
	
	'VALID TABLE
	If (UBound(strSplit) > 1) Then	
		'COLS, ROWS
		ColCount = CInt(strSplit(0))
		RowCount = CInt(strSplit(1))
		
		'SET TABLE
		startx = 2
		ReDim arrTable(ColCount - 1, RowCount - 1)
	
		'EACH ROW
		For x = 0 To RowCount - 1
			'EACH COLUMN
			For y = 0 To UBound(arrTable)
				arrTable(y, x) = Replace(strSplit(startx), "`", "|")
				'INC
				startx = startx + 1
			Next
		Next	
	End If		
End Function


'TABLE TO ARRAY
Function TableToArray(strTable, strTop, strBottom, arrTable, ByRef RowCount, RemoveSepparator)
	'RESULT
	TableToArray = False
	Dim arrColIndex()		
	top = -1
	bottom = -1	

	'POOR-MAN-REGEX
	strSplit = Split(strTable, vbCrLf)
	
	'FIND TOP
	If (Len(strTop) > 0) Then
		For x = 0 To UBound(strSplit) 
			If (InStr(1, strSplit(x), strTop, 1) > 0) Then top = x + 1
		Next
	Else
		top = 0
	End If
	'SEPPARATOR
	If (RemoveSepparator) Then top = top + 1
	
	'BOUNDS
	If (top > -1) Then		
		'FIND BOTTOM
		If (Len(strBottom) > 0) Then
			For x = top To UBound(strSplit) 
				If (InStr(1, strSplit(x), strBottom, 1) > 0) Then bottom = x - 1
			Next	
		Else
			bottom = UBound(strSplit)
		End If
	End If
		
	'BOUNDS
	If (top > -1) And (bottom > -1) Then
		'GET ROWS
		RowCount = bottom - top + 1
		ColCount = 0
		
		'GET COLS				
		x = 1
		strCols = Trim(strSplit(top - 1))		
		Do While (x < Len(strCols))
			'HEADER FOUND
			If (Mid(strCols, x, 1) <> " ") Then
				'UPDATE COLCOUNT AND COLINDEX
				ColCount = ColCount + 1
				ReDim Preserve arrColIndex(ColCount - 1)
				arrColIndex(ColCount - 1) = x
							
				'WALK THE HEADER
				Do While (Mid(strCols, x, 1) <> " ") And (x < Len(strCols))
					'INC
					x = x + 1
				Loop
			End If			
			'INC
			x = x + 1							
		Loop
			
		'BOUNDS
		If (ColCount > 0) And (RowCount > 0) Then		
			'RESULT
			TableToArray = True
			
			'SET TABLE
			ReDim arrTable(ColCount - 1, RowCount - 1)

			'EACH ROW
			For x = 0 To RowCount - 1
				'EACH COLUMN
				For y = 0 To UBound(arrTable)
					start = arrColIndex(y)
					stopx = 1000
					'NOT LAST ROW
					If (y < ColCount - 1) Then stopx = arrColIndex(y + 1) - 1
					
					arrTable(y, x) = Trim(Mid(strSplit(top + x), start, stopx - start))
				Next
			Next		
		End If
	End If
End Function


'LIST TO ARRAY
Function ListToArray(strList, strTop, arrList, arrTable, ByRef RowCount)
	'RESULT
	ListToArray = False
	top = -1
	bottom = -1	

	'POOR-MAN-REGEX
	strSplit = Split(strList, vbCrLf)	
	
	'FIND TOP
	If (Len(strTop) > 0) Then
		For x = 0 To UBound(strSplit) 
			If (InStr(1, strSplit(x), strTop, 1) > 0) Then top = x
		Next
	End If	

	'BOUNDS
	If (top > -1) Then
		'GET ROWS
		RowCount = 0
		ColCount = UBound(arrList) + 2 'STRTOP + LIST OF PROPS
		'RESULT
		ListToArray = True				
	
		Do 
			'FIND NEXT TOP
			For top = bottom + 1 To UBound(strSplit)
				'FOUND
				If (InStr(1, strSplit(top), strTop, 1) > 0) Then
					'FIND NEXT BOTTOM
					For x = top + 1 To UBound(strSplit)
						'FOUND
						If (InStr(1, strSplit(x), strTop, 1) > 0) Then
							bottom = x - 1
							Exit For
						End If
					Next
					
					'NEVER FOUND, BOTTOM IS END OF LIST
					If (x = UBound(strSplit) + 1) Then bottom = UBound(strSplit)
					
					'SET TABLE
					RowCount = RowCount + 1
					ReDim Preserve arrTable(ColCount - 1, RowCount - 1)
					
					'STRTOP COL POOR-MAN-REGEX
					strSplit1 = Split(strSplit(top), strTop)
					'SET THE COL VALUE
					If (UBound(strSplit1) > 0) Then arrTable(0, RowCount - 1) = Trim(strSplit1(1))
					
					'EACH LINE FROM TOP+1 TO BOTTOM
					For z = top + 1 To bottom 					
						'EACH PROP IN ARRLIST
						For y = 0 To UBound(arrList)
							'ARRLIST COL POOR-MAN-REGEX
							strSplit1 = Split(strSplit(z), arrList(y))
							'SET THE COL VALUE
							If (UBound(strSplit1) > 0) Then arrTable(y + 1, RowCount - 1) = Trim(strSplit1(1))					
						Next
					Next
														
					'STOP SEARCH
					Exit For
				End If
			Next
		
		'UNTIL END OF LIST		
		Loop Until (bottom = UBound(strSplit))
	End If
End Function


'OPEN FILE READ
Function OpenFileRead(strDataFile, ByRef strData)
On Error Resume Next

	'RESULT
	OpenFileRead = False
	
	Dim RetryCount
	RetryCount = 0	

	'FILE EXISTS
	If (objFSO.FileExists(strDataFile)) Then
		Set objFile = objFSO.GetFile(strDataFile)
		
		'TRY LOCK FILE FOR 5 SEC
		Do While (OpenFileRead = False) And (RetryCount < 100)
			'CLEAR ERR
			Err.Clear
			'OPEN FILE
			Set objFileLock = objFSO.OpenTextFile(strDataFile, FOR_APPEND, True)
			'ERR CHECK
			If (Err.Number > 0) Then
				RetryCount = RetryCount + 1
				Wscript.Sleep(50)
				'#DEBUG
'					Wscript.echo "Retry no: " & RetryCount
'					Wscript.echo "Error: " & Err.Description
			Else
				Set objFile = objFSO.OpenTextFile(strDataFile, FOR_READ, True)
				strData = objFile.ReadAll
				objFile.Close
				objFileLock.Close					
				
				'RESULT
				OpenFileRead = True										
			End If
		Loop
	Else
		'FILE DOES NOT EXIST, NOT AN ERR
		OpenFileRead = True
	End If
End Function


'OPEN FILE WRITE
Function OpenFileWrite(strDataFile, strData)
On Error Resume Next
	
	'RESULT
	OpenFileWrite = False
	
	Dim RetryCount
	RetryCount = 0
	
	'TRY OPEN FILE WRITE FOR 5 SEC
	Do While (OpenFileWrite = False) And (RetryCount < 100)
		'CLEAR ERR
		Err.Clear
		'OPEN FILE
		Set objFile = objFSO.OpenTextFile(strDataFile, FOR_WRITE, True)
		'ERR CHECK
		If (Err.Number > 0) Then
			RetryCount = RetryCount + 1
			Wscript.Sleep(50)
			'#DEBUG
'			Wscript.echo "Retry no: " & RetryCount
'			Wscript.echo "Error: " & Err.Description
		Else
			objFile.Write(strData)
			objFile.Close
			
			'RESULT
			OpenFileWrite = True
		End If
	Loop
End Function


'ADD SAVED DATA
Function AddSavedData(ByRef strData, Identifiers, Datas)
	'REPLACE
	For x = 0 To UBound(Identifiers)
		If (IsNull(Identifiers(x))) Then Identifiers(x) = ""
		Identifiers(x) = Replace(Identifiers(x), "|", "`")			
	Next
	'REPLACE
	For x = 0 To UBound(Datas)
		If (IsNull(Datas(x))) Then Datas(x) = ""
		Datas(x) = Replace(Datas(x), "|", "`")
	Next		

	'FOUND
	bFound = False
	splitData = Split(strData, vbCrLf)
		
	'EACH LINE
	For x = 0 To UBound(splitData)
		splitLine = Split(splitData(x), "|")
		
		'LINE HAS IDENTIFIER + DATAS MEMBERS
		If (UBound(splitLine) = UBound(Identifiers) + UBound(Datas) + 1) Then		
			'FOUND
			bFound = True
		
			'ALL IDENTIFIERS IN LINE MATCH
			For y = 0 To UBound(Identifiers)
				If (StrComp(Identifiers(y), splitLine(y), 1) <> 0) Then
					'FOUND
					bFound = False
					'EXIT LINE SEARCH
					Exit For
				End If
			Next
			
			'FOUND
			If (bFound = True) Then
				'UPDATE SPLITLINE
				splitData(x) = Join(Identifiers, "|") & "|" & Join(Datas, "|")
				'EXIT ALL LINES SEARCH
				Exit For
			End If
		End If
	Next
	
	strData = Join(splitData, vbCrLf)
	
	'NOT FOUND, ADD NEW LINE
	If (bFound = False) Then
		strLine = Join(Identifiers, "|") & "|" & Join(Datas, "|")
		If (Len(strData)> 0) Then
			strData = strData & vbCrLf & strLine
		Else
			strData = strLine
		End If
	End If	
End Function


'FIND SAVED Data
Function FindSavedData(strData, Identifiers, ByRef Datas)
	'REPLACE
	For x = 0 To UBound(Identifiers)
		Identifiers(x) = Replace(Identifiers(x), "|", "`")
	Next

	'FOUND
	bFound = False
	splitData = Split(strData, vbCrLf)
	
	'EACH LINE
	For x = 0 To UBound(splitData)
		splitLine = Split(splitData(x), "|")
		
		'LINE HAS IDENTIFIER + DATAS MEMBERS
		If (UBound(splitLine) >= UBound(Identifiers) + UBound(Datas) + 1) Then		
			'FOUND
			bFound = True
		
			'ALL IDENTIFIERS IN LINE MATCH
			For y = 0 To UBound(Identifiers)
				If (StrComp(Identifiers(y), splitLine(y), 1) <> 0) Then
					'FOUND
					bFound = False
					'EXIT LINE SEARCH
					Exit For
				End If
			Next
			
			'FOUND
			If (bFound = True) Then
				'RETURN DATAS FOR FOUND LINE
				For y = UBound(Identifiers) + 1 To UBound(Identifiers) + 1 + UBound(Datas)
					Datas(y - UBound(Identifiers) - 1) = Replace(splitLine(y), "`", "|")
				Next
				'EXIT SEARCH
				Exit For				
			End If
		End If
	Next
	
	'RETURN
	FindSavedData = bFound
End Function


'FORMAT SIZE
Function FormatSize(knSize)
	nSize = Round(knSize)

	'BYTES
	If (nSize < 1024) Then 
		FormatSize = nSize & "B"
	Else
		nSize = nSize / 1024

		'KILOBYTES
		If (nSize < 1024) Then
			FormatSize = Round(nSize, 1) & "KB"
		Else
			nSize = nSize / 1024
			
			'MEGABYTES
			If (nSize < 1024) Then
				FormatSize = Round(nSize, 1) & "MB"
			Else	
				nSize = nSize /1024
				
				'GIGABYTES
				FormatSize = Round(nSize, 1) & "GB"
			End If
		End If
	End If
End Function


'FORMAT HTML
Function FormatHtml(strHtml)
	'FORMAT IN HTML		
	strHtml = Replace(strHtml, "&", "&amp;")
	strHtml = Replace(strHtml, "<", "&lt;")
	strHtml = Replace(strHtml, ">", "&gt;")		
	strHtml = Replace(strHtml, vbCrLf, "<br>")
	strHtml = Replace(strHtml, vbCr, "<br>")			
	strHtml = Replace(strHtml, vbLf, "<br>")
	strHtml = Replace(strHtml, "  ", "&nbsp;&nbsp;")	
	strHtml = Replace(strHtml, Chr(34), "&quot;")
	
	'RESULT
	FormatHtml = strHtml
End Function


'GET NAME
Function GetName
On Error Resume Next
	'COMPUTER NAME
	GetName = UCase(objNetwork.ComputerName) 
	'APPEND DOMAIN
	GetName = GetName & "." & LCase(objSysInfo.DomainDNSName)
End Function


'GET NIC INFO
Function GetNicInfo
	'GET NICS
	Set colNics = objWMI.ExecQuery("Select * from Win32_NetworkAdapter")	
	
	'SET TABLE
	Dim arrTable()
	ReDim arrTable(16 - 1, colNics.Count)
	'SET HEADER
	Dim arrHeader()
	ReDim arrHeader(16 - 1)
	
	'HEADER
	arrHeader(0) = "Connection"
	arrHeader(1) = "Device"
	arrHeader(2) = "Status"
	arrHeader(3) = "MAC"
	arrHeader(4) = "IPs"
	arrHeader(5) = "Subnets"
	arrHeader(6) = "GWs"
	arrHeader(7) = "DHCP <br>Enabled"
	arrHeader(8) = "DHCP <br>Lease obtained"
	arrHeader(9) = "DHCP <br>Lease expires"
	arrHeader(10) = "DHCP <br>Server"
	arrHeader(11) = "DNS"
	arrHeader(12) = "DNS <br>Suffixes"
	arrHeader(13) = "DDNS <br>Registration"
	arrHeader(14) = "NETBIOS <br>over TCP/IP"
	arrHeader(15) = "WINS"
	
	'EACH NIC
	c = 0	
	For Each objNic In colNics
		'NIC HAS ASSOCIATED NETWORK CONNECTION
		If (Len(objNic.NetConnectionId) > 0) Then
			'CONNECTION
			arrTable(0, c) = objNic.NetConnectionId 			
		
	        'GET PNP DEVICE FOR NIC
	        strPNPDeviceID = Replace(objNic.PNPDeviceID, "\", "\\")
	        Set colPnpEntities = objWMI.ExecQuery("Select * from Win32_PNPEntity where DeviceID = '" & strPNPDeviceID & "'")
	        
	        'EACH DEVICE NUMBERED NAME
	        For Each objPnp In colPnpEntities
	        	'DEVICE
	        	arrTable(1, c) = objPnp.Caption
	        Next	               	        
	        
			'ENABLED
			If (objNic.NetConnectionStatus = 2) Or (IsNull(objNic.NetConnectionStatus)) Then
				'STATUS
				arrTable(2, c) = "Enabled"
		        'MAC
				arrTable(3, c) = objNic.MACAddress 
			
				'GET CONFIGURATION FOR NIC
				Set colNetConfigs = objWMI.ExecQuery("Associators of {Win32_NetworkAdapter.DeviceID=" & Chr(34) & objNic.DeviceId & Chr(34) & "} where ResultClass = Win32_NetworkAdapterConfiguration")
				
				'EACH CONFIGURATION
				For Each objNetConfig In colNetConfigs
					'IP
					If (Not IsNull(objNetConfig.IPAddress)) Then arrTable(4, c) = Join(objNetConfig.IPAddress, vbCrLf)
					'Mask
					If (Not IsNull(objNetConfig.IPSubnet)) Then arrTable(5, c) = Join(objNetConfig.IPSubnet, vbCrLf)
					'GW
					If (Not IsNull(objNetConfig.DefaultIPGateway)) Then arrTable(6, c) = Join(objNetConfig.DefaultIPGateway, vbCrLf)
					'DHCP ENABLED
					arrTable(7, c) = objNetConfig.DHCPEnabled
					If (objNetConfig.DHCPEnabled) Then
						'DHCP LEASE OPTAINED
						arrTable(8, c) = objNetConfig.DHCPLeaseObtained
						'DHCP LEASE EXPIRES
						arrTable(9, c) = objNetConfig.DHCPLeaseExpires & vbCrLf
						'DHCP SERVER
						arrTable(10, c) = objNetConfig.DHCPServer & vbCrLf
					End If
					
					'DNS
					If (Not IsNull(objNetConfig.DNSServerSearchOrder)) Then arrTable(11, c) = Join(objNetConfig.DNSServerSearchOrder, vbCrLf)
					'DNS SUFFIX
					If (Not IsNull(objNetConfig.DNSDomainSuffixSearchOrder)) Then arrTable(12, c) = Join(objNetConfig.DNSDomainSuffixSearchOrder, vbCrLf)
					'DDNS REGISTRATION ENABLED
					arrTable(13, c) = objNetConfig.FullDNSRegistrationEnabled & vbCrLf											
					'NETBIOS OVER TCPIP
					If (objNetConfig.TcpipNetbiosOptions = 0) Then
						arrTable(14, c) = "Default"	
					ElseIf (objNetConfig.TcpipNetbiosOptions = 1) Then
						arrTable(14, c) = "Enabled"
					ElseIf (objNetConfig.TcpipNetbiosOptions = 2) Then
						arrTable(14, c) = "Disabled"
					End If
					'WINS
					If (Len(objNetConfig.WINSSecondaryServer) > 0) Then
						arrTable(15, c) = objNetConfig.WINSPrimaryServer & vbCrLf & objNetConfig.WINSSecondaryServer
					ElseIf (Len(objNetConfig.WINSPrimaryServer) > 0) Then
						arrTable(15, c) = objNetConfig.WINSPrimaryServer
					End If
				Next	
			Else	
				'STATUS
				arrTable(2, c) = "Disabled"
		        'MAC
				arrTable(3, c) = objNic.MACAddress 
			End If					        
			
			'INC
			c = c + 1
		End If
	Next
	
	'RESULT
	GetNicInfo = "Connections:" & "<br>" & ArrayToHtml(arrHeader, arrTable, c) 
End Function


'GET ROUTE TABLE
Function GetRouteTable
	Dim arrTable()
	
	'USE ROUTE
	strExe = "%windir%\system32\route.exe" 
	strExeArguments = "print"
	
	'RUN EXE WITH 3 SEC TIMEOUT
	If (GetExeOutput(strExe & " " & strExeArguments, strStd, 3, DidTimeout)) Then
		'POOR-MAN-REGEX
		strSplit = Split(strStd, "Persistent Routes:")
		'FOUND
		If (UBound(strSplit) > 0) Then	
			strTable = strSplit(0)
			'REPLACE THE HEADER TO SOMETHING LEFT ALIGNED, NO SPACES			
			strTop0 = "Network Destination        Netmask          Gateway       Interface  Metric"
			strTop1 = "NetworkDestination Netmask         Gateway           Interface       Metric"			
			strTable = Replace(strTable, strTop0, strTop1)
			strBottom = "==========================================================================="
	
			'TABLE TO ARRAY
			If (TableToArray(strTable, strTop1, strBottom, arrTable, RowCount, False)) Then
				'SET HEADER
				Dim arrHeader()
				ReDim arrHeader(UBound(arrTable))
				
				'HEADER
				arrHeader(0) = "Destination"
				arrHeader(1) = "Subnet <br> mask"
				arrHeader(2) = "Gateway"
				arrHeader(3) = "Interface"
				arrHeader(4) = "Metric"			
			
				'RESULT
				GetRouteTable = "Routing table:" & "<br>" & ArrayToHtml(arrHeader, arrTable, RowCount)
			End If
		End If
	
		'POOR-MAN-REGEX
		strSplit = Split(strStd, "Persistent Routes:")
		'FOUND
		If (UBound(strSplit) > 0) Then		
			strTable = strSplit(1)
			
			'2008 HAS IPV6 ROUTE TABLE
			strSplit1 = Split(strSplit(1), "IPv6 Route Table")
			'FOUND
			If (UBound(strSplit1) > 0) Then strTable = strSplit1(0)
			'2008 HAS BAR THAT NEEDS TO BE REMOVED
			strTable = Replace(strTable, "===========================================================================" & vbCrLf, "")
			
			'REPLACE THE HEADER TO SOMETHING LEFT ALIGNED, NO SPACES					
			strTop0 = "Network Address          Netmask  Gateway Address  Metric"
		  	strTop1 = "NetworkAddress     Netmask          GatewayAddress   Metric"
			strTable = Replace(strTable, strTop0, strTop1)
			strBottom = ""
			
			'TABLE TO ARRAY
			If (TableToArray(strTable, strTop1, strBottom, arrTable1, RowCount, False)) Then
				'SET TABLE
				Dim arrHeader1()
				ReDim arrHeader1(UBound(arrTable1))
				
				'HEADER
				arrHeader1(0) = "Destination"
				arrHeader1(1) = "Subnet <br> mask"
				arrHeader1(2) = "Gateway"
				arrHeader1(3) = "Metric"
			
				'RESULT, REMOVE EMPTY ROW
				GetRouteTable = GetRouteTable & "<br>" & "Persistent routes:" & "<br>" & ArrayToHtml(arrHeader1, arrTable1, RowCount - 1)
			End If			
		End If
	End If		
End Function


'GET SOFTWARE
Function GetSoftware
On Error Resume Next

	'NOT INSTALLED BY DEFAULT ON WINDOWS 2003
	'1.In Add or Remove Programs, click Add/Remove Windows Components. 
	'2. In the Windows Components Wizard, select Management and Monitoring Tools and then click Details.
	'3. In the Management and Monitoring Tools dialog box, select WMI Windows Installer Provider and then click OK.
	'4. Click Next.

	'GET SOFTS
	Set colSofts = objWMI.ExecQuery("Select * from Win32_Product")	
	'PROBE THE CLASS FOR ERROR
	x = colSofts.Count

	'NO ERROR	
	If (Err.Number = 0) Then	
		'SET TABLE
		Dim arrTable()
		ReDim arrTable(5 - 1, colSofts.Count)
		'SET HEADER
		Dim arrHeader()
		ReDim arrHeader(5 - 1)
		
		'HEADER
		arrHeader(0) = "Software"
		arrHeader(1) = "Version"
		arrHeader(2) = "License key"
		arrHeader(3) = "Install <br>date"
		arrHeader(4) = "Location"
		
		'EACH SOFT	
		c = 0
		For Each objSoft In colSofts
			'SOFTWARE
			arrTable(0, c) = objSoft.Description
			'VERSION
			arrTable(1, c) = objSoft.Version		
			'KEY
			arrTable(2, c) = ""
			If (Len(objSoft.InstallDate) > 0) Then
				arrTable(3, c) = Mid(objSoft.InstallDate, 1, 4) & "." & Mid(objSoft.InstallDate, 5, 2) & "." & Mid(objSoft.InstallDate, 7, 2)
			Else
				'INSTALL DATE
				If (Not IsNull(objSoft.InstallDate2)) Then arrTable(3, c) = objSoft.InstallDate2.GetVarDate(False)
			End If
			'LOCATION
			arrTable(4, c) = objSoft.InstallLocation		
			'INC
			c = c + 1
		Next	
		
		'BUBBLE SORT
	    For i = c - 2 to 0 Step -1
	        For j = 0 to i
	            If (UCase(arrTable(0, j)) > UCase(arrTable(0, j + 1))) Then
	            	'EACH FIELD
	            	For x = 0 To UBound(arrTable)
	            		'SWITCH
	                	strHolder = arrTable(x, j + 1)
	                	arrTable(x, j + 1) = arrTable(x, j)
	                	arrTable(x, j) = strHolder
	                Next
	            End If
	        Next
	    Next 			
		
		'RESULT
		GetSoftware = "Software:" & "<br>" & ArrayToHtml(arrHeader, arrTable, c) 
	Else
		'WMI INSTALL PROVIDER NOT INSTALLED
		GetSoftware = "1.In Add or Remove Programs, click Add/Remove Windows Components. <br>" & _
					  "2. In the Windows Components Wizard, select Management and Monitoring Tools and then click Details. <br>" & _
					  "3. In the Management and Monitoring Tools dialog box, select WMI Windows Installer Provider and then click OK. <br>" & _
					  "4. Click Next. <br>"
	End If
End Function


'GET INSTALLED COMPONENTS
Function GetComponents
	Set colWin = GetObject("winmgmts:").InstancesOf("Win32_OperatingSystem") 
	
	'GET WINVER
	If (Not colWin Is Nothing) Then
		For Each objWin In colWin
			'VERSION
			strVer = Mid(objWin.Version, 1, 1)
		Next
	End If	

	'ARRAY INDEXES
	IDX_ID = 0
	IDX_PARENTID = 1
	IDX_COMPNAME = 2
	
	'WINDOWS 5
	If (StrComp(strVer, "5") = 0) Then
		'COMPONENTS
		Set dicComp = CreateObject("Scripting.Dictionary")
		Call dicComp.Add("Accessories and Utilities", Array(01, 0, ""))
		Call dicComp.Add("Active Directory Services", Array(02, 0, ""))
		Call dicComp.Add("Application Server", Array(03, 0, ""))
		Call dicComp.Add("Certificate Services", Array(04, 0, ""))
		Call dicComp.Add("Distributed File System", Array(05, 0, ""))
		Call dicComp.Add("E-mail Services", Array(06, 0, ""))
		Call dicComp.Add("Fax Services", Array(07, 0, "fax"))
		Call dicComp.Add("Indexing Service", Array(08, 0, "indexsrv_system"))
		Call dicComp.Add("Internet Explorer Enhanced Security Configuration", Array(09, 0, ""))   
		Call dicComp.Add("Management and Monitoring Tools", Array(10, 0, ""))
		Call dicComp.Add("Microsoft .NET Framework 2.0", Array(11, 0, "netfx20"))
		Call dicComp.Add("Networking Services", Array(12, 0, ""))
		Call dicComp.Add("Other Network File and Print Services", Array(13, 0, ""))
		Call dicComp.Add("Security Configuration Wizard", Array(14, 0, "scw"))
		Call dicComp.Add("Subsystem for UNIX applications", Array(15, 0, ""))''''''''''''
		Call dicComp.Add("Terminal Server", Array(16, 0, "TerminalServer"))
		Call dicComp.Add("Terminal Server Licensing", Array(17, 0, "licenseserver"))
		Call dicComp.Add("Update Root Certificates", Array(18, 0, "rootautoupdate"))
		Call dicComp.Add("Windows Deployment Services", Array(19, 0, "reminst"))
		Call dicComp.Add("Windows Media Services", Array(20, 0, ""))
		Call dicComp.Add("Windows Sharepoint Services", Array(21, 0, ""))
	
		'ACCESSORIES AND UTILITIES
		Call dicComp.Add("Accessibility wizard", Array(0101, 01, "AccessOpt"))
		Call dicComp.Add("Accessories", Array(0102, 01, ""))
		Call dicComp.Add("Communications", Array(0103, 01, ""))
		
		'ACCESSORIES AND UTILITIES/ACCESSORIES
		Call dicComp.Add("Calculator", Array(010201, 0102, "Calc"))
		Call dicComp.Add("Character map", Array(010202, 0102, "Charmap"))
		Call dicComp.Add("Clipboard viewer", Array(010203, 0102, "Clipbook"))
		Call dicComp.Add("Desktop wallpapers", Array(010204, 0102, "Deskpaper"))
		Call dicComp.Add("Document templates", Array(010205, 0102, "Templates"))
		Call dicComp.Add("Mouse pointers", Array(010206, 0102, "Mousepoint"))
		Call dicComp.Add("Paint", Array(010207, 0102, "Paint"))
		Call dicComp.Add("Wordpad", Array(010208, 0102, "Mswordpad"))
		
		'ACCESSORIES AND UTILITIES/COMMUNICATIONS
		Call dicComp.Add("Chat", Array(010301, 0103, "Chat"))
		Call dicComp.Add("HyperTerminal", Array(010302, 0103, "Hypertrm"))
		
		'ACTIVE DIRECTORY SERVICES
		Call dicComp.Add("Active Directory Application Mode (ADAM)", Array(0201, 02, "adam"))
		Call dicComp.Add("Active Directory Federation Services (ADFS)", Array(0202, 02, "adfstraditional"))
		Call dicComp.Add("Identity Managent for Linux", Array(0203, 02, "idmumgmt"))		
		
		'ACTIVE DIRECTORY SERVICES\ACTIVE DIRECTORY FEDERATION SERVICES
		Call dicComp.Add("ADFS Web Agents", Array(020201, 0202, "adfsclaims"))
		
		'ACTIVE DIRECTORY SERVICES\IDENTITY MANAGEMENT FOR LINUX
		'Call dicComp.Add("Administration Components", Array(020201, 0202, ""))
		'Call dicComp.Add("Password Synchronization", Array(020202, 0202, ""))
		'Call dicComp.Add("Server for NIS", Array(020203, 0202, ""))				
		
		'APPLICATION SERVER
		Call dicComp.Add("Application Server Console", Array(0301, 03, "appsrv_console"))
		Call dicComp.Add("Enable Network COM+ Access", Array(0302, 03, "complusnetwork"))
		Call dicComp.Add("Enable Network DTC Access", Array(0303, 03, "dtcnetwork"))
		Call dicComp.Add("Internet Information Services (IIS)", Array(0304, 03, ""))
		Call dicComp.Add("Message Queueing", Array(0305, 03, ""))
		
		'APPLICATION SERVER\IIS
		Call dicComp.Add("BITS Server Extension", Array(030401, 0304, "BitsServerExtensionsISAPI"))
		Call dicComp.Add("IIS Common Files", Array(030402, 0304, "iis_common"))
		Call dicComp.Add("FTP Service", Array(030403, 0304, "iis_ftp"))
		Call dicComp.Add("Internet Information Services Manager", Array(030404, 0304, "iis_inetmgr"))
		Call dicComp.Add("Internet Printing", Array(030405, 0304, "inetprint"))
		Call dicComp.Add("NNTP Service", Array(030406, 0304, "iis_nntp"))		
		Call dicComp.Add("SMTP Service", Array(030407, 0304, "iis_smtp"))
		Call dicComp.Add("WWW Services", Array(030408, 0304, ""))		
		
		'APPLICATION SERVER\IIS\WWW SERVICE
		Call dicComp.Add("Active Server Pages", Array(03040801, 030408, "iis_asp"))
		Call dicComp.Add("Internet Data Connector", Array(03040802, 030408, "iis_internetdataconnector"))
		Call dicComp.Add("Remote Desktop Web Connection", Array(03040803, 030408, "sakit_web"))
		Call dicComp.Add("Server Side Includes", Array(03040804, 030408, "iis_serversideincludes"))
		Call dicComp.Add("WebDAV Publishing", Array(03040805, 030408, "iis_webdav"))
		Call dicComp.Add("WWW Service", Array(03040806, 030408, "iis_www"))
		
		'APPLICATION SERVER\MESSAGE QUEUEING
		Call dicComp.Add("MSMQ AD Integration", Array(030501, 0305, "msmq_ADIntegrated"))
		Call dicComp.Add("MSMQ Core", Array(030502, 0305, "msmq_Core"))
		Call dicComp.Add("MSMQ Downlevel Client Support", Array(030503, 0305, "msmq_MQDSService"))
		Call dicComp.Add("MSMQ HTTP Support", Array(030504, 0305, "msmq_HTTPSupport"))
		Call dicComp.Add("MSMQ Routing Support", Array(030505, 0305, "msmq_RoutingSupport"))
		Call dicComp.Add("MSMQ Triggers", Array(030506, 0305, "msmq_TriggersService"))
		
		'CERTIFICATE SERVICES
		Call dicComp.Add("Certificate Services CA", Array(0401, 04, "certsrv_server"))
		Call dicComp.Add("Certificate Services Web Enrollment Support", Array(0402, 04, "certsrv_client"))
		
		'DISTRIBUTED FILE SYSTEM
		Call dicComp.Add("DFS Management", Array(0501, 05, "dfsfrsui"))
		Call dicComp.Add("DFS Replication Diagnostic and Configuration", Array(0502, 05, "dfsrhelper"))
		Call dicComp.Add("DFS Replication Service", Array(0503, 05, "dfsr"))
		
		'EMAIL SERVICES
		Call dicComp.Add("POP3 Service", Array(0601, 06, "Pop3Service"))
		
		'IE ENHANCED CONFIGURATION
		Call dicComp.Add("Administrator groups", Array(0901, 09, "IEHardenAdmin"))
		Call dicComp.Add("Other User groups", Array(0902, 09, "IEHardenUser"))
		
		'MANAGEMENT AND MONITORING
		'Call dicComp.Add("Connection Manager Administration Kit", Array(1001, 10, ""))
		'Call dicComp.Add("Connection Point Services", Array(1002, 10, ""))
		Call dicComp.Add("File Server Management", Array(1003, 10, "fsrstandard"))
		Call dicComp.Add("File Server Resource Manager", Array(1004, 10, "srm"))
		'Call dicComp.Add("Hardware Management", Array(1005, 10, ""))
		'Call dicComp.Add("Network Monitor Tools", Array(1006, 10, ""))
		Call dicComp.Add("Print Management Component", Array(1007, 10, "pmcsnap"))
		Call dicComp.Add("SNMP", Array(1008, 10, "snmp"))
		'Call dicComp.Add("Storage Manager for SANs", Array(1009, 10, ""))
		'Call dicComp.Add("WMI SNMP Provider", Array(1010, 10, ""))
		Call dicComp.Add("WMI Windows Installer Provider", Array(1011, 10, "WbemMSI"))
		
		'NETWORKING SERVICES
		Call dicComp.Add("DNS Service", Array(1201, 12, "dns"))
		Call dicComp.Add("DHCP Service", Array(1202, 12, "dhcpserver"))
		Call dicComp.Add("IAS Service", Array(1203, 12, "ias"))
		'Call dicComp.Add("Remote Access Quarantine Service", Array(1204, 12, ""))
		Call dicComp.Add("RPC over HTTP proxy", Array(1205, 12, "netcis"))
		Call dicComp.Add("Simple TCPIP Services", Array(1206, 12, "simptcp"))
		Call dicComp.Add("WINS Service", Array(1207, 12, "wins"))
		
		'OTHER NETWORK FILE AND PRINT SERVICES
		'Call dicComp.Add("Common Log File System", Array(1301, 13, ""))
		'Call dicComp.Add("File Services for Macintosh", Array(1302, 13, ""))
		'Call dicComp.Add("Microsoft Services for NFS", Array(1303, 13, ""))
		'Call dicComp.Add("Print Services for Unix", Array(1304, 13, ""))
		
		'WINDOWS MEDIA SERVICES
		Call dicComp.Add("Multicast and Advertisment Logging Agent", Array(2001, 20, "wms_isapi"))
		Call dicComp.Add("Windows Media Services Core", Array(2002, 20, "wms_server"))
		Call dicComp.Add("Windows Media Services Administrator for Web", Array(2003, 20, "wms_admin_asp"))
		Call dicComp.Add("Windows Media Services MMC", Array(2004, 20, "wms_admin_mmc"))
			
		'SET TABLE
		Dim arrTable()
		ReDim arrTable(4 - 1, dicComp.Count)
		'SET HEADER
		Dim arrHeader()
		ReDim arrHeader(4 - 1)
		
		'HEADER
		arrHeader(0) = "Components"
		arrHeader(1) = "Installed services"
		arrHeader(2) = "Installed sub-services"
		arrHeader(3) = "Installed sub-sub-services"	
	
		strKeyPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Setup\OC Manager\Subcomponents"			
		
		'EACH COMPONENT
		c = 0
		x = 0
		colKeys = dicComp.Keys
		colItems = dicComp.Items
		For x = 0 To dicComp.Count - 1
			'GET ITEM
			objItem = colItems(x)
			bInc = False
			
			'FILTER
			If (objItem(IDX_PARENTID) = 0) Then			
				'READ INSTALLED REGKEY
				If (objReg.GetDWordValue(HKLM, strKeyPath, objItem(IDX_COMPNAME), nInstalled) = 0) And (nInstalled = 1) Then
					'ROLE NAME
					arrTable(0, c) = colKeys(x)
					'INC
					bInc = True
					
				'COMPONENT GROUP
				ElseIf (Len(objItem(IDX_COMPNAME)) = 0) Then
					'SERVICES
					y = 0
					For y = 0 To dicComp.Count - 1
						'GET SERVICE						
						objSvc = colItems(y)
						bSubInc = False
						
						'FILTER SERVICES
						If (objSvc(IDX_PARENTID) = objItem(IDX_ID)) Then
							'READ INSTALLED REGKEY
							If (objReg.GetDWordValue(HKLM, strKeyPath, objSvc(IDX_COMPNAME), nInstalled) = 0) And (nInstalled = 1) Then
								'ADD COMMA
								If (Len(arrTable(1, c)) > 0) Then arrTable(1, c) = arrTable(1, c) & "," & vbCrLf 															
								'ROLE NAME
								arrTable(1, c) = arrTable(1, c) & colKeys(y)							
								'INC
								bInc = True
								bSubInc = True
								
							'COMPONENT GROUP
							ElseIf (Len(objSvc(IDX_COMPNAME)) = 0) Then
								'SUB-SERVICES
								z = 0
								For z = 0 To dicComp.Count - 1
									'GET SUB-SERVICE						
									objSubSvc = colItems(z)
									bSubSubInc = False
									
									'FILTER SUB-SERVICES
									If (objSubSvc(IDX_PARENTID) = objSvc(IDX_ID)) Then
										'READ INSTALLED REGKEY
										If (objReg.GetDWordValue(HKLM, strKeyPath, objSubSvc(IDX_COMPNAME), nInstalled) = 0) And (nInstalled = 1) Then
											'ADD COMMA
											If (Len(arrTable(2, c)) > 0) Then arrTable(2, c) = arrTable(2, c) & "," & vbCrLf 															
											'ROLE NAME
											arrTable(2, c) = arrTable(2, c) & colKeys(z)							
											'INC
											bInc = True
											bSubInc = True

										'COMPONENT GROUP
										ElseIf (Len(objSubSvc(IDX_COMPNAME)) = 0) Then										
											'SUB-SUB-SERVICES
											q = 0
											For q = 0 To dicComp.Count - 1
												'GET SUB-SUB-SERVICE						
												objSubSubSvc = colItems(q)
												
												'FILTER SUB-SERVICES
												If (objSubSubSvc(IDX_PARENTID) = objSubSvc(IDX_ID)) Then
													'READ INSTALLED REGKEY
													If (objReg.GetDWordValue(HKLM, strKeyPath, objSubSubSvc(IDX_COMPNAME), nInstalled) = 0) And (nInstalled = 1) Then
														'ADD COMMA
														If (Len(arrTable(3, c)) > 0) Then arrTable(3, c) = arrTable(3, c) & "," & vbCrLf 															
														'ROLE NAME
														arrTable(3, c) = arrTable(3, c) & colKeys(q)							
														'INC
														bInc = True
														bSubSubInc = True
													End If
												End If
											Next
											
											'SUB-SUB-INC
											If (bSubSubInc) Then
												If (Len(arrTable(2, c)) > 0) Then arrTable(2, c) = arrTable(2, c) & "," & vbCrLf 															
												'ROLE NAME
												arrTable(2, c) = arrTable(2, c) & colKeys(z)
											End If																		
										End If
									End If
								Next
								
								'SUB-INC
								If (bSubInc) Then
									If (Len(arrTable(1, c)) > 0) Then arrTable(1, c) = arrTable(1, c) & "," & vbCrLf 															
									'ROLE NAME
									arrTable(1, c) = arrTable(1, c) & colKeys(y)
								End If								
							End If					
						End If
					Next
				End If	
				
				'INC
				If (bInc) Then 
					'ROLE NAME
					arrTable(0, c) = colKeys(x)				
					'INC
					c = c + 1
				End If
			End If
		Next		
		

		'RESULT
		GetComponents = "Components:" & "<br>" & ArrayToHtml(arrHeader, arrTable, c)		
	End If
End Function


'GET SERVER ROLES
Function GetServerRoles
	Set colWin = GetObject("winmgmts:").InstancesOf("Win32_OperatingSystem") 
	
	'GET WINVER
	If (Not colWin Is Nothing) Then
		For Each objWin In colWin
			'VERSION
			strVer = Mid(objWin.Version, 1, 1)
		Next
	End If	
	
	'WINDOWS 6
	If (StrComp(strVer, "6") = 0) Then
		'GET ROLES/FEATURES TOP LEVEL
		Set colRoles = objWMI.ExecQuery("Select * from Win32_ServerFeature where ParentID=0")
		
		'SET TABLE
		Dim arrTable()
		ReDim arrTable(4 - 1, colRoles.Count)
		'SET HEADER
		Dim arrHeader()
		ReDim arrHeader(4 - 1)
		
		'HEADER
		arrHeader(0) = "Roles and features"
		arrHeader(1) = "Installed services"
		arrHeader(2) = "Installed sub-services"
		arrHeader(3) = "Installed sub-sub-services"
		
		'EACH ROLE/FEATURE
		c = 0
		For Each objRole In colRoles
			'ROLE NAME
			arrTable(0, c) = objRole.Name
			
			'GET SERVICES 
			Set colServices = objWMI.ExecQuery("Select * from Win32_ServerFeature where ParentID=" & objRole.ID)						
			'EACH SERVICE
			For Each objService In colServices
				'ADD COMMA
				If (Len(arrTable(1, c)) > 0) Then arrTable(1, c) = arrTable(1, c) & "," & vbCrLf 			
				'SERVICE NAME
				arrTable(1, c) = arrTable(1, c) & objService.Name
				
				'GET SUB-SERVICES 
				Set colSubServices = objWMI.ExecQuery("Select * from Win32_ServerFeature where ParentID=" & objService.ID)							
				'EACH SUB-SERVICE
				For Each objSubService In colSubServices
					'ADD COMMA
					If (Len(arrTable(2, c)) > 0) Then arrTable(2, c) = arrTable(2, c) & "," & vbCrLf 				
					'SERVICE NAME
					arrTable(2, c) = arrTable(2, c) & objSubService.Name
					
					'GET SUB-SUB-SERVICES 
					Set colSubSubServices = objWMI.ExecQuery("Select * from Win32_ServerFeature where ParentID=" & objSubService.ID)							
					'EACH SUB-SUB-SERVICE
					For Each objSubSubService In colSubSubServices
						'ADD COMMA
						If (Len(arrTable(3, c)) > 0) Then arrTable(3, c) = arrTable(3, c) & "," & vbCrLf 					
						'SERVICE NAME
						arrTable(3, c) = arrTable(3, c) & objSubSubService.Name
					Next									
				Next
			Next
			
			'INC
			c = c + 1
		Next
		
		'RESULT
		GetServerRoles = "Roles and features:" & "<br>" & ArrayToHtml(arrHeader, arrTable, c)
	End If
End Function


'GET VIRTUAL
Function GetVirtual
	Set colWin = GetObject("winmgmts:").InstancesOf("Win32_OperatingSystem") 
	
	'GET WINVER
	If (Not colWin Is Nothing) Then
		For Each objWin In colWin
			'VERSION
			strVer = Mid(objWin.Version, 1, 1)
		Next
	End If	
	
	'SET TABLE
	Dim arrTable()
	ReDim arrTable(4 - 1, 0)
	'SET HEADER
	Dim arrHeader
	ReDim arrHeader(4 - 1)	
	
	'HEADER
	arrHeader(0) = "Machine type"
	arrHeader(1) = "Virtualization type"
	arrHeader(2) = "Host/Guest"
	arrHeader(3) = "Host name"		

	'MACHINE TYPE
	arrTable(0, 0) = "Physical machine"
	'VIRTUALIZATION
	arrTable(1, 0) = "-"
	'HOST/GUEST
	arrTable(2, 0) = "-"		
	'HOSTNAME
	arrTable(3, 0) = "-"			
	
	'WINDOWS 6
	If (StrComp(strVer, "6") = 0) Then
		'GET ROLES/FEATURES TOP LEVEL
		Set colRoles = objWMI.ExecQuery("Select * from Win32_ServerFeature where ID=20")
		'IS H-V HOST
		If (colRoles.Count > 0) Then 
			'VIRTUALIZATION
			arrTable(1, 0) = "Hyper-v"
			'HOST/GUEST
			arrTable(2, 0) = "Host"			
		End If
	End If
	
	strKeyPath = "SOFTWARE\Microsoft\Virtual Machine\Guest\Parameters"
	
	'READ VIRTUAL HOST NAME
	If (objReg.GetStringValue(HKLM, strKeyPath, "HostName", strVirtualHostName) = 0) Then
		'VIRTUAL MACHINE
		arrTable(0, 0) = "Virtual machine"
		'VIRTUALIZATION
		arrTable(1, 0) = "Hyper-v"		
		'HOST/GUEST
		arrTable(2, 0) = "Guest"		
		'HOSTNAME
		arrTable(3, 0) = strVirtualHostName		
	End If		
	
	'RESULT
	GetVirtual = "Virtualization roles:" & "<br>" & ArrayToHtml(arrHeader, arrTable, 1) 		
End Function


'GET DC
Function GetDC
	'GET COMPUTERSYSTEM
	Set colComputers = objWMI.ExecQuery("Select * from Win32_ComputerSystem")
	
	'SET TABLE
	Dim arrTable()
	ReDim arrTable(7 - 1, 0)
	DomainRoles = Array("Standalone Workstation", "Member Workstation", "Standalone Server", "Member Server", "Backup Domain Controller", "Primary Domain Controller", "Unknown")
	'SET HEADER
	Dim arrHeader
	ReDim arrHeader(7 - 1)	
	
	'HEADER
	arrHeader(0) = "Computer role"
	arrHeader(1) = "Schema master"
	arrHeader(2) = "Domain naming master"
	arrHeader(3) = "PDC emulator"
	arrHeader(4) = "RID master"
	arrHeader(5) = "Infrastructure master"
	arrHeader(6) = "Global Catalogs"

	'FSMOs	
	arrTable(1, 0) = "-"
	arrTable(2, 0) = "-"
	arrTable(3, 0) = "-"
	arrTable(4, 0) = "-"
	arrTable(5, 0) = "-"
	arrTable(6, 0) = "-"		
	
	'EACH CS
	For Each objComputer In colComputers
		'DOMAIN ROLE
		arrTable(0, 0) = DomainRoles(objComputer.DomainRole)
		
		'DOMAIN CONTROLLER
		If (objComputer.DomainRole = 4) Or (objComputer.DomainRole = 5) Then
			Set objRootDSE = GetObject("LDAP://rootDSE")
			 
			'SCHEMA MASTER
			Set objSchema = GetObject("LDAP://" & objRootDSE.Get("schemaNamingContext"))
			strSchemaMaster = objSchema.Get("fSMORoleOwner")
			Set objNtds = GetObject("LDAP://" & strSchemaMaster)
			Set objComputer = GetObject(objNtds.Parent)
			arrTable(1, 0) = Replace(objComputer.Name, "CN=", "")
			 
			'DOMAIN NAMING MASTER			 
			Set objPartitions = GetObject("LDAP://CN=Partitions," & objRootDSE.Get("configurationNamingContext"))
			strDomainNamingMaster = objPartitions.Get("fSMORoleOwner")
			Set objNtds = GetObject("LDAP://" & strDomainNamingMaster)
			Set objComputer = GetObject(objNtds.Parent)
			arrTable(2, 0) =  Replace(objComputer.Name, "CN=", "")
			
			'PDC EMULATOR 
			Set objDomain = GetObject("LDAP://" & objRootDSE.Get("defaultNamingContext"))
			strPdcEmulator = objDomain.Get("fSMORoleOwner")
			Set objNtds = GetObject("LDAP://" & strPdcEmulator)
			Set objComputer = GetObject(objNtds.Parent)
			arrTable(3, 0) =   Replace(objComputer.Name, "CN=", "")
			 
			'RID MASTER
			Set objRidManager = GetObject("LDAP://CN=RID Manager$,CN=System," & objRootDSE.Get("defaultNamingContext"))
			strRidMaster = objRidManager.Get("fSMORoleOwner")
			Set objNtds = GetObject("LDAP://" & strRidMaster)
			Set objComputer = GetObject(objNtds.Parent)
			arrTable(4, 0) = Replace(objComputer.Name, "CN=", "")
			 
			'INFRASTRUCTURE MASTER
			Set objInfrastructure = GetObject("LDAP://CN=Infrastructure," & objRootDSE.Get("defaultNamingContext"))
			strInfrastructureMaster = objInfrastructure.Get("fSMORoleOwner")
			Set objNtds = GetObject("LDAP://" & strInfrastructureMaster)
			Set objComputer = GetObject(objNtds.Parent)
			arrTable(5, 0) = Replace(objComputer.Name, "CN=", "")
			
			Const NTDSDSA_OPT_IS_GC = 1
			arrTable(6, 0) = "" 
			
			'GLOBAL CATALOGS
			Set colGC = GetObject("LDAP://OU=Domain Controllers," & objRootDSE.Get("defaultNamingContext")) 
			   For Each GC In colGC 
			       'clean up the ldap response 
			       GC = Replace(GC.Name, "CN=", "") 
			       Set objRootDSE = GetObject("LDAP://" & GC & "/rootDSE") 
			       strDsServiceDN = objRootDSE.Get("dsServiceName") 
			       Set objDsRoot  = GetObject("LDAP://" & GC & "/" & strDsServiceDN) 
			       
				   Call Err.Clear
				   On Error Resume Next
				       intOptions = objDsRoot.Get("options") 
			       
				       'CHECK IF PREV COMMAND FAILED
				       If (intOptions And NTDSDSA_OPT_IS_GC) And (Err.Number = 0) Then 
				       		'COMMA
				       		If (Len(arrTable(6, 0)) > 0) Then arrTable(6, 0) = arrTable(6, 0) & "," & vbCrLf 
				       		'ADD
				           	arrTable(6, 0) = arrTable(6, 0) & GC
				       End If 
				   On Error Goto 0				       
			   Next 			
		End If
	Next
	
	'RESULT
	GetDC = "Active Directory roles:" & "<br>" & ArrayToHtml(arrHeader, arrTable, 1) 	
End Function


'GET MOTHERBOARD
Function GetMotherBoard(ByVal strData)
	'GET MOTHERBOARDS
	Set colMBs = objWMI.ExecQuery("Select * from Win32_BaseBoard")
	'GET BIOS
	Set colBIOSs = objWMI.ExecQuery("Select * from Win32_BIOS")	
	
	'SET TABLE
	Dim arrTable()
	ReDim arrTable(10 - 1, colMBs.Count)
	'SET HEADER
	Dim arrHeader
	ReDim arrHeader(10 - 1)
	
	'HEADER
	arrHeader(0) = "Manufacturer"
	arrHeader(1) = "Model"
	arrHeader(2) = "Serial number"
	arrHeader(3) = "Part number"	
	arrHeader(4) = "BIOS Manufacturer"
	arrHeader(5) = "BIOS Type"
	arrHeader(6) = "BIOS Version"
	arrHeader(7) = "Northbridge"
	arrHeader(8) = "Southbridge"
	arrHeader(9) = "Memory Size"
	
	'EACH MOTHERBOARD
	c = 0
	For Each objMB In colMBs
		'MANUFACTURER
		arrTable(0, c) = objMB.Manufacturer
		'MODEL
		If (Len(objMB.Model) > 0) Then
			arrTable(1, c) = objMB.Model
		Else
			arrTable(1, c) = objMB.Product		
		End If		
		'SERIAL NUMBER
		arrTable(2, c) = objMB.SerialNumber	
		'PART NUMBER
		arrTable(3, c) = objMB.PartNumber			
		'INC
		c = c + 1
	Next	
	
	'EACH BIOS
	For Each objBIOS In colBIOSs
		'MANUFACTURER
		arrTable(4, 0) =  objBIOS.Manufacturer	
		'TYPE
		arrTable(5, 0) =  objBIOS.Caption
		'VERSION
		arrTable(6, 0) =  objBIOS.Version		
	Next
	
	'ISOLATE CHIPSET
	strSplit = Split(strData, "Chipset" & vbCrLf & "-------------------------------------------------------------------------")
	'FOUND
	If (UBound(strSplit) > 0) Then strData = strSplit(1)
	'ISOLATE CHIPSET
	strSplit = Split(strData, "Memory SPD" & vbCrLf & "-------------------------------------------------------------------------")
	If (UBound(strSplit) > 0) Then strData = strSplit(0)
	'MAKE LINES
	strSplit = Split(strData, vbCrLf)

	For x = 0 To UBound(strSplit)
		'NORTHBRIDGE
		strSplit1 = Split(strSplit(x), "Northbridge")
		'FOUND
		If (UBound(strSplit1) > 0) Then arrTable(7, 0) = Trim(strSplit1(1))
		'SOUTHBRIDGE
		strSplit1 = Split(strSplit(x), "Southbridge")
		'FOUND
		If (UBound(strSplit1) > 0) Then arrTable(8, 0) = Trim(strSplit1(1))		
		'MEMORY SIZE
		strSplit1 = Split(strSplit(x), "Memory Size")
		'FOUND
		If (UBound(strSplit1) > 0) Then arrTable(9, 0) = Trim(strSplit1(1))
	Next	
	
	'AT LEAST ONE MOTHERBOART
	If (c = 0) Then c = 1
	
	'RESULT
	GetMotherBoard = "Motherboard:" & "<br>" & ArrayToHtml(arrHeader, arrTable, c) 
End Function


'GET PROCESSORS
Function GetProcessors(ByVal strData)
	'ISOLATE PROCESSORS
	strSplit = Split(strData, "Processors Information" & vbCrLf & "-------------------------------------------------------------------------")
	'FOUND
	If (UBound(strSplit) > 0) Then strData = strSplit(1)
	'ISOLATE PROCESSORS
	strSplit = Split(strData, "Thread dumps" & vbCrLf & "-------------------------------------------------------------------------")
	If (UBound(strSplit) > 0) Then strData = strSplit(0)
	
	'TABLEARR
	Dim arrTable()		
	
	strTop = "Processor"
	Dim arrLines(16)
	arrLines(0) = "Name"
	arrLines(1) = "Codename"
	arrLines(2) = "Specification"
	arrLines(3) = "Package (platform ID)"
	arrLines(4) = "Technology"		
	arrLines(5) = "Core Speed"		
	arrLines(6) = "Multiplier x FSB"		
	arrLines(7) = "Rated Bus speed"		
	arrLines(8) = "Stock frequency"		
	arrLines(9) = "Instructions sets"		
	arrLines(10) = "L1 Data cache"		
	arrLines(11) = "L1 Instruction cache"	
	arrLines(12) = "L2 cache"
	arrLines(13) = "FID range"
	arrLines(14) = "Max VID"
	arrLines(15) = "Number of cores"
	arrLines(16) = "Number of threads"

	'LIST TO ARRAY
	If (ListToArray(strData, strTop, arrLines, arrTable, RowCount)) Then		
		'SET HEADER
		Dim arrHeader()
		ReDim arrHeader(UBound(arrTable) + 1)
			
		'HEADER
		arrHeader(0) = "No"
		arrHeader(1) = "Name"
		arrHeader(2) = "Codename"
		arrHeader(3) = "Specification"
		arrHeader(4) = "Socket"
		arrHeader(5) = "Technology"		
		arrHeader(6) = "Core <br>Speed"		
		arrHeader(7) = "Multiplier <br>x FSB"		
		arrHeader(8) = "Rated <br>bus speed"		
		arrHeader(9) = "Stock <br>frequency"		
		arrHeader(10) = "Instructions <br>sets"		
		arrHeader(11) = "L1 Data <br>cache"		
		arrHeader(12) = "L1 Instruction <br>cache"	
		arrHeader(13) = "L2 cache"
		arrHeader(14) = "FID range"
		arrHeader(15) = "Max VID"
		arrHeader(16) = "Cores"
		arrHeader(17) = "Threads"
		
		'TWEAK PROCESSOR
		For x = 0 To RowCount - 1
			'POOR-MAN-REGEX
			strSplit = Split(arrTable(0, x), "ID")
			'FOUND
			If (UBound(strSplit) > 0) Then arrTable(0, x) = strSplit(0)
			'SPECIFICATION
			arrTable(3, x) = Replace(arrTable(3, x), "  ", vbCrLf, 1, 1)
			'INSTRUCTION SETS
			arrTable(10, x) = Replace(arrTable(10, x), ",", "," & vbCrLf)			
			'L1 DATA CACHE
			arrTable(11, x) = Replace(arrTable(11, x), ",", "," & vbCrLf)
			'L1 INSTRUCTION CACHE
			arrTable(12, x) = Replace(arrTable(12, x), ",", "," & vbCrLf)	
			'L2 CACHE
			arrTable(13, x) = Replace(arrTable(13, x), ",", "," & vbCrLf)				
		Next		
	
		'RESULT
		GetProcessors = "Processors:" & "<br>" & ArrayToHtml(arrHeader, arrTable, RowCount)		
	End If
End Function


'GET MEMORY
Function GetMemory(ByVal strData)
	'ISOLATE MEMORY
	strSplit = Split(strData, "Memory SPD" & vbCrLf & "-------------------------------------------------------------------------")
	'FOUND
	If (UBound(strSplit) > 0) Then strData = strSplit(1)
	'ISOLATE MEMORY
	strSplit = Split(strData, "Monitoring" & vbCrLf & "-------------------------------------------------------------------------")
	If (UBound(strSplit) > 0) Then strData = strSplit(0)
	'ISOLATE DIMMS INFO
	strSplit = Split(strData, "DIMM #				1")
	If (UBound(strSplit) > 0) Then strData = "DIMM #				1" & strSplit(1)
	
	'TABLEARR
	Dim arrTable()		
	
	strTop = "DIMM #"
	Dim arrLines(10)
	arrLines(0) = "SMBus address"
	arrLines(1) = "Memory type"
	arrLines(2) = "Manufacturer (ID)"
	arrLines(3) = "Size"
	arrLines(4) = "Max bandwidth"		
	arrLines(5) = "Serial number"		
	arrLines(6) = "Part number"		
	arrLines(7) = "Manufacturing date"		
	arrLines(8) = "Number of banks"		
	arrLines(9) = "JEDEC #1"		

	'LIST TO ARRAY
	If (ListToArray(strData, strTop, arrLines, arrTable, RowCount)) Then		
		'SET HEADER
		Dim arrHeader()
		ReDim arrHeader(UBound(arrTable) + 1)		
		
		'HEADER
		arrHeader(0) = "DIMM"
		arrHeader(1) = "Address"
		arrHeader(2) = "Type"
		arrHeader(3) = "Manufacturer"
		arrHeader(4) = "Size"
		arrHeader(5) = "Bandwidth"
		arrHeader(6) = "Serial number"
		arrHeader(7) = "Part number"
		arrHeader(8) = "Manufacture date"
		arrHeader(9) = "Number <br>of banks"
		arrHeader(10) = "Timings"
	
		'RESULT
		GetMemory = "Memory:" & "<br>" & ArrayToHtml(arrHeader, arrTable, RowCount)		
	End If
End Function


'GET DISKS
Function GetDisks
	'GET DISKS AS PHYSICAL DISKS
	Set colDisks = objWMI.ExecQuery("Select * from Win32_DiskDrive")	
	
	'SET TABLE
	Dim arrTable(), arrVTable()
	ReDim arrTable(7 - 1, colDisks.Count)
	ReDim arrVTable(8 - 1, 0)
	'SET HEADER
	Dim arrHeader, arrVHeader
	ReDim arrHeader(7 - 1)
	ReDim arrVHeader(8 - 1)
	
	'HEADER
	arrHeader(0) = "No"
	arrHeader(1) = "Name"
	arrHeader(2) = "Model"	
	arrHeader(3) = "Serial number"
	arrHeader(4) = "Interface"
	arrHeader(5) = "Media type"
	arrHeader(6) = "Size"
	
	'VOLUME HEADER
	arrVHeader(0) = "Drive letter"
	arrVHeader(1) = "Label"	
	arrVHeader(2) = "Parent <br>disk"	
	arrVHeader(3) = "File system"
	arrVHeader(4) = "Serial"	
	arrVHeader(5) = "Capacity"
	arrVHeader(6) = "Free space"
	arrVHeader(7) = "Used space"
	
	'EACH DISK
	c = 0
	vc = 0
	For Each objDisk In colDisks
		'NUMBER
		arrTable(0, c) = Right(objDisk.DeviceID, 1)
		arrTable(1, c) = objDisk.DeviceID		
		arrTable(2, c) = objDisk.Model
	
		On Error Resume Next
			arrTable(3, c) = objDisk.SerialNumber	'NOT SUPPORTED ON 2003
		On Error Goto 0

		arrTable(4, c) = objDisk.InterfaceType
		arrTable(5, c) = objDisk.MediaType
		arrTable(6, c) = FormatSize(objDisk.Size)
		
		'GET PARTITIONS FOR DISK
	    strDeviceID = Replace(objDisk.DeviceID, "\", "\\")
	    Set colPartitions = objWMI.ExecQuery("Associators of {Win32_DiskDrive.DeviceID=" & Chr(34) & strDeviceID & Chr(34) & "} where AssocClass = Win32_DiskDriveToDiskPartition")
	 
	 	'EACH PARTITION
	 	y = 0
	    For Each objPartition In colPartitions
	        'GET VOLUMES FOR PARTITION
	        Set colLogicalDisks = objWMI.ExecQuery("Associators of {Win32_DiskPartition.DeviceID=" & Chr(34) & objPartition.DeviceID & Chr(34) & "} where AssocClass = Win32_LogicalDiskToPartition")	        
	        
			'EACH VOLUME AS LOGICALDISK	 			
	        For Each objLogicalDisk In colLogicalDisks
	        	'ENLARGE ARRAY
	        	ReDim Preserve arrVTable(8 - 1, vc)
	        	'DRIVE LETTER
	        	arrVTable(0, vc) = objLogicalDisk.DeviceID
	        	'DRIVE LABEL
	        	arrVTable(1, vc) = objLogicalDisk.VolumeName
	        	'PARENT DISK
	        	arrVTable(2, vc) = Right(objDisk.DeviceID, 1)
	        	'FILE SYSTEM
	        	arrVTable(3, vc) = objLogicalDisk.Filesystem
	        	'SERIAL
	        	arrVTable(4, vc) = objLogicalDisk.VolumeSerialNumber	        	
	        	'CAPACITY
	        	arrVTable(5, vc) = FormatSize(objLogicalDisk.Size)
	        	'FREE
	        	arrVTable(6, vc) = FormatSize(objLogicalDisk.FreeSpace)
	        	'USED
	        	arrVTable(7, vc) = FormatSize(objLogicalDisk.Size - objLogicalDisk.FreeSpace)
	        	'INC
	        	vc = vc + 1
	        Next
	        
        	'INC
        	y = y + 1
	    Next
	    
	    'INC
	    c = c + 1	
	 Next
	 
	'RESULT
	GetDisks = "Disks:" & "<br>" & ArrayToHtml(arrHeader, arrTable, c) & "<br>" & "Volumes:" & "<br>" & ArrayToHtml(arrVHeader, arrVTable, vc)		 
End Function


'GETSAVEATTRIBUTES
Function GetSaveAttributes
	'LOAD DATA
	If (OpenFileRead(strCurPath & "\" & strSavedFile, strData)) Then						
		'SET DATA (PART)
		Dim Identifiers(0)
		Dim Datas(1)
		Datas(SAVDATA_TIMEDATE) = Now						

		'---------------------- SOFTWARE ----------------------
		'## STATUS
		WScript.Echo "Collecting software information..."
		
		'GET INSTALLED COMPONENTS
		strComponents = GetComponents
		'GET SERVER ROLES
		strServerRoles = GetServerRoles
		
		'GET SOFTWARE + INSTALLED COMPONENTS
		If (Len(strComponents) > 0) Then
			strResult = GetSoftware & "<br>" & strComponents & "<br>" & GetDC & "<br>"
		ElseIf (Len(strServerRoles) > 0) Then
			'GET SOFTWARE + SERVER ROLES	
			strResult = GetSoftware & "<br>" & strServerRoles & "<br>" & GetDC & "<br>"
		Else
			'SAME - SERVER ROLES
			strResult = GetSoftware & "<br>" & GetDC & "<br>"
		End If

		'SET DATA	
		Identifiers(SAVID_ATTRIBUTE) = OPT_SOFTWARE					
		Datas(SAVDATA_VALUE) = strResult
		'ADD DATA
		Call AddSavedData(strData, Identifiers, Datas)				
		
		'SAVE DATA
		Call OpenFileWrite(strCurPath & "\" & strSavedFile, strData)				
			
		'---------------------- HARDWARE ----------------------
		'## STATUS
		WScript.Echo "Collecting hardware information..."
				
		'GET ENV VARS
		Set objProcEnv = objShell.Environment("Process")
	
		'USE CPUZ 32
		If (StrComp(objProcEnv("PROCESSOR_ARCHITECTURE"), "x86", 1) = 0) Then
			strExe = Chr(34) & strCurPath & "\"	& "cpuz32.exe" & Chr(34)
		Else
			'USE CPUZ 64
			strExe = Chr(34) & strCurPath & "\"	& "cpuz64.exe" & Chr(34)
		End If
		'MAKE IN ROOT, PROGRAM CANT HANDLE SPACES AND DOUBLEQUOTES
		strExeArguments = "-txt=" & "C:\" & objFSO.GetBaseName(strTxtFile)
		
		'RUN EXE WITH WAITFOR, STRANGE WAY IT PASSES STDOUT
		Call objShell.Run(strExe & " " & strExeArguments, 1, True)
	
		'FILE EXISTS
		If (objFSO.FileExists("C:\" & strTxtFile)) Then
			Set objFile = objFSO.OpenTextFile("C:\" & strTxtFile, FOR_READ, True)
			strFile = objFile.ReadAll
			objFile.Close
		End If
	
		'GET MEMORY
		strMemory = GetMemory(strFile)
		'GET MOTHERBOARD, PROCESSORS, DISKS + MEMORY
		If (Len(strMemory) > 0) Then
			strResult = GetMotherBoard(strFile) & "<br>" & GetProcessors(strFile) & "<br>" & strMemory & "<br>" & GetDisks & "<br>" & GetVirtual & "<br>"
		Else
			'SAME - MEMORY
			strResult = GetMotherBoard(strFile) & "<br>" & GetProcessors(strFile) & "<br>" & GetDisks & "<br>" & GetVirtual & "<br>"
		End If
		
		'TXT FILE EXISTS
		If (objFSO.FileExists("C:\" & strTxtFile)) Then
			'DELETE
			Call objFSO.DeleteFile("C:\" & strTxtFile, True)
		End If					

		'SET DATA		
		Identifiers(SAVID_ATTRIBUTE) = OPT_HARDWARE		
		Datas(SAVDATA_VALUE) = strResult
		'ADD DATA
		Call AddSavedData(strData, Identifiers, Datas)				
			
		'SAVE DATA
		Call OpenFileWrite(strCurPath & "\" & strSavedFile, strData)
	End If					
End Function


'GETLOADATTRIBUTE
Function GetLoadAttribute(strAttribute)
	'LOAD DATA
	If (OpenFileRead(strCurPath & "\" & strSavedFile, strData)) Then
		'SET IDENTIFIERS
		Dim Identifiers(0)
		Dim Datas(1)
		Identifiers(SAVID_ATTRIBUTE) = strAttribute
		
		'FIND DATA
		If (FindSavedData(strData, Identifiers, Datas)) Then
			'DATA NOT EXPIRED
			If (DateDiff("h", Datas(SAVDATA_TIMEDATE), Now) < 12) Then		
				'RESULT
				GetLoadAttribute = Datas(SAVDATA_VALUE)
			End If
		End If
	Else
		'ERROR
		Call WScript.echo("Failed to open file " & Chr(34) & strSavedFile & Chr(34) & ".")
		
		'QUIT
		WScript.Quit			
	End If			
End Function








'ARGUMENTS
' 0 : OPTION = [NAME, OS, NETWORK]

'CHECK ARGUMENTS
Call CheckArguments(1)

'CHECK OUT OF PROCESS
Call CheckOutOfProcess

'GET ARGUMENTS
strOption = WScript.Arguments(0)




'NAME (ACTUAL NAME OF THE COMPUTER, FQDN)
If (StrComp(strOption, OPT_NAME, 1) = 0) Then
	'RESULT
	strResult = GetName
	
'OS
ElseIf (StrComp(strOption, OPT_OS, 1) = 0) Then
	
	'CHECK NATIVE BITNESS
	Call CheckNativeBitness

	'GET OPERATING SYSTEMS
	Set colOperatingSystems = objWMI.ExecQuery("Select * from Win32_OperatingSystem")
	'GET ENV VARS
	Set objProcEnv = objShell.Environment("Process")
	
	'EACH OS
	For Each objOperatingSystem in colOperatingSystems
		'RESULT
		strResult = objOperatingSystem.Caption & " (" & objProcEnv("PROCESSOR_ARCHITECTURE") & ")"
	Next
	
'NETWORK
ElseIf (StrComp(strOption, OPT_NETWORK, 1) = 0) Then
	'GET NIC INFO, ROUTING TABLE
	strResult = GetNicInfo & "<br>" & GetRouteTable & "<br"
	
'SOFTWARE
ElseIf (StrComp(strOption, OPT_SOFTWARE, 1) = 0) Then
	'GET ATTRIBUTE FROM SAVED FILE
	strResult = GetLoadAttribute(strOption)			
	
'HARDWARE
ElseIf (StrComp(strOption, OPT_HARDWARE, 1) = 0) Then
	'GET ATTRIBUTE FROM SAVED FILE
	strResult = GetLoadAttribute(strOption)			

'OOP
ElseIf (StrComp(strOption, OPT_OOP, 1) = 0) Then
	'LAUNCH OUT OF PROCESS
	If (LaunchOutOfProcess) Then strResult = Now
	
'STYLE
ElseIf (StrComp(strOption, OPT_STYLE, 1) = 0) Then
	'RETURN STYLE
	strResult = TABLE_STYLE
	
End If




'ZABBIX RESULT
Call WScript.Echo(strResult)