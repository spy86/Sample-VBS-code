'2012.06.25
'fixed return of logexeoutput

'2012.06.17



Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")
strCurPath = objFSO.GetParentFolderName(WScript.ScriptFullName)



'FILE OPERATIONS
FOR_READ = 1
FOR_WRITE = 2
FOR_APPEND = 8




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


'LOG EXE OUTPUT
Function LogExeOutput(strExe, ByRef strStd, objLogFile)
On Error Resume Next

	'RESULT
	LogExeOutput = False
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
		
		'WAIT FOR IT TO FINISH OR TIMEOUT
		Do Until ((objExec.Status = 1) And (objExec.StdOut.AtEndOfStream))
			'STDOUT
			If (Not objExec.StdOut.AtEndOfStream) Then
				strRead = objExec.StdOut.ReadLine
				If (Len(strRead) > 0) Then 					
					strStd = strStd & Now & ": " & strRead & vbCrLf
					'LOG OUTPUT
					objLogFile.Write(vbCrLf & Now & ": " & strRead)
					'## STATUS
					Call WScript.Echo(strRead)				
				End If
			End If
		Loop
			
		'RESULT
		LogExeOutput = True
	End If
End Function





'ARGUMENTS
' 0 : SCRIPT TO RUN
' .. : SCRIPT ARGUMENTS
' LAST : LOG FILE PATH

'CHECK ARGUMENTS
Call CheckArguments(2)

'GET ARGUMENTS
strScript = WScript.Arguments(0)
strLogFile = WScript.Arguments(WScript.Arguments.Count - 1)
For x = 1 To WScript.Arguments.Count - 2
	strArguments = strArguments & " " & Chr(34) & WScript.Arguments(x) & Chr(34)		
Next




'CHECK LOG SIZE
If (objFSO.FileExists(strLogFile)) Then
	Set objFile = objFSO.GetFile(strLogFile)
	
	'FILE SIZE ABOVE 1MB, DELETE
	If (objFile.Size > 1048576) Then Call objFile.Delete(True)
End If

'OPEN LOG FILE
Set objLog = objFSO.OpenTextFile(strLogFile, FOR_APPEND, True)

'RUN THE SCRIPT
strExe = Chr(34) & "cscript.exe" & Chr(34) & " //nologo"
strExeArguments = Chr(34) & strScript & Chr(34) & " " & strArguments

'## STATUS
objLog.Write(vbCrLf & vbCrLf & Now & ": Executing script: " & strExeArguments)

'RUN EXE WITH NO TIMEOUT
If (LogExeOutput(strExe & " " & strExeArguments, strStd, objLog) = True) Then	
	'LOG SUCCESS
	objLog.Write(vbCrLf & "Finished executing " & strExe & " " & strExeArguments & ".")
Else
	'LOG ERROR
	objLog.Write(vbCrLf & "Error executing " & strExe & " " & strExeArguments & ".")
End If




'ZABBIX RESULT
Call WScript.Echo(strResult)