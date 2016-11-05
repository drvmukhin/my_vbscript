Dim strDirectoryWork, D0, objFSO, objShell, objEnvar, CurrentDate, CurrentTime, nDebug, nInfo, objDebug, ShowLog
Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
Set objEnvar = WScript.CreateObject("WScript.Shell")
Set objShell = WScript.CreateObject("WScript.Shell")
Set objApp = CreateObject("Shell.application")
Dim strPID, strCmdLine, strParentPID, strAppName

Main()
Set objFSO = Nothing
set objEnvar = Nothing
Set objShell = Nothing

Sub Main()
    strCmdLine = "test.vbs"
	strAppName = "wscript.exe"
	If not GetWinAppPID(strPID, strParrentID, strCmdLine, strAppName, nDebug) Then 
	    MsgBox "Script is not running. Nothing to stop"
		Exit Sub
	Else
	    If KillWinAppPID(strPID, "None", strAppName, nInfo) then 
		     MsgBox "Process was stopped" & chr(13) &_
			        "App: " & strAppName & chr(13) &_
					"PID: " & strPID & chr(13) &_
					"CmdLine: "& strCmdLine		
		Else 
		     MsgBox "Something when wrong!" & chr(13) &_
			        "I was not able to stop App: " & strAppName & chr(13) &_
					"PID: " & strPID & chr(13) &_
					"CmdLine: "& strCmdLine
		End If 
	End if 
End Sub

'----------------------------------------------------------------
'   Function GetWinAppPID(strPID) Returns focus to the parent Window/Form
'----------------------------------------------------------------
Function GetWinAppPID(ByRef strPID, ByRef strParentPID, ByRef strCommandLine, strAppName, nDebug)
Dim objWMI, colItems
Dim process
Dim strUser, pUser, pDomain, wql
	strUser = GetScreenUserSYS()
	GetWinAppPID = False
	Do 
		On Error Resume Next
		Set objWMI = GetObject("winmgmts:\\127.0.0.1\root\cimv2")
		If Err.Number <> 0 Then 
				'Call TrDebug ("GetMyPID ERROR: CAN'T CONNECT TO WMI PROCESS OF THE SERVER","",objDebug, MAX_LEN, 1, nDebug)
				On error Goto 0 
				Exit Do
		End If 
		wql = "SELECT * FROM Win32_Process WHERE Name = '" & strAppName & "' OR Name = '" & strAppName & " *32'"
		On Error Resume Next
		Set colItems = objWMI.ExecQuery(wql)
		If Err.Number <> 0 Then
				'Call TrDebug ("GetMyPID ERROR: CAN'T READ QUERY FROM WMI PROCESS OF THE SERVER","",objDebug, MAX_LEN, 1, nDebug)
				On error Goto 0 
				Set colItems = Nothing
				Exit Do
		End If 
		On error Goto 0 
		For Each process In colItems
			process.GetOwner  pUser, pDomain 
			'Call TrDebug ("GetWinAppPID: Process Name (PID): " & process.Name & " (" & process.ProcessId & ")", "",objDebug, MAX_LEN, 1, nDebug)
			'Call TrDebug ("GetWinAppPID: Owner: " & process.CSName & "/" & pUser, "",objDebug, MAX_LEN, 1, nDebug) 
			'Call TrDebug ("GetWinAppPID: CMD: " & process.CommandLine, "",objDebug, MAX_LEN, 1, nDebug) 
			'Call TrDebug ("GetWinAppPID: ParentPID:" &  Process.ParentProcessId, "",objDebug, MAX_LEN, 1, nDebug) 			
			Select Case Lcase(strCommandLine)
			    Case "null", "none", ""
					If pUser = strUser then 
						strPID = process.ProcessId
						strParentPID = Process.ParentProcessId
						'Call TrDebug ("GetWinAppPID: Process is already running. Desktop user owns the process: " & strPID , "",objDebug, MAX_LEN, 1, nDebug)
						GetWinAppPID = True
						Exit For
					End If
			    Case Else
					If pUser = strUser and InStr(process.CommandLine,strCommandLine) then 
					    strCommandLine = process.CommandLine
						strPID = process.ProcessId
						strParentPID = Process.ParentProcessId
						'Call TrDebug ("GetWinAppPID: Process is already running. Desktop user owns the process: " & strPID, "",objDebug, MAX_LEN, 1, nDebug)
						GetWinAppPID = True
						Exit For
					End If
			End Select
		Next
		Set colItems = Nothing
		Exit Do
	Loop
	Set objWMI = Nothing
End Function
'----------------------------------------------------------------
'   Function KillWinAppPID(strPID) Returns focus to the parent Window/Form
'----------------------------------------------------------------
Function KillWinAppPID(ByRef strPID, strCommandLine, strAppName, nDebug)
Dim objWMI, colItems
Dim process
Dim strUser, pUser, pDomain, wql
	strUser = GetScreenUserSYS()
	KillWinAppPID = False
	Do 
		On Error Resume Next
		Set objWMI = GetObject("winmgmts:\\127.0.0.1\root\cimv2")
		If Err.Number <> 0 Then 
				'Call TrDebug ("GetMyPID ERROR: CAN'T CONNECT TO WMI PROCESS OF THE SERVER","",objDebug, MAX_LEN, 1, nDebug)
				On error Goto 0 
				Exit Do
		End If 
		wql = "SELECT * FROM Win32_Process WHERE Name = '" & strAppName & "' OR Name = '" & strAppName & " *32'"
		On Error Resume Next
		Set colItems = objWMI.ExecQuery(wql)
		If Err.Number <> 0 Then
				'Call TrDebug ("GetMyPID ERROR: CAN'T READ QUERY FROM WMI PROCESS OF THE SERVER","",objDebug, MAX_LEN, 1, nDebug)
				On error Goto 0 
				Set colItems = Nothing
				Exit Do
		End If 
		On error Goto 0 
		' Select task by PID, UserName, AppName, Command Line
		Do 
			If IsNumeric(strPID) Then nMode = 1 : Exit Do End If
			If strCommandLine <> "" and Lcase(strCommandLine) <> "null" and Lcase(strCommandLine) <> "none" Then nMode = 2 : Exit Do End If
			nMode = 0
			Exit Do
		Loop
		For Each process In colItems
			process.GetOwner  pUser, pDomain 
			'Call TrDebug ("KillWinAppPID: Process Name (PID): " & process.Name & " (" & process.ProcessId & ")", "",objDebug, MAX_LEN, 1, nDebug)
			' 'Call TrDebug ("KillWinAppPID: Owner: " & process.CSName & "/" & pUser, "",objDebug, MAX_LEN, 1, nDebug) 
			Select Case nMode
			    Case 0
					If pUser = strUser then 
						'Call TrDebug ("KillWinAppPID (0): Terminating the Process: Desktop user owns the process: " & strPID , "",objDebug, MAX_LEN, 1, nDebug)
						process.Terminate()
						KillWinAppPID = True
						Exit For
					End If
			    Case 1
					If (InStr(strPID,process.ProcessId) and InStr(process.ProcessId,StrPID)) then 
						'Call TrDebug ("KillWinAppPID (1): Terminating the Process: Desktop user owns the process: " & strPID , "",objDebug, MAX_LEN, 1, nDebug)
						process.Terminate()
						KillWinAppPID = True
						Exit For
					End If				
			    Case 2
					If pUser = strUser and InStr(process.CommandLine,strCommandLine) then 
						'Call TrDebug ("KillWinAppPID (2): Terminating the Process: Desktop user owns the process: " & strPID, "",objDebug, MAX_LEN, 1, nDebug)
						process.Terminate()
						KillWinAppPID = True
						Exit For
					End If
			End Select
		Next
		Set colItems = Nothing
		Exit Do
	Loop
	Set objWMI = Nothing
End Function
'----------------------------------------------------------------------------------
'    Function GetScreenUserSYS
'----------------------------------------------------------------------------------
Function GetScreenUserSYS()
Dim vLine
Dim strScreenUser, strUserProfile
Dim nCount
Dim objEnvar
	Set objEnvar = WScript.CreateObject("WScript.Shell")	
	strUserProfile = objEnvar.ExpandEnvironmentStrings("%USERPROFILE%")
	vLine = Split(strUserProfile,"\")
	nCount = Ubound(vLine)
	strScreenUser = vLine(nCount)
	If InStr(strScreenUser,".") <> 0 then strScreenUser = Split(strScreenUser,".")(0) End If
	set objEnvar = Nothing
	GetScreenUserSYS = strScreenUser
End Function