Dim strDirectoryWork, D0, objFSO, objShell, objEnvar, CurrentDate, CurrentTime, nDebug, nInfo, objDebug, strDebugFile
Const DEBUG_FILE = "vbs_kill_process"
Const ForAppending = 8
Const ForWriting = 2
Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
Set objEnvar = WScript.CreateObject("WScript.Shell")
Set objShell = WScript.CreateObject("WScript.Shell")
Set objApp = CreateObject("Shell.application")
Dim strPID, strCmdLine, strParentPID, strAppName, KillAll
Dim ShowLog, bVerbose,bMultipleInstanceAllowed
bMultipleInstanceAllowed = False
ShowLog = False
bVerbose = True
nInfo = 0 
nDebug = 0
'
'  Run Main sub
Main()
Set objFSO = Nothing
set objEnvar = Nothing
Set objShell = Nothing
'
'  Define Main() sub
Sub Main()
'    strCmdLine = "Kids_schedule_WIN_"
'	strAppName = "wscript.exe"
	'
	'   SET SCRIPT WORK FOLDER
	strDirectoryWork =  objFSO.GetParentFolderName(Wscript.ScriptFullName)
	strDebugFile = strDirectoryWork & "\" & DEBUG_FILE & ".log"
	'
	'  Set default values for main variables
    strCmdLine = ""
	strAppName = ""
    strPID = ""
	UtilsFolder = "C:\UnixUtils"
	KillAll = False
	On Error Resume Next
	For i = 0 to WScript.Arguments.Count - 1
		Select Case WScript.Arguments(i)
			Case "-P", "-p"
				strPID = wscript.Arguments(i+1)
			Case "-C", "-c"
				strCmdLine = wscript.Arguments(i+1)
			Case "-N", "-n"
				strAppName = wscript.Arguments(i+1)
			Case "-U", "-u"
			    UtilsFolder = wscript.Arguments(i+1)
			Case "-A", "-a"
				KillAll = True
		End Select
		If Err.Number > 0 Then 
			MsgBox "ERROR: Wrong number of arguments" & chr(13) &_
			"ARG1: -P <PID>" & chr(13) &_
			"ARG2: -C <process description pattern>" & chr(13) &_
			"ARG3: -N <process/application name>" & chr(13) &_
			"ARG4: -U <path to tail.exe utility>" & chr(13) &_			
			"ARG5: -A Kill all process with name under -A option"
			Exit Sub
		End If			
	Next
	On Error Goto 0 
	If 	strAppName = "" Then
			MsgBox "ERROR: Name of the process is missed" & chr(13) &_
			"ARG1: -P <PID>" & chr(13) &_
			"ARG2: -C <process description pattern>" & chr(13) &_
			"ARG3: -N <process/application name>" & chr(13) &_
			"ARG4: -U <path to tail.exe utility>" & chr(13) &_			
			"ARG5: -A Kill all process with name under -A option"
		Exit Sub
	End If
	'
	'  Open log file 
	If Not OpenLogSession(objDebug, strDebugFile, UtilsFolder, bMultipleInstanceAllowed, ShowLog, bVerbose) Then Exit Sub
	'
	'
	wscript.sleep 200
    '
	'  Start main cycle
	Select case KillAll
		Case False 
			If not GetWinAppAllPID(strPID, strParrentID, strCmdLine, strAppName, vApps,nDebug) Then 
				MsgBox "Process " & strAppName & " was not found. Nothing to stop"
				Exit Sub
			Else
				'strLine = "Following scripts are running:" & chr(13)
				strLine = ""
				For each strItem in vApps
				   If strItem = "" Then exit For
				   strLine = strLine & Replace(strItem,strCmdLine,"") & chr(13)
				Next
				strPID = "None"
				strPID = InputBox("Enter process ID:" & chr(13) & strLine,"Terminate process ")
				If strPID = "" Then Exit Sub
				If Not IsNumeric(strPID) Then Exit Sub
				If KillWinAppPID(strPID, "None", strAppName, nInfo) > 0 then 
					 MsgBox "Process was stopped" & chr(13) &_
							"App: " & strAppName & chr(13) &_
							"PID: " & strPID & chr(13)
				Else 
					 MsgBox "Something when wrong!" & chr(13) &_
							"I was not able to stop App: " & strAppName & chr(13) &_
							"PID: " & strPID & chr(13)
				End If
			End If		
		Case True
			nResult = KillWinAppPID("", "None", strAppName, nInfo)
			If nResult = 0 Then 
				MsgBox "Process " & strAppName & " was not found. Nothing to stop"
				Exit Sub
			Else
				MsgBox nResult & " instances of " & strAppName & " process has been stopped"
				Exit Sub
			End If					
	End Select
End Sub
'----------------------------------------------------------------
'   Function GetWinAppPID(strPID) Returns focus to the parent Window/Form
'----------------------------------------------------------------
Function GetWinAppPID(ByRef strPID, ByRef strParentPID, strCommandLine, strAppName, nDebug)
Dim objWMI, colItems
Dim process
Dim strUser, pUser, pDomain, wql
	strUser = GetScreenUserSYS()
	GetWinAppPID = False
	Do 
		On Error Resume Next
		Set objWMI = GetObject("winmgmts:\\127.0.0.1\root\cimv2")
		If Err.Number <> 0 Then 
				Call TrDebug ("GetMyPID ERROR: CAN'T CONNECT TO WMI PROCESS OF THE SERVER","",objDebug, MAX_LEN, 1, nDebug)
				On error Goto 0 
				Exit Do
		End If 
		wql = "SELECT * FROM Win32_Process WHERE Name = '" & strAppName & "' OR Name = '" & strAppName & " *32'"
		On Error Resume Next
		Set colItems = objWMI.ExecQuery(wql)
		If Err.Number <> 0 Then
				Call TrDebug ("GetMyPID ERROR: CAN'T READ QUERY FROM WMI PROCESS OF THE SERVER","",objDebug, MAX_LEN, 1, nDebug)
				On error Goto 0 
				Set colItems = Nothing
				Exit Do
		End If 
		On error Goto 0 
		For Each process In colItems
			process.GetOwner  pUser, pDomain 
			Call TrDebug ("GetWinAppPID: Process Name (PID): " & process.Name & " (" & process.ProcessId & ")", "",objDebug, MAX_LEN, 1, nDebug)
			Call TrDebug ("GetWinAppPID: Owner: " & process.CSName & "/" & pUser, "",objDebug, MAX_LEN, 1, nDebug) 
			Call TrDebug ("GetWinAppPID: CMD: " & process.CommandLine, "",objDebug, MAX_LEN, 1, nDebug) 
			Call TrDebug ("GetWinAppPID: ParentPID:" &  Process.ParentProcessId, "",objDebug, MAX_LEN, 1, nDebug) 			
			Select Case Lcase(strCommandLine)
			    Case "null", "none", ""
					If pUser = strUser then 
						strPID = process.ProcessId
						strParentPID = Process.ParentProcessId
						Call TrDebug ("GetWinAppPID: Process is already running. Desktop user owns the process: " & strPID , "",objDebug, MAX_LEN, 1, nDebug)
						GetWinAppPID = True
						Exit For
					End If
			    Case Else
					If pUser = strUser and InStr(process.CommandLine,strCommandLine) then 
						strPID = process.ProcessId
						strParentPID = Process.ParentProcessId
						Call TrDebug ("GetWinAppPID: Process is already running. Desktop user owns the process: " & strPID, "",objDebug, MAX_LEN, 1, nDebug)
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
'   Function GetWinAppPID(strPID) Returns focus to the parent Window/Form
'----------------------------------------------------------------
Function GetWinAppAllPID(ByRef strPID, ByRef strParentPID, ByRef strCommandLine, strAppName, byRef vApp, nDebug)
Dim objWMI, colItems
Dim process
Dim strUser, pUser, pDomain, wql, objFSO, MyScriptName
	Set objFSO = wscript.CreateObject("Scripting.FileSystemObject")
	MyScriptName = objFSO.GetFile(wscript.ScriptFullName).Name
	Call TrDebug ("GetWinAppAll : ScriptName Runnig: ",MyScriptName,objDebug, MAX_LEN, 1, nDebug)
	Set objFSO = Nothing

    Redim vApp(1)
	strUser = GetScreenUserSYS()
	GetWinAppAllPID = False
	Do 
		On Error Resume Next
		Set objWMI = GetObject("winmgmts:\\127.0.0.1\root\cimv2")
		If Err.Number <> 0 Then 
				Call TrDebug ("GetWinAppAllPID ERROR: CAN'T CONNECT TO WMI PROCESS OF THE SERVER","",objDebug, MAX_LEN, 1, nDebug)
				On error Goto 0 
				Exit Do
		End If 
		wql = "SELECT * FROM Win32_Process WHERE Name = '" & strAppName & "' OR Name = '" & strAppName & " *32'"
		On Error Resume Next
		Set colItems = objWMI.ExecQuery(wql)
		If Err.Number <> 0 Then
				Call TrDebug ("GetWinAppAllPID ERROR: CAN'T READ QUERY FROM WMI PROCESS OF THE SERVER","",objDebug, MAX_LEN, 1, nDebug)
				On error Goto 0 
				Set colItems = Nothing
				Exit Do
		End If 
		On error Goto 0 
		nCount = 0
		On Error Resume Next
		Err.Clear
		For Each process In colItems
			process.GetOwner  pUser, pDomain 
			If Err.Number > 0 Then exit for
			Call TrDebug ("GetWinAppPID: Process Name (PID): " & process.Name & " (" & process.ProcessId & ")", "",objDebug, MAX_LEN, 1, nDebug)
			Call TrDebug ("GetWinAppPID: Owner: " & process.CSName & "/" & pUser, "",objDebug, MAX_LEN, 1, nDebug) 
			Call TrDebug ("GetWinAppPID: CMD: " & process.CommandLine, "",objDebug, MAX_LEN, 1, nDebug) 
			Call TrDebug ("GetWinAppPID: ParentPID:" &  Process.ParentProcessId, "",objDebug, MAX_LEN, 1, nDebug) 			
			Select Case Lcase(strCommandLine)
			    Case "null", "none", ""
					If pUser = strUser and InStr(process.CommandLine,MyScriptName) = 0 then 
						strPID = process.ProcessId
						nLine = Ubound(Split(process.CommandLine,"\"))
						Redim Preserve vApp(nCount + 1)
						vApp(nCount) = strAppName & " PID: " & strPID 
						nCount = nCount + 1
						strParentPID = Process.ParentProcessId
						Call TrDebug ("GetWinAppPID: Process is already running. Desktop user owns the process: " & strPID, "",objDebug, MAX_LEN, 1, nDebug)
						GetWinAppAllPID = True
					End If
			    Case Else
					If pUser = strUser and InStr(process.CommandLine,strCommandLine) and InStr(process.CommandLine,MyScriptName) = 0 then 
					    'strCommandLine = process.CommandLine
						strPID = process.ProcessId
						nLine = Ubound(Split(process.CommandLine,"\"))
						Redim Preserve vApp(nCount + 1)
						vApp(nCount) = Split(process.CommandLine,"\")(nLine) & " : " & strPID 
						nCount = nCount + 1
						strParentPID = Process.ParentProcessId
						Call TrDebug ("GetWinAppPID: Process is already running. Desktop user owns the process: " & strPID, "",objDebug, MAX_LEN, 1, nDebug)
						GetWinAppAllPID = True
					End If
			End Select
		Next
		On Error Goto 0 
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
Dim process, MyScriptName
Dim strUser, pUser, pDomain, wql
Dim objFSO
	Set objFSO = wscript.CreateObject("Scripting.FileSystemObject")
	MyScriptName = objFSO.GetFile(wscript.ScriptFullName).Name
	Set objFSO = Nothing
	strUser = GetScreenUserSYS()
	KillWinAppPID = 0
	Do 
		On Error Resume Next
		Set objWMI = GetObject("winmgmts:\\127.0.0.1\root\cimv2")
		If Err.Number <> 0 Then 
				Call TrDebug ("GetMyPID ERROR: CAN'T CONNECT TO WMI PROCESS OF THE SERVER","",objDebug, MAX_LEN, 1, nDebug)
				On error Goto 0 
				Exit Do
		End If 
		wql = "SELECT * FROM Win32_Process WHERE Name = '" & strAppName & "' OR Name = '" & strAppName & " *32'"
		On Error Resume Next
		Set colItems = objWMI.ExecQuery(wql)
		If Err.Number <> 0 Then
				Call TrDebug ("GetMyPID ERROR: CAN'T READ QUERY FROM WMI PROCESS OF THE SERVER","",objDebug, MAX_LEN, 1, nDebug)
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
		On Error Resume Next
		Err.clear
		For Each process In colItems
			process.GetOwner  pUser, pDomain
			If Err.Number > 0 Then 
				Call TrDebug ("KillWinAppPID: Process Name (PID): " & process.Name & " (" & process.ProcessId & ")", "",objDebug, MAX_LEN, 1, 1)
				Call TrDebug ("KillWinAppPID: Owner: " & process.CSName & "/" & pUser, "",objDebug, MAX_LEN, 1, 1) 
				Exit For 
			End If 
			Call TrDebug ("KillWinAppPID: Process Name (PID): " & process.Name & " (" & process.ProcessId & ")", "",objDebug, MAX_LEN, 1, nDebug)
			Call TrDebug ("KillWinAppPID: Owner: " & process.CSName & "/" & pUser, "",objDebug, MAX_LEN, 1, nDebug) 
			Select Case nMode
			    Case 0
					If pUser = strUser and InStr(process.CommandLine,MyScriptName) = 0 then 
						Call TrDebug ("KillWinAppPID (0): Terminating the Process: Desktop user owns the process: " & strPID , "",objDebug, MAX_LEN, 1, nDebug)
						process.Terminate()
						KillWinAppPID = KillWinAppPID + 1
						'Exit For
					End If
			    Case 1
					If InStr(strPID,process.ProcessId) and InStr(process.ProcessId,strPID) then 
						Call TrDebug ("KillWinAppPID (1): Terminating the Process: Desktop user owns the process: " & strPID , "",objDebug, MAX_LEN, 1, nDebug)
						process.Terminate()
						KillWinAppPID = KillWinAppPID + 1
						Exit For
					End If				
			    Case 2
					If pUser = strUser and InStr(process.CommandLine,strCommandLine) then 
						Call TrDebug ("KillWinAppPID (2): Terminating the Process: Desktop user owns the process: " & strPID, "",objDebug, MAX_LEN, 1, nDebug)
						process.Terminate()
						KillWinAppPID =  KillWinAppPID + 1
						Exit For
					End If
			End Select
		Next
		On Error goto 0
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
'------------------------------------------------------------------
'   Function OpenLogSession(ByRef objDebug, ByRef strDebugFile, bMultipleInstanceAllowed, bShowLog, bVerbose)
'------------------------------------------------------------------
Function OpenLogSession(ByRef objDebug, ByRef strDebugFile, UtilsFolder, bMultipleInstanceAllowed, bShowLog, bVerbose)
Dim nIndex, strErrorLog,strNewInstanceLog, objError, INSTANCE_LOG, nError, DEBUG_FILE
Dim my_objShell
	Set my_objShell = CreateObject("WScript.Shell")
    nError = 0
	nLenEnd = InStrRev(strDebugFile,"\")
	strErrorLog = Left(strDebugFile,nLenEnd) & Split(Right(strDebugFile,Len(strDebugFile) - nLenEnd),".")(0) & "_Error.log"
	If InStr(Right(strDebugFile,Len(strDebugFile) - nLenEnd),".") Then 
	   INSTANCE_LOG = Left(strDebugFile,nLenEnd) & Split(Right(strDebugFile,Len(strDebugFile) - nLenEnd),".")(0) & "-inst-<index>." & Split(Right(strDebugFile,Len(strDebugFile) - nLenEnd),".")(1)
	Else 
	   INSTANCE_LOG = Left(strDebugFile,nLenEnd) & Right(strDebugFile,Len(strDebugFile) - nLenEnd) & "-inst-<index>"
	End If   
	nIndex = 0
    Set objError = objFSO.OpenTextFile(strErrorLog,ForWriting,True)
	Do
		On Error Resume Next
		Err.Clear
		Set objDebug = objFSO.OpenTextFile(strDebugFile,ForWriting,True)
		Select Case Err.Number
			Case 0
				Exit Do
			Case 70
				nIndex = nIndex + 1
				Select Case nIndex
                   	Case 1
                       	If bMultipleInstanceAllowed Then 
					        strDebugFile = Replace(INSTANCE_LOG,"<index>",nIndex)
						Else 
							If bVerbose Then MsgBox "Another instance of the script is allready running. Exit now"
							objError.WriteLine Date() & " " & Time() & ": ERROR:  Another instances of the script is allready running. Exit now"
                            nError = 1
							Exit Do
						End if 
                    Case 2					
                           strDebugFile = Replace(INSTANCE_LOG,"<index>",nIndex)
					Case 3
					    If bVerbose Then MsgBox "Two other instances of the script are allready running. Exit now"
						objError.WriteLine Date() & " " & Time() & ": ERROR:  Two other instances of the script are allready running. Exit now"
						nError = 3
						Exit Do
				End Select
				wscript.sleep 500
			Case Else 
			    If bVerbose Then MsgBox "Can't open log file" & chr(13) & "Error: #" & Err.Number & ": " &  Err.Description 
			    objError.WriteLine Date() & " " & Time() & "ERROR: Can't open log file" 
				objError.WriteLine Date() & " " & Time() & "Error: #" & Err.Number & ": " &  Err.Description 
			    nError = 1000
				Exit Do
		End Select
	Loop
	On Error goto 0
    If nError > 0 Then 
	   OpenLogSession = False
	   If IsObject(objError) Then objError.Close : End If
	   If bShowLog Then my_objShell.Run "notepad.exe " & strErrorLog,1
	   set my_objShell = Nothing
	   Exit Function
	End If 
	' Open tail -f to stream log mesages into desktop
	If bShowLog and objFSO.FileExists(UtilsFolder & "\tail.exe") Then 
		strLaunch = UtilsFolder & "\tail.exe -n 10 -f " & strDebugFile
		DEBUG_FILE = Split(Right(strDebugFile,Len(strDebugFile) - nLenEnd),".")(0) 
		If Not GetWinAppPID(strPID, strParrentID, DEBUG_FILE, "tail.exe", nDebug) Then 
			my_objShell.run (strLaunch)
		Else
			Call FocusToParentWindow(strPID)
		End If
	End If
	' Exit Function
	Set my_objShell = Nothing
    OpenLogSession = True
    If IsObject(objError) Then objError.Close : End If
End Function 
'---------------------------------------------------------------
'   Function FocusToParentWindow(strPID) Returns focus to the parent Window/Form
'----------------------------------------------------------------
Function FocusToParentWindow(strPID)
Dim objShell
Call TrDebug ("FocusToParentWindow: RESTORE IE WINDOW:", "PID: " & strPID, objDebug, MAX_LEN, 1, 1) 
Const IE_PAUSE = 70
	Set objShell = WScript.CreateObject("WScript.Shell")
	wscript.sleep IE_PAUSE  
	objShell.SendKeys "%"
	wscript.sleep IE_PAUSE
	objShell.AppActivate strPID			
	wscript.sleep IE_PAUSE  
	objShell.SendKeys "% "
	wscript.sleep IE_PAUSE  
	objShell.SendKeys "r"
	Set objShell = Nothing
End Function
' ----------------------------------------------------------------------------------------------
'   Function  TrDebug (strTitle, strString, objDebug)
'   nFormat: 
'	0 - As is
'	1 - Strach
'	2 - Center
' ----------------------------------------------------------------------------------------------
Function  TrDebug (strTitle, strString, objDebug, nChar, vFormat, nDebug)
Dim strLine
Dim nFormat
	If IsArray(vFormat) Then 
	    nFormat = vFormat(0)
		Set_A_Date = vFormat(1)
	Else 
	    nFormat = vFormat
		Set_A_Date = True
    End If	
	strLine = ""
	If nDebug <> 1 Then Exit Function End If
	If IsObject(objDebug) Then 
		Select Case nFormat
			Case 0
				If Set_A_Date Then strLine = GetMyDate() & " " & FormatDateTime(Time(), 3) Else strLine = ""
				strLine = strLine & ":  " & strTitle
				strLine = strLIne & strString
				objDebug.WriteLine strLine
				
			Case 1
				If Set_A_Date Then strLine = GetMyDate() & " " & FormatDateTime(Time(), 3) Else strLine = ""
				strLine = strLine & ":  " & strTitle
				If nChar - Len(strLine) - Len(strString) > 0 Then 
					strLine = strLine & Space(nChar - Len(strLine) - Len(strString)) & strString
				Else 
					strLine = strLine & " " & strString
				End If
				objDebug.WriteLine strLine
			Case 2
				If Set_A_Date Then strLine = GetMyDate() & " " & FormatDateTime(Time(), 3) & ":  " Else strLine = ""
				
				If nChar - Len(strLine & strTitle & strString) > 0 Then 
						strLine = strLine & Space(Int((nChar - 1 - Len(strLine & strTitle & strString))/2)) & strTitle & " " & strString			
				Else 
						strLine = strLine & strTitle & " " & strString	
				End If
				objDebug.WriteLine strLine
			Case 3
				If Set_A_Date Then strLine = GetMyDate() & " " & FormatDateTime(Time(), 3) & ":  " Else strLine = ""
				For i = 0 to nChar - Len(strLine)
					strLIne = strLIne & "-"
				Next
				objDebug.WriteLine strLine
				If Set_A_Date Then strLine = GetMyDate() & " " & FormatDateTime(Time(), 3) & ":  " Else strLine = ""
				If nChar - 1 - Len(strLine & strTitle & strString) > 0 Then 
						strLine = strLine & Space(Int((nChar - 1 - Len(strLine & strTitle & strString))/2)) & strTitle & " " & strString			
				Else 
						strLine = strLine & strTitle & " " & strString	
				End If
				objDebug.WriteLine strLine
				If Set_A_Date Then strLine = GetMyDate() & " " & FormatDateTime(Time(), 3) & ":  " Else strLine = ""
				For i = 0 to nChar - Len(strLine)
					strLine = strLine & "-"
				Next
				objDebug.WriteLine strLine
		End Select
	End If
End Function
'-----------------------------------------------------------------
'     Function GetMyDate()
'-----------------------------------------------------------------
Function GetMyDate()
	GetMyDate = Month(Date()) & "/" & Day(Date()) & "/" & Year(Date()) 
End Function