#$language = "VBScript"
#$interface = "1.0"
'----------------------------------------------------------------------------------
'	KIDSSHIELD SCRIPT. SSH TO JUNIPER SSG FW. 
'----------------------------------------------------------------------------------
	Dim strDirectoryLCL
strDirectoryLCL = "C:\VBScript"
'----------------------------------------------------------------------------------
Const ForAppending = 8
Const ForWriting = 2
Const MAX_LEN = 130
Dim nResult
Dim strLine
Dim nOverwrite
Dim strMonthMaxFileName, strFileString, strSkip, strFileButton, strFileInventory, strFileSession
Dim strDirectory, strDirectoryUpdate, strDirectoryWork, strDirectoryVandyke, strLogFile
Dim strDeviceID, strAccountID
Dim nDebug, ShowDebug
Dim nIndex, nInd, nCount
Dim objDebug, objSession, objFSO, objEnvar, objButtonFile, objShell, objFolder
Dim vSession
Dim nStartHH, nEndHH, n, i, nRetries
Dim strUserProfile, vLine, strScreenUser
Dim nCommand, vCommand, vCmdSrx
Dim strAction, strTab
Const CRT_REG_INSTALL = "HKEY_LOCAL_MACHINE\SOFTWARE\VanDyke\SecureCRT\Install\Main Directory"
Const CRT_REG_SESSION = "HKEY_CURRENT_USER\Software\VanDyke\SecureCRT\Config Path"

strFileSession = "sessions.txt"
strDirectory = "\\MEDIA\_PublicFolder\KidsSchild\"
strDirectoryWork = "C:\KidsSchild\DVLP"
strDirectoryUpdate = "\\HIGS\Install_My\Tools_Networking\KidsSchild"
nDebug = 1
ShowDebug = True
strVersion = "None"
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")
' Set objEnvar = CreateObject("WScript.Shell")
Sub Main()
'----------------------------------------------------------------
'	Open log File
'----------------------------------------------------------------
			n = 5
			i = 0
			nRetries = 5
				Do While i < nRetries
					On Error Resume Next
					Err.Clear
					    strLogFile = strDirectoryLCL & "\" & "debug-terminal.log"
						Set objDebug = objFSO.OpenTextFile(strLogFile,ForAppending,True)
						Select Case Err.Number
							Case 0
								Exit Do
							Case 70
								i =  i + 1
								n = 3
								crt.sleep 100 * n
							Case Else 
								Exit Do		
						End Select
				Loop
				On Error goto 0
	'------------------------------------------------------------
	'   Start Monitoring Log file
	'------------------------------------------------------------
	strWinUtilsFolder = "C:\UnixUtils"
	If ShowDebug Then 
	    Set objBat = objFSO.OpenTextFile(Replace(strLogFile,"log","bat"),ForWriting,True)
		objBat.WriteLine strWinUtilsFolder & "\tail.exe -f " &  """" & strLogFile & """"
		objBat.close
		Set objBat = Nothing  
		objShell.run """" & Replace(strLogFile,"log","bat") & """",1
	End If	
				
'------------------------------------------------------------------
'	LOAD TELNET SESSION CONFIGURATION
'------------------------------------------------------------------
	If objFSO.FileExists(strDirectoryLCL & "\" & strFileSession) Then 
		nSession = GetFileLineCountSelect(strDirectoryLCL & "\" & strFileSession, vSession,"#","NULL","NULL", 0)
		vLine = Split(vSession(2), ",")
		strFolder = vLine(0)
		strSession = vLine(1)
		strHost = vLine(2)
		If IsObject(objDebug) and nDebug = 1 Then objDebug.WriteLine GetMyDate() & " " & FormatDateTime(Time(), 3) & ": Folder: " & strFolder & " Session: " & strSession & " Host: " & strHost End If
		strDirectory = vSession(1)
		If Right(strDirectory,1) = "\" Then 
			strDirectory = Left(strDirectory,Len(strDirectory) - 1)
			If IsObject(objDebug) and nDebug = 1 Then objDebug.WriteLine GetMyDate() & " " & FormatDateTime(Time(), 3) & ": Remote Server Folder: " & strDirectory End If
		End If
		If nSession > 3 Then	strDirectoryWork = vSession(3)		End If	' - Work directory scripts are installed to 
		If nSession > 4 Then    strDirectoryUpdate = vSession(4)	End If	' - Source directory to take updates from 
		If nSession > 5 Then    strVersion = vSession(5)			End If	' - Current version of the Package/Launcer
	End If
	'-----------------------------------------------------------------
	'  CHECK SECURECRT IS INSTALLED ON THE SYSTEM
	'-----------------------------------------------------------------
	Dim strCRT_InstallFolder, strCRT_SessionFolder
	On Error Resume Next
	    SecureCRT_Installed = True
		Err.Clear
		strCRT_InstallFolder = objShell.RegRead(CRT_REG_INSTALL)
		if Err.Number <> 0 Then 
            SecureCRT_Installed = False
			strCRT_InstallFolder = vSession(0)
		End If
		If Right(strCRT_InstallFolder,1) = "\" Then strCRT_InstallFolder = Left(strCRT_InstallFolder,Len(strCRT_InstallFolder)-1)
		Call TrDebug("Session Folder: " & strCRT_InstallFolder, "", objDebug, MAX_LEN, 1, 1)
		'--------------------------------------------------------------------------------
		strCRT_SessionFolder = objShell.RegRead(CRT_REG_SESSION)
		if Err.Number <> 0 Then 
            SecureCRT_Installed = False
			strCRT_SessionFolder = "C:"
		End If
		If Right(strCRT_SessionFolder,1) = "\" Then strCRT_SessionFolder = Left(strCRT_SessionFolder,Len(strCRT_SessionFolder)-1)
		strCRT_SessionFolder = strCRT_SessionFolder & "\Sessions"
		Call TrDebug("Session Folder: " & strCRT_SessionFolder, "", objDebug, MAX_LEN, 1, 1)		
	On Error Goto 0
    strDirectoryVandyke = strCRT_InstallFolder
	strCRTexe = """" & strDirectoryVandyke & strCRTexe	
	'------------------------------------------------------
	'  Copy Session Files to the SecureCrt Database
	'------------------------------------------------------
	If Not objFSO.FolderExists(strCRT_SessionFolder & "\" & strFolder) Then 
	    objFSO.CreateFolder strCRT_SessionFolder & "\" & strFolder
    End If
	If Not objFSO.FileExists(strCRT_SessionFolder & "\" & strFolder & "\" & strSession & ".ini" ) Then 
	   objFSO.CopyFile strDirectoryLCL & "\" & strSession & ".ini", strCRT_SessionFolder & "\" & strFolder & "\" & strSession & ".ini" , True
	   ' objFSO.CopyFile strDirectoryLCL & "\" & "__FolderData__.ini",strCRT_SessionFolder & "\" & strFolder & "\" & "__FolderData__.ini" , True
	End If 
    Mode = 2
	Dim vTerm
	Redim vTerm(0)
	Dim strCmd
	strCmd = ""
	Select Case Mode
	        Case 2
			    strFilter = "FW_KIDSSHIELD"
			    Set objTab = crt.Session.ConnectInTab("/S Home\SRX-Home")
				'--------------------------------------------------------------------------------
				'	SEND <RETURN> KEY
				'--------------------------------------------------------------------------------
				objTab.Caption = strTab
				objTab.Screen.Synchronous = True
				'--------------------------------------------------------------------------------
				'  Get actual host name of the box
				'--------------------------------------------------------------------------------
				objTab.Screen.Send chr(13)
				strLine = objTab.Screen.ReadString (">")
				If InStr(strLine,"@") Then strHost = Split(strLine,"@")(1)
				objTab.Screen.Send chr(13)
				nResult = objTab.Screen.WaitForString ("@" & strHost & ">",2)
				If nResult = 0  Then
					Call TrDebug ("ERROR: CAN'T GET RESPONSE FROM NODE", "", objDebug, MAX_LEN, 1, 1)
					objTab.Session.Disconnect
					Exit Sub 
				End If	
				objTab.Screen.Send chr(13)
				objTab.Screen.WaitForString "@" & strHost & ">"
				objTab.Screen.Send "edit" & chr(13)
				objTab.Screen.WaitForString "@" & strHost & "#"
			    objTab.Screen.Send "show firewall family inet filter " & strFilter & " |match ""inactive: term""" & chr(13)
             	strLine = objTab.Screen.ReadString ("@" & strHost & "#")
				vLine = Split(strLine,chr(13))
				nTerm = 0 
				For i = 0 to UBound(vLine)
				    If InStr(vLine(i),"inactive: term ") > 0 Then 
					   Redim Preserve vTerm(nTerm + 1)
				       vTerm(nTerm) = Split(vLine(i)," ")(2)
					   Call TrDebug ("[" & vTerm(nTerm) & "]", "", objDebug, MAX_LEN, 1, 1)
					   nTerm = nTerm + 1
				    End If
				Next
				For nTerm = 0 to UBound(vTerm) - 1
						strCmd = "activate firewall family inet filter " & strFilter & " term " & vTerm(nTerm)
						Call TrDebug ("SEND COMMAND: " & strCmd , "", objDebug, MAX_LEN, 1, nDebug)
						objTab.Screen.Send strCmd & chr(13)
						objTab.Screen.WaitForString "@" & strHost & "#"
				Next

			Case 1
                Set objTab = crt.Session.ConnectInTab("/S Home\SRX-Home")
				'--------------------------------------------------------------------------------
				'	SEND <RETURN> KEY
				'--------------------------------------------------------------------------------
				objTab.Caption = strTab
				objTab.Screen.Synchronous = True
				'--------------------------------------------------------------------------------
				'  Get actual host name of the box
				'--------------------------------------------------------------------------------
				objTab.Screen.Send chr(13)
				strLine = objTab.Screen.ReadString (">")
				If InStr(strLine,"@") Then strHost = Split(strLine,"@")(1)
				objTab.Screen.Send chr(13)
				nResult = objTab.Screen.WaitForString ("@" & strHost & ">",2)
				If nResult = 0  Then
					Call TrDebug ("ERROR: CAN'T GET RESPONSE FROM NODE", "", objDebug, MAX_LEN, 1, 1)
					objTab.Session.Disconnect
					Exit Sub 
				End If	
				objTab.Screen.Send chr(13)
				objTab.Screen.WaitForString "@" & strHost & ">"
				objTab.Screen.Send "show version" & chr(13)
				objTab.Screen.WaitForString "@" & strHost & ">"
			End Select		

	If IsObject(objDebug) Then objDebug.close : End If
	If objFSO.FileExists(stdOutFile) Then 
		objFSO.DeleteFile stdOutFile, True
	End If
'	crt.quit
	Set objFSO = Nothing
    Set objShell = Nothing
End Sub
'------------------------------------------------------------------------------------------
'    Function CheckSRXFilters(strFolder, strSession, Byref strHost, strTab, ByRef vCommand, strFilter, nDebug)		
'-------------------------------------------------------------------------------------------
Function CheckSRXFilters(strFolder, strSession, Byref strHost, strTab, ByRef vCommand, strFilter, nDebug)		
	Dim  nResult, strStdOut, vStdOut
	Dim vCmd, nCmd
	Dim vCmdList
	Dim objTab
	Dim strCmd
	Dim vWaitForCommit
	Const MAX_LEN = 140	
	vWaitForCommit = Array("error: configuration check-out failed","error: commit failed","commit complete")
	crt.Screen.Synchronous = True
	'--------------------------------------------------------------------------------
    '  Start Telnet session to Home Gateway
    '--------------------------------------------------------------------------------
   nInd = 0
	'--------------------------------------------------------------------------------
	'	OPEN NEW SECURECRT TAB
	'--------------------------------------------------------------------------------
    On Error Resume Next
	Err.Clear
	Set objTab = crt.Session.ConnectInTab("/S " & strFolder & strSession)
	If Err.Number <> 0 Then 
		Call  TrDebug (strTab & "ERROR:", Err.Number & " Srce: " & Err.Source & " Desc: " &  Err.Description , objDebug, 1, MAX_LEN, 1)
		CheckSRXFilters = False
		Exit Function
	End If
	On Error Goto 0
	'--------------------------------------------------------------------------------
	'	SEND <RETURN> KEY
	'--------------------------------------------------------------------------------
	objTab.Caption = strTab
	objTab.Screen.Synchronous = True
	'--------------------------------------------------------------------------------
    '  Get actual host name of the box
    '--------------------------------------------------------------------------------
	objTab.Screen.Send chr(13)
	strLine = objTab.Screen.ReadString (">")
    If InStr(strLine,"@") Then strHost = Split(strLine,"@")(1)
	objTab.Screen.Send chr(13)
	nResult = objTab.Screen.WaitForString ("@" & strHost & ">",2)
    If nResult = 0  Then
		Call TrDebug ("ERROR: CAN'T GET RESPONSE FROM NODE", "", objDebug, MAX_LEN, 1, 1)
		objTab.Session.Disconnect
        CheckSRXFilters = False
		Exit Function 
	End If	
	objTab.Screen.Send chr(13)
	objTab.Screen.WaitForString "@" & strHost & ">"
	objTab.Screen.Send "edit" & chr(13)
	objTab.Screen.WaitForString "@" & strHost & "#"
	'--------------------------------------------------------------------------------
	'	CHECK CURRENT FILTER STATUS
	'--------------------------------------------------------------------------------
    objTab.Screen.Send "show firewall family inet filter " & strFilter & " |display set |match deactivate" & chr(13)
	strLine = objTab.Screen.ReadString ("@" & strHost & "#")
	For i = 0 to UBound(vCommand) - 1
	    If InStr(strLine,vCommand(i)) <> 0 Then Exit For
		If i = UBound(vCommand) - 1 Then 
			Call TrDebug("All Filters are alredy active " ,"EXIT", objDebug, MAX_LEN, 1, 1)
			CheckSRXFilters = True
			objTab.Screen.Send "exit" & chr(13)
			objTab.Screen.WaitForString "@" & strHost & ">"
			objTab.Session.Disconnect
			Exit Function
		End If 
	Next
	'--------------------------------------------------------------------------------
	'	SEND COMMANDS 
	'--------------------------------------------------------------------------------
	For i = 0 to UBound(vCommand) - 1
		strCmd = "activate firewall family inet filter " & strFilter & " term " & vCommand(i)
		Call TrDebug ("SEND COMMAND: " & strCmd , "", objDebug, MAX_LEN, 1, nDebug)
		objTab.Screen.Send strCmd & chr(13)
        objTab.Screen.WaitForString "@" & strHost & "#"
'		crt.sleep 50
	Next
	'--------------------------------------------------------------------------------
	'	COMMIT CONFIGURATION
	'--------------------------------------------------------------------------------
	Dstart = Date()
	Tstart = Time()
	Call  TrDebug ("COMMIT " & strHost, "......IN PROGRESS", objDebug, MAX_LEN, 1, 1)   
	objTab.Screen.Send "commit" & chr(13)
	nResult = objTab.Screen.WaitForStrings (vWaitForCommit, 30)
    Select Case nResult
        Case 0
			Call  TrDebug ("COMMIT " & strHost, "TIME OUT", objDebug, MAX_LEN, 1, 1) 
            CheckSRXFilters = False
	        objTab.Session.Disconnect			
			Exit Function 
        Case 1 
			Call  TrDebug ("COMMIT " & strHost, "ERROR 1", objDebug, MAX_LEN, 1, 1)
            objTab.Screen.Send chr(13)
            objTab.Screen.WaitForString "@" & strHost & "#"			
            objTab.Screen.Send "rollback" & chr(13)
            objTab.Screen.WaitForString "@" & strHost & "#"
			objTab.Screen.Send "exit" & chr(13)
			objTab.Screen.WaitForString "@" & strHost & ">"
			objTab.Session.Disconnect			
            CheckSRXFilters = False
			Exit Function			
        Case 2 
			Call  TrDebug ("COMMIT " & strHost, "ERROR 2", objDebug, MAX_LEN, 1, 1)
            objTab.Screen.Send chr(13)
            objTab.Screen.WaitForString "@" & strHost & "#"			
            objTab.Screen.Send "rollback" & chr(13)
            objTab.Screen.WaitForString "@" & strHost & "#"
			objTab.Screen.Send "exit" & chr(13)
			objTab.Screen.WaitForString "@" & strHost & ">"
            CheckSRXFilters = False			
			Exit Function
		Case Else
			Call  TrDebug ("COMMIT " & strHost, "OK", objDebug, MAX_LEN, 1, 1)   
    End Select		
	Tcompiler = DateDiff("s",Dstart & " " & Tstart,Date() & " " & Time()) 
	Call TrDebug("Commit time: " & Tcompiler & " sec" ,"", objDebug, MAX_LEN, 1, 1)
    CheckSRXFilters = True
	objTab.Screen.Send "exit" & chr(13)
	objTab.Screen.WaitForString "@" & strHost & ">"
	objTab.Session.Disconnect
End Function
'----------------------------------------------------------------
' Function CrtWriteProgressToFile 
'----------------------------------------------------------------
 Function CrtWriteProgressToFile(strProcessName, pProgress)
	Dim objProgressFile, f_objFSO, g_objShell
	Set g_objShell = CreateObject("WScript.Shell")	
	strWork = g_objShell.ExpandEnvironmentStrings("%USERPROFILE%")
	Set f_objFSO = CreateObject("Scripting.FileSystemObject")
	Set objProgressFile = f_objFSO.OpenTextFile(strWork & "\" & strProcessName & ".dat",ForAppending,True)
		objProgressFile.WriteLine pProgress
    objProgressFile.close
	Set f_objFSO = Nothing
	Set g_objShell = Nothing
End Function
'#######################################################################
' Function GetFileLineCountSelect - Returns number of lines int the text file
'#######################################################################
 Function GetFileLineCountSelect(strFileName, ByRef vFileLines,strChar1, strChar2, strChar3, nDebug)
    Dim nIndex
	Dim strLine
	Dim objDataFileName
    strFileWeekStream = ""	
	If objFSO.FileExists(strFileName) Then 
		On Error Resume Next
		Err.Clear
		Set objDataFileName = objFSO.OpenTextFile(strFileName)
		If Err.Number <> 0 Then 
			Call TrDebug("GetFileLineCountSelect: ERROR: CAN'T OPEN FILE:", strFileName, objDebug, MAX_LEN, 0, 1)
			On Error Goto 0
			Redim vFileLines(0)
			GetFileLineCountSelect = 0
			Exit Function
		End If
	Else
	    Call TrDebug("GetFileLineCountSelect: ERROR: CAN'T FIND FILE:", strFileName, objDebug, MAX_LEN, 0, 1)
		Redim vFileLines(0)
		GetFileLineCountSelect = 0
		Exit Function
	End If 
    Redim vFileLines(0)
	Set objDataFileName = objFSO.OpenTextFile(strFileName)	
	If nDebug = 1 Then objDebug.WriteLine "           NOW TRYING TO RIGHT INTO AN ARRAY        "
	nIndex = 0
    Do While objDataFileName.AtEndOfStream <> True
		strLine = objDataFileName.ReadLine
		Select Case Left(strLine,1)
			Case strChar1
			Case strChar2
			Case strChar3
			Case Else
					Redim Preserve vFileLines(nIndex + 1)
					vFileLines(nIndex) = strLine
					If nDebug = 1 Then objDebug.WriteLine "GetFileLineCountSelect: vFileLines(" & nIndex & ")="  & vFileLines(nIndex) End If  
					nIndex = nIndex + 1
		End Select
	Loop
	objDataFileName.Close
    GetFileLineCountSelect = nIndex
End Function
'-----------------------------------------------------------------
'     Function GetMyDate()
'-----------------------------------------------------------------
Function GetMyDate()
	GetMyDate = Month(Date()) & "/" & Day(Date()) & "/" & Year(Date()) 
End Function
'-----------------------------------------------------------------
'     Function GetDateFormat(nFormat)
'     nFormat: 1 = m/d/yyyy
'              2 = yyyy-mm-dd
'-----------------------------------------------------------------
Function GetDateFormat(MyDate, nFormat)
Dim strDate, strDay, strMonth, strYear
    If IsDate(MyDate) Then 
	    strDate = MyDate
	Else
	    strDate = Date()
		Call TrDebug("ERROR: Can't translate Date: " & MyDate, "", objDebug, 80 , 1, 1)
		Call TrDebug("CURRENT DATE WILL BE USED: " & Date(), "", objDebug, 80 , 1, 1)
	End If
    Select Case nFormat
	    Case 1
	        GetDateFormat = Month(strDate) & "/" & Day(strDate) & "/" & Year(strDate) 
		Case 2
		    If Year(strDate) <= 9 Then strYear = "0" & Year(strDate) Else strYear = Year(strDate)
			If Month(strDate) <= 9 Then strMonth = "0" & MOnth(strDate) Else strMonth = MOnth(strDate)
			If Day(strDate) <= 9 Then strDay = "0" & Day(strDate) Else strDay = Day(strDate)
			GetDateFormat = strYear  & "-" & strMonth & "-" & strDay
		Case Else 
		    GetDateFormat = Month(strDate) & "/" & Day(strDate) & "/" & Year(strDate) 
	End Select
End Function
'--------------------------------------------------------------
' Function returns a random intiger between min and max
'--------------------------------------------------------------
Function My_Random(min, max)
	Randomize
	My_Random = (Int((max-min+1)*Rnd+min))
End Function
' ----------------------------------------------------------------------------------------------
'   3 - Center and Mark
'	2 - Center
'	1 - Strach
'	0 - As is
'   nFormat: 
'   Function  TrDebug (strTitle, strString, objDebug)
' ----------------------------------------------------------------------------------------------
Function TrDebug(strTitle, strString, objDebug, nChar, nFormat, nDebug)
Dim strLine
strLine = ""
If nDebug <> 1 Then Exit Function End If
If IsObject(objDebug) Then 
	Select Case nFormat
		Case 0
			strLine = GetMyDate() & " " & FormatDateTime(Time(), 3) 
			strLine = strLine & ":  " & strTitle
			strLine = strLIne & strString
			objDebug.WriteLine strLine
			
		Case 1
			strLine = GetMyDate() & " " & FormatDateTime(Time(), 3)
			strLine = strLine & ":  " & strTitle
			If nChar - Len(strLine) - Len(strString) > 0 Then 
				strLine = strLine & Space(nChar - Len(strLine) - Len(strString)) & strString
			Else 
				strLine = strLine & " " & strString
			End If
			objDebug.WriteLine strLine
		Case 2
			strLine = GetMyDate() & " " & FormatDateTime(Time(), 3) & ":  "
			
			If nChar - Len(strLine & strTitle & strString) > 0 Then 
					strLine = strLine & Space(Int((nChar - 1 - Len(strLine & strTitle & strString))/2)) & strTitle & " " & strString			
			Else 
					strLine = strLine & strTitle & " " & strString	
			End If
			objDebug.WriteLine strLine
		Case 3
			strLine = GetMyDate() & " " & FormatDateTime(Time(), 3) & ":  "
			For i = 0 to nChar - Len(strLine)
				strLIne = strLIne & "-"
			Next
			objDebug.WriteLine strLine
			strLine = GetMyDate() & " " & FormatDateTime(Time(), 3) & ":  "
			If nChar - 1 - Len(strLine & strTitle & strString) > 0 Then 
					strLine = strLine & Space(Int((nChar - 1 - Len(strLine & strTitle & strString))/2)) & strTitle & " " & strString			
			Else 
					strLine = strLine & strTitle & " " & strString	
			End If
			objDebug.WriteLine strLine
			strLine = GetMyDate() & " " & FormatDateTime(Time(), 3) & ":  "
			For i = 0 to nChar - Len(strLine)
				strLIne = strLIne & "-"
			Next
			objDebug.WriteLine strLine
	End Select
						
End If
End Function
