Dim strDirectoryWork, D0, objFSO, objShell, objEnvar, CurrentDate, CurrentTime, nDebug, nInfo, objDebug, ShowLog, objIE
Const ForAppending = 8
Const ForWriting = 2
Const HttpTextColor1 = "#292626"
Const HttpTextColor2 = "#F0BC1F"
Const HttpTextColor3 = "#EBEAF7"
Const HttpTextColor4 = "#A4A4A4"
Const HttpBgColor1 = "Grey"
Const HttpBgColor2 = "#292626" 
Const HttpBgColor3 = "#2C2A23" 
Const HttpBgColor4 = "#504E4E"
Const HttpBgColor5 = "#0D057F"
Const HttpBgColor6 = "#8B9091"
Const HttpBdColor1 = "Grey"
Const DEBUG_FILE = "debug-test-script"
Const MAX_LEN = 140
' strDirectoryWork = "C:\VBScript"
D0 = DateSerial(2015,1,1)
Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
Set objEnvar = WScript.CreateObject("WScript.Shell")
Set objShell = WScript.CreateObject("WScript.Shell")
Set objApp = CreateObject("Shell.application")
strDirectoryWork =  objFSO.GetParentFolderName(Wscript.ScriptFullName)
nDebug = 0
nInfo = 1
ShowLog = True
CurrentDate = Date()
CurrentTime = Time()
D0 = DateSerial(2015,1,1)

Class Devices
	Private device_id
	Private device_activity
	'
	'
	Public Property Let Id(strID)
		device_id = strID
	End Property
	'
	'
	Public Property Let Activity(strAct)
		device_activity = strAct
	End Property
	'
	'
	Public Property Get Id()
		Id = device_id
	End Property
	'
	'
	Public Property Get Activity()
		Activity = device_activity
	End Property
End Class

Main()
	Call TrDebug("SCRIPT END", "", objDebug, MAX_LEN , 3, nInfo)
If IsObject(objDebug) Then objDebug.Close : End If
Set objFSO = Nothing
set objEnvar = Nothing
Set objShell = Nothing

Sub Main()

	'-------------------------------------------------------------------------------------------
	'  OPEN LOG FILE
	'-------------------------------------------------------------------------------------------
	If Not objFSO.FolderExists(strDirectoryWork & "\Log") Then 
			objFSO.CreateFolder(strDirectoryWork & "\Log") 
	End If
	'-----------------------------------------------------------------
	'  	CHECK IF START SCRIPT IS ALREADY RUNNING AND OPEN LOG FILE
	'-----------------------------------------------------------------
	strDebugFile = strDirectoryWork & "\Log\" & DEBUG_FILE & ".log"
	UtilsFolder = "C:\UnixUtils"
	bVerbose = False
	bMultipleInstanceAllowed = False
	If Not OpenLogSession(objDebug, strDebugFile, UtilsFolder, bMultipleInstanceAllowed, ShowLog, bVerbose) Then Exit Sub
	'
	' Print out PID of the script process
	Call GetWinAppPID(strPID, strParrentID, "test.vbs", "wscript.exe", nDebug)
	Call TrDebug("test.vbs script is running with PID: " & strPID, "", objDebug, MAX_LEN , 1, nInfo)
    '------------------------------------------------
	'   MAIN SYCLE
	'------------------------------------------------
	Dim colListOfServices, ServiceName, objAccount, colAccounts, objSession, colList
	Dim objWMIService, strMsg, strComputer
	Dim colItems, objItem, vScreen
	Dim vIncl, vExcl	
	Dim vCmdOut, strCMD, Line
	Call TrDebug("SCRIPT BEGINS", "", objDebug, MAX_LEN , 3, nInfo)
	Dim objDom, objGr, objUsr, objEnv, colGroups
	Dim objLocalGroup, objDomainUser, DomName
	Dim objWindows
	Dim PID, strParentPID
	
	nBlock = 141
	Select Case nBlock
		Case 146
		    MsgBox "Start"
			temp = "xxxxxxxxx"
			key = "huasHIYhkasdh11"
			temp = encrypt(temp,key)
			WScript.Echo temp
			key = "huasHIYhkasdh12"
			temp = Decrypt(temp,key)
			WScript.Echo temp
		
		Case 145
			Dim x, y, z
			y=Array(1,2,3)
			x=Array(y,Array(3,4,5), Array(6,7,8))
			z=Array(x, Array(7,8,9), Array(10,11,12)) 
			Call TrDebug("z(0)(0)(0)= " & z(0)(0)(0), "", objDebug, MAX_LEN , 1, nInfo)
	    Case 144 
				Call GetFileLineCountSelect("C:\VBScript\accounts.dat", vFileLines,"Ivan", "[End", "ljkhjkhj", 1)
	    Case 143
				vArray = Array(1,2,3)
				strLine = Join(vArray,",")
				MsgBox strLine
		Case 141
'			Dim objRegEx
			Set objRegEx = CreateObject("VBScript.RegExp")
			objRegEx.Global = False			
			Const IE_PAUSE = 200
			Dim vAcademics(9,20)
			nRetries = 4
			vAcademics(0,0) = "Classroom Name"
			vAcademics(1,0) = "Grade"
			vAcademics(2,0) = "Progress"
			vAcademics(3,0) = "Date List"			
			vAcademics(4,0) = "Score List"
			vAcademics(5,0) = "Category List"
			vAcademics(6,0) = "Weight List"			
			vAcademics(7,0) = "Count"
			vAcademics(8,0) = "Progress Report url"
			nAcademics = UBound(vAcademics,1)-1
			vCred = Array(_
						"https://evhs.schoolloop.com/portal/parent_home",_
						"vmukhin",_
						"Null",_
						"ANASTASIA",_
						"MUKHINA")
			If LoginEVHS(objIE, vCred,nRetries,nInfo) Then 
				Call TrDebug("Portal Page loaded OK" , "", objDebug, MAX_LEN , 1, nInfo)
			Else 
				Exit Sub
			End If
			Call TrDebug("Location Origin: " & objIE.Document.Location.origin, "", objDebug, MAX_LEN , 1, nInfo)
			Call TrDebug("Location Hostname: " & objIE.Document.Location.hostname, "", objDebug, MAX_LEN , 1, nInfo)
			Exit Sub
			Call SelectStudentPage(objIE, vCred, nInfo)
			nIndex = 0
			For Each oTable in objIE.Document.getElementsByTagName("table")
				If ProgressReportExists(oTable, nInfo) Then 
					nIndex = nIndex + 1
					Call GetProgressReportHref(oTable, vAcademics,nIndex,nInfo)
					Call GetClassroom(oTable, vAcademics,nIndex,nInfo)
					Call GetGrade(oTable, vAcademics,nIndex,nInfo)
					Call GetScore(oTable, vAcademics,nIndex,nInfo)
				End If 
			Next
			nIndex = 1
			Do While vAcademics(0,nIndex) <> ""
				Call TrDebug(vAcademics(0,0) & ": " & vAcademics(0,nIndex),"", objDebug, MAX_LEN , 3, nInfo)
				Call TrDebug(vAcademics(1,0) & ": " & vAcademics(1,nIndex),"" , objDebug, MAX_LEN , 1, nInfo)
				Call TrDebug(vAcademics(2,0) & ": " & vAcademics(2,nIndex),"", objDebug, MAX_LEN , 1, nInfo)
				Call TrDebug(vAcademics(nAcademics,0) & ": " & Left(vAcademics(nAcademics,nIndex),30),"", objDebug, MAX_LEN , 1, nInfo)
				nIndex = nIndex + 1
			Loop
			'
			'  Load Student Scores
			nIndex = 1
			Do While vAcademics(0,nIndex) <> ""
				objRegEx.Pattern = "https://.+\.com/"
				' Validate progress report link
				If objRegEx.Test(vAcademics(nAcademics,nIndex)) Then 
					objIE.navigate vAcademics(nAcademics,nIndex)
					nTimer = 0
					Do
						WScript.Sleep 200
						nTimer = nTimer + 0.2
						If nTimer > 10 Then exit do
					Loop While objIE.Busy
					Call TrDebug("Page 3 loaded in: " & nTimer & "sec.", "", objDebug, MAX_LEN , 1, nInfo)
					wscript.sleep 2000
					Call GetAssessmentsList(objIE, vAcademics, nIndex, nInfo)
				End If
				nIndex = nIndex + 1
			Loop
			nIndex = 1
			Do While vAcademics(0,nIndex) <> ""
				Call TrDebug(vAcademics(0,0) & ": " & vAcademics(0,nIndex),"", objDebug, MAX_LEN , 3, nInfo)
				For nRow = 1 to nAcademics - 1
					Call TrDebug(vAcademics(nRow,0) & ": " & vAcademics(nRow,nIndex),"", objDebug, MAX_LEN , 1, nInfo)
				Next
				nIndex = nIndex + 1
			Loop	
		Case 142 ' Downloading files from Jnpr sharepoint page or any hhtp table with links to pptx files
			' see: https://msdn.microsoft.com/en-us/library/windows/desktop/bb773974(v=vs.85).aspx
			Set objWindows = objApp.Windows
			If IsObject(objWindows) Then
				strLine = ""
				'MsgBox objWindows.Count
				For Each Window in objWindows
					'$Wpid = ObjName($Window,3)
					strLine = Window.LocationName & " ; " & Window.FullName & " ; " & Wpid 
					Call TrDebug(strLine, "", objDebug, MAX_LEN , 1, nInfo)						
					If InStr(Lcase(Window.FullName), "iexplore") > 0 Then				
						strURL = """" & Window.Document.Location.href & """"
						Call TrDebug("You browsing the URL: ", strURL, objDebug, MAX_LEN , 1, nInfo)
						nTable = 0
						For each table in Window.Document.getElementsByTagName("table")
						    nTable = nTable + 1
							If nTable > 1 Then Exit for
						    Call TrDebug("TABLE " & nTable , "", objDebug, MAX_LEN , 1, nInfo)
						    nTr = 0
							nFile = 0
						    For each Row in table.getElementsByTagName("tr")
						        nTr = nTr + 1
								nCell = 0
								For each Cell in Row.getElementsByTagName("td")
								    nCell = nCell + 1
									If nCell = 5 then 
									    For each Anchor in Cell.getElementsByTagName("a")
										    nFile = nFile + 1
											Do
												URL = Anchor.href
												strSaveAs = Anchor.InnerText
												If Len(URL) > 255 Then
												   Call TrDebug("File #" & nFile & ": Url is longer then 255 characters. " , "SKIP", objDebug, MAX_LEN , 1, nInfo)
												   Exit Do
												End If
												If objFSO.FileExists(strDirectoryWork & "\" & strSaveAs) Then 
												   Call TrDebug(Left(strSaveAs,30) & ": File already exists. " , "SKIP", objDebug, MAX_LEN , 1, nInfo)
												   Exit Do
												End If
												If InStr(strSaveAs,".pptx") Then 
													Call TrDebug("Download: " & Left(strSaveAs,50)  , "", objDebug, MAX_LEN , 1, nInfo)
													'Call DownloadingFile(URL, strDirectoryWork, strSaveAs)
													'    Call Set_IE_obj (g_objIE)
														'g_objIE.Offline = True
													'	g_objIE.navigate URL
													'	Do
													'		WScript.Sleep 200
													'	Loop While g_objIE.Busy
													Set objPPT = CreateObject("PowerPoint.Application")
													objPPT.Visible = True
													Set objPresentation = objPPT.presentations.Open(URL, , , msoFalse)								
													wscript.sleep 15000
													objPresentation.SaveAs (strDirectoryWork & "\" & strSaveAs)
													objPPT.Quit
													Set objPPT = Nothing
													Set objPresentation = Nothing
												End If
												Exit Do
											Loop
										Next
									End If
								Next
						    Next
						    Call TrDebug("Parser found " & nTr & " rows in table " & nTable, "", objDebug, MAX_LEN , 1, nInfo)
						Next
						Call TrDebug("Parser found " & nTable & " tables", "", objDebug, MAX_LEN , 1, nInfo)
						Set objDataFileName = objFSO.OpenTextFile(strDirectoryWork & "\sharepoint_jnpr.txt",2,True)
		                objDataFileName.WriteLine "Test"
						objDataFileName.Close
						' Call WriteArrayToFile(strDirectoryWork & "\sharepoint_jnpr.txt",vIncl, UBound(vIncl),1,0)
						If InStr(Window.Document.Location.href, "cnn") > 0 Then 
							Window.GoHome
							wscript.sleep 5000
						End If
					End If
				Next
			End If		

		Case 54
			Dim objRegEx
			Set objRegEx = CreateObject("VBScript.RegExp")
			objRegEx.Global = False			
			strLin1 = "'#   Function IE_MSG_Internal (g_objIE, vIE_Scale, strTitle, vLine, ByVal nLine)"
			strLin2 = "END OF LIST"
			strLin3 = "function IE_MSG_Internal (g_objIE, vIE_Scale, strTitle, vLine, ByVal nLine)"
			strLin4 = "end   function"
			objRegEx.Pattern = "'#   Function "
			If objRegEx.Test(strLin1) Then Call TrDebug("Patter1: OK", "",objDebug, MAX_LEN , 1, 1)	 else Call TrDebug("Patter1: FAILED", "",objDebug, MAX_LEN , 1, 1)
			objRegEx.Pattern = "END OF LIST"
			If objRegEx.Test(strLin2) Then Call TrDebug("Patter2: OK", "",objDebug, MAX_LEN , 1, 1)	 else Call TrDebug("Patter2: FAILED", "",objDebug, MAX_LEN , 1, 1)
			objRegEx.Pattern = "^\s{0,2}[Ff]unction\s{1,2}\w*\s{0,3}\(.*\)"
			If objRegEx.Test(strLin3) Then Call TrDebug("Patter3: OK", "",objDebug, MAX_LEN , 1, 1)	 else Call TrDebug("Patter3: FAILED", "",objDebug, MAX_LEN , 1, 1)
			objRegEx.Pattern = "[Ee]nd\s{1,3}[Ff]unction"
			If objRegEx.Test(strLin4) Then Call TrDebug("Patter4: OK", "",objDebug, MAX_LEN , 1, 1)	 else Call TrDebug("Patter4: FAILED", "",objDebug, MAX_LEN , 1, 1)

		Case 52
		   MsgBox Now
		Case 51
		   strLine = CopyFileToString("C:\VBScript\test.txt")
		   Call TrDebug("Text: " & strLine, "",objDebug, MAX_LEN , 1, 1)	
		Case 50
			Set objFile = objFSO.GetFile("C:\VBScript\task_find_and_kill.vbs.new")
			Call TrDebug("Name:      " & objFile.Name, "",objDebug, MAX_LEN , 1, 1)
			Call TrDebug("Path:      " & objFile.Path, "",objDebug, MAX_LEN , 1, 1)
			Call TrDebug("Ext:       " &  objFSO.GetExtensionName(objFile.Path), "",objDebug, MAX_LEN , 1, 1)
			Call TrDebug("Base Name: " &  objFSO.GetBaseName(objFile.Path), "",objDebug, MAX_LEN , 1, 1)
			Call TrDebug("ShortName: " & objFile.ShortName, "",objDebug, MAX_LEN , 1, 1)
			Call TrDebug("Type:      " & objFile.Type, "",objDebug, MAX_LEN , 1, 1)
				   
		Case 49
		    Dim KidsDevice, vDevices
		    Redim KidsDevices(1)
			Set KidsDevices(0) = New Devices
			KidsDevices(0).Id = "IVAN-PC"
			KidsDevices(0).Activity = "ACTIVE"
			Redim preserve KidsDevices(2)
			Set KidsDevices(1) = New Devices
			KidsDevices(1).Id = "MEDIA-PC"
			KidsDevices(1).Activity = "INACTIVE"
			For each host in KidsDevices
				If IsObject(host) Then 
					Call TrDebug("strDeviceID: " & host.id & " Status: " & host.activity, "",objDebug, MAX_LEN , 1, 1)	
				End If 
			Next
			
	    Case 48
		        If IsDate("11/5/2016 12:11:50 AM") Then MsgBox "Ok"
		Case 47
				Set objNet = WScript.CreateObject("WScript.Network")
                Wscript.Echo "Your Computer Name is " & objNet.ComputerName
                WScript.Echo "Your Username is " & objNet.UserName
		Case 46
				strServer ="127.0.0.1"
				ServiceName = "ClientReportSvc"
				StartMode = "Manual"
				On Error Resume Next
					Err.Clear
					Set objWMIService = GetObject("winmgmts:\\" & strServer & "\root\cimv2")
					If Err.Number <> 0 Then 
						Call TrDebug("SetServiceStartMode: " & Err.Description, "Error" & Err.Number, objDebug, MAX_LEN , 1, 1)	
						Exit Sub
					End If 
					Set oService = objWMIService.Get("Win32_Service.Name='" & ServiceName & "'")
					Call TrDebug("SetServiceStartMode: Path Name: " & oService.PathName, "", objDebug, MAX_LEN , 1, 1)	
					If Err.Number <> 0 Then 
						Call TrDebug("SetServiceStartMode: " & Err.Description, "Error" & Err.Number, objDebug, MAX_LEN , 1, 1)	
						Exit Sub
					End If 				
					nResult = oService.Change( , , , ,StartMode)
					Call TrDebug("SetServiceStartMode: Operation Error Code: " & nResult, "", objDebug, MAX_LEN , 1, 1)	
					Err.Clear
				On Error goto 0
        Case 45
			Call GetAllAppPID(strPID, strParentPID, strCommandLine, "chrome.exe", 1)
		
		Case 44
			strDebugFile = "C:\Tmp\Log\mydebug-01-02-2016.log"
			nLenEnd = InStrRev(strDebugFile,"\")
			strErrorLog = Left(strDebugFile,nLenEnd) & Split(Right(strDebugFile,Len(strDebugFile) - nLenEnd),".")(0) & "_Error.log"
			MsgBox strErrorLog
			Call TrDebug("Error log file name: " & strErrorLog, "", objDebug, MAX_LEN , 1, nInfo)
	    Case 43
		    strDirectoryLCL = "C:\Users\All Users\Vandyke"
			strFileSession = "sessions.txt"
			If objFSO.FileExists(strDirectoryLCL & "\" & strFileSession) Then 
				nSession = GetFileLineCountSelect(strDirectoryLCL & "\" & strFileSession, vSession,"NONE","NONE", "NONE", 0)
				strSrvDirectory = vSession(1)
				If Right(strSrvDirectory,1) = "\" Then 
					strSrvDirectory = Left(strSrvDirectory,Len(strSrvDirectory) - 1)
					If nDebug = 1 Then objDebug.WriteLine FormatDateTime(Date(), 0) & " " & FormatDateTime(Time(), 3) & ": Remote Server Folder: " & strSrvDirectory End If
				End If
				If nSession > 0 Then strDirectoryVandyke = vSession(0)	End If	' - Vandyke folder to run SecureCRT from 
				If nSession > 3 Then strDirectoryWork = vSession(3)		End If	' - Work directory scripts are installed to 
				If nSession > 4 Then strDirectoryUpdate = vSession(4)	End If	' - Source directory to take updates from 
				If nSession > 5 Then strVersion = vSession(5)			End If	' - Current version of the Package/Launcer
				If nSession > 6 Then strLclDeviceID = vSession(6)    	End If	' - Name of the Local PC is stored in session.txt line 7
				If nSession > 7 Then strOwnerID = vSession(7)    		End If	' - Name of the Account used for last time logon line 8
				If nSession > 8 Then strServerIP = vSession(8)  		End If	' - Server IP address is stored in session.txt line 9
				If nSession >=10 Then strGatewayIP = vSession(9) 		End If	' - Gateway IP address is stored in session.txt line 10
				If nSession >=11 Then strModeID = vSession(10)		    End If  ' - GET Host type: CLIENT or SERVER
			End If		
			Const REG_CRT_INSTALL = "HKEY_LOCAL_MACHINE\SOFTWARE\VanDyke\SecureCRT\Install\Main Directory"
			Const REG_CRT_SESSION = "HKEY_CURRENT_USER\Software\VanDyke\SecureCRT\Config Path"
			Const ALLOW_REMOTE_RPC = "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Terminal Server\AllowRemoteRPC"
			Const TOKEN_FILTER_POLICY = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System\LocalAccountTokenFilterPolicy"
			Const REG_KIDSHIELD = "HKEY_LOCAL_MACHINE\SOFTWARE\KidsShield\"
			Const REG_DIRECTORY_LCL = "\Local Directory"
			Const REG_DIRECTORY_WRK = "\Work Directory"
			Const REG_DIRECTORY_SRV = "\Server Directory"
			Const REG_DIRECTORY_UPD = "\Update Directory"
			Const REG_VERSION = "\Version"
			Const REG_LCL_DEVICE_ID = "\Local Device ID"
			Const REG_LCL_OWNER_ID = "\Device Owner ID"
			Const REG_GW_IP = "\Gateway IP"
			Const REG_SRV_IP = "\Server IP"
			objShell.RegWrite REG_KIDSHIELD & REG_DIRECTORY_LCL,strDirectoryLCL & "\","REG_SZ"
			objShell.RegWrite REG_KIDSHIELD & REG_DIRECTORY_SRV,strSrvDirectory & "\","REG_SZ"
			objShell.RegWrite REG_KIDSHIELD & REG_DIRECTORY_WRK,strDirectoryWork & "\","REG_SZ"
			objShell.RegWrite REG_KIDSHIELD & REG_DIRECTORY_UPD,strDirectoryUpdate & "\","REG_SZ"
			objShell.RegWrite REG_KIDSHIELD & REG_VERSION,strVersion,"REG_SZ"
			objShell.RegWrite REG_KIDSHIELD & REG_LCL_DEVICE_ID,strLclDeviceID,"REG_SZ"
			objShell.RegWrite REG_KIDSHIELD & REG_LCL_OWNER_ID,strOwnerID,"REG_SZ"
			objShell.RegWrite REG_KIDSHIELD & REG_GW_IP,strGatewayIP,"REG_SZ"
			objShell.RegWrite REG_KIDSHIELD & REG_SRV_IP,strServerIP,"REG_SZ"	
			
	    
		Case 42 ' Downloading files from Jnpr sharepoint page or any hhtp table with links to pptx files
			' see: https://msdn.microsoft.com/en-us/library/windows/desktop/bb773974(v=vs.85).aspx
			Set objWindows = objApp.Windows
			If IsObject(objWindows) Then
				strLine = ""
				'MsgBox objWindows.Count
				For Each Window in objWindows
					'$Wpid = ObjName($Window,3)
					strLine = Window.LocationName & " ; " & Window.FullName & " ; " & Wpid 
					Call TrDebug(strLine, "", objDebug, MAX_LEN , 1, nInfo)						
					If InStr(Lcase(Window.FullName), "iexplore") > 0 Then				
						strURL = """" & Window.Document.Location.href & """"
						Call TrDebug("You browsing the URL: ", strURL, objDebug, MAX_LEN , 1, nInfo)
						nTable = 0
						For each table in Window.Document.getElementsByTagName("table")
						    nTable = nTable + 1
							If nTable > 1 Then Exit for
						    Call TrDebug("TABLE " & nTable , "", objDebug, MAX_LEN , 1, nInfo)
						    nTr = 0
							nFile = 0
						    For each Row in table.getElementsByTagName("tr")
						        nTr = nTr + 1
								nCell = 0
								For each Cell in Row.getElementsByTagName("td")
								    nCell = nCell + 1
									If nCell = 5 then 
									    For each Anchor in Cell.getElementsByTagName("a")
										    nFile = nFile + 1
											Do
												URL = Anchor.href
												strSaveAs = Anchor.InnerText
												If Len(URL) > 255 Then
												   Call TrDebug("File #" & nFile & ": Url is longer then 255 characters. " , "SKIP", objDebug, MAX_LEN , 1, nInfo)
												   Exit Do
												End If
												If objFSO.FileExists(strDirectoryWork & "\" & strSaveAs) Then 
												   Call TrDebug(Left(strSaveAs,30) & ": File already exists. " , "SKIP", objDebug, MAX_LEN , 1, nInfo)
												   Exit Do
												End If
												If InStr(strSaveAs,".pptx") Then 
													Call TrDebug("Download: " & Left(strSaveAs,50)  , "", objDebug, MAX_LEN , 1, nInfo)
													'Call DownloadingFile(URL, strDirectoryWork, strSaveAs)
													'    Call Set_IE_obj (g_objIE)
														'g_objIE.Offline = True
													'	g_objIE.navigate URL
													'	Do
													'		WScript.Sleep 200
													'	Loop While g_objIE.Busy
													Set objPPT = CreateObject("PowerPoint.Application")
													objPPT.Visible = True
													Set objPresentation = objPPT.presentations.Open(URL, , , msoFalse)								
													wscript.sleep 15000
													objPresentation.SaveAs (strDirectoryWork & "\" & strSaveAs)
													objPPT.Quit
													Set objPPT = Nothing
													Set objPresentation = Nothing
												End If
												Exit Do
											Loop
										Next
									End If
								Next
						    Next
						    Call TrDebug("Parser found " & nTr & " rows in table " & nTable, "", objDebug, MAX_LEN , 1, nInfo)
						Next
						Call TrDebug("Parser found " & nTable & " tables", "", objDebug, MAX_LEN , 1, nInfo)
						Set objDataFileName = objFSO.OpenTextFile(strDirectoryWork & "\sharepoint_jnpr.txt",2,True)
		                objDataFileName.WriteLine "Test"
						objDataFileName.Close
						' Call WriteArrayToFile(strDirectoryWork & "\sharepoint_jnpr.txt",vIncl, UBound(vIncl),1,0)
						If InStr(Window.Document.Location.href, "cnn") > 0 Then 
							Window.GoHome
							wscript.sleep 5000
						End If
					End If
				Next
			End If		

	    Case 41
		        Dim StrRecord, nLen
				Dim vWebApplication, vWebApplicationPattern
				'
				'  WebApplications Catalog
				vWebApplication = Array("YouTube",_
										"Google Drive",_
										"Google Search",_
										"Google Docs",_
										"Wikipedia",_
										"Schoolloop",_
										"Amazon")
				vWebApplicationPattern = Array("youtube",_
										"google drive",_
										"google search",_
										"google docs",_
										"wikipedia",_
										"(\S+) (\S+)'s portal:",_
										"amazon.com")
				' Start block cycle
				nInd = 0
		        strWinUser = "Vasily"
				Do While nInd < 1
					strCmd = "tasklist /V /fo csv /fi ""USERNAME eq " & strWinUser & """"
					Call RunCmd("127.0.0.1", "", vCmdOut, strCMD,"", nDebug)
					strPID = ""
					For Each strLine in vCmdOut
					    strRecord = ""
						If UBound(Split(strLine,""",""")) = 8 Then 
							If InStr(strLine,"chrome") or InStr(strLine,"firefox") or InStr(strLine,"iexplore")   then 
								strPID = Split(strLine,""",""")(1)
								strAppName = Split(strLine,""",""")(0)
								strAppName = Right(strAppName,Len(strAppName)-1)
								strRecord = strAppName & ":" & strPID 
								strTitle = Split(strLine,""",""")(8)
								
								If ((strTitle <> "N/A""") and (Len(Trim(strTitle)) > 5)) Then 
								    Call TrDebug(strTitle,"", objDebug, MAX_LEN , 1, nInfo)
								    strApp = GetInetApplication(strTitle,vWebApplication,vWebApplicationPattern)
									Select Case strApp
										Case "Amazon"
											strRecord = strRecord & ":Amazon"
											nLen = InStrRev(strTitle," - ")
											Select Case nLen 
												Case 0
												   strTitle = "Main Amazon page"
												Case Else 
												   strTitle = Left(strTitle,nLen)
											End Select
											strRecord = strRecord & ":" & strTitle
										Case "YouTube"
											strRecord = strRecord & ":YouTube"
											nLen = InStrRev(strTitle," - ")
											Select Case nLen 
												Case 0
												   strTitle = "Main youtube page"
												Case Else 
												   strTitle = Left(strTitle,nLen)
											End Select
											strRecord = strRecord & ":" & strTitle
										Case Else 
											strRecord = strRecord & ":" & strApp
											nLen = InStrRev(strTitle," - ")
											Select Case nLen 
												Case 0
												   strTitle = "Title page"
												Case Else 
												   strTitle = Left(strTitle,nLen)
											End Select
											strRecord = strRecord & ":" & strTitle
									End Select
									Call TrDebug(strRecord,"", objDebug, MAX_LEN , 1, nInfo)
								End If
							'	Call KillWinAppPID(strPID, "None", strAppName, nInfo)
							End If
						End If
					Next
					nInd=nInd+1
					Call TrDebug("CHECK #" & nInd , "", objDebug, MAX_LEN , 3, nInfo)
					wscript.sleep 10000
				Loop
		Case 40
	        Set obj = CreateObject("APIWrapperCOM.APIWrapper")
            Set winHandles = obj.FindWindow()
			For each winHandle in winHandles
			    Call TrDebug(winHandle, "", objDebug, MAX_LEN , 1, nInfo)									
            Next
			
	    Case 39
            Call GetBrowserOpenUrl(PID, strParentPID, "none", "chrome.exe", nInfo)
	    Case 38
			' see: https://msdn.microsoft.com/en-us/library/windows/desktop/bb773974(v=vs.85).aspx
			Set objWindows = objApp.Windows
			If IsObject(objWindows) Then
				strLine = ""
	'			MsgBox objWindows.Count
				For Each Window in objWindows
					'$Wpid = ObjName($Window,3)
					strLine=Window.FullName & " ; " & Window.LocationName  
'					Call TrDebug(strLine, "", objDebug, MAX_LEN , 1, nInfo)
					On Error Resume Next
					If InStr(Window.FullName, "iexplore") > 0 Then						
						Call TrDebug("Found " & Window.FullName, "", objDebug, MAX_LEN , 1, nInfo)									
						Call TrDebug("With title: " &  Window.Document.Title, "", objDebug, MAX_LEN , 1, nInfo)									
						strURL = """" & Window.Document.Location.href & """"
						Call TrDebug("You browsing the URL: ", strURL, objDebug, MAX_LEN , 1, nInfo)									
						If InStr(Window.Document.Title, "Admin Panel") = 0 Then 
							Window.quit
							Set Window = Nothing
						Else 
							Call TrDebug("This is KSLD Admin Panel: ", "Skip closing", objDebug, MAX_LEN , 1, nInfo)									
						End If
					End If
					On Error Goto 0
				Next
			End If		
	    Case 37
		   boolTest = IsDate("23:21")
		   Call TrDebug("BoolTest: " & boolTest, "", objDebug, MAX_LEN , 1, nInfo)						
		   nResult = StrComp("", "23:00" )
		   Call TrDebug("nResult: " & nResult, "", objDebug, MAX_LEN , 1, nInfo)						
	    Case 36
		   strContact = InputBox("Please enter the new contact folder name")
		   WScript.Echo strContact
		   WScript.Echo "Done"
	    Case 35
		    Dim Array1(3,3), Array2
		    Array1(0,0) = "1" : Array1(1,0) = "4" : Array1(2,0) = "7" 
			Array1(0,1) = "2" : Array1(1,1) = "5" : Array1(2,1) = "8" 
			Array1(0,2) = "3" : Array1(1,2) = "6" : Array1(2,2) = "9" 
			Call CopyArray(Array1, Array2,1)
	    Case 33
		   MsgBox DateDiff("s",Date(), Date() & " " & Time())
	    Case 32
		    WeekDate = DateAdd("d",-(Weekday(Date()) - 1),Date())
			StopDate = DateAdd("n",10,Date() & " " & Time())
			Call TrDebug("StopDate: " & StopDate, "", objDebug, MAX_LEN , 1, nInfo)
			StopTime = FormatDateTime(StopDate,4)
            Call TrDebug("StopTime: " & StopTime, "", objDebug, MAX_LEN , 1, nInfo)						
			Call TrDebug("StopHH: " & Hour(StopTime) & " StopMM=" & Minute(StopTime), "", objDebug, MAX_LEN , 1, nInfo)
			Call TrDebug("Filename: <Account>-" & Year(WeekDate) & Month(weekdate) & day(weekdate) & "-week.dat", "", objDebug, MAX_LEN , 1, nInfo)
            Call TrDebug("Filename: <Account>-" & Year(Date()) & Month(date()) & "01-month.dat", "", objDebug, MAX_LEN , 1, nInfo)			
			
			
		Case 31
		    Call TrDebug("DATE ADD", FormatDateTime(DateAdd("n",100 + 720,Date() & " " & Time()),4), objDebug, MAX_LEN , 1, nInfo)
			
	    Case 30
		    nResult = GetWinLogonWMI("10.168.6.15", vWinUsers, 1)
			Call TrDebug( "nResult = " & nResult, "", objDebug, MAX_LEN , 1, 1)
			For i = 0 to nResult - 1
			    Call TrDebug( "User Logon: " & vWinUsers(i), "", objDebug, MAX_LEN , 1, 1)
			Next
			Call TrDebug( "TOTAL LOG ON USERS: " & nResult, "", objDebug, MAX_LEN , 3, 1)
		Case 29
		    Call GetDefaultGwWMI(".", 1)
		Case 28
		    StartTime = FormatDateTime("2:0", 4)
			Call TrDebug( "START TIME: " & StartTime, "", objDebug, MAX_LEN , Array(3,False), nInfo)
	    Case 27
		     MsgBox CLng("&h" & "000008ae")
	    Case 26
		    strComputer = "127.0.0.1"
			GroupName = "HelpDesk"
			strFolder = "C:\test_folder"
			strPermissions = "(OI)(CI)F"
		    If SetCmdPermission(strComputer, "", GroupName, strFolder, strPermissions, 1)	Then 
	            Call TrDebug( "SUCCESS", "", objDebug, MAX_LEN, Array(3,False), nInfo)
			Else 
			    Call TrDebug( "FAILED", "", objDebug, MAX_LEN , Array(3,False), nInfo)
			End If
		Case 25 
		    strComputer = "."
			GroupName = "HelpDesk"
			UserName = "LocalSystem"
     		Call AddWinUserToGroup(strComputer,GroupName,UserName, nDebug)
	    Case 24 
		    strComputer = "."
			GroupName = "HelpDesk"
			UserName = "Vasily"
			' DomName = "K-MESON-W7-32"
		    If Not WinGroupExists(strComputer, GroupName, DomName, 1) Then
				'----------------------------------------------
				'   CREATE GROUP
				'----------------------------------------------
				Set objEnv = objShell.Environment("Process")
				strComputer = objEnv("COMPUTERNAME")
				Set objDom = Getobject("WinNT://" & strComputer & ",computer" )
				Set objGr = objDom.Create("group","HelpDesk")
				objGr.SetInfo
				Call TrDebug( "GROUP " & GroupName & " WAS CREATED", "", objDebug, MAX_LEN , 1, nInfo)
				Call WinGroupExists(strComputer, GroupName, DomName, 1)
			Else 
			    Call TrDebug( "GROUP " & GroupName & " ALREADY EXISTS", "", objDebug, MAX_LEN , 1, nInfo)
			End If
            Call TrDebug( "ALTERNATE METHOD OF GETTING GROUP LIST", "", objDebug, MAX_LEN , 3, nInfo)
			Set colGroups = GetObject("WinNT://" & strComputer & "")
            colGroups.Filter = Array("group")
			For Each objGroup In colGroups
				Call TrDebug( "Group: " &  objGroup.Name, "", objDebug, MAX_LEN , 1, nInfo)
				For Each objUser in objGroup.Members
					Call TrDebug( "      User: " &  objUser.Name, "", objDebug, MAX_LEN , 1, nInfo) 
				Next
			Next
			Set objLocalGroup = GetObject("WinNT://" & strComputer & "/" & GroupName & ",group")
            Call TrDebug( "CHECK IF USER IS A MEMBER OF THE GROUP", "", objDebug, MAX_LEN , 3, nInfo)			
			Set objLocalGroup = GetObject("WinNT://" & strComputer & "/" & GroupName & ",group")
			' Set objDomainUser = GetObject("WinNT://" & DomName & "/" & UserName & ",user")
			If Not objLocalGroup.IsMember("WinNT://" & DomName & "/" & UserName) Then
                Call TrDebug( "ADDING USER TO GROUP", "", objDebug, MAX_LEN , 3, nInfo)			
			    objLocalGroup.Add("WinNT://" & DomName & "/" & UserName)
				Call TrDebug( "USER " & UserName & " ADDED TO GROUP " & GroupName, "", objDebug, MAX_LEN , 1, nInfo)
			Else 
			    Call TrDebug( "USER " & DomName & "/" & UserName & " IS ALREADY A MEMBER OF THE GROUP " & GroupName, "", objDebug, MAX_LEN , 1, nInfo)
			End If
            Call TrDebug( "ALTERNATE METHOD OF GETTING GROUP LIST", "", objDebug, MAX_LEN , 3, nInfo)
			Set colGroups = GetObject("WinNT://" & strComputer & "")
            colGroups.Filter = Array("group")
			For Each objGroup In colGroups
				Call TrDebug( "Group: " &  objGroup.Name, "", objDebug, MAX_LEN , 1, nInfo)
				For Each objUser in objGroup.Members
					Call TrDebug( "      User: " &  objUser.Name, "", objDebug, MAX_LEN , 1, nInfo) 
				Next
			Next
		
	    Case 23
		    strCmd = "wmic useraccount get name"
			Call RunCmd("127.0.0.1", "", vCmdOut, strCMD, strAny, nInfo)
			For each Line in vCmdOut
			    Call TrDebug( "Caption: " & Line, "", objDebug, MAX_LEN , 1, nInfo)
			Next
			Call TrDebug( "DONE ", "", objDebug, MAX_LEN , 3, nInfo)
	    Case 22
		    ServiceName = "ClientReportSvc"
		    Call WinSvcExists(ServiceName)
			
	    Case 21 ' CMD Manage TaskSchedulers
	        strCmd = "schtasks /Change /TN ""Set_Screen_User_Vasily"" /RU ""NT AUTHORITY\SYSTEM"""
			strCmd = "schtasks /create /TN ""Clear_Task_System"" /RU ""NT AUTHORITY\SYSTEM"" /XML """ & strDirectoryWork & "\Clear_Screen_User_Vasily.xml"""
			strCmd = "schtasks /delete /TN ""Clear_Task_System"" /F"			
			strCmd = "schtasks /Query /FO LIST"						
'			"/Create /S " & strComputerName & " /RU " & strComputerName & "\" & Administrator & " /RP " & Password &" /XML " & strDirectoryWork & "\Tasks\" & vTaskList(i) & " /TN " & strTaskName
'			"/Create /S " & strComputerName & " /RU " & strComputerName & "\" & Administrator & " /RP " & Password &" /XML " & strDirectoryWork & "\Tasks\" & vTaskList(i) & " /TN " & strTaskName
		    Call RunCmd("127.0.0.1", "", vCmdOut, strCMD, strAny, nInfo)	
			For each Line in vCmdOut
			    If InStr(Line, "Vasily") Then Call TrDebug( "Caption: " & Line, "", objDebug, MAX_LEN , 1, nInfo)
			Next
	    Case 20 ' WMI Manage Taskschedulers
					strComputer = "."
			Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
			Set colScheduledJobs = objWMIService.ExecQuery ("Select * from Win32_ScheduledJob")

			For Each objJob in colScheduledJobs
				Call TrDebug( "Caption: " & objJob.Caption, "", objDebug, MAX_LEN , 1, nInfo)
				Call TrDebug( "Command: " & objJob.Command, "", objDebug, MAX_LEN , 1, nInfo)
				Call TrDebug( "Days Of Month: " & objJob.DaysOfMonth, "", objDebug, MAX_LEN , 1, nInfo)
				Call TrDebug( "Days Of Week: " & objJob.DaysOfWeek, "", objDebug, MAX_LEN , 1, nInfo)
				Call TrDebug( "Description: " & objJob.Description, "", objDebug, MAX_LEN , 1, nInfo)
				Call TrDebug( "Elapsed Time: " & objJob.ElapsedTime, "", objDebug, MAX_LEN , 1, nInfo)
				Call TrDebug( "Install Date: " & objJob.InstallDate, "", objDebug, MAX_LEN , 1, nInfo)
				Call TrDebug( "Interact with Desktop: " & objJob.InteractWithDesktop, "", objDebug, MAX_LEN , 1, nInfo)
				Call TrDebug( "Job ID: " & objJob.JobID, "", objDebug, MAX_LEN , 1, nInfo)
				Call TrDebug( "Job Status: " & objJob.JobStatus, "", objDebug, MAX_LEN , 1, nInfo)
				Call TrDebug( "Name: " & objJob.Name, "", objDebug, MAX_LEN , 1, nInfo)
				Call TrDebug( "Notify: " & objJob.Notify, "", objDebug, MAX_LEN , 1, nInfo)
				Call TrDebug( "Owner: " & objJob.Owner, "", objDebug, MAX_LEN , 1, nInfo)
				Call TrDebug( "Priority: " & objJob.Priority, "", objDebug, MAX_LEN , 1, nInfo)
				Call TrDebug( "Run Repeatedly: " & objJob.RunRepeatedly, "", objDebug, MAX_LEN , 1, nInfo)
				Call TrDebug( "Start Time: " & objJob.StartTime, "", objDebug, MAX_LEN , 1, nInfo)
				Call TrDebug( "Status: " & objJob.Status, "", objDebug, MAX_LEN , 1, nInfo)
				Call TrDebug( "Time Submitted: " & objJob.TimeSubmitted, "", objDebug, MAX_LEN , 1, nInfo)
				Call TrDebug( "Until Time: " & objJob.UntilTime, "", objDebug, MAX_LEN , 1, nInfo)
			Next
		Case 19 ' DELETE SIMPLE INTERNET ACCESS
			vIncl = Array("All")
            vExcl = Array("{", "}")		
		    Call GetFileLineCountInclude("C:\KidsSchild\DVLP\Project-001\SRX\address_book_delete.txt", vMap, vIncl, vExcl,0)
			For i = 0 to UBound(vMap) - 1
			   ADDR_NAME = split(vMap(i),"address ")(1)
			   ADDR_NAME = Left(ADDR_NAME, Len(ADDR_NAME) - 1)
			   vMap(i) =  "delete security address-book MY-BOOK-UNTRUST address " & ADDR_NAME
            Next
			Call WriteArrayToFile("C:\KidsSchild\DVLP\Project-001\SRX\srx_delete_addresses.conf",vMap, UBound(vMap),1,0)
		    
		Case 18 ' DELETE UNUSED SCHEDULERS
			vIncl = Array("All")
            vExcl = Array("Null")		
		    Call GetFileLineCountInclude("C:\KidsSchild\DVLP\Project-001\SRX\delet-scheduler.txt", vMap, vIncl, vExcl,0)
			For i = 0 to UBound(vMap) - 1
			   Redim Preserve vFileLines(3 * i + 3)
			   SHED_NAME = vMap(i)
			   vFileLines(3 * i) =  "delete schedulers scheduler " & SHED_NAME
			   vFileLines(3 * i + 1) = "delete schedulers scheduler " & SHED_NAME & "_TV"
			   vFileLines(3 * i + 2) = "delete schedulers scheduler " & SHED_NAME & "_Games"
            Next
			Call WriteArrayToFile("C:\KidsSchild\DVLP\Project-001\SRX\srx_delete_filters.conf",vFileLines, UBound(vFileLines),1,0)
	    Case 17 ' GENERATE FW FILTER CONFIG
			vIncl = Array("All")
            vExcl = Array("Null")
			Redim vFileLines(3)
    		Call GetFileLineCountInclude("C:\KidsSchild\DVLP\Project-001\SRX\Device-scheduler-map.txt", vMap, vIncl, vExcl,0)
			For i = 0 to UBound(vMap) - 1
			   Redim Preserve vFileLines(3 * i + 3)
			   DEVICE_NAME = Split(vMap(i),";")(0)
			   SHED_NAME = Split(vMap(i),";")(1)
			   vFileLines(3 * i) =     "set policy-options prefix-list " & DEVICE_NAME & " apply-path ""security address-book MY-BOOK-TRUST address " & DEVICE_NAME & " <*>"""
			   vFileLines(3 * i + 1) = "set firewall family inet filter FW_KIDSSHIELD term " & SHED_NAME & " from destination-prefix-list " & DEVICE_NAME
			   vFileLines(3 * i + 2) = "set firewall family inet filter FW_KIDSSHIELD term " & SHED_NAME & " then policer BW_128K"
            Next
			Call WriteArrayToFile("C:\KidsSchild\DVLP\Project-001\SRX\srx_filter_cfg.conf",vFileLines, UBound(vFileLines),1,0)
	    Case 16 ' REPLACE (NAME) with _NAME patterns
			vIncl = Array("All")
            vExcl = Array("Null")
    		Call GetFileLineCountInclude("C:\KidsSchild\DVLP\Project-001\SRX\_cfg.txt", vFileLines, vIncl, vExcl,0)
			For i = 0 to UBound(vFileLines) - 1
			   If InStr(vFileLines(i),"(Alex)") Then vFileLines(i) = Replace(vFileLines(i),"(Alex)","_Alex")
			   If InStr(vFileLines(i),"(Nast)") Then vFileLines(i) = Replace(vFileLines(i),"(Nast)","_Nast")	
			   If InStr(vFileLines(i),"(Ivan)") Then vFileLines(i) = Replace(vFileLines(i),"(Ivan)","_Ivan")			   
			Next
			Call WriteArrayToFile("C:\KidsSchild\DVLP\Project-001\SRX\_new_scheduler_cfg.dat",vFileLines, UBound(vFileLines),1,0)
		    
		Case 15 ' ADD ALIAS TO DEVICES FILE
			vIncl = Array("All")
            vExcl = Array("|")
    		Call GetFileLineCountInclude("C:\KidsSchild\DVLP\Project-001\SRX\devices.dat", vFileLines, vIncl, vExcl,0)
			For i = 0 to UBound(vFileLines) - 1
			   nCount = Ubound(Split(vFileLines(i),","))
			   For n = nCount + 1 to 9
			      vFileLines(i) = vFileLines(i) & ",-" 
			   Next
			   vFileLines(i) = vFileLines(i) & "," & Split(vFileLines(i),",")(0)
			Next
			Call WriteArrayToFile("C:\KidsSchild\DVLP\Project-001\SRX\devices_alias.dat",vFileLines, UBound(vFileLines),1,0)
		Case 14
			vIncl = Array("All")
            vExcl = Array("_TV ", "_Games ", "Night-Closed", "Evening_and_Morning")
		    Call GetFileLineCountInclude("C:\KidsSchild\DVLP\Project-001\SRX\schedulers.txt", vFileLines, vIncl, vExcl,0)
			For i=0 to UBound(vFileLines) - 1
			   vFileLines(i) = Split(vFileLines(i)," ")(3)
			Next 
			Call WriteArrayToFile("C:\KidsSchild\DVLP\Project-001\SRX\schedulers_list.txt",vFileLines, UBound(vFileLines),1,0)
		Case 13
		   Call TrDebug(FormatDateTime(Time(), 4), "", objDebug, MAX_LEN , 1, nInfo)
		   Call TrDebug(FormatDateTime("0:1", 4), "", objDebug, MAX_LEN , 1, nInfo)
		   Call TrDebug(FormatDateTime("7:9", 4), "", objDebug, MAX_LEN , 1, nInfo)
		   Call TrDebug(FormatDateTime(Date(), 0), "", objDebug, MAX_LEN , 1, nInfo)
		   Call TrDebug(FormatDateTime(Date(), 1), "", objDebug, MAX_LEN , 1, nInfo)		   
		   Call TrDebug(GetDateFormat(Date(), 2), "", objDebug, MAX_LEN , 1, nInfo)		   		   
		   Call TrDebug(GetDateFormat("01/02/2015", 2), "", objDebug, MAX_LEN , 1, nInfo)		   		   
		   Call TrDebug(GetDateFormat("07", 2), "", objDebug, MAX_LEN , 1, nInfo)		   		   
           
	    Case 12

			vIncl = Array("All")
            vExcl = Array("wlan", "pppoe", "attack", "set url protocol sc-cpa", "log session-init", "vpn", "ipsec", "ike", "pki", """Trust"" ""10.168" )
		    Call GetFileLineCountInclude("Y:\Basil_5\My_Papers\My_Network\SSG_Config_Nov_2015_edited.txt", vFileLines, vIncl, vExcl,0)
			Call WriteArrayToFile("Y:\Basil_5\My_Papers\My_Network\SSG_Config_Nov_2015_short.txt",vFileLines, UBound(vFileLines),1,0)
        Case 11
		    Set objFile = objFSO.OpenTextFile("Y:\Basil_5\My_Papers\My_Network\dhcp-bindings-srx.txt",ForWriting,True)
		    Call GetFileLineCountSelect("Y:\Basil_5\My_Papers\My_Network\dhcp-bindings.txt", vFileLines,"", "", "",0)
			For Each strLine in vFileLines
			   If Ubound(Split(strLine," "))>3 Then 
			       strMAC = Split(strLine," ")(8)
				   strIP = Split(strLine," ")(6)
				   objFile.WriteLine "set system services dhcp static-binding " & strMAC & " fixed-address " & strIP
			   End If 
			Next
			objFile.Close
			Set objFile = Nothing
 	    Case 10
	        strDirectoryWork =  objFSO.GetParentFolderName(Wscript.ScriptFullName)
		    strDirectoryUser = objEnvar.ExpandEnvironmentStrings("%USERPROFILE%")
	        strComputerName = objEnvar.ExpandEnvironmentStrings("%COMPUTERNAME%")
	        strSysFolder = Split(objEnvar.ExpandEnvironmentStrings("%PATH%"),";")(0)
			Call TrDebug( "ScriptName: "  & Split(Wscript.ScriptFullName,"\")(UBound(Split(Wscript.ScriptFullName,"\")))  , "", objDebug, MAX_LEN , 1, nInfo)			
			Call TrDebug( "strDirectoryWork: "  & strDirectoryWork  , "", objDebug, MAX_LEN , 1, nInfo)
			Call TrDebug( "strDirectoryUser: "  & strDirectoryUser  , "", objDebug, MAX_LEN , 1, nInfo)
			Call TrDebug( "strComputerName: "  & strComputerName  , "", objDebug, MAX_LEN , 1, nInfo)
			Call TrDebug( "strSysFolder: "  & strSysFolder  , "", objDebug, MAX_LEN , 1, nInfo)			
			
	   Case 9 
	        strComputer = "."
	        Dim vWinUsers
			Redim vWinUsers(1)
	        Call GetWinUsersWMI(strComputer, vWinUsers, 1)
			For each User in vWinUsers
			   Call TrDebug( User, "", objDebug, MAX_LEN , 1, nInfo)			
			Next 
			
       Case 8
    	   strComputer = "MEDIA" 
			Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2") 
			Set colAccounts = objWMIService.ExecQuery("SELECT * FROM Win32_UserAccount") 
			For Each objAccount in colAccounts
    			Call TrDebug( objAccount.Name , " LocalAccount : " & objAccount.LocalAccount & "  Status: " & objAccount.Status, objDebug, MAX_LEN , 1, nInfo)
			Next
'			MsgBox "Win32_LogonSession"
'			Set objWMIService = Nothing
'			Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2") 			
			Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_LogonSession WHERE LogonType='10'") 
			On Error Resume Next
			For Each objSession in colItems 
			    If Err.Number > 0 Then Call TrDebug( "ERROR 320: " & Err.Number & ": " & Err.Description, "", objDebug, MAX_LEN , 1, nInfo) : Err.Clear : End If
			    ' Call TrDebug( "Session ID: " & objSession.LogonID, " LogonType : " & objSession.LogonType, objDebug, MAX_LEN , 3, nInfo)
			    Set colList = objWMIService.ExecQuery("Associators of {Win32_LogonSession.LogonId=" & objSession.LogonId & "} Where AssocClass=Win32_LoggedOnUser Role=Dependent" )
			    If Err.Number > 0 Then Call TrDebug( "ERROR 323: " & Err.Number & ": " & Err.Description, "", objDebug, MAX_LEN , 1, nInfo) : Err.Clear : End If
				For Each objItem in colList
				   If Err.Number > 0 Then Call TrDebug( "ERROR 325: " & Err.Number & ": " & Err.Description, "", objDebug, MAX_LEN , 1, nInfo) : Err.Clear : End If
				   Call TrDebug( "Session ID: " & objSession.LogonID & " LogonType : " & objSession.LogonType &  "  Username: " & objItem.Name & " Domain: " & objItem.Domain, "", objDebug, MAX_LEN , 1, nInfo)
				   If Err.Number > 0 Then Call TrDebug( "ERROR 327: " & Err.Number & ": " & Err.Description, "", objDebug, MAX_LEN , 1, nInfo) : Err.Clear : End If
				  ' Call TrDebug( "Full Name: " & objItem.FullName, "", objDebug, MAX_LEN , 1, nInfo)
				   ' Call TrDebug( "Domain: " & objItem.Domain, "", objDebug, MAX_LEN , 1, nInfo)					   
				 Next
			Next
			On Error Goto 0
			Call TrDebug( "DONE", "", objDebug, MAX_LEN , 3, nInfo)
'	        MsgBox "Account"
'			set objItem = Nothing
'			Set colItems = Nothing
'			Set objWMIService = Nothing
'	     	strComputer = "." 
'			Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2") 
'			Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_Account") 
'			For Each objItem in colItems 
 '   			Call TrDebug( objItem.Name & " LocalAccount : " & objItem.LocalAccount & "  Status: " & objItem.Status, "", objDebug, MAX_LEN , 1, nInfo)
	'		Next
			  
       Case 7 
             Call GetWin32Screen(".", vScreen, 1) 
'	     	strComputer = "." 
'			Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2") 
'			Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_DesktopMonitor") 
'			For Each objItem in colItems 
'				Call TrDebug( objItem.DeviceID & " Hight : " & objItem.ScreenHeight & "  Width: " & objItem.ScreenWidth, "", objDebug, MAX_LEN , 1, nInfo)
'			Next
	    Case 6
            ServiceName = "ClientReportSvc"
			strComputer = "."
			Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
			Set colListOfServices = objWMIService.ExecQuery ("Select * from Win32_Service Where Name = '" & ServiceName & "'")
			Set colListOfServices = objWMIService.ExecQuery ("Select * from Win32_Service")
			For Each objService in colListOfServices
				Call TrDebug( "Service : " & objService.Name & " " & objService.Status & " " & objService.State & ", " & objService.StartName, "", objDebug, MAX_LEN , 1, nInfo)
			Next
			
			' If StartService(".", ServiceName, True, 1) Then MsgBox "Service " & ServiceName & "Successfully restarted"
		case 5
		    Dim cimv2
            ServiceName = "ClientReportSvc"
			strComputer = "."
     		Set cimv2 = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
            Set objService = cimv2.Get("Win32_Service.Name='" & ServiceName & "'")
			Call TrDebug( "Service : " & objService.Name & " " & objService.Status & " " & objService.State & ", " & objService.StartName, "", objDebug, MAX_LEN , 1, nInfo)

	    Case 1
		   
			strMsg = ""
			strComputer = "."
			Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
			Set IPConfigSet = objWMIService.ExecQuery("Select * from Win32_NetworkAdapterConfiguration Where IPEnabled = True")
			For Each IPConfig in IPConfigSet
				'If Not IsNull(IPConfig.IPAddress) Then
				'	If Not Instr(IPConfig.IPAddress, ":") > 0 Then
						Call TrDebug(IPConfig.Description & ": ", IPConfig.IPAddress(0), objDebug, MAX_LEN , 1, nInfo)
				'	End If
				'End If
			Next
		Case 2
			vModels = Array("acx5096","acx5048","acx1100","acx1000","acx2100","acx2200","mx80","mx104","mx240","mx480","mx960")
			MsgBox UBound(vModels)
		Case 3
		    Const NOTEPAD_PP = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\notepad++.exe\"
			On Error Resume Next
				Err.Clear
				strEditor = objShell.RegRead(NOTEPAD_PP)
				if Err.Number <> 0 Then 
					strDirectoryWork = "Not Set"
				End If
				MsgBox strEditor & chr(13) & "Error: " & Err.Description
				
			On Error Goto 0
		Case 4
		    Dim strPID, strAppName
			strAppName = "iexplore.exe"
		    Call GetAppPID(strPID, strAppName)
			MsgBox "PID =" & strPID
	End Select
End Sub
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


'----------------------------------------------------------------
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
'----------------------------------------------------------------
'   Function GetMyPID(strPID) Returns focus to the parent Window/Form
'----------------------------------------------------------------
Function GetAppPID(ByRef strPID, strAppName)
Dim objWMI, colItems
Const IE_PAUSE = 70
Dim process
Dim strUser, pUser, pDomain, wql
	strUser = GetScreenUserSYS()
	Do 
		On Error Resume Next
		Set objWMI = GetObject("winmgmts:\\127.0.0.1\root\cimv2")
		If Err.Number <> 0 Then 
				Call TrDebug ("GetMyPID ERROR: CAN'T CONNECT TO WMI PROCESS OF THE SERVER","",objDebug, MAX_LEN, 1, 1)
				On error Goto 0 
				Exit Do
		End If 
'		wql = "SELECT ProcessId FROM Win32_Process WHERE Name = 'Launcher Ver.'"  WHERE Name = 'iexplore.exe' OR Name = 'wscript.exe'
		wql = "SELECT * FROM Win32_Process WHERE Name = '" & strAppName & "' OR Name = '" & strAppName & " *32'"
		On Error Resume Next
		Set colItems = objWMI.ExecQuery(wql)
		If Err.Number <> 0 Then
				Call TrDebug ("GetMyPID ERROR: CAN'T READ QUERY FROM WMI PROCESS OF THE SERVER","",objDebug, MAX_LEN, 1, 1)
				On error Goto 0 
				Set colItems = Nothing
				Exit Do
		End If 
		On error Goto 0 
		For Each process In colItems
			process.GetOwner  pUser, pDomain 
			Call TrDebug ("GetMyPID: RESTORE IE WINDOW:", "PName: " & process.Name & ", PID " & process.ProcessId & ", OWNER: " & pUser & ", Parent PID: " &  Process.ParentProcessId,objDebug, MAX_LEN, 1, 1) 
			If pUser = strUser then 
				strPID = process.ProcessId
				MsgBox "I'm Here"
				Call TrDebug ("GetMyPID: ", "PName: " & process.Name & ", PID " & process.ProcessId & ", OWNER: " & pUser & ", Parent PID: " &  Process.ParentProcessId,objDebug, MAX_LEN, 1, 1) 
				Call TrDebug ("GetMyPID: ", "Caption: " & process.Caption & ", CSName " & process.CSName & ", Description: " & process.Description & ", Handle: " &  Process.Handle,objDebug, MAX_LEN, 1, 1) 
			GetMyPID = True
				Exit For
			End If
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


Function StartService(Computer, ServiceName, Wait, nDebug)
  Dim cimv2, oService, nResult, nTimeOut
  StartService = False
  Set cimv2 = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & Computer & "\root\cimv2")
  Set oService = cimv2.Get("Win32_Service.Name='" & ServiceName & "'")
    If oService.Started Then
		Call TrDebug( "Service : " & ServiceName & " is started. Will stop it now ", "", objDebug, MAX_LEN , 1, nDebug)
		'-----------------------------------
		'   STOP SERVICE FIRST
		'-----------------------------------
		nResult =  oService.StopService
		If nResult <> 0 Then 
			Call TrDebug( "StartService ERROR: Can't Stop service: " & ServiceName, "", objDebug, MAX_LEN , 1, 1)
			Exit Function
		End If
		nTimeOut = 0
		Do While InStr(Lcase(oService.State),"stopped") = 0 And Wait And nTimeOut < 5000
			Set oService = cimv2.Get("Win32_Service.Name='" & ServiceName & "'")
			Wscript.Sleep 200
			nTimeOut = nTimeOut + 200
		Loop
		If nTimeOut => 5000 Then 
			Call TrDebug( "StartService ERROR: Service stopping timeout (>5sec): " & ServiceName, "", objDebug, MAX_LEN , 1, 1)
			Exit Function
		End If 		
		Call TrDebug( "StartService: Service: " & ServiceName, "STOPPED", objDebug, MAX_LEN , 1, nDebug)
    End If
    'Start the service
    nResult = oService.StartService
	If nResult <> 0 Then 
		Call TrDebug( "StartService ERROR: Can't Start service: " & ServiceName, "", objDebug, MAX_LEN , 1, 1)
		Exit Function
	End If
	nTimeOut = 0
	Do While InStr(Lcase(oService.State),"running") = 0 And Wait And nTimeOut < 5000
		Set oService = cimv2.Get("Win32_Service.Name='" & ServiceName & "'")
		Wscript.Sleep 200
		nTimeOut = nTimeOut + 200
	Loop
	If nTimeOut => 5000 Then 
		Call TrDebug( "StartService ERROR: Service starting timeout (>5sec): " & ServiceName, "", objDebug, MAX_LEN , 1, 1)
		Exit Function
	End If 		
	Call TrDebug( "StartService: Service: " & ServiceName, "STARTED", objDebug, MAX_LEN , 1, nDebug)
    StartService = True
End Function
'-------------------------------------------------------
'  Function GetWin32Screen(strComputer, vScreen, nDebug) GetScreen Resolution from WMI 
'-------------------------------------------------------
Function GetWin32Screen(strComputer, vScreen, nDebug)
Dim objWMIService, colItems, objItem
GetWin32Screen = False
Redim vScreen(2)
	Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2") 
	Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_DesktopMonitor Where DeviceID = ""DesktopMonitor1""") 
	On Error Resume Next
	For Each objItem in colItems
		If Err.Number <> 0 Then Call TrDebug( "Screen Resolution Error ", "", objDebug, MAX_LEN , 1, nDebug) : Exit For : End If	
		Call TrDebug( "Screen Resolution: " & objItem.DeviceID & " Hight : " & objItem.ScreenHeight & "  Width: " & objItem.ScreenWidth, "", objDebug, MAX_LEN , 1, nDebug)
		vScreen(0) = objItem.ScreenWidth
		vScreen(1) = objItem.ScreenHeight
		GetWin32Screen = True
	Next
	On Error Goto 0
End Function
 
'#######################################################################
 ' Function WriteArrayToFile - Returns number of lines int the text file
 ' nMode = 1  Then Rewire all File content
 ' nMode = 2  Then Append
 ' Creates File if it doesn't exists
 '#######################################################################
 Function WriteArrayToFile(strFile,vFileLine, nFileLine,nMode,nDebug)
    Dim i, nCount
	Dim strLine
	Dim objDataFileName, objFSO
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	If Not objFSO.FileExists(strFile) Then 	
		On Error Resume Next
		Err.Clear
		Set objDataFileName = objFSO.CreateTextFile(strFile)
		If Err.Number = 0 Then 
			objDataFileName.close
			On Error Goto 0
		Else
			Set objFSO = Nothing
			If IsObject(objDebug) Then 
				objDebug.WriteLine GetMyDate() & " " & FormatDateTime(Time(), 3) & ": WriteArrayToFile: ERROR: CAN'T CREATE FILE " & strFile
				objDebug.WriteLine GetMyDate() & " " & FormatDateTime(Time(), 3) & ": WriteArrayToFile:  Error: " & Err.Number & " Srce: " & Err.Source & " Desc: " &  Err.Description
			End If
			WriteArrayToFile = False
			On Error Goto 0
			Exit Function
		End If
	End If
	
	Select Case nMode
		Case 1 
			Set objDataFileName = objFSO.OpenTextFile(strFile,2,True)
		Case 2 	
			Set objDataFileName = objFSO.OpenTextFile(strFile,8,True)
	End Select 

	i = 0
	On Error Resume Next
	Err.Clear
	Do While i < nFileLine
		objDataFileName.WriteLine vFileLine(i)
		If Err.Number <> 0 Then 
			If IsObject(objDebug) Then 
				objDebug.WriteLine GetMyDate() & " " & FormatDateTime(Time(), 3) & ": WriteArrayToFile: ERROR: CAN'T WRITE TO FILE " & strFile
				objDebug.WriteLine GetMyDate() & " " & FormatDateTime(Time(), 3) & ": WriteArrayToFile:  Error: " & Err.Number & " Srce: " & Err.Source & " Desc: " &  Err.Description
			End If
			WriteArrayToFile = False
			Exit Do 			
		End If
		i = i + 1
	Loop
	On Error Goto 0
	If i = nFileLine Then WriteArrayToFile = True End If
	objDataFileName.close
	Set objFSO = Nothing
End Function
'#######################################################################
 ' Function GetFileLineCountInclude - Returns number of lines int the text file
 '#######################################################################
 Function GetFileLineCountInclude(strFileName, ByRef vFileLines,vIncl, vExcl,nDebug)
    Dim nIndex
	Dim strLine, nCount, nSize
	Dim objDataFileName, nResult, nSymbol
	
    strFileWeekStream = ""	
	Set objDataFileName = objFSO.OpenTextFile(strFileName)
	Redim vFileLines(0)
	Set objDataFileName = objFSO.OpenTextFile(strFileName)	
	nIndex = 0
    Do While objDataFileName.AtEndOfStream <> True
	    ' Check if string contains symbol or substric to exclude from coping to array
		nResult = True
		strLine = LTrim(objDataFileName.ReadLine)
        ' Additional check of the fully commented lines
        if InStr(strLine,"'") Then 
		    nSymbol = 1
			Do 
			    strChar = Mid(strLine,nSymbol,1)
			    Select Case strChar
				    Case "'"
						nResult = False
						Exit Do
					Case chr(9)
					    nSymbol = nSymbol + 1
					Case Else 
						nResult = True
						Exit Do
			    End Select
           Loop			
        End If 		
        ' Check if string contains symbol or substring which must be excluded 
	    If nResult = True and  vExcl(0) <> "Null" Then 
			nExcl = 0
		    Do while nExcl <= UBound(vExcl)
		        If InStr(strLine, vExcl(nExcl)) <> 0 Then 
				    nResult = False
					Exit Do
				End If 
				nExcl = nExcl + 1
		    Loop
		End If
        ' Check if string contains symbol or substring which must be included 
    	If nResult = True and vIncl(0) <> "All" Then 
			nResult = false
			nIncl = 0
			Do while nIncl <= UBound(vIncl)
				If InStr(strLine, vIncl(nIncl)) <> 0 Then 
					
					nResult = True
					Exit Do
				End If 
				nIncl = nIncl + 1
			Loop
		End If 
        If nResult Then 
			Redim Preserve vFileLines(nIndex + 1)
			vFileLines(nIndex) = strLine
			If nDebug = 1 Then objDebug.WriteLine "GetFileLineCountSelect: vFileLines(" & nIndex & ")="  & vFileLines(nIndex) End If  
			nIndex = nIndex + 1
		End if
	Loop
	objDataFileName.Close
    GetFileLineCountInclude = nIndex
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
'------------------------------------------------------------------------------------------------------------------
' Function returns the number of the line from 1 to N which contains string strObject. Returns 0 if nothing found
'------------------------------------------------------------------------------------------------------------------
Function GetObjectLineNumber( byRef vArray, nArrayLen, strObjectName, bCaseSensitive)
 Dim nInd, strLine, strPattern
	nInd = 0
	GetObjectLineNumber = 0
	Do While nInd < nArrayLen
		Select Case bCaseSensitive
		    Case True
		        strLine =  vArray(nInd)
                strPattern = strObjectName				
		    Case False
   		        strLine =  LCase(vArray(nInd))
                strPattern = LCase(strObjectName)
		End Select
		If InStr(strLine, strPattern) <> 0	Then 
			GetObjectLineNumber = nInd + 1
			Exit Do
		End If
		nInd = nInd + 1
    Loop
End Function
'--------------------------------------------------------------------
' Function Runs MS CMD Command on local or remote PC
'--------------------------------------------------------------------
Function RunCmd(strHost, strPsExeFolder, ByRef vCmdOut, strCMD, strAny, nDebug)	
	Dim nResult
	Dim nCmd, stdOutFile, objCmdFile, cmdFile, f_objShell
	Set f_objFSO = CreateObject("Scripting.FileSystemObject")
	Set f_objShell = WScript.CreateObject("WScript.Shell")
	strRnd = My_Random(1,999999)
	stdOutFile = "svc-" & strRnd & ".dat"
	cmdFile = "run-" & strRnd & ".bat"
    strWork = f_objShell.ExpandEnvironmentStrings("%USERPROFILE%")
	If strHost = f_objShell.ExpandEnvironmentStrings("%COMPUTERNAME%") or strHost = "127.0.0.1" Then 
		strPsExec = ""
	Else 
		strPsExec = strPsExeFolder & "\psexec \\" & strHost & " -s "
	End If
	'-------------------------------------------------------------------
	'       CREATE A NEW TERMINAL SESSION IF REQUIRED
	'-------------------------------------------------------------------
	Set objCmdFile = objFSO.OpenTextFile(strWork & "\" & cmdFile,ForWriting,True)
	Call TrDebug ("COMMAND: ", strPsExec & strCMD & " >" & strWork & "\" & stdOutFile, objDebug, MAX_LEN, 1, nDebug)
	objCmdFile.WriteLine strPsExec & strCMD & " >" & strWork & "\" & stdOutFile
	objCmdFile.WriteLine "Exit"
	objCmdFile.close
	f_objShell.run strWork & "\" & cmdFile,0,True
	Call TrDebug ("BATCH FILE EXECUTED: ", strWork & "\" & cmdFile, objDebug, MAX_LEN, 1, nDebug)
	wscript.sleep 100
	'-----------------------------------------
	' READ OUTPUT FILE AND DELETE WHEN DONE
	'-----------------------------------------
	RunCmd = GetFileLineCountSelect(strWork & "\" & stdOutFile, vCmdOut,"NULL","NULL","NULL",0)
	If f_objFSO.FileExists(strWork & "\" & stdOutFile) Then
		On Error Resume Next
		Err.Clear
		f_objFSO.DeleteFile strWork & "\" & stdOutFile, True
 		If Err.Number <> 0 Then 
			Call TrDebug ("RunCmd: ERROR CAN'T DELET FILE:",stdOutFile, objDebug, MAX_LEN, 1, 1)
			On Error goto 0
		End If	
	End If
	If f_objFSO.FileExists(strWork & "\" & cmdFile) Then 
		On Error Resume Next
		Err.Clear
		f_objFSO.DeleteFile strWork & "\" & cmdFile, True
 		If Err.Number <> 0 Then 
			Call TrDebug ("RunCmd: ERROR CAN'T DELET FILE:",cmdFile, objDebug, MAX_LEN, 1, 1)
			On Error goto 0
		End If		
	End If
	Set f_objFSO = Nothing
	Set f_objShell = Nothing
	If RunCmd = 0 Then 
		Call TrDebug ("RunCmd: " & strCMD & " ERROR: ", "CAN'T WRITE TO OUTPUT FILE OR EMPTY FILE" , objDebug, MAX_LEN, 1, 1)
		Exit Function 
	End If
End Function
'--------------------------------------------------------------
' Function returns a random intiger between min and max
'--------------------------------------------------------------
Function My_Random(min, max)
	Randomize
	My_Random = (Int((max-min+1)*Rnd+min))
End Function
'--------------------------------------------------------------
' Function WinSvcExists(ServiceName)
'--------------------------------------------------------------
Function WinSvcExists(ServiceName)
  Dim objWMIService, colListOfServices, objService, bResult
  bResult = False
	strComputer = "."
	Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	Set colListOfServices = objWMIService.ExecQuery ("Select * from Win32_Service Where Name = """ & ServiceName & """")
	
	For Each objService in colListOfServices
		If Not IsNull(objService.State) Then 
			Call TrDebug( "Service : " & objService.Name, objService.Status & ", " & objService.State, objDebug, MAX_LEN , 1, nInfo)
			WinSvcExists = objService.State
            bResult = True			
		End If	
	Next
	If Not bResult Then  Call TrDebug( "SERVICE " & ServiceName,"DOESN'T EXIST", objDebug, MAX_LEN , 1, nInfo) : WinSvcExists = False
End Function
'--------------------------------------------------------------
' Function GetWinUsersWMI(strComputer)
'--------------------------------------------------------------
Function GetWinUsersWMI(strComputer, vWinUsers, nDebug)
Dim objWMIService, colAccounts, i   
   ' strComputer = "MEDIA" 
	Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2") 
	Set colAccounts = objWMIService.ExecQuery("SELECT * FROM Win32_UserAccount") 
	i = 1
	For Each objAccount in colAccounts
	    Redim Preserve vWinUsers(i)
		vWinUsers(i-1) = objAccount.Name & ", " & objAccount.SID 
		Call TrDebug( objAccount.Name , " LocalAccount : " & objAccount.AccountType & "  Status: " & objAccount.Status, objDebug, MAX_LEN , 1, nDebug)
		i = i + 1
	Next
End Function
'-----------------------------------------------------------------------
'   WinGroupExists(strComputer, GroupName)
'-----------------------------------------------------------------------
Function WinGroupExists(strComputer, GroupName, byRef DomName, nDebug)
  Dim objWMIGroup, colListOfGroups, objGroup, bResult
  bResult = False
	Set objWMIGroup = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	Set colListOfGroups = objWMIGroup.ExecQuery ("Select * from Win32_Group Where Name = """ & GroupName & """")
	For Each objGroup in colListOfGroups
		If Not IsNull(objGroup.Name) Then 
			Call TrDebug( "GROUP : " & objGroup.Name, objGroup.Domain, objDebug, MAX_LEN , 1, nDebug)
			WinGroupExists = True
			DomName = objGroup.Domain
            bResult = True			
		End If	
	Next
	If Not bResult Then  Call TrDebug( "GROUP " & GroupName,"DOESN'T EXIST", objDebug, MAX_LEN , 1, 1) : WinGroupExists = False
End Function
'-----------------------------------------------------------------------
'   Function AddWinUserToGroup(strComputer,GroupName,UserName, nDebug)
'-----------------------------------------------------------------------
Function AddWinUserToGroup(strComputer,GroupName,UserName, nDebug)
Dim DomName, objDom, objGr, objLocalGroup
    AddWinUserToGroup = False
	If Not WinGroupExists(strComputer, GroupName, DomName, 1) Then
		'----------------------------------------------
		'   CREATE GROUP
		'----------------------------------------------
'		Set objEnv = objShell.Environment("Process")
		Set objDom = Getobject("WinNT://" & strComputer & ",computer" )
		On Error resume next
			Err.Clear
			Set objGr = objDom.Create("group",GroupName)
			objGr.SetInfo
		On Error Goto 0
		If Not WinGroupExists(strComputer, GroupName, DomName, 1) Then 
			Call TrDebug( "CREATING GROUP " & GroupName , "ERROR", objDebug, MAX_LEN , 1, 1)
			Call TrDebug( "Error Code " & Err.Number, Err.Description, objDebug, MAX_LEN , 1, 1)			
		    Exit Function 
		Else 
		    Call TrDebug( "GROUP " & GroupName & " WAS CREATED", "", objDebug, MAX_LEN , 1, 1)
		End If
	Else 
		Call TrDebug( "GROUP " & GroupName & " ALREADY EXISTS", "", objDebug, MAX_LEN , 1, nDebug)
	End If
	Set objLocalGroup = GetObject("WinNT://" & strComputer & "/" & GroupName & ",group")
	Call TrDebug( "CHECK IF USER IS A MEMBER OF THE GROUP", "", objDebug, MAX_LEN , 1, nDebug)			
	Set objLocalGroup = GetObject("WinNT://" & strComputer & "/" & GroupName & ",group")
	' Set objDomainUser = GetObject("WinNT://" & DomName & "/" & UserName & ",user")
	If Not objLocalGroup.IsMember("WinNT://" & DomName & "/" & UserName) Then
		Call TrDebug( "ADDING USER TO GROUP", "", objDebug, MAX_LEN , 1, nDebug)			
		On Error resume next
			Err.Clear
			objLocalGroup.Add("WinNT://" & DomName & "/" & UserName)
		On Error Goto 0
		If Not objLocalGroup.IsMember("WinNT://" & DomName & "/" & UserName) Then 
    		AddWinUserToGroup = False 
			Call TrDebug( "ADDING USER " & UserName & " TO GROUP " & GroupName, "ERROR", objDebug, MAX_LEN , 1, 1)
			Call TrDebug( "Error Code " & Err.Number, Err.Description, objDebug, MAX_LEN , 1, 1)
		Else 
		    AddWinUserToGroup = True
		    Call TrDebug( "USER " & UserName & " ADDED TO GROUP " & GroupName, "", objDebug, MAX_LEN , 1, 1)			
		End If
	Else 
		Call TrDebug( "USER " & DomName & "/" & UserName & " IS ALREADY A MEMBER OF THE GROUP " & GroupName, "", objDebug, MAX_LEN , 1, nDebug)
		AddWinUserToGroup = True
	End If
End Function
' --------------------------------------------------
'    Function SetCmdPermission(strHost, strPsExeFolder, GroupName, strFolder, strPermissions, nDebug)	
' --------------------------------------------------
Function SetCmdPermission(strHost, strPsExeFolder, GroupName, strFolder, strPermissions, nDebug)	
	Dim strCmd
	Dim vResult
	'-----------
	strCmd = strPsExec & "icacls " & strFolder & " /grant " & GroupName & ":" & strPermissions
    If RunCmd(strHost, strPsExeFolder, vResult, strCmd, strAny, nDebug) = 0 Then SetCmdPermission = False : Exit Function : End If 
	'------------------------------------------
	' CHECK FOR RESULT
	'------------------------------------------
	SetCmdPermission = False
	If GetObjectLineNumber(vResult, UBound(vResult), "Successfully processed 0", False) > 0 Then 
		Call TrDebug ("SetCmdPermission: ERROR: SOMETHING WENT WRONG", GroupName & ": " & StrFolder, objDebug, MAX_LEN, 1, nDebug)
		SetCmdPermission = False
	End If
	If GetObjectLineNumber(vResult, UBound(vResult), "Failed processing 0", False) > 0 Then 
		Call TrDebug ("SetCmdPermission:", "PERMISSIONS ASSIGNED" , objDebug, MAX_LEN, 1, nDebug)
		SetCmdPermission = True
	End If
End Function
'--------------------------------------------------------------
' Function GetDefaultGwWMI(strComputer,nDebug)
'--------------------------------------------------------------
Function GetDefaultGwWMI(strComputer, nDebug)
Dim objWMIService, IPConfigSet, IPConfig, nGw, strGw , nMetric
    nMetric = 1000
	Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
	Set IPConfigSet = objWMIService.ExecQuery("Select * from Win32_NetworkAdapterConfiguration Where IPEnabled = True")
	On Error resume next
	For Each IPConfig in IPConfigSet
		' strLine = strLine &	"<option value=" & IPConfig.IPAddress(0) & ">" & IPConfig.Description & "</option>"
		nGw = 0
        For each strGw in IPConfig.DefaultIPGateway
     		Call TrDebug ("Default Gateway:" & strGw & " Metrc: " & IPConfig.GatewayCostMetric(nGw), "" , objDebug, MAX_LEN, 1, nDebug)
			If Err.Number > 0 Then 
			    GetDefaultGwWMI = 0 
				Call TrDebug ("GetDefaultGwWMI: CAN'T GET Default Gateway. Error: " & Err.Description, "ERROR" , objDebug, MAX_LEN, 1, 1)
				Exit For 
			End If
			If IPConfig.GatewayCostMetric(nGw) < nMetric Then 
			   GetDefaultGwWMI = strGw
			   nMetric = IPConfig.GatewayCostMetric(nGw)
			End If 
			nGw = nGw + 1
		Next
	    nAdapter = nAdapter + 1	
	Next
    On Error goto 0
    Call TrDebug ("BEST GATEWAY:" & GetDefaultGwWMI & " Best Metric: " & nMetric, "" , objDebug, MAX_LEN, 1, nDebug)
End Function 
'--------------------------------------------------------------
'   Function GetWinLogonWMI(strComputer, vWinUsers, nDebug)
'--------------------------------------------------------------
Function GetWinLogonWMI(strComputer, ByRef vWinUsers, nDebug)
Dim objWMIService, i  , strWinUser , strLogonName, objItem, colItems
	Redim vWinUsers(1) 
	vWinUsers(0) = "NOT_FOUND"
	On Error Resume Next
	Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2") 
	If Err.Number > 0 Then 
	    Call TrDebug( "ERROR 1: " & Err.Number & ": " & Err.Description, "", objDebug, MAX_LEN , 1, 1) 
		Err.Clear 
    	GetWinLogonWMI = UBound(vWinUsers)
        Exit Function		
	End If
	Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_LogonSession WHERE LogonType='2'") 
	i = 0
	strWinUser = ","
	For Each objSession in colItems
		Set colList = objWMIService.ExecQuery("Associators of {Win32_LogonSession.LogonId=" & objSession.LogonId & "} Where AssocClass=Win32_LoggedOnUser Role=Dependent" )
		For Each objItem in colList
		    strLogonName = objItem.Name
			Select Case Err.Number
			    Case 424
				     Err.Clear
					 Exit For
				Case 0
				Case Else
				   Call TrDebug( "ERROR 2: " & Err.Number & ": " & Err.Description, "", objDebug, MAX_LEN , 1, 1) 
				   Err.Clear 
				   Exit For
			End Select
		    If InStr(strWinUser,"," & strLogonName & ",") = 0 Then 
		        strWinUser = strWinUser & objItem.Name & ","
				Redim Preserve vWinUsers(i + 1)
				vWinUsers(i) = objItem.Name
				i = i + 1
			End If
		   ' Call TrDebug( "Session ID: " & objSession.LogonID & " LogonType : " & objSession.LogonType &  "  Username: " & objItem.Name & " Domain: " & objItem.Domain, "", objDebug, MAX_LEN , 1, nDebug)
		Next
		If Err.Number > 0 Then Call TrDebug( "ERROR 3: " & Err.Number & ": " & Err.Description, "", objDebug, MAX_LEN , 1, 1) : Err.Clear : Exit For : End If
	Next
	On Error Goto 0
	GetWinLogonWMI = UBound(vWinUsers)
	Set objWMIService = Nothing
End Function
'------------------------------------------------------------------------
'	Function CopyArray(ByRef Array1, ByRef Array2)
'------------------------------------------------------------------------
Function CopyArray(ByRef Array1, ByRef Array2, nDebug)
Dim vDim(3), nDim
	nDim = 0
	vDim(0) = 0 : vDim(1) = 0 : vDim(2) = 0 
	On Error resume next
    Err.Clear
	Do While nDim < 3
	    vDim(nDim) = UBound(Array1,nDim + 1)
	    If Err.Number <> 0 Then Exit Do
		nDim = nDim + 1
	Loop
	Call TrDebug("nDim = " & nDim & " Error: " & Err.Description, "", objDebug, MAX_LEN , 1, nDebug)
	' nDim equals to number Array1 dimentions: 1, 2 or 3
	On Error goto 0
	CopyArray = nDim
    Select Case nDim
        Case 1
		    Redim Array2(UBound(Array1,1))
			For i = 0 to UBound(Array1,1) - 1 
			    Array2(i) = Array1(i)
				Call TrDebug("Array(" & i & ") = " & Array2(i), "", objDebug, MAX_LEN , 1, nDebug)
			Next
        Case 2
		    Redim Array2(UBound(Array1,1),UBound(Array1,2))
			For i = 0 to UBound(Array1,1) - 1 
			    For n = 0 to UBound(Array1,2) - 1 
					Array2(i,n) = Array1(i,n)
					Call TrDebug("Array(" & i & "," & n & ") = " & Array2(i,n), "", objDebug, MAX_LEN , 1, nDebug)
				Next
			Next
			
        Case 3		
	        Redim Array2(UBound(Array1,1),UBound(Array1,2),UBound(Array1,3))
			For i = 0 to UBound(Array1,1) - 1 
			    For n = 0 to UBound(Array1,2) - 1 
					For k = 0 to UBound(Array1,3) - 1 
						Array2(i,n,k) = Array1(i,n,k)
						Call TrDebug("Array(" & i & "," & n & "," & k & ") = " & Array2(i,n,k), "", objDebug, MAX_LEN , 1, nDebug)
					Next
				Next
			Next
	End Select 
End Function
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
		MsgBox "I'm Here"
		For Each process In colItems
			process.GetOwner  pUser, pDomain 
			Call TrDebug ("KillWinAppPID: Process Name (PID): " & process.Name & " (" & process.ProcessId & ")", "",objDebug, MAX_LEN, 1, nDebug)
			' Call TrDebug ("KillWinAppPID: Owner: " & process.CSName & "/" & pUser, "",objDebug, MAX_LEN, 1, nDebug) 
			Select Case nMode
			    Case 0
					If pUser = strUser then 
						Call TrDebug ("KillWinAppPID (0): Terminating the Process: Desktop user owns the process: " & strPID , "",objDebug, MAX_LEN, 1, nDebug)
						process.Terminate()
						KillWinAppPID = True
						Exit For
					End If
			    Case 1
					If Cint(strPID) = CInt(process.ProcessId) then 
						Call TrDebug ("KillWinAppPID (1): Terminating the Process: Desktop user owns the process: " & strPID , "",objDebug, MAX_LEN, 1, nDebug)
						process.Terminate()
						KillWinAppPID = True
						Exit For
					End If				
			    Case 2
					If pUser = strUser and InStr(process.CommandLine,strCommandLine) then 
						Call TrDebug ("KillWinAppPID (2): Terminating the Process: Desktop user owns the process: " & strPID, "",objDebug, MAX_LEN, 1, nDebug)
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
'----------------------------------------------------------------
'   Function GetBrowserOpenUrl(strPID) Returns focus to the parent Window/Form
'----------------------------------------------------------------
Function GetBrowserOpenUrl(ByRef strPID, ByRef strParentPID, strCommandLine, strAppName, nDebug)
Dim objWMI, colItems
Dim process
Dim strUser, pUser, pDomain, wql
	strUser = GetScreenUserSYS()
	GetBrowserOpenUrl = False
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
			Call TrDebug ("Process Name: " & process.Name & " (" & process.ProcessId & ")", "",objDebug, MAX_LEN, 1, nDebug)
			Call TrDebug ("Owner:        " & process.CSName & "/" & pUser, "",objDebug, MAX_LEN, 1, nDebug) 
			Call TrDebug ("ParentPID:    " & process.ParentProcessId, "",objDebug, MAX_LEN, 1, nDebug)
            Call TrDebug ("Description:  " & process.Description, "",objDebug, MAX_LEN, 1, nDebug)			
            Call TrDebug ("Handle:       " & process.Handle, "",objDebug, MAX_LEN, 1, nDebug)						
            Call TrDebug ("SessionId:    " & process.MainWindowHandle, "",objDebug, MAX_LEN, 1, nDebug)									
'			Call TrDebug ("Command Line:          " , "",objDebug, MAX_LEN, 1, nDebug) 			
'			vCMD = Split(process.CommandLine," --")
'			For Each strLine in vCMD
'			   If strLine = "" Then Exit For
'			   Call TrDebug ("  --" & strLine, "",objDebug, MAX_LEN, 1, nDebug) 
'			Next 
			Call TrDebug ("    " , "",objDebug, MAX_LEN, 1, nDebug) 						
			Select Case Lcase(strCommandLine)
			    Case "null", "none", ""
					If pUser = strUser then 
						strPID = process.ProcessId
						strParentPID = Process.ParentProcessId
						Call TrDebug ("GetBrowserOpenUrl: Process ID: " & strPID , "",objDebug, MAX_LEN, 1, nDebug)
						GetBrowserOpenUrl = True
					End If
			    Case Else
					If pUser = strUser and InStr(process.CommandLine,strCommandLine) then 
						strPID = process.ProcessId
						strParentPID = Process.ParentProcessId
						Call TrDebug ("GetBrowserOpenUrl: Process ID: " & strPID, "",objDebug, MAX_LEN, 1, nDebug)
						GetBrowserOpenUrl = True
					End If
			End Select
		Next
		Set colItems = Nothing
		Exit Do
	Loop
	Set objWMI = Nothing
End Function
'----------------------------------------------------------------
'   Function GetBrowserOpenUrl(strPID) Returns focus to the parent Window/Form
'----------------------------------------------------------------
Function DownloadingFile(URL, strDownloadFolder, SaveAs)
Dim objFSO,objXMLHTTP,Tab,strHDLocation,objADOStream,Command,Start,File
Dim MsgTitre,MsgAttente,StartTime,DurationTime,ProtocoleHTTP
	Set objFSO = Createobject("Scripting.FileSystemObject")
    ProtocoleHTTP = "https://"
	If URL = "" Then 
	   DownloadingFile = False
	   Exit Function
	End If
	If Left(URL,8) <> ProtocoleHTTP Then
		URL = ProtocoleHTTP & URL
		MsgBox "Source :" & chr(13) & URL,64
	End if
	If SaveAs = "" Then
		Tab = split(url,"/")
		File =  Tab(UBound(Tab))
		File = Replace(File,"%20"," ")
		File = Replace(File,"%28","(")
		File = Replace(File,"%29",")")
		SaveAs = File
	End If
    Set objXMLHTTP = CreateObject("MSXML2.ServerXMLHTTP.3.0")
    strHDLocation = strDownloadFolder & "\" & SaveAs
    'msgbox strHDLocation
	StartTime = Timer
    On Error Resume Next
	objXMLHTTP.SetOption(2) = (objXMLHTTP.GetOption(2) - SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS)
    objXMLHTTP.open "GET",URL,false,"vmukhin",""
    objXMLHTTP.send()
	If Err.number <> 0 Then
	   MsgBox err.description,16,err.description
	   DownloadingFile = False
	   Exit Function
	Else
		If objXMLHTTP.Status = 200 Then
			strHDLocation = strDownloadFolder & "\" & SaveAs
			Set objADOStream = CreateObject("ADODB.Stream")
			objADOStream.Open
			objADOStream.Type = 1 'adTypeBinary
			objADOStream.Write objXMLHTTP.ResponseBody
			objADOStream.Position = 0    'Set the stream position to the start
			If objFSO.Fileexists(strHDLocation) Then objFSO.DeleteFile strHDLocation
			Msgbox  "Destination Path : " & Dblquote(strHDLocation),64
			objADOStream.SaveToFile strHDLocation
			objADOStream.Close
		    Set objADOStream = Nothing
		End If
	End if
	Set objXMLHTTP = Nothing
	Set objFSO = Nothing
'   Voice.Speak "The Download of " & Dblquote(File) & " is finished in " & DurationTime &" !"
    DurationTime = FormatNumber(Timer - StartTime, 0) & " seconds."
    MsgBox "The Download of " & Dblquote(SaveAs) & " is finished in " & DurationTime &" !",64,"The Download of " & Dblquote(SaveAs) & " is finished in " & DurationTime &" !"
End Function

'Function to add double quotes in a variable
Function DblQuote(Str)
    DblQuote = Chr(34) & Str & Chr(34)
End Function
'----------------------------------------------------------
'   Function Set_IE_obj (byRef objIE)
'----------------------------------------------------------
Function Set_IE_obj (byRef objIE)
	Dim nCount
	Set_IE_obj = False
	nCount = 0
	Do 
		On Error Resume Next
		Err.Clear
		Set objIE = CreateObject("InternetExplorer.Application")
		Select Case Err.Number
			Case &H800704A6 
				wscript.sleep 1000
				nCount = nCount + 1
				Call  TrDebug ("Set_IE_obj ERROR:" & Err.Number & " " & Err.Description, "", objDebug, MAX_LEN, 1, 1)
				If nCount > 4 Then
					On Error goto 0
					Exit Function
				End If
			Case 0 
				Set_IE_obj = True
				On Error goto 0
				Exit Function
			Case Else 
				Call  TrDebug ("Set_IE_obj ERROR:" & Err.Number & " " & Err.Description, "", objDebug, MAX_LEN, 1, 1)
				On Error goto 0
				Exit Function
		End Select
	On Error goto 0
	Loop
End Function
'---------------------------------------------------------------
'	Function  ClearTitle(byRef strTitle)
'---------------------------------------------------------------
Function  ClearTitle(byRef strTitle)
Dim vTmp
     vTmp = Split(strTitle,"-")
	 strTitle = ""
	 For i = 0 to UBound(vTmp)
	    If InStr(LCase(vTmp(i)),"mozilla firefox")=0 and InStr(LCase(vTmp(i)),"internet explorer")=0 and InStr(LCase(vTmp(i)),"google chrome")=0 and InStr(LCase(vTmp(i)),"youtube")=0 Then  
		    strTitle = strTitle & vTmp(i) & " "
		End If     
	 Next
End Function
'-------------------------------------------------------------------------
' Function AppendStringToFile - Returns number of lines int the text file
'-------------------------------------------------------------------------
 Function WriteStringToFile(strFile,strLine, bAppend,nDebug)
	Dim objDataFileName, objFSO	
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	If Not objFSO.FileExists(strFile) Then 	
		On Error Resume Next
		Err.Clear
		Set objDataFileName = objFSO.CreateTextFile(strFile)
		If Err.Number = 0 Then 
			objDataFileName.close
			On Error Goto 0
		Else
			Set objFSO = Nothing
			If IsObject(objDebug) Then 
				objDebug.WriteLine GetMyDate() & " " & FormatDateTime(Time(), 3) & ": AppendStringToFile: ERROR: CAN'T CREATE FILE " & strFile
				objDebug.WriteLine GetMyDate() & " " & FormatDateTime(Time(), 3) & ": AppendStringToFile:  Error: " & Err.Number & " Srce: " & Err.Source & " Desc: " &  Err.Description
			End If
			AppendStringToFile = False
			On Error Goto 0
			Exit Function
		End If
	End If
	If bAppend Then 
	    Set objDataFileName = objFSO.OpenTextFile(strFile,8,True)
	Else 
	    Set objDataFileName = objFSO.OpenTextFile(strFile,2,True)
	End If
	objDataFileName.WriteLine strLine
	objDataFileName.close
	Set objFSO = Nothing
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
	If bShowLog Then 
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
'----------------------------------------------------------------
'   Function GetWinAllPID(strPID) Returns focus to the parent Window/Form
'----------------------------------------------------------------
Function GetAllAppPID(ByRef strPID, ByRef strParentPID, strCommandLine, strAppName, nDebug)
Dim objWMI, colItems
Dim process
Dim strUser, pUser, pDomain, wql
	strUser = GetScreenUserSYS()
	GetWinAllPID = False
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
			Call TrDebug ("GetWinAllPID: Process Name (PID): " & process.Name & " (" & process.ProcessId & ")", "",objDebug, MAX_LEN, 1, nDebug)
			Call TrDebug ("GetWinAllPID: " & process.Caption , "",objDebug, MAX_LEN, 1, nDebug)
			'Call TrDebug ("GetWinAllPID: Owner: " & process.CSName & "/" & pUser, "",objDebug, MAX_LEN, 1, nDebug) 
			'Call TrDebug ("GetWinAllPID: CMD: " & process.CommandLine, "",objDebug, MAX_LEN, 1, nDebug) 
			'Call TrDebug ("GetWinAllPID: ParentPID:" &  Process.ParentProcessId, "",objDebug, MAX_LEN, 1, nDebug) 			
			'Select Case Lcase(strCommandLine)
			'    Case "null", "none", ""
			'		If pUser = strUser then 
			'			strPID = process.ProcessId
			'			strParentPID = Process.ParentProcessId
			'			Call TrDebug ("GetWinAllPID: Process is already running. Desktop user owns the process: " & strPID , "",objDebug, MAX_LEN, 1, nDebug)
			'			GetWinAllPID = True
			'			Exit For
			'		End If
			'   Case Else
			'		If pUser = strUser and InStr(process.CommandLine,strCommandLine) then 
			'			strPID = process.ProcessId
			'			strParentPID = Process.ParentProcessId
			'			Call TrDebug ("GetWinAllPID: Process is already running. Desktop user owns the process: " & strPID, "",objDebug, MAX_LEN, 1, nDebug)
			'			GetWinAllPID = True
			'			Exit For
			'		End If
			'End Select
		Next
		Set colItems = Nothing
		Exit Do
	Loop
	Set objWMI = Nothing
End Function
'--------------------------------------------------
'   Function GetInetApplication()
'--------------------------------------------------
Function GetInetApplication(strLine,vApplication, vPattern)
Dim objRegEx, Pattern, nIndex
	Set objRegEx = CreateObject("VBScript.RegExp")
	objRegEx.Global = True
	nIndex = 0
    GetInetApplication = "Inet"
	For each Pattern in vPattern
		If Pattern = "" Then exit for
		objRegEx.Pattern = Pattern
		If objRegEx.Test((LCase(strLine))) Then
		    GetInetApplication = vApplication(nIndex)
			Set objRegEx = Nothing
			Exit Function
		End If
		nIndex = nIndex + 1
	Next
	Set objRegEx = Nothing
End Function
'------------------------------------------------------------------------
'	Function CopyFileToString(strSourceFile)
'------------------------------------------------------------------------
Function CopyFileToString(strSourceFile)
Dim strFileString
Dim objFSO,objSourceFile
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Const ForAppending = 8
	Const ForWriting = 2
	Const ForReading = 1
	If objFSO.FileExists(strSourceFile) Then 	
		On Error Resume Next
			Err.Clear
			Set objSourceFile = objFSO.OpenTextFile(strSourceFile,ForReading,True)
			Select Case Err.Number
				Case 0 ' Do Nothing
				Case Else 
					Err.Clear
					CopyFileToString = "-1"
					Set objFSO = Nothing
					On Error Goto 0
					Exit Function
			End Select	
        Err.Clear
		strFileString = objSourceFile.ReadAll
		objSourceFile.close		
		Set objSourceFile = Nothing
		If Err.Number > 0 Then
					Err.Clear
					CopyFileToString = "-2"
					Set objFSO = Nothing
					On Error Goto 0
					Exit Function		
		End If 
		On Error Goto 0	
    End If		
	Set objFSO = Nothing
	CopyFileToString = strFileString
End Function 
	
'----------------------------------------------------------
'   Function Set_IE_obj (byRef objIE)
'----------------------------------------------------------
Function Set_IE_obj (byRef objIE)
	Dim nCount
	Set_IE_obj = False
	nCount = 0
	Do 
		On Error Resume Next
		Err.Clear
		Set objIE = CreateObject("InternetExplorer.Application")
		Select Case Err.Number
			Case &H800704A6 
				wscript.sleep 1000
				nCount = nCount + 1
				Call  TrDebug ("Set_IE_obj ERROR:" & Err.Number & " " & Err.Description, "", objDebug, MAX_LEN, 1, 1)
				If nCount > 4 Then
					On Error goto 0
					Exit Function
				End If
			Case 0 
				Set_IE_obj = True
				On Error goto 0
				Exit Function
			Case Else 
				Call  TrDebug ("Set_IE_obj ERROR:" & Err.Number & " " & Err.Description, "", objDebug, MAX_LEN, 1, 1)
				On Error goto 0
				Exit Function
		End Select
	On Error goto 0
	Loop
End Function
'-----------------------------------------
'  Function LoginEVHS(ByRef g_objIE)
'-----------------------------------------
Function LoginEVHS(ByRef g_objIE, vCred, nRetries, nInfo)
Dim Anchore, nTimer, URL, nCount, Login_url, vIE_Scale, strLogin, strPassword
Dim vvMsg(10,10)
Dim objRegEx
	Set objRegEx = CreateObject("VBScript.RegExp")
	objRegEx.Global = False
	objRegEx.Pattern = "https://.+\.com/"
    LoginEVHS = False
	URL = vCred(0)
	HostURL = objRegEx.Execute(URL).Item(0).Value
	'
	'   Open iexplore.exe
	Call Set_IE_obj (g_objIE)
	g_objIE.Visible = True
	'
	'   Navigate to schoolloop portal page
	nCount = 1
	Action = "LOAD PAGE"
	Do
		If nCount > nRetries Then 
			Exit Do
			
		End If 
		Select Case Action
			Case "LOAD PAGE"
					g_objIE.navigate URL
					nTimer = 0
					Do
						WScript.Sleep 200
						nTimer = nTimer + 0.2
						If nTimer > 10 Then exit do
					Loop While g_objIE.Busy
					Call TrDebug(Action & ": Page " & nCount & " loaded in: " & nTimer & "sec.", "", objDebug, MAX_LEN , 1, nInfo)
					'
					'  Validate portal page
					If g_objIE.Document.Location.href = URL Then 			
						LoginEVHS = True
						Exit Do
					End If 
					Action = "LOGIN WITH SAVED CRED"
			Case "LOGIN WITH SAVED CRED"
					If InStr(g_objIE.Document.Location.href,"login") Then
						For each Anchore in g_objIE.Document.getElementsByTagName("a")
							If Anchore.InnerText = "Login" Then
								Call TrDebug("Found Login Button: ", "", objDebug, MAX_LEN , 1, nInfo)
								Anchore.Click()
								nCount = nCount + 1
								Exit For
							End if 
						Next
						nTimer = 0
						' Wait until page loaded
						Do
							WScript.Sleep 200
							nTimer = nTimer + 0.2
							If nTimer > 10 Then exit do
						Loop While g_objIE.Busy		
						Call TrDebug(Action & ": Page " & nCount & " loaded in: " & nTimer & "sec.", "", objDebug, MAX_LEN , 1, nInfo)
						'  Validate portal page
						If g_objIE.Document.Location.href = URL Then 			
							LoginEVHS = True
							Exit Do
						End If
						Action = "ENTER CRED"
					Else 
						Login_url = HostURL & "/portal/login"
						Call TrDebug("Didn't find correct Login Page", "", objDebug, MAX_LEN , 1, nInfo)					
						Call TrDebug("Loading default login portal url: " & Login_url, "", objDebug, MAX_LEN , 1, nInfo)											
						g_objIE.navigate Login_url
						nCount = nCount + 1
						Do
							WScript.Sleep 200
							nTimer = nTimer + 0.2
							If nTimer > 10 Then exit do
						Loop While g_objIE.Busy		
						Call TrDebug(Action & ": Page " & nCount & " loaded in: " & nTimer & "sec.", "", objDebug, MAX_LEN , 1, nInfo)
						Action = "LOGIN WITH SAVED CRED"
					End If
			Case "ENTER CRED"
					If InStr(g_objIE.Document.Location.href,"login") Then
						strLogin = vCred(1)
						'
						'  GET SCREEN RESOLUTION
						Call WriteScreenResolution(vIE_Scale, 0)
						vvMsg(0,0) = "Sign In Schoolloop" 				: vvMsg(0,1) = "normal" : vvMsg(0,2) = HttpTextColor1
						Call IE_PromptLoginPassword (objParentWin,vIE_Scale, vvMsg, 1,strLogin, strPassword, False, 0 )
						g_objIE.document.getElementById("login_name").Value = vCred(1)
						g_objIE.document.getElementById("password").Value = strPassword
						For each Anchore in g_objIE.Document.getElementsByTagName("a")
							If Anchore.InnerText = "Login" Then
								Call TrDebug("Found Login Button: ", "", objDebug, MAX_LEN , 1, nInfo)
								Anchore.Click()
								nCount = nCount + 1
								Exit For
							End if 
						Next
						nTimer = 0
						' Wait until page loaded
						Do
							WScript.Sleep 200
							nTimer = nTimer + 0.2
							If nTimer > 10 Then exit do
						Loop While g_objIE.Busy		
						Call TrDebug(Action & ": Page " & nCount & " loaded in: " & nTimer & "sec.", "", objDebug, MAX_LEN , 1, nInfo)
						'  Validate portal page
						If g_objIE.Document.Location.href = URL Then 			
							LoginEVHS = True
							Exit Do
						Else 
							Exit Do
							Call TrDebug(Action & ": Can't Login to Schoolloop Portal" , "ERROR", objDebug, MAX_LEN , 1, 1)
						End If
						Action = "ENTER CRED"
					End If
		End Select
	Loop
End Function
'-----------------------------------------
'  Function ProgressReportExists(ByRef oTable, nInfo)
'-----------------------------------------
Function ProgressReportExists(ByRef oTable, nInfo)
Dim Anchore
		ProgressReportExists = False
		For Each Anchore in oTable.getElementsByTagName("a")
			If Anchore.InnerText = "Progress Report" Then
				ProgressReportExists = True
				Exit For
			End If
		Next
End Function
'-----------------------------------------
'  Function GetClassroom(ByRef oTable, ByRef vAcademics, )
'-----------------------------------------
Function GetClassroom(ByRef oTable, ByRef vAcademics, nIndex, nInfo)
Dim Anchore, objRegEx
	Set objRegEx = CreateObject("VBScript.RegExp")
	objRegEx.Global = True
	objRegEx.Pattern = "Progress Report"
	GetClassroom	= False
	For Each Anchore in oTable.getElementsByTagName("a")
		If Not objRegEx.Test(Anchore.InnerText) Then
			GetClassroom = True
			vAcademics(0,nIndex) = Anchore.InnerText
			Exit For 
		End If
	Next
End Function
'-----------------------------------------
'  Function GetProgressReportHref(ByRef oTable, ByRef vAcademics, )
'-----------------------------------------
Function GetProgressReportHref(ByRef oTable, ByRef vAcademics, nIndex, nInfo)
Dim Anchore, objRegEx
	Set objRegEx = CreateObject("VBScript.RegExp")
	objRegEx.Global = True
	objRegEx.Pattern = "Progress Report"
	GetProgressReportHref	= False
	For Each Anchore in oTable.getElementsByTagName("a")
		If objRegEx.Test(Anchore.InnerText) Then
			GetProgressReportHref = True
			vAcademics(UBound(vAcademics,1)-1,nIndex) = Anchore.Href
			Exit For 
		End If
	Next
End Function
'-----------------------------------------
'  Function GetGrade(ByRef oTable, ByRef vAcademics, )
'-----------------------------------------
Function GetGrade(ByRef oTable, ByRef vAcademics, nIndex, nInfo)
Dim oDiv, objRegEx
	Set objRegEx = CreateObject("VBScript.RegExp")
	objRegEx.Global = True
	GetGrade	= False
	objRegEx.Pattern = "^[ABCDEF][-\+]?"
	For Each oDiv in oTable.getElementsByTagName("div")
	'Call TrDebug("-->" & oDiv.InnerText, "", objDebug, MAX_LEN , 1, nInfo)
		If objRegEx.Test(oDiv.InnerText) Then
			GetGrade = True
			vAcademics(1,nIndex) = Trim(oDiv.InnerText)
			Exit For
		End If
	Next
End Function
'-----------------------------------------
'  Function GetScore(ByRef oTable, ByRef vAcademics, )
'-----------------------------------------
Function GetScore(ByRef oTable, ByRef vAcademics, nIndex, nInfo)
Dim oDiv, objRegEx
	Set objRegEx = CreateObject("VBScript.RegExp")
	objRegEx.Global = True
	GetScore	= False
	objRegEx.Pattern = "^\d.{0,5}%"
	For Each oDiv in oTable.getElementsByTagName("div")
		'Call TrDebug("-->" & oDiv.InnerText, "", objDebug, MAX_LEN , 1, nInfo)
		If objRegEx.Test(oDiv.InnerText) Then
			GetScore = True
			vAcademics(2,nIndex) = Trim(oDiv.InnerText)
			Exit For
		End If
	Next
End Function
'-------------------------------------------------------------
'    Function WriteScreenResolution(vIE_Scale, intX,intY)
'-------------------------------------------------------------
Function WriteScreenResolution(ByRef vIE_Scale, nDebug)
Dim g_objIE, intX, intY, intXreal, intYreal, vScr
Dim vScreen(6), stdOutFile
	stdOutFile = "ks-screen.dat"
    Redim vIE_Scale(2,3)
	nInd = 0
	Call Set_IE_obj(g_objIE)
	With g_objIE
		.Visible = False
		.Offline = True	
		.navigate "about:blank"
		Do
			WScript.Sleep 200
		Loop While g_objIE.Busy	
		.Document.Body.innerHTML = "<p>TEST</p>"
		.MenuBar = False
		.StatusBar = False
		.AddressBar = False
		.Toolbar = False		
		.Document.body.scroll = "no"
		.Document.body.Style.overflow = "hidden"
		.Document.body.Style.border = "None " & HttpBdColor1
		.Height = 100
		.Width = 100
    	OffsetX = .Width - .Document.body.clientWidth
		OffsetY = .Height - .Document.body.clientHeight
		If GetWin32Screen(".", vScr, nDebug) Then 
			 intXreal = vScr(0)
			 intYreal = vScr(1)
		Else 
			.FullScreen = True
			.navigate "about:blank"	
			 intXreal = .width
			 intYreal = .height
	    End If 
		.Quit
	End With
	If intXreal => 1440 Then intX = 1920 else intX = intXreal
	If intYreal => 900 Then intY = 1080  else intY = intYreal
	vIE_Scale(0,0) = intX : vIE_Scale(0,1) = OffsetX : vIE_Scale(0,2) = intXreal 
	vIE_Scale(1,0) = intY : vIE_Scale(1,1) = OffsetY : vIE_Scale(1,2) = intYreal
	Set g_objIE = Nothing
End Function
'------------------------------------------------------------------------------
'      Function ASKS USER TO ENTER PASSWORD
'------------------------------------------------------------------------------
 Function IE_PromptLoginPassword (objParentWin, vIE_Scale, vLine, nLine, ByRef strUsername, ByRef strPassword, Confirm, nDebug )
    Dim strPID
	Dim intX
    Dim intY
	Dim WindowH, WindowW
	Dim nFontSize_Def, nFontSize_10, nFontSize_12
	Dim g_objIE, g_objShell
	intX = 1920
	intY = 1080
	Dim IE_Menu_Bar
	Dim  IE_Border
	Const IE_REG_KEY = "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main\Window Title"
	'-----------------------------------------------------------------
	'  GET THE TITLE NAME USED BY IE EXPLORER WINDOW
	'-----------------------------------------------------------------
	On Error Resume Next
		Err.Clear
		IE_Window_Title =  objShell.RegRead(IE_REG_KEY)
		if Err.Number <> 0 Then 
			IE_Window_Title = "Internet Explorer"
		End If
	On Error Goto 0
	IE_Window_Title = "KSLD - " & IE_Window_Title
	strPassword = "DO NOT MATCH"
	IE_PromptLoginPassword = False	
	
	'----------------------------------------
	' SCREEN RESOLUTION
	'----------------------------------------
	intX = 1920
	intY = 1080
	intX = vIE_Scale(0,2) : IE_Border = vIE_Scale(0,1) : intY = vIE_Scale(1,2) : IE_Menu_Bar = vIE_Scale(1,1)
	nRatioX = vIE_Scale(0,0)/1920
    nRatioY = vIE_Scale(1,0)/1080
	Call Set_IE_obj (g_objIE)
	g_objIE.Offline = True
	g_objIE.navigate "about:blank"
	' This loop is required to allow the IE object to finish loading...
	Do
		WScript.Sleep 200
	Loop While g_objIE.Busy
	nHeader = Round (12 * nRatioY,0)
	LineH = Round (12 * nRatioY,0)
	nTab = 20
	nFontSize_10 = Round(10 * nRatioY,0)
	nFontSize_12 = Round(12 * nRatioY,0)
	nFontSize_14 = Round(14 * nRatioY,0)
	nFontSize_Def = Round(16 * nRatioY,0)
	nButtonX = Round(80 * nRatioX,0)
	nButtonY = Round(40 * nRatioY,0)
	If nButtonX < 50 then nButtonX = 50 End If
	If nButtonY < 30 then nButtonY = 30 End If
	CellW = Round(330 * nRatioX,0)
	ColumnW1 = Round(150 * nRatioX,0)
	CellH = 2 * (nLine + 7) * LineH
	WindowW = IE_Border + CellW
	WindowH = IE_Menu_Bar + CellH
	If Confirm Then 
	    CellH = CellH + 3 * 2 * LineH 
		nOrder = 1
    Else 
	    nOrder = 0
	End If
	WindowW = IE_Border + CellW
	WindowH = IE_Menu_Bar + CellH
    '----------------------------------------------
    '   MAIN COLORS OF THE FORM
    '----------------------------------------------		
	BackGroundColor = "grey"
	ButtonColor = HttpBgColor2
	InputBGColor = HttpBgColor4
	MainTextColor = HttpTextColor1
	g_objIE.Document.body.Style.FontFamily = "Helvetica"
	g_objIE.Document.body.Style.FontSize = nFontSize_Def
	g_objIE.Document.body.scroll = "no"
	g_objIE.Document.body.Style.overflow = "hidden"
	g_objIE.Document.body.Style.border = "none " & BackGroundColor
	g_objIE.Document.body.Style.background = BackGroundColor
	g_objIE.Document.body.Style.color = BackGroundColor
	g_objIE.Top = (intY - WindowH)/2
	g_objIE.Left = (intX - WindowW)/2
	'----------------------------------------------------------
	'    TITLE
	'----------------------------------------------------------
	strHTMLBody = strHTMLBody &_
		"<table border=""1"" cellpadding=""1"" cellspacing=""1"" style="" position: absolute; left: 0px; top: 0px;" &_
		" border-collapse: collapse; border-style: none; border width: 1px; border-color: " & HttpBgColor2 & "; background-color: "& HttpBgColor2 & ";" &_
		"width: " & CellW & "px;"">" & _
		"<tbody>"	
	For nInd = 0 to nLine - 1
		 If vLine(nInd,2) = HttpTextColor1 Then vLine(nInd,2) = MainTextColor
		strHTMLBody = strHTMLBody &_
		"<tr>" &_
			"<td style=""border-style: none; background-color: " & HttpBgColor2 & ";"" class=""oa1"" height=""" &  2 * LineH & """ width=""" & CellW & """>" & _
				"<p style=""text-align: center; font-family: 'arial narrow';font-size: " & nFontSize_12 & ".0pt; font-weight: " & vLine(nInd,1) & "; color: " & vLine(nInd,2) & """>" & vLine(nInd,0) & "</p>" &_
			"</td>" &_
		"</tr>"
	Next
	strHTMLBody = strHTMLBody & "</tbody></table>"
	
	'----------------------------------------------------------
	'    MAIN FORM FOR ENTERING LOGON AND PASSWORD
	'----------------------------------------------------------
	TableW = CellW
	ColumnW_1 = 3 * Int(TableW/3)
	ColumnW_2 = TableW - ColumnW_1
	strHTMLBody = strHTMLBody &_
		"<table border=""1"" cellpadding=""1"" cellspacing=""1"" style="" position: absolute; left: 0px; top: " & (nLine + 1) * LineH * 2 & "px;" &_
		" border-collapse: collapse; border-style: none; border width: 1px; border-color: " & HttpBgColor5 & "; background-color: none;;" &_
		"width: " & TableW & "px;"">" & _
		"<tbody>"		
	'----------------------------------------------------	
	'  ROW 1
	'----------------------------------------------------
	strHTMLBody = strHTMLBody & _
	"<tr>" &_
		"<td style=""border-style: none; background-color: none;""class=""oa1"" height=""" &  2 * LineH & """ width=""" & ColumnW_1 & """>" & _
			"<p style=""position: relative; left: " & Int(nTab/2) & "px; bottom: -3px; font-size: " & nFontSize_12 & ".0pt; font-family: 'arial narrow'; color: " & MainTextColor &_
			"; font-weight: bold;"">LOGIN NAME</p>" &_
		"</td>" &_
		"<td style=""border-style: none; background-color: none;""class=""oa1"" height=""" &  2 * LineH & """ width=""" & ColumnW_1 & """>" & _
		"</td>" &_
	"</tr>"		
	'----------------------------------------------------	
	'  ROW 2
	'----------------------------------------------------
	strHTMLBody = strHTMLBody & _
	"<tr>" &_
		"<td style=""border-style: none; background-color: none;"" align=""center"" class=""oa1"" height=""" & 2 * LineH & """ >" & _
			"<input name=UserName style=""text-align: center;font-size: " & nFontSize_12 & ".0pt; border-style: None; font-family: 'Helvetica'; color: " & HttpTextColor2 &_
			"; border-radius: 10px " &_
			"; background-color: " & InputBGColor & "; font-weight: Normal;"" AccessKey=p size=20 maxlength=25 tabindex=1>" &_
		"</td>" &_
		"<td style=""border-style: none; background-color: none;"" align=""center"" class=""oa1"" height=""" &  2 * LineH & """ width=""" & ColumnW_1 & """>" & _
		   "<button style=""font-weight: bold; border-style: None; background-color: " & HttpBgColor2 & "; color: " & HttpTextColor2 &_
			"; width:" & nButtonX & ";height:" & 2 * LineH & "; font-family: 'arial narrow';""" & _
			"id='EXIT' name='Cancel' AccessKey='C' tabindex=" & nOrder + 4 & " onclick=document.all('ButtonHandler').value='Cancel';>CANCEL</button>" & _		    
		"</td>" &_		
	"</tr>"
	'----------------------------------------------------	
	'  ROW 3 (EMPTY)
	'----------------------------------------------------
	strHTMLBody = strHTMLBody & _
	"<tr>" &_
		"<td style=""border-style: none; background-color: none;""class=""oa1"" height=""" &  LineH & """>" & _
		"</td>" &_
		"<td style=""border-style: none; background-color: none;""class=""oa1"" height=""" &  LineH & """>" & _
		"</td>" &_
	"</tr>"
	'----------------------------------------------------	
	'  ROW 4
	'----------------------------------------------------
	strHTMLBody = strHTMLBody & _
	"<tr>" &_
		"<td style=""border-style: none; background-color: none;""class=""oa1"" height=""" &  2 * LineH & """ >" & _
			"<p style=""position: relative; left: " & Int(nTab/2) & "px; bottom: -3px; font-size: " & nFontSize_12 & ".0pt; font-family: 'arial narrow'; color: " & MainTextColor &_
			"; font-weight: bold;"">PASSWORD</p>" &_
		"</td>" &_
		"<td style=""border-style: none; background-color: none;""class=""oa1"" height=""" &  2 * LineH & """ >" & _
		"</td>" &_
	"</tr>"			
	'----------------------------------------------------	
	'  ROW 5
	'----------------------------------------------------
	strHTMLBody = strHTMLBody & _
	"<tr>" &_
		"<td style=""border-style: none; background-color: none;"" align=""center"" class=""oa1"" height=""" & 2 * LineH & """>" & _
			"<input id='PASSWD' name=Password style=""text-align: center;font-size: " & nFontSize_12 & ".0pt; border-style: None; font-family: 'Helvetica'; color: " & HttpTextColor2 &_
			"; border-radius: 10px " &_
			"; background-color: " & InputBGColor & "; font-weight: Normal;"" AccessKey=p size=20 maxlength=32 tabindex=2 " & _
			"type=password onkeydown=""if (event.keyCode == 13) document.all('ButtonHandler').value='OK'"" > " &_
		"</td>" &_
		"<td style=""border-style: none; background-color: none;"" align=""center"" class=""oa1"" height=""" &  2 * LineH & """ width=""" & ColumnW_1 & """>" & _
		   "<button style=""font-weight: bold; border-style: None; background-color: " & HttpBgColor2 & "; color: " & HttpTextColor2 &_
			"; width:" & nButtonX & ";height:" & 2 * LineH & "; font-family: 'arial narrow';""" & _
			"id='OK' name='OK' AccessKey='C' tabindex=" & nOrder + 3 & " onclick=document.all('ButtonHandler').value='OK';>SIGN IN</button>" & _		    
		"</td>" &_				
	"</tr>"
	'----------------------------------------------------	
	'  ROW 6 (EMPTY)
	'----------------------------------------------------
	strHTMLBody = strHTMLBody & _
	"<tr>" &_
		"<td style=""border-style: none; background-color: none;""class=""oa1"" height=""" &  LineH & """>" & _
		"</td>" &_
		"<td style=""border-style: none; background-color: none;""class=""oa1"" height=""" &  LineH & """>" & _
		"</td>" &_
	"</tr>"
	'----------------------------------------------------	
	'  CONFIRM PASSWORD ROW
	'----------------------------------------------------
	If Confirm Then 
		'----------------------------------------------------	
		'  ROW 7
		'----------------------------------------------------
		strHTMLBody = strHTMLBody & _
		"<tr>" &_
			"<td style=""border-style: none; background-color: none;""class=""oa1"" height=""" &  2 * LineH & """ >" & _
				"<p style=""position: relative; left: " & Int(nTab/2) & "px; bottom: -3px; font-size: " & nFontSize_12 & ".0pt; font-family: 'arial narrow'; color: " & MainTextColor &_
				"; font-weight: bold;"">CONFIRM PASSWORD</p>" &_
			"</td>" &_
			"<td style=""border-style: none; background-color: none;""class=""oa1"" height=""" &  2 * LineH & """ >" & _
			"</td>" &_
		"</tr>"			
		'----------------------------------------------------	
		'  ROW 8
		'----------------------------------------------------
		strHTMLBody = strHTMLBody & _
		"<tr>" &_
			"<td style=""border-style: none; background-color: none;"" align=""center"" class=""oa1"" height=""" & 2 * LineH & """>" & _
				"<input id='PASSWD2' name=Password2 style=""text-align: center;font-size: " & nFontSize_12 & ".0pt; border-style: None; font-family: 'Helvetica'; color: " & HttpTextColor2 &_
				"; border-radius: 10px " &_
				"; background-color: " & InputBGColor & "; font-weight: Normal;"" AccessKey=p size=20 maxlength=32 tabindex=3 " & _
				"type=password onkeydown=""if (event.keyCode == 13) document.all('ButtonHandler').value='OK'"" > " &_
			"</td>" &_
			"<td style=""border-style: none; background-color: none;"" align=""center"" class=""oa1"" height=""" &  2 * LineH & """>" & _
			"</td>" &_				
		"</tr>"
	End If
	strHTMLBody = strHTMLBody & "</tbody></table>"
    strHTMLBody = strHTMLBody &_
                "<input name='ButtonHandler' type='hidden' value='Nothing Clicked Yet'>"
			
	g_objIE.Document.Body.innerHTML = strHTMLBody
	g_objIE.MenuBar = False
	g_objIE.StatusBar = False
	g_objIE.AddressBar = False
	g_objIE.Toolbar = False
	g_objIE.height = WindowH
	g_objIE.width = WindowW
	g_objIE.document.Title = "Login and Password"
	g_objIE.document.getElementById("OK").style.borderRadius = "10px"
	g_objIE.document.getElementById("EXIT").style.borderRadius = "10px"
	g_objIE.document.getElementById("OK").style.backgroundcolor = ButtonColor
	g_objIE.document.getElementById("EXIT").style.backgroundcolor = ButtonColor
	If Confirm Then
	    g_objIE.Document.getElementById("OK").innerHTML = "OK"
	Else 
	   	g_objIE.Document.getElementById("OK").innerHTML = "SIGN IN"
	End If
	
	g_objIE.Visible = False
	Do
		WScript.Sleep 100
	Loop While g_objIE.Busy	
	Set g_objShell = WScript.CreateObject("WScript.Shell")
	g_objIE.Visible = True
	g_objIE.Document.All("UserName").Focus
	g_objIE.Document.All("UserName").Value = strUsername
'    g_objIE.Document.body.addeventlistener "keydown", GetRef("KeyLA"), false
	Do
		On Error Resume Next
		Err.Clear
		strNothing = g_objIE.Document.All("ButtonHandler").Value
		if Err.Number <> 0 then exit do
		On Error Goto 0
		Select Case strNothing
			Case "Cancel"
				' The user clicked Cancel. Exit the loop
				IE_PromptLoginPassword = False				
				Exit Do
			Case "OK"
				' strUsername = g_objIE.Document.All("Username").Value
				Select Case Confirm
					Case True
						if g_objIE.Document.All("Password").Value = g_objIE.Document.All("Password2").Value  and _
						   InStr(g_objIE.Document.All("Password").Value," ") = 0 and _
						   g_objIE.Document.All("Password").Value <> "" Then 
							strUsername = g_objIE.Document.All("UserName").Value
							strPassword = g_objIE.Document.All("Password").Value
							IE_PromptLoginPassword = True
							Exit Do
						Else
							strUsername = g_objIE.Document.All("UserName").Value
							strPassword = "DO NOT MATCH"
							IE_PromptLoginPassword = True
							Exit Do
						End If 
					Case False
							strUsername = g_objIE.Document.All("UserName").Value
							strPassword = g_objIE.Document.All("Password").Value
							IE_PromptLoginPassword = True
							Exit Do
				End Select
		End Select
	    Wscript.Sleep 200
    Loop
	g_objIE.quit
	Wscript.Sleep 200
	Set g_objIE = Nothing
	Set g_objShell = Nothing
End Function
'-------------------------------------------------
'   Function SelectStudentPage(objIE, nInfo)
'-------------------------------------------------
Function SelectStudentPage(ByRef g_objIE, ByRef vCred, nInfo)
Dim Anchore, Div, FoundPortalTitle, objDiv
	Set objDiv = g_objIE.Document.getElementById("container_content")
    FoundPortalTitle = False 
	'
	'   Validate Student's name for currently Loaded page
	For each Div in objDiv.getElementsByTagName("div")
		If FoundPortalTitle Then 
			Call TrDebug("Currently Displayed Student: " & Div.InnerText, "", objDebug, MAX_LEN , 1, nInfo)
			If InStr(Lcase(Div.InnerText), Lcase(vCred(3))) and InStr(Div.InnerText, Lcase(vCred(4))) Then 
				SelectStudentPage = True
				Exit Function
			End If
			Exit For
		End If
		If InStr(Div.InnerText, "Portal:") Then
			FoundPortalTitle = True
		End If 
	Next
	'
	'   If wrong student page loaded then look for right studen's page link and click it
	For each Anchore in g_objIE.Document.getElementsByTagName("a")
		If InStr(Lcase(Anchore.InnerText), Lcase(vCred(3))) > 0 and InStr(Lcase(Anchore.InnerText), Lcase(vCred(4))) > 0 Then
			Call TrDebug("Found Student Button: ", "", objDebug, MAX_LEN , 1, nInfo)
			Anchore.Click()
			' Wait until page loaded
			Do
				WScript.Sleep 200
				nTimer = nTimer + 0.2
				If nTimer > 10 Then exit do
			Loop While g_objIE.Busy		
			Call TrDebug("Page loaded OK ", "", objDebug, MAX_LEN , 1, nInfo)
			nCount = nCount + 1
			Exit For
		End if 
	Next
	'
	'   Validate Student's name for currently Loaded page
	FoundPortalTitle = False
	Set objDiv = g_objIE.Document.getElementById("container_content")
	For each Div in objDiv.getElementsByTagName("div")
		If FoundPortalTitle Then 
			Call TrDebug("Currently Displayed Student: " & Div.InnerText, "", objDebug, MAX_LEN , 1, nInfo)
			If InStr(Lcase(Div.InnerText), Lcase(vCred(3))) and InStr(Div.InnerText, Lcase(vCred(4))) Then 
				SelectStudentPage = True
				Exit Function
			End If
			Exit For
		End If
		If InStr(Div.InnerText, "Portal:") Then
			FoundPortalTitle = True
		End If 
	Next
End Function 
'-----------------------------------------
'  Function GetClassroom(ByRef oTable, ByRef vAcademics, )
'-----------------------------------------
Function GetAssessmentsList(ByRef g_objIE, ByRef vAcademics, nIndex, nInfo)
Dim Anchore, objRegEx, oRow, oTable, nColDate, MyTable
Dim nColAssessment, nColScore,nColWeight,nColWScore,nColCategory
Dim FoundCategoryTable, FoundAssessmentTable
Dim oString
Dim vCategory
Set objRegEx = CreateObject("VBScript.RegExp")
	' 
	' Define variables
	objRegEx.Global = False	
	GetAssessmentsList = False		
	FoundAssessmentTable = False
	nColCategory = -1
	nColWeight = -1
	nColWScore = -1
	Redim vCategory(1,3)
	'
	'  Look for Category list
	If GetHTMLTableByRowName(g_objIE,oTable,"Category:") Then 
		Set oRow = oTable.rows.item(0)
		For each oCell in oRow.cells
			If InStr(oCell.InnerText,"Category:")     Then nColCategory = oCell.cellIndex 
			If InStr(oCell.InnerText,"Weight:")       Then nColWeight = oCell.cellIndex
			If InStr(oCell.InnerText,"Weight Score:") Then nColWScore = oCell.cellIndex
		Next
		'
		'  Validate if there is any category listed in the table
		If oTable.rows.length > 1 Then 
			Redim vCategory(oTable.rows.length-1,3)
			For nRow = 1 to oTable.rows.length-1
				Set oRow = oTable.rows.item(nRow)
				If nColCategory <> -1 Then vCategory(nRow-1,0) = oRow.Cells.Item(nColCategory).InnerText
				If nColWeight <> -1   Then vCategory(nRow-1,1) = oRow.Cells.Item(nColWeight).InnerText
				If nColWScore <> -1   Then vCategory(nRow-1,2) = oRow.Cells.Item(nColWScore).InnerText
			Next
		End If		
	End If
	'
	'  Look for Assessment table with scores
	If GetHTMLTableByRowName(g_objIE,oTable,"Assessment:") Then 
		If oTable.rows.length = 1 Then 
			Call TrDebug("Assessment Table for : " & vAcademics(3,nIndex) & " has no data", "", objDebug, MAX_LEN , 1, nInfo)
			vAcademics(7,nIndex) = 0
			Exit Function
		End If 
		Set oRow = oTable.rows.item(0)
		'   Validate assessment table 
		If oRow.cells.length = 1 Then 
			Call TrDebug("Assessment Table for : " & vAcademics(3,nIndex) & " has no data", "", objDebug, MAX_LEN , 1, nInfo)
			vAcademics(7,nIndex) = 0
			Exit Function
		End If 
		'   Get columns' index
		For each oCell in oRow.cells
			If InStr(oCell.InnerText,"Assessment:") Then nColAssessment = oCell.cellIndex 
			If InStr(oCell.InnerText,"Score:")      Then nColScore = oCell.cellIndex
			If InStr(oCell.InnerText,"Due:")        Then nColDate = oCell.cellIndex
		Next
	Else 
		Call TrDebug("Can't Find Assessment Table for : " & vAcademics(3,nIndex), "", objDebug, MAX_LEN , 1, nInfo)
		vAcademics(7,nIndex) = 0
		Exit Function
	End If
	'
	' Set initial values
	vAcademics(3,nIndex) = ""
	vAcademics(4,nIndex) = ""
	vAcademics(5,nIndex) = ""
	vAcademics(6,nIndex) = ""	
	vAcademics(7,nIndex) = oTable.rows.length - 1
	'
	' Read Assessments list
	For nRow = 1 to oTable.rows.length - 1
		Set oRow = oTable.rows.Item(nRow)
		Set oCell = oRow.Cells.Item(nColAssessment)
		'  Read from Asesment Cell
		Call ReadAssessmentCategory(oCell.InnerText,vCategory,vAcademics, nIndex)
		'  Read from Due Cell		
		Set oCell = oRow.Cells.Item(nColScore)
		objRegEx.Pattern = "\d{1,3}\.\d{1,2}%"
		Set oString = objRegEx.Execute(oCell.InnerText)
		If oString.Count > 0 _
			Then vAcademics(4,nIndex) = vAcademics(4,nIndex) & oString.Item(0) & "," _
			Else vAcademics(4,nIndex) = vAcademics(4,nIndex) & ",,"		
		'  Read from Score Cell
		Set oCell = oRow.Cells.Item(nColDate)
		objRegEx.Pattern = "\d{1,2}/\d{1,2}/\d{1,2}"
		Set oString = objRegEx.Execute(oCell.InnerText)
		If oString.Count > 0 _
			Then vAcademics(3,nIndex) = vAcademics(3,nIndex) & oString.Item(0) & "," _
			Else vAcademics(3,nIndex) = vAcademics(3,nIndex) & ",,"
	Next
	GetAssessmentsList = True
End Function
'-----------------------------------------------------
'   Function GetHTMLTableByRowName(g_objIE, oTable, strTitle)
'----------------------------------------------------
Function GetHTMLTableByRowName(ByRef g_objIE, ByRef oTable, strTitle)
	GetHTMLTableByRowName = False
    For each oTable in g_objIE.Document.getElementsByTagName("table")
		If oTable.getElementsByTagName("table").length = 0 Then 
			If oTable.rows.length > 0 Then 
				Set oRow = oTable.rows.item(0)
				If oRow.cells.length > 0 Then 
					For each oCell in oRow.cells
						If InStr(oCell.InnerText,strTitle) Then 
							GetHTMLTableByRowName = True 
							Exit Function
						End If
					Next
				End If
			End If
		End If
	Next
End Function
'-------------------------------------------------------------
' Function ReadAssessmentCategory(strText,vCategory,vAcademics, nIndex)
'------------------------------------------------------------
Function ReadAssessmentCategory(strText,vCategory,vAcademics, nIndex)
Dim objRegEx, nInd
    ReadAssessmentCategory = False
	Set objRegEx = CreateObject("VBScript.RegExp")
	' 
	' Define variables
	objRegEx.Global = False	
	For nInd = 0 to UBound(vCategory,1)	
		Do
			If vCategory(nInd,0) = "" Then Exit Do
			objRegEx.Pattern = "[\t\r\n\v\f]"
			vCategory(nInd,0) = objRegEx.Replace(vCategory(nInd,0),"")
			objRegEx.Pattern = vCategory(nInd,0)
			If objRegEx.Execute(strText).Count > 0 Then 
				vAcademics(5,nINdex) = vAcademics(5,nIndex) & vCategory(nInd,0) & ","
				vAcademics(6,nINdex) = vAcademics(6,nIndex) & vCategory(nInd,1) & ","
				ReadAssessmentCategory = True
				Exit Function
			End If
			Exit Do
		Loop
	Next
	vAcademics(5,nIndex) = vAcademics(5,nIndex) & ","
	vAcademics(6,nIndex) = vAcademics(6,nIndex) & ","
End Function
'-----------------------------------------------------------------------------------
' Function GetFileLineCountSelect - Returns number of lines int the text file
'-----------------------------------------------------------------------------------
  Function GetFileLineCountSelect(strFileName, ByRef vFileLines,strChar1, strChar2, strChar3, nDebug)
    Dim nIndex
	Dim strLine
	Dim objDataFileName
	Dim objRegEx
	Const EMPTY_STRING = "e-m-p-t-y"
	Set objRegEx = CreateObject("VBScript.RegExp")
	objRegEx.Global = False
	
	Call NormalizePatterm(strChar1, EMPTY_STRING)
	Call NormalizePatterm(strChar2, EMPTY_STRING)
	Call NormalizePatterm(strChar3, EMPTY_STRING)
	
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
		TestLine = strLine
		objRegEx.Pattern =".+"
		If Not objRegEx.Test(TestLine) Then TestLine = EMPTY_STRING
		objRegEx.Pattern ="(^" & strChar1 & ")|(^" & strChar2 & ")|(^" & strChar3 & ")"
		If 	objRegEx.Test(TestLine) = False Then 
					Redim Preserve vFileLines(nIndex + 1)
					vFileLines(nIndex) = strLine
					If nDebug = 1 Then objDebug.WriteLine "GetFileLineCountSelect: vFileLines(" & nIndex & ")="  & vFileLines(nIndex) End If  
					nIndex = nIndex + 1
					bResult = True
		End If
		
	Loop
	objDataFileName.Close
	Set objDataFileName = Nothing
    GetFileLineCountSelect = nIndex
End Function
'-------------------------------------------------
'  Function NormalizePatterm(strPattern)
'-------------------------------------------------
Function NormalizePatterm(ByRef strPattern, EmptyString)
Dim vSpecualChars, Special
	vSpecualChars = Array("\","$",".","[","]","{","}","(",")","?","*","+","|")
	If strPattern = "" Then strPattern = EmptyString : Exit Function : End If
	' If not an empty string
	For Each Special in vSpecualChars
		strPattern = Replace(strPattern,Special,"\" & Special)
	Next
	MsgBox strPattern
End Function


Function encrypt(Str, key)
 Dim lenKey, KeyPos, LenStr, x, Newstr
 
 Newstr = ""
 lenKey = Len(key)
 KeyPos = 1
 LenStr = Len(Str)
 str = StrReverse(str)
 For x = 1 To LenStr
	  On Error Resume Next
      Newstr = Newstr & chr(Asc(Mid(str,x,1)) + Asc(Mid(key,KeyPos,1)))
	  If Err.Number > 0 Then  MsgBox "1.=" & Mid(str,x,1) & chr(13) & "2.=" & Mid(key,KeyPos,1) : Exit For : End If
	  On Error Goto 0
      KeyPos = keyPos+1
      If KeyPos > lenKey Then KeyPos = 1
 Next
 encrypt = Newstr
End Function

Function Decrypt(str,key)
 Dim lenKey, KeyPos, LenStr, x, Newstr
 
 Newstr = ""
 lenKey = Len(key)
 KeyPos = 1
 LenStr = Len(Str)
 
 str=StrReverse(str)
 For x = LenStr to 1 Step - 1
      Newstr = Newstr & chr(asc(Mid(str,x,1)) - Asc(Mid(key,KeyPos,1)))
      KeyPos = KeyPos+1
      If KeyPos > lenKey Then KeyPos = 1
      Next
      Newstr=StrReverse(Newstr)
      Decrypt = Newstr
End Function

Function Crypt(key, str)
	' str, key - strings, containing char from 32 to 126 ASCII code table
	' return encrypt/decrypt string, on input error return empty string
	Dim x, keyCharNum, strCharNum, diffCharNum, diffSum

	For x = 1 To (Len(str) + Len(key) - Abs(Len(str) - Len(key)))/2
		keyCharNum = Asc( Mid (key, (x-1) Mod Len(key) + 1, 1))
		strCharNum = Asc( Mid (str, (x-1) Mod Len(str) + 1, 1))
		If (keyCharNum > 126 Or keyCharNum < 32 Or strCharNum > 126 Or strCharNum < 32) Then
			Crypt = ""
			Exit For
		End If
		diffCharNum = keyCharNum - strCharNum
		If (diffCharNum < 0) Then diffCharNum = diffCharNum + 126 - 32 + 1
		diffSum = chr(diffCharNum + 32)
		Crypt = Crypt & diffSum
	Next
End Function