Set objShell = wscript.createobject( "wscript.shell")
Set objFSO = wscript.CreateObject("Scripting.FileSystemObject")
strDirectoryWork = objFSO.GetParentFolderName(wscript.ScriptFullName)
strLaunch = "wscript " & strDirectoryWork & "\VBS_kill_process.vbs -n SecureCRT.exe -a" 
								
objShell.run strLaunch,0,False
set objShell = Nothing
set objFSO = Nothing
