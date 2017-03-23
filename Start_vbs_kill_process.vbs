Set objShell = wscript.createobject( "wscript.shell")
strProjectFolder =     "C:\KidsSchild\DVLP\Project-001"
strDestinationFolder = "C:\KidsSchild\DVLP\Script"
strProjectFile =       "project.txt"
strLaunch = "wscript C:\KidsSchild\DVLP\Tools\VBS_kill_process.vbs -n wscript.exe -c Kids_schedule_WIN_ " 
								
objShell.run strLaunch,0,False
set objShell = Nothing
