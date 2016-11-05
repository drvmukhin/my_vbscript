Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
Set objShellApp = CreateObject("Shell.Application")
scriptFile = "test.vbs"
strDirectoryWork =  objFSO.GetParentFolderName(Wscript.ScriptFullName)
strParam = strDirectoryWork & "\" & scriptFile
objShellApp.ShellExecute "wscript", strParam, "", "runas", 1
Set objShellApp = Nothing
Set objFSO = Nothing