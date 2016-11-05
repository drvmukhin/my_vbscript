' Using customized COM Component "APIWrapperCOM.dll"

Set obj = CreateObject("APIWrapperCOM.APIWrapper")

winHandle = obj.FindWindow("test.txt - Notepad")

obj.KillWindow(winHandle)