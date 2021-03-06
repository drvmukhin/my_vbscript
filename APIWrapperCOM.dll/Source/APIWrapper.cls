VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "APIWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Declare Function FxFindWindow Lib "user32" Alias "FindWindowA" _
(ByVal lpszClassName As String, ByVal lpszWindow As String) As Long

Private Declare Function FxSendMessage Lib "user32" Alias "SendMessageA" _
   (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam _
   As Long, ByVal lParam As Long) As Long
   
Private Const WM_CLOSE = &H10

Public Function FindWindow(ByVal CaptionString As String) As Long
    FindWindow = FxFindWindow(vbNullString, CaptionString)
End Function

Public Function KillWindow(ByVal winHandle As Long) As Long
    KillWindow = FxSendMessage(winHandle, WM_CLOSE, 0, 0)
End Function

