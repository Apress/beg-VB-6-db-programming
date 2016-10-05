Attribute VB_Name = "Globals"
Option Explicit

Public Declare Function sendMessageByString& Lib "user32" _
    Alias "SendMessageA" (ByVal hwnd As Long, _
    ByVal wMsg As Long, ByVal wParam As Long, _
     ByVal lParam As String)

Public Const LB_SELECTSTRING = &H18C
Public gFindString As String
Public Const gDataBaseName = "C:\BegDB\Biblio.mdb"


