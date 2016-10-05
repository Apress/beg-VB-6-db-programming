VERSION 5.00
Begin VB.Form frmFind 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find Record"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4830
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Left            =   3840
      Top             =   720
   End
   Begin VB.Data dtaFind 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   1800
      TabIndex        =   5
      Top             =   4440
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   360
      TabIndex        =   3
      Top             =   1200
      Width           =   4095
   End
   Begin VB.TextBox txtFind 
      Height          =   285
      Left            =   360
      TabIndex        =   2
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      Caption         =   "Double-click to select"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   3480
      Width           =   4095
   End
   Begin VB.Label lblCount 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblCount"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   3840
      Width           =   4095
   End
   Begin VB.Label lblWhichTable 
      Caption         =   "lblWhichTable"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   4095
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub Form_Activate()
List1.Enabled = False
dtaFind.DatabaseName = gDataBaseName
dtaFind.Refresh
If (dtaFind.Recordset.RecordCount > 0) Then
  Screen.MousePointer = vbHourglass
  dtaFind.Recordset.MoveFirst
  While Not dtaFind.Recordset.EOF
        List1.AddItem dtaFind.Recordset.Fields(0) & ""
        dtaFind.Recordset.MoveNext
  Wend
  List1.Enabled = True
  DoEvents
End If

lblCount = "There are " & dtaFind.Recordset.RecordCount & " records"
Screen.MousePointer = vbDefault

End Sub



Public Property Let RecordSource(ByVal sNewValue As String)
   dtaFind.RecordSource = sNewValue
End Property


Public Property Let addCaption(ByVal sNewValue As String)
   lblWhichTable = sNewValue
End Property


Private Sub Form_Load()
Timer1.Enabled = False
Timer1.Interval = 1

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmFind = Nothing
End Sub

Private Sub List1_DblClick()
'get the item the user clicks on and assign it
If (InStr(List1, "'")) Then
    gFindString = SrchReplace(List1)
Else
    gFindString = List1
End If

Timer1.Enabled = True

End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = False
Unload Me

End Sub

Private Sub txtFind_Change()
Dim entryNum As Long
Dim txtToFind As String
txtToFind = txtFind.Text

entryNum = sendMessageByString(List1.hwnd, _
           LB_SELECTSTRING, 0, txtToFind)

End Sub

Function SrchReplace(ByVal sStringToFix As String) As String

Dim iPosition As Integer        'where is the offending char?
Dim sCharToReplace As String    'which char do we want to replace?
Dim sReplaceWith As String      'what should it be replaced with?
Dim sTempString As String       'build the correct returned string

sCharToReplace = "'"
sReplaceWith = "''"

iPosition = InStr(sStringToFix, sCharToReplace)
sTempString = ""

Do While iPosition
  sTempString = sTempString & Left$(sStringToFix, iPosition - 1)
  sTempString = sTempString & sReplaceWith
  sTempString = sTempString & _
    Mid$(sStringToFix, iPosition + 1, Len(sStringToFix))
  iPosition = InStr(iPosition + 1, sStringToFix, sCharToReplace)
Loop

SrchReplace = sTempString

End Function


