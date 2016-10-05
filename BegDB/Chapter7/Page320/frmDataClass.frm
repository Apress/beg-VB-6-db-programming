VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDataClass 
   Caption         =   "Template form for our dataClass"
   ClientHeight    =   3540
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   3540
   ScaleWidth      =   6585
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text4 
      DataField       =   "PubID"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   16
      Text            =   "PubID"
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      DataField       =   "ISBN"
      DataSource      =   "Data1"
      Height          =   405
      Left            =   2160
      TabIndex        =   15
      Tag             =   "1"
      Text            =   "ISBN"
      Top             =   360
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      DataField       =   "Year Published"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   5280
      TabIndex        =   14
      Tag             =   "1"
      Text            =   "Year Published"
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      DataField       =   "Title"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Tag             =   "1"
      Text            =   "Title"
      Top             =   1440
      Width           =   4815
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   840
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "|<<"
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "<<"
      Height          =   495
      Index           =   1
      Left            =   1080
      TabIndex        =   9
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   ">>"
      Height          =   495
      Index           =   2
      Left            =   2040
      TabIndex        =   8
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   ">>|"
      Height          =   495
      Index           =   3
      Left            =   3000
      TabIndex        =   7
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "&Add New"
      Height          =   495
      Index           =   4
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "&Edit"
      Height          =   495
      Index           =   5
      Left            =   3360
      TabIndex        =   5
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "&Save"
      Height          =   495
      Index           =   6
      Left            =   1200
      TabIndex        =   4
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "&Delete"
      Height          =   495
      Index           =   7
      Left            =   4440
      TabIndex        =   3
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "&Undo"
      Height          =   495
      Index           =   8
      Left            =   2280
      TabIndex        =   2
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "&Find"
      Height          =   495
      Index           =   9
      Left            =   5520
      TabIndex        =   1
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "D&one"
      Height          =   495
      Index           =   10
      Left            =   5520
      TabIndex        =   0
      Top             =   2640
      Width           =   975
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   3285
      Width           =   6585
      _ExtentX        =   11615
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label4 
      Caption         =   "Book Title"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Year Published"
      Height          =   255
      Left            =   5280
      TabIndex        =   19
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "ISBN"
      Height          =   255
      Left            =   2160
      TabIndex        =   18
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Publisher's ID"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblRecordCount 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblRecordCount"
      Height          =   375
      Left            =   4080
      TabIndex        =   12
      Top             =   2760
      Width           =   1335
   End
End
Attribute VB_Name = "frmDataClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim myDataClass As dataClass

Private Sub cmdButton_Click(Index As Integer)
Text3.SetFocus
myDataClass.ProcessCMD Index
End Sub

Private Sub Form_Activate()

Static blnIsOld As Boolean  'the default value of a Boolean is False

If blnIsOld = False Then
    Set myDataClass = New dataClass
    
    With myDataClass
 Set .FormName = Me     'pass in the current form
 Set .dataCtl = Data1   'the data control to manage
 Set .ProgressBar = ProgressBar1
 .dbName = gDataBaseName
 .Buttons = "cmdButton"
 .RecordSource = "SELECT * FROM Titles"
 .LabelToUpdate = lblRecordCount
 .FindCaption = "Select a Book Title"
 .FindRecordSource = "SELECT Title FROM Titles ORDER BY Title"
 .FindMatchField = "Title"

 .Tag = "1"        'identifies the controls
 .ProcessCMD 0     'default to the 1st record
End With

Text3.SetFocus
blnIsOld = True
End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim iMessage As Integer

If (Data1.Recordset.EditMode <> dbEditNone) Then
  iMessage = MsgBox("You must complete editing the current record", _
                           vbInformation, App.EXEName)
  Cancel = True
End If

End Sub


Private Sub Text1_GotFocus()
  highLight
End Sub

Private Sub Text2_GotFocus()
  highLight
End Sub

Private Sub Text3_GotFocus()
  highLight
End Sub

Private Sub Text4_GotFocus()
  highLight
End Sub
