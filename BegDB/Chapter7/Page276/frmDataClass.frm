VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDataClass 
   Caption         =   "Template form for our dataClass"
   ClientHeight    =   2310
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   2310
   ScaleWidth      =   6585
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      DataSource      =   "Data1"
      Height          =   375
      Left            =   2400
      TabIndex        =   13
      Tag             =   "1"
      Text            =   "Text1"
      Top             =   240
      Width           =   1935
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   240
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "|<<"
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "<<"
      Height          =   495
      Index           =   1
      Left            =   1080
      TabIndex        =   9
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   ">>"
      Height          =   495
      Index           =   2
      Left            =   2040
      TabIndex        =   8
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   ">>|"
      Height          =   495
      Index           =   3
      Left            =   3000
      TabIndex        =   7
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "&Add New"
      Height          =   495
      Index           =   4
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "&Edit"
      Height          =   495
      Index           =   5
      Left            =   3360
      TabIndex        =   5
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "&Save"
      Height          =   495
      Index           =   6
      Left            =   1200
      TabIndex        =   4
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "&Delete"
      Height          =   495
      Index           =   7
      Left            =   4440
      TabIndex        =   3
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "&Undo"
      Height          =   495
      Index           =   8
      Left            =   2280
      TabIndex        =   2
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "&Find"
      Height          =   495
      Index           =   9
      Left            =   5520
      TabIndex        =   1
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "D&one"
      Height          =   495
      Index           =   10
      Left            =   5520
      TabIndex        =   0
      Top             =   1440
      Width           =   975
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   2055
      Width           =   6585
      _ExtentX        =   11615
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label lblRecordCount 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblRecordCount"
      Height          =   375
      Left            =   4080
      TabIndex        =   12
      Top             =   1560
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
myDataClass.ProcessCMD (Index)
End Sub

Private Sub Form_Activate()

Static blnIsOld As Boolean 'the default value of a Boolean is False

If blnIsOld = False Then


    Set myDataClass = New dataClass
    
    With myDataClass
 Set .FormName = Me     'pass in the current form
 Set .dataCtl = Data1   'the data control to manage
 Set .ProgressBar = ProgressBar1
 .dbName = gDataBaseName
 .Buttons = "cmdButton"
 .RecordSource = "SELECT * FROM Publishers"
 .LabelToUpdate = lblRecordCount
 .FindCaption = "Select a Publisher's ID"
 .FindRecordSource = "SELECT PubID FROM Publishers"
 .FindMatchField = "PubID"

 .Tag = "1"        'identifies the controls
 .ProcessCMD 0     'default to the 1st record
End With

blnIsOld = True
End If

Text1.DataField = "Name"

If (Data1.Recordset.EditMode = dbEditNone) Then
   Text1.Locked = True
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
