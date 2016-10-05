VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDataClass 
   Caption         =   "Easy Data Entry Grid Example"
   ClientHeight    =   5280
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7020
   LinkTopic       =   "Form1"
   ScaleHeight     =   5280
   ScaleWidth      =   7020
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   27
      Top             =   5025
      Width           =   7020
      _ExtentX        =   12383
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmDataClass.frx":0000
      Height          =   2175
      Left            =   120
      OleObjectBlob   =   "frmDataClass.frx":0014
      TabIndex        =   26
      Tag             =   "2"
      Top             =   1440
      Width           =   6735
   End
   Begin VB.CommandButton cmdTitles 
      Caption         =   "D&one"
      Height          =   495
      Index           =   10
      Left            =   5880
      TabIndex        =   24
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton cmdTitles 
      Caption         =   "&Find"
      Height          =   495
      Index           =   9
      Left            =   120
      TabIndex        =   23
      Top             =   5400
      Width           =   975
   End
   Begin VB.CommandButton cmdTitles 
      Caption         =   "&Undo"
      Height          =   495
      Index           =   8
      Left            =   3960
      TabIndex        =   22
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton cmdTitles 
      Caption         =   "&Delete"
      Height          =   495
      Index           =   7
      Left            =   3000
      TabIndex        =   21
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton cmdTitles 
      Caption         =   "&Save"
      Height          =   495
      Index           =   6
      Left            =   2040
      TabIndex        =   20
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton cmdTitles 
      Caption         =   "&Edit"
      Height          =   495
      Index           =   5
      Left            =   1080
      TabIndex        =   19
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton cmdTitles 
      Caption         =   "&Add New"
      Height          =   495
      Index           =   4
      Left            =   120
      TabIndex        =   18
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton cmdTitles 
      Caption         =   ">>|"
      Height          =   495
      Index           =   3
      Left            =   3000
      TabIndex        =   17
      Top             =   3720
      Width           =   975
   End
   Begin VB.CommandButton cmdTitles 
      Caption         =   ">>"
      Height          =   495
      Index           =   2
      Left            =   2040
      TabIndex        =   16
      Top             =   3720
      Width           =   975
   End
   Begin VB.CommandButton cmdTitles 
      Caption         =   "<<"
      Height          =   495
      Index           =   1
      Left            =   1080
      TabIndex        =   15
      Top             =   3720
      Width           =   975
   End
   Begin VB.CommandButton cmdTitles 
      Caption         =   "|<<"
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   14
      Top             =   3720
      Width           =   975
   End
   Begin VB.TextBox Text2 
      DataField       =   "Name"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   1320
      TabIndex        =   13
      Text            =   "Name"
      Top             =   240
      Width           =   2055
   End
   Begin VB.Data Data2 
      Caption         =   "Titles"
      Connect         =   "Access"
      DatabaseName    =   "C:\BegDB\Biblio.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Enabled         =   0   'False
      Exclusive       =   0   'False
      Height          =   300
      Left            =   4320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Titles"
      Top             =   6120
      Width           =   2220
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000004&
      DataField       =   "PubID"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Tag             =   "1"
      Text            =   "PubID"
      Top             =   240
      Width           =   855
   End
   Begin VB.Data Data1 
      Caption         =   "Publishers"
      Connect         =   "Access"
      DatabaseName    =   "C:\BegDB\Biblio.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Enabled         =   0   'False
      Exclusive       =   0   'False
      Height          =   300
      Left            =   4320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Publishers"
      Top             =   5640
      Visible         =   0   'False
      Width           =   2220
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "|<<"
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "<<"
      Height          =   495
      Index           =   1
      Left            =   1200
      TabIndex        =   9
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   ">>"
      Height          =   495
      Index           =   2
      Left            =   2280
      TabIndex        =   8
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   ">>|"
      Height          =   495
      Index           =   3
      Left            =   3360
      TabIndex        =   7
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "&Add New"
      Height          =   495
      Index           =   4
      Left            =   120
      TabIndex        =   6
      Top             =   6600
      Width           =   975
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "&Edit"
      Height          =   495
      Index           =   5
      Left            =   3360
      TabIndex        =   5
      Top             =   6600
      Width           =   975
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "&Save"
      Height          =   495
      Index           =   6
      Left            =   1200
      TabIndex        =   4
      Top             =   6600
      Width           =   975
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "&Delete"
      Height          =   495
      Index           =   7
      Left            =   4440
      TabIndex        =   3
      Top             =   6600
      Width           =   975
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "&Undo"
      Height          =   495
      Index           =   8
      Left            =   2280
      TabIndex        =   2
      Top             =   6600
      Width           =   975
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "&Find"
      Height          =   495
      Index           =   9
      Left            =   5880
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
      Top             =   6600
      Width           =   975
   End
   Begin VB.Label lblTitles 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblTitles"
      Height          =   375
      Left            =   5160
      TabIndex        =   25
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Label lblRecordCount 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblRecordCount"
      Height          =   375
      Left            =   5520
      TabIndex        =   11
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "frmDataClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim publishersClass As dataClass
Dim titlesClass As dataClass

Private Sub cmdButton_Click(Index As Integer)
Text2.SetFocus
publishersClass.ProcessCMD Index
titlesClass.RecordSource = "SELECT * FROM Titles WHERE PubID = " _
    & Data1.Recordset!PubID
End Sub

Private Sub cmdTitles_Click(Index As Integer)
titlesClass.ProcessCMD Index

If (Index = 4) Then         ' we know it is an add
  Data2.Recordset!PubID = Data1.Recordset!PubID
End If

End Sub

Private Sub Data1_Reposition()
If (TypeName(titlesClass) <> "Nothing") Then
    titlesClass.RecordSource = _
        "SELECT * FROM Titles WHERE PubID = " _
        & Data1.Recordset!PubID
End If

End Sub


Private Sub Data2_Validate(Action As Integer, Save As Integer)
'======================================================
'== Place Business rules for the Titles records here ==
'======================================================
If (Action = vbDataActionUpdate) Then
  If (Len(DBGrid1.Columns("ISBN").Text) < 5) Then
    DBGrid1.Col = 3
    MsgBox ("Please enter a valid ISBN number.")
    Action = 0
    Save = False
    Exit Sub
  ElseIf (Len(DBGrid1.Columns("Title").Text) < 3) Then
    MsgBox ("Please enter a valid Title.")
    Action = 0
    Save = False
    Exit Sub
  ElseIf (Len(DBGrid1.Columns("Year Published").Text) < 4) Then
    MsgBox ("Please enter a valid Year Published.")
    Action = 0
    Save = False
    Exit Sub
  End If
End If

End Sub

Private Sub Form_Load()

Set publishersClass = New dataClass
Set titlesClass = New dataClass

With publishersClass
Set .FormName = Me    'pass in the current form
Set .dataCtl = Data1  'the data control to manage
Set .ProgressBar = ProgressBar1
   .dbName = gDataBaseName
   .Buttons = "cmdButton"
   .RecordSource = "SELECT * FROM Publishers"
   .LabelToUpdate = lblRecordCount
   .FindCaption = "Select a Publisher"
   .FindRecordSource = "SELECT Name FROM Publishers ORDER BY Name"
   .FindMatchField = "Name"
   .Tag = "1"        'identifies the controls
   .ProcessCMD 0     'default to the 1st record
End With

With titlesClass
Set .FormName = Me
Set .dataCtl = Data2
Set .ProgressBar = ProgressBar1
   .dbName = gDataBaseName
   .Buttons = "cmdTitles"
   .RecordSource = "SELECT * FROM Titles WHERE PubID = " _
    & Data1.Recordset!PubID
   .LabelToUpdate = lblTitles
   .Tag = "2"
   .ProcessCMD 0
End With

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim iMessage As Integer

If (Data1.Recordset.EditMode <> dbEditNone) Then
  iMessage = MsgBox("You must complete editing the current record", _
                           vbInformation, App.EXEName)
  Cancel = True
End If

End Sub
