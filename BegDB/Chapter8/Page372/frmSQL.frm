VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmSQL 
   Caption         =   "SQL Query Tester"
   ClientHeight    =   4695
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6450
   LinkTopic       =   "Form1"
   ScaleHeight     =   4695
   ScaleWidth      =   6450
   StartUpPosition =   3  'Windows Default
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmSQL.frx":0000
      Height          =   1815
      Left            =   120
      OleObjectBlob   =   "frmSQL.frx":0014
      TabIndex        =   0
      Top             =   120
      Width           =   6135
   End
   Begin VB.TextBox Text1 
      Height          =   1575
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   3000
      Width           =   4095
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "&Execute"
      Height          =   495
      Left            =   4440
      TabIndex        =   7
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\BegDB\Biblio.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4200
      Width           =   1140
   End
   Begin VB.Label lblUpdatable 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   4440
      TabIndex        =   6
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label lblCurrentRecord 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   2280
      TabIndex        =   5
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label lblRecordCount 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Updatable"
      Height          =   255
      Left            =   4440
      TabIndex        =   3
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Current Record"
      Height          =   255
      Left            =   2280
      TabIndex        =   2
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Record Count"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2040
      Width           =   1695
   End
End
Attribute VB_Name = "frmSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdRun_Click()
On Error GoTo SQLError:

Data1.RecordSource = Text1
Data1.Refresh

If Data1.RecordSource <> "" Then
  If (Data1.Recordset.RecordCount > 0) Then
    With Data1.Recordset
      .MoveLast
      .MoveFirst
      lblRecordCount = .RecordCount
      lblUpdatable = IIf(.Updatable, "Yes", "No")
    End With
  Else
   lblRecordCount = "Records Returned: 0"
   lblCurrentRecord = "No records"
   lblUpdatable = ""
  End If
Else
  MsgBox ("Please enter an SQL statement")
End If

Exit Sub
SQLError:
  Dim sError As String
  sError = "Error Number: " & Err.Number & vbCrLf
  sError = sError & Err.Description
  MsgBox (sError)
  Exit Sub

End Sub

Private Sub Data1_Reposition()

lblCurrentRecord = Data1.Recordset.AbsolutePosition + 1

End Sub

