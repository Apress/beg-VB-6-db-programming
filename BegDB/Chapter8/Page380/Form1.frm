VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Form1 
   Caption         =   "Simple Dynamic SQL Statement"
   ClientHeight    =   4140
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5595
   LinkTopic       =   "Form1"
   ScaleHeight     =   4140
   ScaleWidth      =   5595
   StartUpPosition =   3  'Windows Default
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "Form1.frx":0000
      Height          =   2775
      Left            =   120
      OleObjectBlob   =   "Form1.frx":0014
      TabIndex        =   0
      Top             =   120
      Width           =   5295
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "C:\BegDB\Biblio.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   3480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3600
      Visible         =   0   'False
      Width           =   1980
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\BegDB\Biblio.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   3480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3000
      Visible         =   0   'False
      Width           =   1980
   End
   Begin VB.ListBox List1 
      Height          =   1035
      Left            =   120
      TabIndex        =   1
      Top             =   3000
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()
Dim mySQL As String

mySQL = "SELECT City FROM Publishers GROUP BY City"

Data1.RecordSource = mySQL
Data1.Refresh

While Not Data1.Recordset.EOF 'loop through all of the records returned
  With Data1.Recordset
       If (Not IsNull(!city)) Then 'ensure that the field is not null
           List1.AddItem !city      'if the field is not null, add it
       End If
       .MoveNext
  End With
Wend

End Sub

Private Sub List1_Click()

Data2.RecordSource = "SELECT * FROM Publishers WHERE City = '" & _
                                       List1 & "' ORDER BY Name"
Data2.Refresh

End Sub
