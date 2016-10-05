VERSION 5.00
Begin VB.Form frmADOPublishers 
   Caption         =   "Programming with ADO"
   ClientHeight    =   4260
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4350
   LinkTopic       =   "Form1"
   ScaleHeight     =   4260
   ScaleWidth      =   4350
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Fill List"
      Height          =   615
      Left            =   1200
      TabIndex        =   1
      Top             =   3480
      Width           =   1935
   End
   Begin VB.ListBox List1 
      Height          =   3180
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "frmADOPublishers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim adoConnection As ADODB.Connection
Dim adoRecordset As ADODB.Recordset
Dim connectString As String

'-Create a new connection --
Set adoConnection = New ADODB.Connection
'-Create a new recordset --
Set adoRecordset = New ADODB.Recordset

'-Build our connection string to use when we open the connection --
connectString = "Provider=Microsoft.Jet.OLEDB.3.51;" _
                 & "Data Source=C:\BegDB\Biblio.mdb"

adoConnection.Open connectString
adoRecordset.Open "Publishers", adoConnection

Do Until adoRecordset.EOF
  List1.AddItem adoRecordset!Name
  adoRecordset.MoveNext
Loop

adoRecordset.Close
adoConnection.Close
   Set adoRecordset = Nothing
   Set adoConnection = Nothing

End Sub
