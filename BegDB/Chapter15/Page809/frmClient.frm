VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmClient 
   Caption         =   "Client Side Cursors for AbsolutePosition"
   ClientHeight    =   3420
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4620
   LinkTopic       =   "Form1"
   ScaleHeight     =   3420
   ScaleWidth      =   4620
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar progressBar1 
      Height          =   255
      Left            =   743
      TabIndex        =   1
      Top             =   1920
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdClient 
      Caption         =   "Client Side Cursors"
      Height          =   855
      Left            =   983
      TabIndex        =   0
      Top             =   480
      Width           =   2655
   End
   Begin VB.Label lblAbsolute 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   1170
      TabIndex        =   2
      Top             =   2520
      Width           =   2295
   End
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClient_Click()
Dim adoRecordset As ADODB.Recordset
Dim sConnectionString As String
Dim sMessage As String

sConnectionString = "Provider=Microsoft.Jet.OLEDB.3.51;" & _
                  "Data Source=C:\begdb\Biblio.mdb"


Set adoRecordset = New ADODB.Recordset
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open "Titles", sConnectionString, , , adCmdTable

progressBar1.Min = 0
progressBar1.Max = adoRecordset.RecordCount

While Not adoRecordset.EOF
     progressBar1.Value = adoRecordset.AbsolutePosition
     lblAbsolute = "Record: " & adoRecordset.AbsolutePosition & _
                                     " of " & adoRecordset.RecordCount
     DoEvents
    adoRecordset.MoveNext
Wend

End Sub
