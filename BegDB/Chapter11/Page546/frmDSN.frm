VERSION 5.00
Begin VB.Form frmDSN 
   Caption         =   "DSN Connection to DB"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExecute 
      Caption         =   "Test &Execute"
      Height          =   615
      Left            =   1320
      TabIndex        =   1
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CommandButton cmdTestDSN 
      Caption         =   "&Test the DSN"
      Height          =   615
      Left            =   1320
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
End
Attribute VB_Name = "frmDSN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExecute_Click()
Dim myConnection As ADODB.Connection
Dim myRecordSet As ADODB.Recordset

Set myConnection = New ADODB.Connection


myConnection.ConnectionString = "Provider=Microsoft.Jet.OLEDB.3.51;Data Source=C:\BegDB\Biblio.mdb"

myConnection.Open

' Create a Recordset by executing a SQL statement
Set myRecordSet = myConnection.Execute("Select * From Titles")

' Show the first title in the recordset.
MsgBox myRecordSet("Title")

' Close the recordset and connection.
myRecordSet.Close
myConnection.Close

End Sub

Private Sub cmdTestDSN_Click()
Dim myConnection As ADODB.Connection
Set myConnection = New ADODB.Connection

'If we wanted, we could set the provider property to the OLE
'DB Provider for ODBC. However we will set it in the connect 'string.

' Open a connection using an ODBC DSN. The MS OLE DB for
' SQL is MSDASQL. We gave our new data source the name "Our ADO Example DSN"
' so let's use it.

myConnection.ConnectionString = "Provider=MSDASQL.1;Persist Security Info=False;User ID=Admin;Data Source=Our ADO Example DSN;Mode=Share Deny None"

myConnection.Open

' Determine if we conected.
If myConnection.State = adStateOpen Then
   MsgBox "Welcome to the Biblio Database!"
Else
  MsgBox "The connection could not be made."
End If

' Close the connection.
myConnection.Close

End Sub
