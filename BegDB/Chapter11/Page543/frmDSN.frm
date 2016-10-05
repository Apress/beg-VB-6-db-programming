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
   Begin VB.CommandButton cmdTestDSN 
      Caption         =   "&Test the DSN"
      Height          =   615
      Left            =   1320
      TabIndex        =   0
      Top             =   1200
      Width           =   1815
   End
End
Attribute VB_Name = "frmDSN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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
