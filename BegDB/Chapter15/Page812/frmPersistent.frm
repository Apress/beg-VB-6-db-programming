VERSION 5.00
Begin VB.Form frmPersistent 
   Caption         =   "Persistent Recordsets"
   ClientHeight    =   2850
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4065
   LinkTopic       =   "Form1"
   ScaleHeight     =   2850
   ScaleWidth      =   4065
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRetrieve 
      Caption         =   "&Retrieve Recordset"
      Height          =   735
      Left            =   765
      TabIndex        =   1
      Top             =   1560
      Width           =   2535
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save Recordset"
      Height          =   735
      Left            =   765
      TabIndex        =   0
      Top             =   360
      Width           =   2535
   End
End
Attribute VB_Name = "frmPersistent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdRetrieve_Click()
Dim adoRecordset As ADODB.Recordset

On Error GoTo retrieveError

Set adoRecordset = New ADODB.Recordset

ChDir App.Path

adoRecordset.Open "myADORecordset", , , , adCmdFile
MsgBox ("Recordset successfully retrieved")
Exit Sub

retrieveError:
MsgBox Err.Description
Exit Sub

End Sub

Private Sub cmdSave_Click()
Dim adoConnection As ADODB.Connection
Dim adoRecordset As ADODB.Recordset
Dim sConnString As String
Dim sSqlString As String
Dim sMyFile As String

On Error GoTo createError:

sConnString = "Provider=Microsoft.Jet.OLEDB.3.51;" & _
"Persist Security Info=False;" & _
  "Data Source=C:\begdb\Biblio.mdb"

Set adoConnection = New ADODB.Connection
adoConnection.Open sConnString

sSqlString = "SELECT * FROM Authors"

Set adoRecordset = New ADODB.Recordset

adoRecordset.CursorType = adOpenDynamic
adoRecordset.LockType = adLockOptimistic
adoRecordset.CursorLocation = adUseClient

adoRecordset.Open sSqlString, adoConnection

ChDir App.Path

sMyFile = Dir(App.Path & "\myAdoRecordset")
If (Len(sMyFile)) Then
  Kill sMyFile
End If

adoRecordset.Save ("myADORecordset")

MsgBox ("Recordset successfully saved")

Exit Sub

createError:
MsgBox Err.Description
Set adoRecordset = Nothing
Set adoConnection = Nothing
Exit Sub

End Sub

