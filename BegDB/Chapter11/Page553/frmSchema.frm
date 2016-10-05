VERSION 5.00
Begin VB.Form frmSchema 
   Caption         =   "Show Table Schema"
   ClientHeight    =   2640
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3600
   LinkTopic       =   "Form1"
   ScaleHeight     =   2640
   ScaleWidth      =   3600
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDataTypes 
      Caption         =   "&Data Types"
      Height          =   615
      Left            =   960
      TabIndex        =   1
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton cmdSchema 
      Caption         =   "&Schema"
      Height          =   615
      Left            =   960
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "frmSchema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDataTypes_Click()
Dim adoConnection As ADODB.Connection
Dim adoRsFields As ADODB.Recordset
Dim sConnection As String
Set adoConnection = New ADODB.Connection
sConnection = "Provider=Microsoft.Jet.OLEDB.3.51;Data Source=c:\BegDB\Biblio.mdb"
adoConnection.Open sConnection
Set adoRsFields = adoConnection.OpenSchema(adSchemaProviderTypes)
Do Until adoRsFields.EOF
  Debug.Print "Data Type: " & adoRsFields!TYPE_NAME & vbTab _
                   & "Column Size: " & adoRsFields!COLUMN_SIZE
  adoRsFields.MoveNext
Loop
adoRsFields.Close
Set adoRsFields = Nothing
adoConnection.Close
Set adoConnection = Nothing

End Sub

Private Sub cmdSchema_Click()
Dim adoConnection As ADODB.Connection
Dim adoRsFields As ADODB.Recordset
Dim sConnection As String
Dim sCurrentTable As String
Dim sNewTable As String

Set adoConnection = New ADODB.Connection

sConnection = "Provider=Microsoft.Jet.OLEDB.3.51;Data Source=c:\BegDB\Biblio.mdb"

adoConnection.Open sConnection

Set adoRsFields = adoConnection.OpenSchema(adSchemaColumns)

sCurrentTable = ""
sNewTable = ""

Do Until adoRsFields.EOF
  sCurrentTable = adoRsFields!TABLE_NAME
  If (sCurrentTable <> sNewTable) Then
    sNewTable = adoRsFields!TABLE_NAME
    Debug.Print "Current Table: " & adoRsFields!TABLE_NAME
  End If
  Debug.Print "   Field: " & adoRsFields!COLUMN_NAME
  adoRsFields.MoveNext
Loop

adoRsFields.Close
Set adoRsFields = Nothing
adoConnection.Close
Set adoConnection = Nothing

End Sub
