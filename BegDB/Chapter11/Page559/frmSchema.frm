VERSION 5.00
Begin VB.Form frmSchema 
   Caption         =   "Show Table Schema"
   ClientHeight    =   4335
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3495
   LinkTopic       =   "Form1"
   ScaleHeight     =   4335
   ScaleWidth      =   3495
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdProvider 
      Caption         =   "&Provider"
      Height          =   615
      Left            =   960
      TabIndex        =   3
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton cmdErrors 
      Caption         =   "&Error Collection"
      Height          =   615
      Left            =   960
      TabIndex        =   2
      Top             =   2280
      Width           =   1455
   End
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

Private Sub cmdErrors_Click()
Dim adoConnection As ADODB.Connection
Dim adoErrors As ADODB.Errors

Dim i As Integer
Dim StrTmp

On Error GoTo AdoError

Set adoConnection = New ADODB.Connection

' Open connection to Bogus ODBC Data Source for BIBLIO.MDB
adoConnection.ConnectionString = "DBQ=BIBLIO.MDB;" & _
                  "DRIVER={Microsoft Access Driver (*.mdb)};" & _
                  "DefaultDir=C:\OhNooo\Directory\Path;"

adoConnection.Open

' Remaining code goes here, but of course our program
' will never reach it because the connection string
' will generate an error because of the bogus directory

' Close the open objects
adoConnection.Close

' Destroy anything not destroyed yet
Set adoConnection = Nothing

Exit Sub

AdoError:

  Dim errorCollection As Variant
  Dim errLoop As Error
  Dim strError As String
  Dim iCounter As Integer
  
 ' In case our adoConnection is not set or
 ' there were other initialization problems
  On Error Resume Next

  iCounter = 1

  ' Enumerate Errors collection and display properties of
  ' each Error object.
   strError = ""
   Set errorCollection = adoConnection.Errors
   For Each errLoop In errorCollection
         With errLoop
            strError = "Error #" & iCounter & vbCrLf
            strError = strError & " ADO Error #" & .Number & vbCrLf
            strError = strError & " Description  " & .Description & vbCrLf
            strError = strError & " Source       " & .Source & vbCrLf
            Debug.Print strError
            iCounter = iCounter + 1
         End With
      Next

End Sub

Private Sub cmdProvider_Click()
Dim adoConnection As ADODB.Connection
Dim sConnection As String
Set adoConnection = New ADODB.Connection
sConnection = "Provider=Microsoft.Jet.OLEDB.3.51;Data Source=c:\BegDB\Biblio.mdb"
adoConnection.Open sConnection
'Output all of the version information to the debug window.
Debug.Print "ADO Version: " & adoConnection.Version & vbCrLf
Debug.Print "Database Name: " & adoConnection.Properties("DBMS Name") & vbCrLf
Debug.Print "Database Version: " & adoConnection.Properties("DBMS Version") & vbCrLf
Debug.Print "OLE DB Version: " & adoConnection.Properties("OLE DB Version") & vbCrLf
Debug.Print "Provider Name: " & adoConnection.Properties("Provider Name") & vbCrLf
Debug.Print "Provider Version: " & adoConnection.Properties("Provider Version") & vbCrLf

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
