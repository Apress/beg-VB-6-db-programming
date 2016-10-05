VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmExport 
   Caption         =   "Form1"
   ClientHeight    =   4215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4215
   LinkTopic       =   "Form1"
   ScaleHeight     =   4215
   ScaleWidth      =   4215
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CD1 
      Left            =   3480
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdExcel 
      Caption         =   "Export to Excel"
      Height          =   735
      Left            =   1140
      TabIndex        =   2
      Top             =   2760
      Width           =   2175
   End
   Begin VB.CommandButton cmdHTML 
      Caption         =   "Export to HTML"
      Height          =   735
      Left            =   1140
      TabIndex        =   1
      Top             =   1680
      Width           =   2175
   End
   Begin VB.CommandButton cmdCSV 
      Caption         =   "Export to CSV"
      Height          =   735
      Left            =   1140
      TabIndex        =   0
      Top             =   480
      Width           =   2175
   End
End
Attribute VB_Name = "frmExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adoConnection As ADODB.Connection
Dim adoRecordset As ADODB.Recordset
Dim connectString As String
Dim objExcel As Object
Dim objTemp As Object
Public Function csv(adoRecordset As ADODB.Recordset) As Boolean

Dim iTotalRecords As Integer
Dim sFileToExport As String
Dim iFileNum As Integer
Dim msg As String
Dim iIndx As Integer
Dim iNumberOfFields As Integer

Screen.MousePointer = vbDefault

On Error Resume Next

With CD1
  .CancelError = True
  .FileName = "Export.csv"
  .InitDir = App.Path
  .DialogTitle = "Save Comma Delimited Export File"
  .Filter = "Export Files (*.CSV)|*.CSV"
  .DefaultExt = "CSV"
  .Flags = cdlOFNOverwritePrompt Or cdlOFNCreatePrompt
  .ShowSave
End With

'--------------------------------
'-- User cancels the operation --
'--------------------------------
If (Err = 32755) Then  'operation canceled
   Screen.MousePointer = vbDefault
   Beep
   msg = "The export operation was canceled." & vbCrLf
   iIndx = MsgBox(msg, vbOKOnly + vbInformation, "Comma Delimited Export File")
   csv = False
   Exit Function
Else
  On Error GoTo expError
End If

'---------------------------------------
'-- Let's save the data now.          --
'-- Get the name of the file to save. --
'---------------------------------------
Screen.MousePointer = vbHourglass

iTotalRecords = 0
sFileToExport = CD1.FileName
iFileNum = FreeFile()
Open sFileToExport For Output As #iFileNum   ' Open file for output.

'-------------------------
'-- Stream out the data --
'-------------------------

iNumberOfFields = adoRecordset.Fields.Count - 1

adoRecordset.MoveFirst
Do Until adoRecordset.EOF
  iTotalRecords = iTotalRecords + 1
  For iIndx = 0 To iNumberOfFields
    If (IsNull(adoRecordset.Fields(iIndx))) Then
       Print #iFileNum, ","; 'simply a comma delimited string
    Else
       If iIndx = iNumberOfFields Then
         Print #iFileNum, Trim$(CStr(adoRecordset.Fields(iIndx)));
       Else
         Print #iFileNum, Trim$(CStr(adoRecordset.Fields(iIndx))); ",";
       End If
    End If
  Next
  Print #iFileNum,
  adoRecordset.MoveNext
  DoEvents
Loop

'----------------
Close iFileNum
Screen.MousePointer = vbDefault
Beep
msg = "Export File " & sFileToExport & vbCrLf
msg = msg & "successfully created." & vbCrLf
msg = msg & iTotalRecords & " records written to disk." & vbCrLf
iIndx = MsgBox(msg, vbOKOnly + vbInformation, "Comma Delimited File")
csv = True
Exit Function

expError:

Screen.MousePointer = vbDefault
MsgBox (Err & " " & Err.Description)
csv = False

End Function

Private Sub cmdCSV_Click()
Call csv(adoRecordset)
End Sub

Private Sub cmdExcel_Click()
Call excel(adoRecordset)
End Sub


Private Sub cmdHTML_Click()
Call html(adoRecordset)
End Sub

Private Sub Form_Activate()

Dim sSqlString As String

Set adoConnection = New ADODB.Connection
Set adoRecordset = New ADODB.Recordset

connectString = "Provider=Microsoft.Jet.OLEDB.3.51;" _
                 & "Data Source=C:\begdb\biblio.mdb"

sSqlString = "SELECT * FROM Publishers where PubID <= 50"

adoConnection.Open connectString
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open sSqlString, adoConnection
MsgBox adoRecordset.RecordCount

End Sub

Public Function html(adoRecordset As ADODB.Recordset) As Boolean

Dim fileToExport As String
Dim iFileNumber As Integer
Dim outerloop As Integer
Dim innerloop As Integer
Dim sMsg As String
Dim iIndx As Integer

Screen.MousePointer = vbDefault

On Error Resume Next

With CD1
  .CancelError = True
  .FileName = "Export.htm"
  .InitDir = App.Path
  .DialogTitle = "Save (H)yper (T)ext (M)arkup (L)anguage Export File"
  .Filter = "Export Files (*.HTM)|*.HTM"
  .DefaultExt = "HTM"
  .Flags = cdlOFNOverwritePrompt Or cdlOFNCreatePrompt
  .ShowSave
End With

'--------------------------------
'-- User cancels the operation --
'--------------------------------
If (Err = 32755) Then  'operation canceled
   Screen.MousePointer = vbDefault
   Beep
   sMsg = "The export operation was canceled." & vbCrLf
   iIndx = MsgBox(sMsg, vbOKOnly + vbInformation, "HTML Export File")
   html = False
   Exit Function
Else
  On Error GoTo htmlError
End If

'---------------------------------------
'-- Let's save the data now.          --
'-- Get the name of the file to save. --
'---------------------------------------
Screen.MousePointer = vbHourglass

fileToExport = CD1.FileName
iFileNumber = FreeFile()

Open fileToExport For Output As #iFileNumber

adoRecordset.MoveFirst

Print #iFileNumber, "<HTML><HEAD><TITLE>ADO Recordset HTML Data Export</TITLE></HEAD>"
Print #iFileNumber, "<BODY BGCOLOR=""FFFFFF"">"
Print #iFileNumber, "<TABLE BGCOLOR=""00AAFF"" WIDTH=""100%"">"
Print #iFileNumber, "<TR><TD>"
Print #iFileNumber, "<FONT FACE=ARIAL SIZE+=3><B>ADO Recordset HTML Export</B></FONT></TD></TR>"
Print #iFileNumber, "<TR>"
For iIndx = 0 To adoRecordset.Fields.Count - 1
  Print #iFileNumber, "<TD BGCOLOR=CCCCC>"
  Print #iFileNumber, "<B> &nbsp"; adoRecordset.Fields(iIndx).Name; "&nbsp </B>"
  Print #iFileNumber, "</TD>"
Next
Print #iFileNumber, "</TR>"

With adoRecordset
  .MoveFirst
  While Not .EOF
    Print #iFileNumber, "<TR>"
    For innerloop = 0 To .Fields.Count - 1
      Print #iFileNumber, "<TD BGCOLOR=CCCCC>"
      Print #iFileNumber, "&nbsp"; .Fields(innerloop); "&nbsp"
      Print #iFileNumber, "</TD>"
    Next
    Print #iFileNumber, "</TR>"
    .MoveNext
  Wend
End With

Print #iFileNumber, "</TABLE></BODY></HTML>"

Close #iFileNumber

MsgBox "Done"
Screen.MousePointer = vbDefault
html = True

Exit Function

htmlError:
Screen.MousePointer = vbDefault
MsgBox Err.Description
html = False

End Function


Public Sub excel(adoRecordset As ADODB.Recordset)

Dim iIndx As Integer
Dim iRowIndex As Integer
Dim iColIndex As Integer
Dim iRecordCount As Integer
Dim iFieldCount As Integer
Dim sMessage As String
Dim avRows As Variant
Dim excelVersion As Integer

'-- Read all of the records into our array
avRows = adoRecordset.GetRows()

'-- Determine how many fields and records
iRecordCount = UBound(avRows, 2) + 1
iFieldCount = UBound(avRows, 1) + 1

'-- Create reference variable for the spreadsheet
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
objExcel.Workbooks.Add

'-- We need this line to insure Excel remains visible if we switch
Set objTemp = objExcel

excelVersion = Val(objExcel.Application.Version)
If (excelVersion >= 8) Then
 Set objExcel = objExcel.ActiveSheet
End If

'-- Place the names of the fields as column headers --
iRowIndex = 1
iColIndex = 1
For iColIndex = 1 To iFieldCount
  With objExcel.Cells(iRowIndex, iColIndex)
     .Value = adoRecordset.Fields(iColIndex - 1).Name
     With .Font
       .Name = "Arial"
       .Bold = True
       .Size = 9
     End With
  End With
Next

'-- memory management --
adoRecordset.Close
Set adoRecordset = Nothing

'-- Just add data --
With objExcel
  For iRowIndex = 2 To iRecordCount + 1
   For iColIndex = 1 To iFieldCount
    .Cells(iRowIndex, iColIndex).Value = avRows(iColIndex - 1, iRowIndex - 2)
   Next
  Next
End With

objExcel.Cells(1, 1).CurrentRegion.EntireColumn.AutoFit

End Sub

