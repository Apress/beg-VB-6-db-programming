VERSION 5.00
Begin VB.Form frmMarshal 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMarshal 
      Caption         =   "Marshal ADO Recordset"
      Height          =   615
      Left            =   720
      TabIndex        =   0
      Top             =   720
      Width           =   2955
   End
End
Attribute VB_Name = "frmMarshal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sConnString As String

Private Sub cmdMarshal_Click()
Dim adoMainRecordset As ADODB.Recordset
Dim strMessage As String
Dim iIndx As Integer

Set adoMainRecordset = createRecordSet()

Set adoMainRecordset = changeData(adoMainRecordset)

'-The data has been changed. Which marshaling option?

strMessage = "Edit in progress." & vbCr
strMessage = "Would you like to update the database?"

iIndx = MsgBox(strMessage, vbYesNo + vbQuestion, _
                      "MarshalOptions")

If (iIndx = vbYes) Then

   strMessage = "Would you like to send all the rows " & _
                       "in the recordset back to the server?"
   
   iIndx = MsgBox(strMessage, vbYesNo + vbQuestion, _
                         "MarshalOptions")
   
   If (iIndx = vbYes) Then
      adoMainRecordset.MarshalOptions = adMarshalAll
   Else
      adoMainRecordset.MarshalOptions = adMarshalModifiedOnly
   End If
   
   If (updateData(adoMainRecordset)) Then
     If (iIndx = vbYes) Then
        MsgBox "Database updated - All records marshaled."
     Else
        MsgBox "Database updated - Modified records marshaled."
     End If
     Exit Sub
   Else
    MsgBox "Not updated."
   End If
   
Else

  MsgBox "Database not updated."
End If

End Sub



Private Sub Form_Load()
sConnString = "Provider=Microsoft.Jet.OLEDB.3.51;" & _
"Data Source=C:\begdb\Biblio.mdb"

End Sub
Public Function createRecordSet() As ADODB.Recordset

Dim adoRecordset As ADODB.Recordset
Dim sSqlString As String

On Error GoTo createError:

sSqlString = "SELECT * FROM Authors"

Set adoRecordset = New ADODB.Recordset

adoRecordset.CursorType = adOpenKeyset
adoRecordset.LockType = adLockOptimistic
adoRecordset.CursorLocation = adUseClientBatch

adoRecordset.Open sSqlString, sConnString

Set createRecordSet = adoRecordset
Exit Function

createError:
MsgBox Err.Description
Set createRecordSet = Nothing
Exit Function

End Function
Public Function changeData(adoRecordset As _
                             ADODB.Recordset) As ADODB.Recordset

On Error GoTo changeError:

With adoRecordset
   !author = "New Data"
End With

Set changeData = adoRecordset
Exit Function

changeError:
MsgBox Err.Description
Set changeData = Nothing


End Function
Public Function updateData(adoRecordset As ADODB.Recordset) As Boolean

Dim adoTemp As ADODB.Recordset

On Error GoTo updateError:

Set adoTemp = New ADODB.Recordset

adoTemp.Open adoRecordset, sConnString
adoTemp.UpdateBatch
updateData = True
Exit Function

updateError:
MsgBox "In Update:  " & Err.Description
updateData = False

End Function

