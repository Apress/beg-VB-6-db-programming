VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 1  'vbDataSource
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "myBoundClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private showInfo As Boolean
Private WithEvents adoPublishers As ADODB.Recordset
Attribute adoPublishers.VB_VarHelpID = -1
Private adoConnection As ADODB.Connection

Private Sub adoPublishers_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)

Dim sMessage As String

Dim iIndx As Integer

If showInfo = True Then
  
  sMessage = "Reason: "
  Select Case adReason
    Case adRsnAddNew
      sMessage = sMessage & "Adding New Record" & vbCrLf
    Case adRsnClose
      sMessage = sMessage & "Closing Recordset" & vbCrLf
    Case adRsnDelete
      sMessage = sMessage & "Deleting Record" & vbCrLf
    Case adRsnFirstChange
      sMessage = sMessage & "First Change" & vbCrLf
    Case adRsnMove
      sMessage = sMessage & "Move" & vbCrLf
    Case adRsnMoveFirst
      sMessage = sMessage & "Move First Record" & vbCrLf
    Case adRsnMoveLast
      sMessage = sMessage & "Move Last Record" & vbCrLf
    Case adRsnMoveNext
      sMessage = sMessage & "Move Next Record" & vbCrLf
    Case adRsnMovePrevious
      sMessage = sMessage & "Move Previous" & vbCrLf
    Case adRsnRequery
      sMessage = sMessage & "Requering" & vbCrLf
    Case adRsnResynch
      sMessage = sMessage & "Resynch" & vbCrLf
    Case adRsnUndoAddNew
      sMessage = sMessage & "Undo Adding New Record" & vbCrLf
    Case adRsnUndoUpdate
      sMessage = sMessage & "Undoing Update" & vbCrLf
    Case adRsnUpdate
      sMessage = sMessage & "Updating record" & vbCrLf
  End Select

  sMessage = sMessage & "Status: "
  Select Case adStatus
    Case adStatusCancel
      sMessage = sMessage & "Cancel" & vbCrLf
    Case adStatusCantDeny
      sMessage = sMessage & "Cant Deny" & vbCrLf
    Case adStatusErrorsOccurred
      sMessage = sMessage & "Errors Occured" & vbCrLf
    Case adStatusOK
      sMessage = sMessage & "Ok" & vbCrLf
    Case adStatusUnwantedEvent
      sMessage = sMessage & "Unwanted Event" & vbCrLf
  End Select
    
  iIndx = MsgBox(sMessage, vbOKOnly + vbInformation, _
"Bound Data Class MoveComplete")
  
End If
  

End Sub

Private Sub Class_GetDataMember(DataMember As String, Data As Object)
Set Data = adoPublishers
End Sub

Private Sub Class_Initialize()
showInfo = False
DataMembers.Add "Publishers"
  
Set adoPublishers = New ADODB.Recordset
Set adoConnection = New ADODB.Connection

With adoConnection
  .Provider = "Microsoft.Jet.OLEDB.3.51"
  .Open "C:\BegDb\Biblio.mdb"
End With

adoPublishers.Open "Select * FROM Publishers", _
                   adoConnection, _
                   adOpenStatic, _
                   adLockOptimistic

End Sub


Public Sub MoveNext()

If adoPublishers.EOF Then
   adoPublishers.MoveFirst
Else
   adoPublishers.MoveNext
End If


End Sub

Public Property Let displayMoveComplete(ByVal vNewValue As Boolean)
  showInfo = vNewValue
End Property

