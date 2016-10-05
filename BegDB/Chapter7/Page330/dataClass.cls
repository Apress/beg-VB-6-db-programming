VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "dataClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "This class is used to manage the functionality of a data control. It can instantiated in a form and will then handle an associated data control."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvardataCtl As Object 'local copy
Private mvarButtons As String 'local copy
Private mvardbName As String 'local copy
Private mvarFormName As Object 'local copy
Private mvarLabelToUpdate As String 'local copy
Private mvarProgressBar As Object 'local copy
Private mvarRecordSource As String 'local copy
Private mvarTag As String 'local copy
Private mvarTotalRecords As Long

Private Enum cmdButtons
  cmdMoveFirst = 0
  cmdMovePrevious = 1
  cmdMoveNext = 2
  cmdMoveLast = 3
  cmdAddNew = 4
  cmdEdit = 5
  cmdSave = 6
  cmdDelete = 7
  cmdUndo = 8
  cmdFind = 9
  cmdDone = 10
End Enum

'local variable(s) to hold property value(s)
Private mvarFindCaption As String 'local copy
Private mvarFindRecordSource As String 'local copy
Private mvarFindMatchField As String 'local copy

Private Function BuildCriteria() As String

 Dim sCriteria As String
 Dim fField As Field
 Dim sMessage As String
 Dim iIndx As Integer
 
 Set fField = mvardataCtl.Recordset.Fields(mvarFindMatchField)
 
 Select Case fField.Type
   Case dbInteger, dbLong, dbCurrency, dbSingle, dbDouble
     BuildCriteria = "" & mvarFindMatchField & " = " & gFindString
   Case dbDate
     BuildCriteria = "" & mvarFindMatchField & " = #" & gFindString & "#"
   Case dbText
     BuildCriteria = "" & mvarFindMatchField & " = '" & gFindString & "'"
   Case Else
     sMessage = "Sorry, you can't use the find feature on fields"
     sMessage = sMessage & " of type: " & fField.Type
     iIndx = MsgBox(sMessage, vbCritical, App.EXEName)
End Select
End Function

Public Property Let FindMatchField(ByVal vData As String)
    mvarFindMatchField = vData
End Property

Public Property Let FindRecordSource(ByVal vData As String)
    mvarFindRecordSource = vData
End Property

Public Property Let FindCaption(ByVal vData As String)
    mvarFindCaption = vData
End Property

Private Sub updateLabel(sLabelString As String)

Dim iIndx As Integer

With mvarFormName
  For iIndx = 0 To .Controls.Count - 1
    If (.Controls(iIndx).Name = mvarLabelToUpdate) Then
        .Controls(iIndx).Caption = sLabelString
        Exit Sub
    End If
  Next
End With

End Sub

Private Sub updateButtons(Optional vLockEm As Variant)
On Error Resume Next
Dim sLabelString As String
 
Select Case mvardataCtl.Recordset.EditMode

   Case dbEditNone
     If (mvarTotalRecords > 1) Then
       If (mvardataCtl.Recordset.BOF) Or _
          (mvardataCtl.Recordset.AbsolutePosition = 0) Then
             navigateButtons ("00111101011")
       ElseIf (mvardataCtl.Recordset.EOF) Or _
(mvardataCtl.Recordset.AbsolutePosition = mvarTotalRecords - 1) Then
             navigateButtons ("11001101011")
       Else
             navigateButtons ("11111101011")
       End If
     ElseIf (mvarTotalRecords > 0) Then
       navigateButtons ("00001101001")
     Else
       navigateButtons ("00001000001")
     End If
     If (Not IsMissing(vLockEm)) Then
       lockTheControls (vLockEm)
     End If
  Case dbEditInProgress    'we are editing a current record
     Call lockTheControls(False)
     navigateButtons ("00000010100")
  Case dbEditAdd           'we are adding a new record
     Call lockTheControls(False)
     navigateButtons ("00000010100")
 End Select

mvarProgressBar.Value = mvardataCtl.Recordset.PercentPosition

sLabelString = mvardataCtl.Recordset.AbsolutePosition + 1
sLabelString = sLabelString & " of " & mvarTotalRecords
Call updateLabel(sLabelString)
End Sub

Private Sub navigateButtons(sButtonString As String)
Dim iIndx As Integer
Dim iButtonLength As Integer

sButtonString = Trim$(sButtonString)
iButtonLength = Len(sButtonString)

For iIndx = 1 To iButtonLength
  If (Mid$(sButtonString, iIndx, 1) = "1") Then
    mvarFormName.Controls(mvarButtons)(iIndx - 1).Enabled = True
  Else
    mvarFormName.Controls(mvarButtons)(iIndx - 1).Enabled = False
  End If
Next

DoEvents

End Sub

Private Sub lockTheControls(bLocked As Boolean)

On Error Resume Next

Dim iIndx As Integer

With mvarFormName
  For iIndx = 0 To .Controls.Count - 1
    If (.Controls(iIndx).Tag = mvarTag) Then
      If (TypeOf .Controls(iIndx) Is TextBox) Then
        If (bLocked) Then
          .Controls(iIndx).Locked = True
          .Controls(iIndx).BackColor = vbWhite
        Else
          .Controls(iIndx).Locked = False
          .Controls(iIndx).BackColor = vbYellow
        End If
      ElseIf (TypeOf .Controls(iIndx) Is DBGrid) Then
       If (bLocked) Then
        .Controls(iIndx).AllowAddNew = False
        .Controls(iIndx).AllowDelete = False
        .Controls(iIndx).AllowUpdate = False
        .Controls(iIndx).BackColor = vbWhite
       Else
        .Controls(iIndx).AllowAddNew = True
        .Controls(iIndx).AllowDelete = True
        .Controls(iIndx).AllowUpdate = True
        .Controls(iIndx).BackColor = vbYellow
       End If

      End If
    End If
  Next
End With

End Sub

Public Sub ProcessCMD(Index As Integer)
Static vMyBookMark As Variant
On Error Resume Next

Select Case Index
    
 Case cmdMoveFirst      '--- move first ---
    mvardataCtl.Recordset.MoveFirst
    Call updateButtons

 Case cmdMovePrevious       '--- move previous ---
    mvardataCtl.Recordset.MovePrevious
    Call updateButtons

 Case cmdMoveNext       '--- move next ---
    mvardataCtl.Recordset.MoveNext
    Call updateButtons

 Case cmdMoveLast   '-- move last ---
    mvardataCtl.Recordset.MoveLast
    Call updateButtons

 '-- Now we are modifying the database --
 Case cmdAddNew  '-- add a new record
    With mvardataCtl.Recordset
        If (.EditMode = dbEditNone) Then
             If (mvarTotalRecords > 0) Then
               vMyBookMark = .Bookmark
             Else
               vMyBookMark = ""
             End If
             .AddNew
             Call updateButtons
             updateLabel ("Adding New Record")
        End If
    End With

 Case cmdEdit '-- edit the current record
     With mvardataCtl.Recordset
        If (.EditMode = dbEditNone) Then
            vMyBookMark = .Bookmark
            .Edit
            Call updateButtons
            updateLabel ("Editing")
        End If
    End With

 Case cmdSave '-- save the current record
     Dim bMoveLast As Boolean
     With mvardataCtl.Recordset
     If (.EditMode <> dbEditNone) Then
        If .EditMode = dbEditAdd Then
          bMoveLast = True
        Else
          bMoveLast = False
        End If
        .Update
        If (.EditMode = dbEditNone) Then
          mvarTotalRecords = .RecordCount
          If (bMoveLast = True) Then
             .MoveLast
          Else
             .Move 0
          End If
          updateButtons True
        End If
     Else
        .Move 0
     End If
     End With

 Case cmdDelete  '-- delete the current record
    Dim iResponse As Integer
    Dim sAskUser As String
    sAskUser = "Are you sure you want to delete this record?"
    iResponse = MsgBox(sAskUser, vbQuestion + vbYesNo + _
       vbDefaultButton2, "Publishers Table")
    If (iResponse = vbYes) Then
      With mvardataCtl.Recordset
         .Delete
         mvarTotalRecords = .RecordCount
         If (mvarTotalRecords > 0) Then
           If mvarTotalRecords = 1 Then

                .MoveNext


            ElseIf .BOF Then

             .MoveFirst
           Else
             .MovePrevious
           End If
         End If
      End With
    End If
    Call updateButtons

 Case cmdUndo '-- undo changes to the current record
     With mvardataCtl.Recordset
       If (.EditMode <> dbEditNone) Then
           .CancelUpdate
           If (Len(vMyBookMark)) Then
              .Bookmark = vMyBookMark
           End If
           updateButtons True
       Else
           .Move 0
       End If
     End With
     updateButtons True

  Case cmdFind '-- find a specific record
     Dim iReturn As Integer
     gFindString = ""

     With frmFind
       .addCaption = mvarFindCaption
       .RecordSource = mvarFindRecordSource
       .Show vbModal
     End With

    If (Len(gFindString) > 0) Then
        Dim sResult As String
        With mvardataCtl.Recordset
            sResult = BuildCriteria()
            If Len(sResult) > 0 Then
                .FindFirst sResult
                If (.NoMatch) Then
                    iReturn = MsgBox(mvarFindMatchField & " " & gFindString & _
                        " was not found.", vbCritical, App.EXEName)
                Else
                    iReturn = MsgBox(mvarFindMatchField & " " & gFindString & _
                        " was retrieved.", vbInformation, App.EXEName)
                End If
            End If
        End With
    End If
    updateButtons

  Case cmdDone  '-- Done. Unload the form
     Unload mvarFormName
End Select
End Sub


Public Property Let Tag(ByVal vData As String)
Attribute Tag.VB_Description = "The ID number of the bound controls"
    mvarTag = vData
End Property

Public Property Let RecordSource(ByVal vData As String)
Attribute RecordSource.VB_Description = "The SQL command to retrieve the recordset"
    mvarRecordSource = vData
    
    If (mvardataCtl Is Nothing) Then
       Exit Property
    End If
    
    mvardataCtl.DatabaseName = mvardbName
    mvardataCtl.RecordSource = mvarRecordSource
    mvardataCtl.Refresh
    
'-- Are there any records? --
    If (mvardataCtl.Recordset.RecordCount < 1) Then
      mvarTotalRecords = 0
      Exit Property
    End If

'-- There is at least 1 record --
    Screen.MousePointer = vbHourglass
    mvardataCtl.Recordset.MoveLast
    mvarTotalRecords = mvardataCtl.Recordset.RecordCount
    mvardataCtl.Recordset.MoveFirst
    updateButtons  'added for child
    Screen.MousePointer = vbDefault

End Property

Public Property Set ProgressBar(ByVal vData As Object)
Attribute ProgressBar.VB_Description = "Will manage the progress bar"
    Set mvarProgressBar = vData
End Property

Public Property Let LabelToUpdate(ByVal vData As String)
Attribute LabelToUpdate.VB_Description = "The label on the form to put status information"
    mvarLabelToUpdate = vData
End Property

Public Property Set FormName(ByVal vData As Object)
Attribute FormName.VB_Description = "The form that hosts the data control"
    Set mvarFormName = vData
End Property

Public Property Let dbName(ByVal vData As String)
Attribute dbName.VB_Description = "The name of the database used"
    mvardbName = vData
End Property

Public Property Let Buttons(ByVal vData As String)
Attribute Buttons.VB_Description = "The string to handle the buttons in the array"
    mvarButtons = vData
End Property

Public Property Set dataCtl(ByVal vData As Object)
Attribute dataCtl.VB_Description = "The data control that the class is operating on."
    Set mvardataCtl = vData
    mvardataCtl.RecordsetType = 1  'ensure a dynaset for absposition
End Property
