VERSION 5.00
Begin VB.UserControl dbCtl 
   ClientHeight    =   1455
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6030
   DataSourceBehavior=   1  'vbDataSource
   PropertyPages   =   "dbCtl.ctx":0000
   ScaleHeight     =   97
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   402
   ToolboxBitmap   =   "dbCtl.ctx":0016
   Begin VB.CommandButton cmdButton 
      Caption         =   "&Undo"
      Height          =   495
      Index           =   8
      Left            =   4920
      TabIndex        =   9
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "&Delete"
      Height          =   495
      Index           =   7
      Left            =   3720
      TabIndex        =   8
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "&Save"
      Height          =   495
      Index           =   6
      Left            =   2520
      TabIndex        =   7
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "&Edit"
      Height          =   495
      Index           =   5
      Left            =   1320
      TabIndex        =   6
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "&Add New"
      Height          =   495
      Index           =   4
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   ">>|"
      Height          =   495
      Index           =   3
      Left            =   3000
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   ">>"
      Height          =   495
      Index           =   2
      Left            =   2040
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "<<"
      Height          =   495
      Index           =   1
      Left            =   1080
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "|<<"
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblControl 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   4440
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "dbCtl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit
'-- Default Property Values
Const m_def_connectionString = ""
Const m_def_recordSource = ""

'-- Property Variables. These will be read from the property bag
Dim m_connectionString As String
Dim m_recordSource As String
Dim m_form As Object        '-the form that hosts our control
Dim lTotalRecords As Long   '-holds the current number of records

'-- Keep our control a constant size --
Private Const m_def_Height = 97
Private Const m_def_Width = 402


'-Values for our navigation and editing buttons
Public Enum cmdButtons
  cmdMoveFirst = 0
  cmdMovePrevious = 1
  cmdMoveNext = 2
  cmdMoveLast = 3
  cmdAddNew = 4
  cmdEdit = 5
  cmdSave = 6
  cmdDelete = 7
  cmdUndo = 8
End Enum

'-Values for our current edit status
Private Enum editMode
  nowStatic = 0
  nowEditing = 1
  nowadding = 2
End Enum

Dim editStatus As editMode


'Declare our object variables for the ADO connection
'and the recordset used in the control
Private adoConnection As ADODB.Connection
Private adoRecordset As ADODB.Recordset
Public Event validateRecord(ByVal operation As String, ByRef cancel As Boolean)

Const m_def_Tag = "no tag"
Private m_Tag As String


Private Sub cmdButton_Click(Index As Integer)
Static vMyBookmark As Variant
Dim bCancel As Boolean

'-- sanity check here --
If adoRecordset Is Nothing Then Exit Sub

Select Case Index
 Case cmdMoveFirst      '--- move first ---
    adoRecordset.MoveFirst
    editStatus = nowStatic
    Call updateButtons
    lblControl = "Record " & adoRecordset.AbsolutePosition & _
    " of " & lTotalRecords
 Case cmdMovePrevious  '--- move previous ---
    adoRecordset.MovePrevious
    editStatus = nowStatic
    Call updateButtons
    lblControl = "Record " & adoRecordset.AbsolutePosition & _
    " of " & lTotalRecords

 Case cmdMoveNext      '--- move next ---
    adoRecordset.MoveNext
    editStatus = nowStatic
    Call updateButtons
    lblControl = "Record " & adoRecordset.AbsolutePosition & _
    " of " & lTotalRecords

 Case cmdMoveLast      '-- move last ---
    adoRecordset.MoveLast
    editStatus = nowStatic
    Call updateButtons
    lblControl = "Record " & adoRecordset.AbsolutePosition & _
    " of " & lTotalRecords

 '-- Now we are modifying the database --
 Case cmdAddNew       '-- add a new record
    RaiseEvent validateRecord("Add", bCancel)
    If (bCancel = True) Then Exit Sub

    editStatus = nowadding
    With adoRecordset
      If (.RecordCount > 0) Then
        If (.BOF = False) And (.EOF = False) Then
          vMyBookmark = .Bookmark
        Else
          vMyBookmark = ""
         End If
      Else
          vMyBookmark = ""
      End If
      .AddNew
      lblControl = "Adding New Record"
      Call updateButtons
    End With

 Case cmdEdit '-- edit the current record
    RaiseEvent validateRecord("Edit", bCancel)
    If (bCancel = True) Then Exit Sub
     editStatus = nowEditing
     With adoRecordset
        vMyBookmark = adoRecordset.Bookmark
       'We just change the value with ado
        lblControl = "Editing Record"
        Call updateButtons
    End With

 Case cmdSave '-- save the current record
     Dim bMoveLast As Boolean
     RaiseEvent validateRecord("Save", bCancel)
     If (bCancel = True) Then Exit Sub
     
     With adoRecordset
         If .editMode = adEditAdd Then
             bMoveLast = True
         Else
             bMoveLast = False
         End If
         .Move 0
         .Update
         editStatus = nowStatic
         If (bMoveLast = True) Then
            .MoveLast
         Else
            .Move 0
         End If
         editStatus = nowStatic
         lTotalRecords = adoRecordset.RecordCount
         updateButtons True
         lblControl = "New Record Saved"
     End With '

 Case cmdDelete  '-- delete the current record
    Dim iResponse As Integer
    Dim sAskUser As String
    
    RaiseEvent validateRecord("Delete", bCancel)
    If (bCancel = True) Then Exit Sub
    
    sAskUser = "Are you sure you want to delete this record?"
    iResponse = MsgBox(sAskUser, vbQuestion + vbYesNo _
       + vbDefaultButton2, Ambient.DisplayName)
    If (iResponse = vbYes) Then
      With adoRecordset
          .Delete
          If (adoRecordset.RecordCount > 0) Then
            If .BOF Then
              .MoveFirst
           Else
             .MovePrevious
          End If
          lTotalRecords = adoRecordset.RecordCount
          lblControl = "Record Deleted"
        End If
      End With
   End If
   editStatus = nowStatic
   Call updateButtons '
   
 Case cmdUndo '-- undo changes to the current record
    RaiseEvent validateRecord("Undo", bCancel)
    If (bCancel = True) Then Exit Sub
    
    With adoRecordset
        
       If editStatus = nowEditing Then
           .Move 0
           .Bookmark = vMyBookmark
        End If
        .CancelUpdate
        If editStatus = nowEditing Then
           .Move 0
        Else
          If Len(vMyBookmark) Then
            .Bookmark = vMyBookmark
          Else
            If .RecordCount > 0 Then
              .MoveFirst
            End If
          End If
        End If
        lblControl = "Cancelled"
     End With
     editStatus = nowStatic
     updateButtons True
     
End Select

End Sub



Private Sub UserControl_GetDataMember(DataMember As String, Data As Object)
Dim iReturn As Integer

On Error GoTo ohno

'-Reasonability test --
If (adoRecordset Is Nothing) Or (adoConnection Is Nothing) Then
  If Trim$(m_connectionString) = "" Then
    iReturn = MsgBox("There is no connection string!", _
    vbCritical, Ambient.DisplayName)
    Exit Sub
  End If

  If Trim$(m_recordSource) = "" Then
    iReturn = MsgBox("There is no recordsource!", vbCritical, _
                                  Ambient.DisplayName)
    Exit Sub
  End If
  
Set adoConnection = New ADODB.Connection
adoConnection.Open m_connectionString
  
Set adoRecordset = New ADODB.Recordset
adoRecordset.CursorLocation = adUseClient
adoRecordset.CursorType = adOpenDynamic
adoRecordset.LockType = adLockBatchOptimistic

adoRecordset.Open m_recordSource, adoConnection, , , adCmdTable
  
lTotalRecords = adoRecordset.RecordCount
    
adoRecordset.MoveFirst
  
Call cmdButton_Click(cmdMoveFirst)
 
 End If
Set Data = adoRecordset
Exit Sub

ohno:
MsgBox Err.Description
Exit Sub
  

End Sub

Private Sub updateButtons(Optional bLockem As Variant)

'-------------------------------------
'Position   Button
'   0       move first
'   1       move previous
'   2       move next
'   3       move last
'   4       add a new record
'   5       edit the current record
'   6       save the current record
'   7       delete the current record
'   8       undo any current changes
'--------------------------------------
'

'Either we are Editing / Adding or we are not
If (editStatus = nowEditing) Or (editStatus = nowadding) Then
   Call lockTheControls(False)
   navigateButtons ("000000101")
Else
   If (adoRecordset.RecordCount > 2) Then
      If (adoRecordset.BOF) Or _
         (adoRecordset.AbsolutePosition = 1) Then
           navigateButtons ("001111010")
       ElseIf (adoRecordset.EOF) Or _
          (adoRecordset.AbsolutePosition = lTotalRecords) Then
             navigateButtons ("110011010")
       Else
             navigateButtons ("111111010")
       End If
   ElseIf (adoRecordset.RecordCount > 0) Then
       navigateButtons ("000011010")
   Else
       navigateButtons ("000010000")
   End If
     
   If (Not IsMissing(bLockem)) Then
      lockTheControls (bLockem)
   End If
        
End If

End Sub

Private Sub navigateButtons(buttonString As String)

''--------------------------------------------------
''-- This routine handles setting the enabled --
''-- to true / false on the buttons.                --
''-------------------------------------------------
''-- A string of 0101 passed. If 0, disabled   --
''-------------------------------------------------

Dim indx As Integer

buttonString = Trim$(buttonString)

For indx = 1 To Len(buttonString)
  If (Mid$(buttonString, indx, 1) = "1") Then
    cmdButton(indx - 1).Enabled = True
  Else
    cmdButton(indx - 1).Enabled = False
  End If
Next

DoEvents

End Sub

Private Sub lockTheControls(bLocked As Boolean)

On Error Resume Next

Dim iindx As Integer

With m_form

For iindx = 0 To .Controls.Count - 1
If (.Controls(iindx).Tag = Me.Tag) Then
If (TypeOf .Controls(iindx) Is TextBox) Then
If (bLocked) Then
.Controls(iindx).Locked = True
.Controls(iindx).BackColor = vbWhite
Else
.Controls(iindx).Locked = False
.Controls(iindx).BackColor = vbYellow
End If
End If
End If
Next
End With



End Sub


Private Sub UserControl_Initialize()
   editStatus = nowStatic
End Sub

Private Sub UserControl_InitProperties()
   m_recordSource = m_def_recordSource
   m_connectionString = m_def_connectionString

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
m_recordSource = PropBag.ReadProperty("RecordSource", _
m_def_recordSource)
m_connectionString = PropBag.ReadProperty _
("ConnectionString", m_def_connectionString)
m_Tag = PropBag.ReadProperty("Tag", m_def_Tag)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Call PropBag.WriteProperty("RecordSource", _
m_recordSource, m_def_recordSource)
Call PropBag.WriteProperty("ConnectionString", _
m_connectionString, m_def_connectionString)
Call PropBag.WriteProperty("Tag", m_Tag, m_def_Tag)
End Sub


Private Sub UserControl_Resize()
   Width = UserControl.ScaleX(m_def_Width, vbPixels, vbTwips)
   Height = UserControl.ScaleX(m_def_Height, vbPixels, vbTwips)
   Set m_form = UserControl.Parent

End Sub

Private Sub UserControl_Terminate()
On Error Resume Next
If Not adoRecordset Is Nothing Then
  Set adoRecordset = Nothing
End If

If Not adoConnection Is Nothing Then
  Set adoConnection = Nothing
End If

Err.Clear

End Sub



Public Property Get RecordSource() As String
Attribute RecordSource.VB_ProcData.VB_Invoke_Property = "VBDBDataControl"
   RecordSource = m_recordSource
End Property

Public Property Let RecordSource(ByVal New_RecordSource As String)
    m_recordSource = New_RecordSource
    PropertyChanged "RecordSource"
End Property

Public Property Get ConnectionString() As String
Attribute ConnectionString.VB_ProcData.VB_Invoke_Property = "VBDBDataControl"
   ConnectionString = m_connectionString
End Property

Public Property Let ConnectionString(ByVal New_ConnectionString As String)
   m_connectionString = New_ConnectionString
   PropertyChanged "ConnectionString"
End Property

Public Sub showAbout()
Attribute showAbout.VB_Description = "This show the about box for our DB data control."
Attribute showAbout.VB_UserMemId = -552
   frmAbout.Show vbModal
End Sub

Public Property Get Tag() As String
Tag = m_Tag
End Property

Public Property Let Tag(ByVal vNewValue As String)
m_Tag = vNewValue
PropertyChanged "Tag"
End Property


