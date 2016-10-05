VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form frmCall 
   Caption         =   "New Call"
   ClientHeight    =   6075
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9060
   Icon            =   "frmCall.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   6075
   ScaleWidth      =   9060
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCallCancel 
      Caption         =   "Canc&el"
      Height          =   495
      Left            =   4800
      TabIndex        =   11
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton cmdCallAdd 
      Caption         =   "Add Ca&ll"
      Height          =   495
      Left            =   2520
      TabIndex        =   10
      Top             =   5400
      Width           =   1215
   End
   Begin MSACAL.Calendar Calendar1 
      Height          =   3135
      Left            =   4080
      TabIndex        =   9
      Top             =   2040
      Width           =   4935
      _Version        =   524288
      _ExtentX        =   8705
      _ExtentY        =   5530
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   1998
      Month           =   7
      Day             =   23
      DayLength       =   1
      MonthLength     =   2
      DayFontColor    =   0
      FirstDay        =   1
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtNotesOnPhone 
      Height          =   2055
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   3000
      Width           =   3735
   End
   Begin VB.Frame Frame1 
      Caption         =   "Call Types"
      Height          =   1335
      Left            =   2880
      TabIndex        =   2
      Top             =   240
      Width           =   6015
      Begin VB.CommandButton cmdCancelType 
         Caption         =   "&Cancel"
         Height          =   495
         Left            =   4560
         TabIndex        =   6
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton cmdSaveType 
         Caption         =   "&Save"
         Height          =   495
         Left            =   3120
         TabIndex        =   5
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton cmdDeleteType 
         Caption         =   "&Delete"
         Height          =   495
         Left            =   1680
         TabIndex        =   4
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton cmdAddType 
         Caption         =   "&Add"
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.ComboBox cboDescription 
      Height          =   315
      Left            =   360
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   720
      Width           =   2295
   End
   Begin VB.Line Line2 
      X1              =   8880
      X2              =   120
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Line Line1 
      X1              =   8880
      X2              =   120
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      Caption         =   "lblCaption"
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   240
      TabIndex        =   7
      Top             =   2160
      Width           =   3495
   End
   Begin VB.Label lblDescription 
      Alignment       =   2  'Center
      Caption         =   "Call Description"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   2055
   End
End
Attribute VB_Name = "frmCall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsCallType As Recordset
Dim currentSelection As Long
Dim bDescriptionChanged As Boolean
Dim sCurrentDescription As String
Dim iCurrentState As Integer
Public sContactName As String
    'looks like a property from the outside
Public lContactNumber As Long
    'also looks like a property


Private Sub cboDescription_Change()
 bDescriptionChanged = True   '-set our form-level Boolean
 If (iCurrentState <> NOW_ADDING) Then
   iCurrentState = NOW_EDITING
 End If
 Call updateButtons         '-our centralised button handler

End Sub

Private Sub cboDescription_Click()
 sCurrentDescription = Trim$(Left$(cboDescription.Text, 20))
 currentSelection = cboDescription.ItemData _
    (cboDescription.ListIndex)

End Sub

Private Sub cmdAddType_Click()

iCurrentState = NOW_ADDING
Call updateButtons
cboDescription.Locked = False
cboDescription.Text = ""
cboDescription.SetFocus

End Sub

Private Sub cmdCallAdd_Click()
Dim indx As Integer
Dim rsNotes As Recordset

'-Sanity check. Don't save a blank message
If (Len(txtNotesOnPhone) < 1) Then
  indx = MsgBox("Please enter the discussion in the text box", _
                 vbOKOnly + vbInformation, progname)
  Exit Sub
End If

Screen.MousePointer = vbHourglass
DoEvents

'-Open the Notes table to accept our new entry
Set rsNotes = dbContact.OpenRecordset("Notes", dbOpenTable)
rsNotes.AddNew
rsNotes!ContactID = lContactNumber
rsNotes!DateOfCall = Calendar1.Value
rsNotes!CallTypeID = currentSelection
rsNotes!NotesOnPhoneCall = txtNotesOnPhone
rsNotes.Update

Me.Hide
DoEvents

Unload Me
Screen.MousePointer = vbDefault

End Sub

Private Sub cmdCallCancel_Click()
Unload Me
End Sub

Private Sub cmdCancelType_Click()
iCurrentState = NOW_IDLE
Call updatedescriptionCombo

End Sub

Private Sub cmdDeleteType_Click()
Dim indx As Integer
Dim sMessage As String

iCurrentState = NOW_DELETING
Call updateButtons

If (Len(cboDescription.Text) < 1) Then
  indx = MsgBox("Please choose an entry to delete", _
    vbOKOnly + vbInformation, progname)
  iCurrentState = NOW_IDLE
  Call updateButtons
  Exit Sub
ElseIf (cboDescription.ListCount < 1) Then
  indx = MsgBox("There are no entries to delete", _
      vbOKOnly + vbInformation, progname)
  iCurrentState = NOW_IDLE
  Call updateButtons
  Exit Sub
Else
  sMessage = "Do you wish to delete the call " & vbCrLf
  sMessage = sMessage & "type " & cboDescription.Text & vbCrLf
  sMessage = sMessage & "and all associated calls in your"
  sMessage = sMessage & " database?" & vbCrLf
  indx = MsgBox(sMessage, vbYesNo + vbCritical, progname)
  
  If indx <> vbYes Then
    iCurrentState = NOW_IDLE
    Call updateButtons
    Exit Sub
  End If
  
'OK, now we delete the record
  Dim lKeyToDelete As Long
  Dim sDeleteSQL As String
  lKeyToDelete = cboDescription.ItemData _
    (cboDescription.ListIndex)
  sDeleteSQL = "DELETE * FROM CallType WHERE CallTypeID =  " _
    & lKeyToDelete
  dbContact.Execute (sDeleteSQL)
  
  iCurrentState = NOW_IDLE
  Call updatedescriptionCombo
  Call updateButtons
  
End If

End Sub

Private Sub cmdSaveType_Click()
Dim indx As Integer
Dim sqlMaxID As String
Dim rsMaxIDNumber As Recordset
Dim lNewTypeID As Long
Dim sMessage As String

If (Len(cboDescription.Text) < 1) Then
  sMessage = "Please enter the Call Type Description in "
  sMessage = sMessage & "the combo box."
  indx = MsgBox(sMessage, vbOKOnly + vbInformation, progname)
  Exit Sub
End If

'-If the user is not adding a new description, we know that
'-the user wishes to edit a current entry. Checking the
'-iCurrentState tells us which mode

If (iCurrentState <> NOW_ADDING) Then
  indx = MsgBox("Change all entries for '" & _
    sCurrentDescription & "' to '" & cboDescription.Text & _
    "'?", vbYesNo + vbQuestion, progname)
Else
  indx = MsgBox("Add call type: " & cboDescription.Text, _
    vbYesNo + vbQuestion, progname)
End If

'-the user aborts the change. Restore the form
If (indx <> vbYes) Then
  iCurrentState = NOW_IDLE
  Call updatedescriptionCombo
  Exit Sub
End If

'-Otherwise update the table with either a new record or
'-change the description field if this is an edit.
Set rsCallType = dbContact.OpenRecordset _
    ("CallType", dbOpenTable)

If (iCurrentState = NOW_ADDING) Then
  sqlMaxID = "SELECT Max(CallType.CallTypeID) AS LastType"
  sqlMaxID = sqlMaxID & " FROM CallType"
  Set rsMaxIDNumber = dbContact.OpenRecordset(sqlMaxID)

  rsCallType.AddNew
lNewTypeID = rsCallType!CallTypeID

  rsCallType!CallDescription = Trim$ _
    (Left$(cboDescription.Text, 20))
  currentSelection = lNewTypeID
Else
  rsCallType.Edit
  rsCallType!CallDescription = sCurrentDescription
  currentSelection = rsCallType!CallTypeID
End If
rsCallType.Update

cmdSaveType.Enabled = False

Call updatedescriptionCombo

iCurrentState = NOW_IDLE

End Sub

Private Sub Form_Activate()
 currentSelection = 0
 
 '-open our form-level recordset using dbOpenTable
 Set rsCallType = dbContact.OpenRecordset("CallType", dbOpenTable)
 
 '-Set the focus to our Notes text box
 txtNotesOnPhone.SetFocus
 
 iCurrentState = NOW_IDLE
 
 '-Load our combo box with all current call descriptions
 Call updatedescriptionCombo
 
 '-Initialize our calendar control to today
 Calendar1.Day = Day(Now)
 Calendar1.Month = Month(Now)
 Calendar1.Year = Year(Now)
 
 lblCaption = "Enter the notes on the call to " & sContactName
 
 '-force all visual changes to the screen to look snappy
 DoEvents

End Sub


 Public Sub updatedescriptionCombo()
 
 Dim indx As Integer
 
 cboDescription.Clear
 
 If (rsCallType.RecordCount < 1) Then
   cboDescription.AddItem "<No Entries>"
   cboDescription.Locked = True
   cmdDeleteType.Enabled = False
   txtNotesOnPhone.Enabled = False
   cmdCallAdd.Enabled = False
 Else
   rsCallType.MoveFirst
   While (Not rsCallType.EOF)
     cboDescription.AddItem rsCallType!CallDescription
     cboDescription.ItemData(cboDescription.NewIndex) = _
         rsCallType!CallTypeID
     rsCallType.MoveNext
   Wend
   txtNotesOnPhone.Enabled = True
   cmdCallAdd.Enabled = True
 End If
 
 cmdSaveType.Enabled = False
 cmdCancelType.Enabled = False
 
 If iCurrentState = NOW_ADDING Then
   For indx = 0 To (cboDescription.ListCount - 1)
      cboDescription.ListIndex = indx
      If (cboDescription.ItemData(cboDescription.ListIndex) = _
             currentSelection) Then
           Exit For
      End If
    Next
 Else
    cboDescription.ListIndex = 0
 End If
 
 bDescriptionChanged = False
 sCurrentDescription = cboDescription.Text
 iCurrentState = NOW_IDLE
 Call updateButtons
 
 End Sub


Public Sub updateButtons()

Select Case iCurrentState
  Case NOW_ADDING
    cmdAddType.Enabled = False
    cmdDeleteType.Enabled = False
    cmdSaveType.Enabled = True
    cmdCancelType.Enabled = True
  Case NOW_DELETING
    cmdAddType.Enabled = False
    cmdDeleteType.Enabled = False
    cmdSaveType.Enabled = True
    cmdCancelType.Enabled = True
  Case NOW_EDITING
    cmdAddType.Enabled = False
    cmdDeleteType.Enabled = False
    cmdSaveType.Enabled = True
    cmdCancelType.Enabled = True
  Case NOW_IDLE
    cmdAddType.Enabled = True
    If (rsCallType.RecordCount < 1) Then
      cmdDeleteType.Enabled = False
    Else
      cmdDeleteType.Enabled = True
    End If
    cmdSaveType.Enabled = False
    cmdCancelType.Enabled = False
End Select

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmCall = Nothing
End Sub

