VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPublishers 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Publishers"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6840
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   6840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdButton 
      Caption         =   "D&one"
      Height          =   495
      Index           =   10
      Left            =   5400
      TabIndex        =   20
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "&Find"
      Height          =   495
      Index           =   9
      Left            =   2280
      TabIndex        =   15
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "&Undo"
      Height          =   495
      Index           =   8
      Left            =   2280
      TabIndex        =   12
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "&Delete"
      Height          =   495
      Index           =   7
      Left            =   1200
      TabIndex        =   14
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "&Save"
      Height          =   495
      Index           =   6
      Left            =   1200
      TabIndex        =   11
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "&Edit"
      Height          =   495
      Index           =   5
      Left            =   120
      TabIndex        =   13
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "&Add New"
      Height          =   495
      Index           =   4
      Left            =   120
      TabIndex        =   10
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   ">>|"
      Height          =   495
      Index           =   3
      Left            =   3000
      TabIndex        =   19
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   ">>"
      Height          =   495
      Index           =   2
      Left            =   2040
      TabIndex        =   18
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "<<"
      Height          =   495
      Index           =   1
      Left            =   1080
      TabIndex        =   17
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "|<<"
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   16
      Top             =   4560
      Width           =   975
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   31
      Top             =   5310
      Width           =   6840
      _ExtentX        =   12065
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.TextBox Text1 
      DataField       =   "Comments"
      DataSource      =   "Data1"
      Height          =   855
      Index           =   9
      Left            =   3480
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Tag             =   "1"
      Top             =   3480
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      DataField       =   "Fax"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   8
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   8
      Tag             =   "1"
      Text            =   "XXXXXXXXXXXXXXX"
      Top             =   2760
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      DataField       =   "Telephone"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   7
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   7
      Tag             =   "1"
      Text            =   "XXXXXXXXXXXXXXX"
      Top             =   2760
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      DataField       =   "Zip"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   6
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   6
      Tag             =   "1"
      Text            =   "XXXXXXXXXXXXXXX"
      Top             =   2760
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      DataField       =   "State"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   5
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   5
      Tag             =   "1"
      Text            =   "XXXXXXXXXX"
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      DataField       =   "City"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   4
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   4
      Tag             =   "1"
      Text            =   "XXXXXXXXXXXXXXXXXXXX"
      Top             =   2040
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      DataField       =   "Address"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   3
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   3
      Tag             =   "1"
      Text            =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
      Top             =   2040
      Width           =   4095
   End
   Begin VB.TextBox Text1 
      DataField       =   "Company Name"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   2
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   2
      Tag             =   "1"
      Text            =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
      Top             =   1320
      Width           =   6615
   End
   Begin VB.TextBox Text1 
      DataField       =   "Name"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   1
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   1
      Tag             =   "1"
      Text            =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
      Top             =   600
      Width           =   5415
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0C0C0&
      DataField       =   "PubID"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "XXXXXX"
      Top             =   600
      Width           =   735
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\BegDB\Biblio.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Publishers"
      Top             =   5220
      Visible         =   0   'False
      Width           =   6840
   End
   Begin VB.Label lblRecordCount 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Height          =   255
      Left            =   4560
      TabIndex        =   32
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Comments"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   9
      Left            =   3480
      TabIndex        =   30
      Top             =   3240
      Width           =   3255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Fax"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   8
      Left            =   5040
      TabIndex        =   29
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Telephone"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   7
      Left            =   3240
      TabIndex        =   28
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Zip"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   6
      Left            =   1440
      TabIndex        =   27
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "State"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   26
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "City"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   4
      Left            =   4440
      TabIndex        =   25
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Address"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   24
      Top             =   1800
      Width           =   4095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Company Name"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   23
      Top             =   1080
      Width           =   6615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Name"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   1
      Left            =   1320
      TabIndex        =   22
      Top             =   360
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Publisher's ID"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   21
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "frmPublishers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lTotalRecords As Long
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

Private iEditMode As Integer

Public Sub highLight()

With Screen.ActiveForm
   If (TypeOf .ActiveControl Is TextBox) Then
     .ActiveControl.SelStart = 0
     .ActiveControl.SelLength = Len(.ActiveControl)
   End If
End With

End Sub



Private Sub cmdButton_Click(Index As Integer)

Static vMyBookMark As Variant 'used to bookmark the current record

Select Case Index         'what is the value of the key pressed?
   Case cmdMoveFirst
       Data1.Recordset.MoveFirst
       Call updateButtons
   Case cmdMovePrevious
       Data1.Recordset.MovePrevious
         Call updateButtons
   Case cmdMoveNext
       Data1.Recordset.MoveNext
       Call updateButtons
   Case cmdMoveLast
      Data1.Recordset.MoveLast
       Call updateButtons
   Case cmdAddNew       '-add a new record
       With Data1.Recordset
        If (.EditMode = dbEditNone) Then
             If (lTotalRecords > 0) Then
                  vMyBookMark = .Bookmark
             Else
                 vMyBookMark = ""
             End If
            .AddNew
            iEditMode = 2
            Call updateButtons
            lblRecordCount = "Adding New Record"
        End If
    End With

   
   Case cmdEdit         '-- edit the current record
       With Data1.Recordset
        If (.EditMode = dbEditNone) Then
             vMyBookMark = .Bookmark
            .Edit
            iEditMode = 1
            Call updateButtons
            lblRecordCount = "Editing"
        End If
    End With

   Case cmdSave         '-- save the current record
        Dim bMoveLast As Boolean
     With Data1.Recordset
       If (iEditMode <> dbEditNone) Then
            If iEditMode = dbEditAdd Then
                bMoveLast = True
            Else
                bMoveLast = False
            End If

        iEditMode = 0

        If (iEditMode = dbEditNone) Then
            lTotalRecords = .RecordCount
            If (bMoveLast = True) Then
                .MoveLast
            Else
                .Move 0
            End If

       .Edit
       .Update

       updateButtons True
       End If
       Else
           .Move 0
       End If
     End With

   Case cmdDelete       '-- delete the current record
       Dim iResponse As Integer
    Dim sAskUser As String
    sAskUser = "Are you sure you want to delete this record?"
    iResponse = MsgBox(sAskUser, vbQuestion + vbYesNo + _
              vbDefaultButton2, "Publishers Table")
    If (iResponse = vbYes) Then
      With Data1.Recordset
          .Delete
          lTotalRecords = .RecordCount
          If (lTotalRecords > 0) Then
                If lTotalRecords = 1 Then

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

   Case cmdUndo         '-- undo changes to the current record
   With Data1.Recordset
       If (.EditMode <> dbEditNone) Then
           .CancelUpdate
           iEditMode = 0
           If (Len(vMyBookMark)) Then
              .Bookmark = vMyBookMark
           End If
           updateButtons True
       Else
           .Move 0
       End If
     End With

   Case cmdFind         '-- find a specific record
     Dim iReturn As Integer
  gFindString = ""

  With frmFind
   .addCaption = "Type Publisher Name to find"
   .recordSource = "SELECT Name FROM Publishers ORDER BY Name"
   .Show vbModal
  End With

  If (Len(gFindString) > 0) Then
   With Data1.Recordset
     .FindFirst "Name = '" & gFindString & "' "
       If (.NoMatch) Then
         iReturn = MsgBox("Publisher Name " & gFindString & _
            " was not found.", vbCritical, "Publisher")
       Else
         iReturn = MsgBox("Publisher Name " & gFindString & _
             " was retrieved.", vbInformation, "Publisher")
       End If
   End With
  End If
  updateButtons

   Case cmdDone         '-- Done. Unload the form
        Unload Me
End Select

End Sub

Private Sub Data1_Reposition()
With Data1.Recordset
      lTotalRecords = .RecordCount
      lblRecordCount.Caption = "Publisher " & _
        (.AbsolutePosition + 1) & " of " & lTotalRecords
    ProgressBar1.Value = .PercentPosition
  If (Text1(1).Visible) Then Text1(1).SetFocus
End With
End Sub

Private Sub Data1_Validate(Action As Integer, Save As Integer)
Dim iResponse As Integer

If (Action = vbDataActionUpdate) Then
  If (Len(Text1(1)) = 0) Then
     iResponse = MsgBox("Please enter a Company name.", _
                     vbInformation + vbOKOnly, "Publisher's Table")
     Text1(1).SetFocus
     Save = 0
     Action = 0
   End If
ElseIf (Action = vbDataActionUnload) Then
    Save = 0
End If

End Sub

Private Sub Form_Activate()

Static blnIsOld As Boolean ' the default value of a boolean value is false

If blnIsOld = False Then


    With Data1.Recordset
        .MoveLast
        lTotalRecords = .RecordCount
        .MoveFirst
    End With

    updateButtons True
    blnIsOld = True

End If

End Sub

Public Sub updateButtons(Optional bLockEm As Variant)

'-----------------------------------------------------------------
'- The position of the 0 or 1 in the string represents
'- a specific button in our cmdButton control array.
'----------------------------------------------------------------
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
'   9       find a specific record
'  10       done. Unload the form
'--------------------------------------

Select Case Data1.Recordset.EditMode

   Case dbEditNone   '-no editing taking place, just handle navigation
     If (lTotalRecords > 1) Then
       If (Data1.Recordset.BOF) Or _
          (Data1.Recordset.AbsolutePosition = 0) Then
             navigateButtons ("00111101011")
       ElseIf (Data1.Recordset.EOF) Or _
          (Data1.Recordset.AbsolutePosition = lTotalRecords - 1) Then
             navigateButtons ("11001101011")
       Else
             navigateButtons ("11111101011")
       End If
     ElseIf (lTotalRecords > 0) Then
       navigateButtons ("00001101001")
     Else
       navigateButtons ("00001000001")
     End If
     If (Not IsMissing(bLockEm)) Then
       lockTheControls (bLockEm)
     End If
  Case dbEditInProgress    'we are editing a current record
      Call lockTheControls(False)
      Text1(1).SetFocus
      navigateButtons ("00000010100")
  Case dbEditAdd              'we are adding a new record
      Call lockTheControls(False)
      navigateButtons ("00000010100")
      Text1(1).SetFocus
 End Select

End Sub


Public Sub navigateButtons(sButtonString As String)

'-------------------------------------------------
'-- This routine handles setting the enabled    --
'-- property to true/false on the buttons.  --
'-------------------------------------------------
'-- A string of 0101 passed. If 0, disabled --
'-------------------------------------------------

Dim iIndx As Integer
Dim iButtonLength As Integer

sButtonString = Trim$(sButtonString)
iButtonLength = Len(sButtonString)

For iIndx = 1 To iButtonLength
  If (Mid$(sButtonString, iIndx, 1) = "1") Then
    cmdButton(iIndx - 1).Enabled = True
  Else
    cmdButton(iIndx - 1).Enabled = False
  End If
Next

DoEvents

End Sub


Public Sub lockTheControls(bLocked As Boolean)

Dim iIndx As Integer

With Screen.ActiveForm
For iIndx = 0 To .Controls.Count - 1
  If (.Controls(iIndx).Tag = "1") Then
    If (TypeOf .Controls(iIndx) Is TextBox) Then
      If (bLocked) Then
        .Controls(iIndx).Locked = True
        .Controls(iIndx).BackColor = vbWhite
      Else
        .Controls(iIndx).Locked = False
        .Controls(iIndx).BackColor = vbYellow
      End If
    End If
  End If
Next
End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim iMessage As Integer

If (iEditMode <> dbEditNone) Then
   iMessage = MsgBox("You must complete editing the current record", _
                        vbInformation, "Publishers")
  Cancel = True
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmPublishers = Nothing
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   highLight
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then    ' The ENTER key.
      SendKeys "{tab}"               ' Send the focus to next control.
      KeyAscii = 0                   ' Throw this key away
   End If

End Sub

Private Sub Text1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
With Screen.ActiveForm
  If (Len(.ActiveControl.Text) = .ActiveControl.MaxLength) Then
    SendKeys "{Tab}"
  End If
End With

End Sub
