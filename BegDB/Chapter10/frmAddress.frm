VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmAddress 
   Caption         =   "VB 6.0 Address Book"
   ClientHeight    =   6885
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9885
   Icon            =   "frmAddress.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6885
   ScaleWidth      =   9885
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   600
      Top             =   4080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddress.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddress.frx":0896
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddress.frx":0CEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddress.frx":113E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddress.frx":1592
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddress.frx":19E6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sbStatus 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   6390
      Width           =   9885
      _ExtentX        =   17436
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9313
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "18:02"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "14/07/99"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvContact 
      Height          =   5415
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   9551
      _Version        =   393217
      LineStyle       =   1
      Style           =   6
      Appearance      =   1
   End
   Begin TabDlg.SSTab tbContact 
      Height          =   5415
      Left            =   3600
      TabIndex        =   1
      Top             =   840
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   9551
      _Version        =   393216
      Style           =   1
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Contact"
      TabPicture(0)   =   "frmAddress.frx":1E3A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "mskBirthday"
      Tab(0).Control(1)=   "txtHomeEmail"
      Tab(0).Control(2)=   "mskHomeCellPhone"
      Tab(0).Control(3)=   "mskHomeFax"
      Tab(0).Control(4)=   "mskHomePhone"
      Tab(0).Control(5)=   "mskHomeZip"
      Tab(0).Control(6)=   "txtHomeState"
      Tab(0).Control(7)=   "txtHomeCity"
      Tab(0).Control(8)=   "txtHomeStreet"
      Tab(0).Control(9)=   "txtLastName"
      Tab(0).Control(10)=   "txtMiddleInitial"
      Tab(0).Control(11)=   "txtFirstName"
      Tab(0).Control(12)=   "Label13"
      Tab(0).Control(13)=   "Label12"
      Tab(0).Control(14)=   "Label11"
      Tab(0).Control(15)=   "Label10"
      Tab(0).Control(16)=   "Label9"
      Tab(0).Control(17)=   "Label8"
      Tab(0).Control(18)=   "Label7"
      Tab(0).Control(19)=   "Label6"
      Tab(0).Control(20)=   "Label5"
      Tab(0).Control(21)=   "Label4"
      Tab(0).Control(22)=   "Label3"
      Tab(0).Control(23)=   "Label2"
      Tab(0).Control(24)=   "lblBirthday"
      Tab(0).ControlCount=   25
      TabCaption(1)   =   "Call Log"
      TabPicture(1)   =   "frmAddress.frx":1E56
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lvCalls"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "txtNotes"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "DB Statistics"
      TabPicture(2)   =   "frmAddress.frx":1E72
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblDateCreated"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "lblLastUpdated"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "lblRecordCount"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label14"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label15"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Label16"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Frame1"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).ControlCount=   7
      Begin VB.Frame Frame1 
         Caption         =   "Database Stats"
         Height          =   1935
         Left            =   -74760
         TabIndex        =   34
         Top             =   3120
         Width           =   5655
         Begin VB.Label Label22 
            Caption         =   "Locks Released"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2880
            TabIndex        =   49
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label Label21 
            Caption         =   "Locks Placed"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2880
            TabIndex        =   48
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label Label20 
            Caption         =   "Read Ahead"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2880
            TabIndex        =   47
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label Label19 
            Caption         =   "Read Cache"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   240
            TabIndex        =   46
            Top             =   1440
            Width           =   975
         End
         Begin VB.Label Label18 
            Caption         =   "Disk Writes"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   240
            TabIndex        =   45
            Top             =   960
            Width           =   975
         End
         Begin VB.Label Label17 
            Caption         =   "Disk Reads"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   240
            TabIndex        =   44
            Top             =   480
            Width           =   975
         End
         Begin VB.Label lblReleaseLocks 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   4200
            TabIndex        =   40
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label lblLocksPlaced 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   4200
            TabIndex        =   39
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label lblReadAhead 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   4200
            TabIndex        =   38
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label lblReadCache 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1320
            TabIndex        =   37
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label lblDiskWrites 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1320
            TabIndex        =   36
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label lblDiskReads 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1320
            TabIndex        =   35
            Top             =   480
            Width           =   1215
         End
      End
      Begin VB.TextBox txtNotes 
         Height          =   2175
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   2880
         Width           =   5775
      End
      Begin MSComctlLib.ListView lvCalls 
         Height          =   2055
         Left            =   120
         TabIndex        =   29
         Top             =   600
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   3625
         View            =   2
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSMask.MaskEdBox mskBirthday 
         Height          =   285
         Left            =   -74760
         TabIndex        =   15
         Tag             =   "1"
         Top             =   4800
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtHomeEmail 
         Height          =   285
         Left            =   -74760
         MaxLength       =   30
         TabIndex        =   14
         Tag             =   "1"
         Text            =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
         Top             =   3840
         Width           =   3255
      End
      Begin MSMask.MaskEdBox mskHomeCellPhone 
         Height          =   285
         Left            =   -70680
         TabIndex        =   13
         Tag             =   "1"
         Top             =   2880
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   13
         Mask            =   "(###)###-####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskHomeFax 
         Height          =   285
         Left            =   -72720
         TabIndex        =   12
         Tag             =   "1"
         Top             =   2880
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   13
         Mask            =   "(###)###-####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskHomePhone 
         Height          =   285
         Left            =   -74760
         TabIndex        =   11
         Tag             =   "1"
         Top             =   2880
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   13
         Mask            =   "(###)###-####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskHomeZip 
         Height          =   285
         Left            =   -70320
         TabIndex        =   10
         Tag             =   "1"
         Top             =   1920
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "#####-####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtHomeState 
         Height          =   285
         Left            =   -70800
         MaxLength       =   2
         TabIndex        =   9
         Tag             =   "1"
         Text            =   "XX"
         Top             =   1920
         Width           =   375
      End
      Begin VB.TextBox txtHomeCity 
         Height          =   285
         Left            =   -72360
         MaxLength       =   12
         TabIndex        =   8
         Tag             =   "1"
         Text            =   "XXXXXXXXXXXX"
         Top             =   1920
         Width           =   1455
      End
      Begin VB.TextBox txtHomeStreet 
         Height          =   285
         Left            =   -74760
         MaxLength       =   20
         TabIndex        =   7
         Tag             =   "1"
         Text            =   "XXXXXXXXXXXXXXXXXXXX"
         Top             =   1920
         Width           =   2295
      End
      Begin VB.TextBox txtLastName 
         Height          =   285
         Left            =   -71400
         MaxLength       =   20
         TabIndex        =   6
         Tag             =   "1"
         Text            =   "XXXXXXXXXXXXXXXXXXXX"
         Top             =   960
         Width           =   2295
      End
      Begin VB.TextBox txtMiddleInitial 
         Height          =   285
         Left            =   -72360
         MaxLength       =   1
         TabIndex        =   5
         Tag             =   "1"
         Text            =   "X"
         Top             =   960
         Width           =   255
      End
      Begin VB.TextBox txtFirstName 
         Height          =   285
         Left            =   -74760
         MaxLength       =   15
         TabIndex        =   4
         Tag             =   "1"
         Text            =   "XXXXXXXXXXXXXXX"
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label16 
         Caption         =   "Contact Records"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   -74400
         TabIndex        =   43
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label Label15 
         Caption         =   "Last Updated"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   -74400
         TabIndex        =   42
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label14 
         Caption         =   "Database Created"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   -74400
         TabIndex        =   41
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label lblRecordCount 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   -72720
         TabIndex        =   33
         Top             =   2400
         Width           =   3615
      End
      Begin VB.Label lblLastUpdated 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   -72720
         TabIndex        =   32
         Top             =   1800
         Width           =   3615
      End
      Begin VB.Label lblDateCreated 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   -72720
         TabIndex        =   31
         Top             =   1200
         Width           =   3615
      End
      Begin VB.Label Label13 
         Caption         =   "Last Name"
         Height          =   255
         Left            =   -71400
         TabIndex        =   28
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label Label12 
         Caption         =   "M. I."
         Height          =   255
         Left            =   -72360
         TabIndex        =   27
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label11 
         Caption         =   "First Name"
         Height          =   255
         Left            =   -74760
         TabIndex        =   26
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label10 
         Caption         =   "Zip"
         Height          =   255
         Left            =   -70320
         TabIndex        =   25
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "State"
         Height          =   255
         Left            =   -70800
         TabIndex        =   24
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label8 
         Caption         =   "City"
         Height          =   255
         Left            =   -72360
         TabIndex        =   23
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "Home Street"
         Height          =   255
         Left            =   -74760
         TabIndex        =   22
         Top             =   1680
         Width           =   2295
      End
      Begin VB.Label Label6 
         Caption         =   "Cell Phone"
         Height          =   255
         Left            =   -70680
         TabIndex        =   21
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Home Fax"
         Height          =   255
         Left            =   -72720
         TabIndex        =   20
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Home Phone"
         Height          =   255
         Left            =   -74760
         TabIndex        =   19
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Home Email Address"
         Height          =   255
         Left            =   -74760
         TabIndex        =   18
         Top             =   3600
         Width           =   3255
      End
      Begin VB.Label Label2 
         Caption         =   "Birthday"
         Height          =   255
         Left            =   -74760
         TabIndex        =   17
         Top             =   4560
         Width           =   1815
      End
      Begin VB.Label lblBirthday 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   -73080
         TabIndex        =   16
         Tag             =   "1"
         Top             =   4800
         Width           =   3015
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   810
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9885
      _ExtentX        =   17436
      _ExtentY        =   1429
      ButtonWidth     =   1349
      ButtonHeight    =   1376
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Add New"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cancel"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Edit"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Quit"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "myPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuAddNew 
         Caption         =   "&Add New"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete"
      End
   End
End
Attribute VB_Name = "frmAddress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim contactNode As Node
Dim rsNotesTable As Recordset
Dim rsCallType As Recordset

Dim iCurrentState As Integer
Dim lCurrentContactKey As Long
Dim sCurrentContactName As String
Dim bFieldsPopulated As Boolean  'flag to see if the fields have data

Private Sub Form_Activate()
Static bLoadedAlready As Boolean

sbStatus.Panels.Item(2).Text = "Loading...."
  
If (Not bLoadedAlready) Then
  Call initializeForm
  bLoadedAlready = True
End If

sbStatus.Panels.Item(2).Text = "Ready."

End Sub

Private Sub Form_Load()
If (Not openTheDatabase()) Then
  MsgBox "Sorry - the database could not be opened."
  End      'terminate the program unconditionally
End If

'-- Remove the Xs in our text boxes --
Call clearFields
bFieldsPopulated = False
iCurrentState = NOW_IDLE

End Sub

Public Sub initializeForm()

  Screen.MousePointer = vbHourglass  '-show activity is occurring
  iCurrentState = NOW_IDLE  '-set the current state of the prog.
  sbStatus.Panels.Item(2).Text = "Loading..."
  tbContact.Tab = 0          '-- make the 1st of the 3 tabs current
  DoEvents              '-- ensure the visual components are updated
  Call clearFields
  Call lockFields(True)
  Call updateTree
  Call updateForm
  Call setUpListView
  tbContact.Enabled = False
  Screen.MousePointer = vbDefault
  sbStatus.Panels.Item(2).Text = "Ready."

End Sub

Public Sub clearFields()

Dim indx As Integer
Dim tempMask As String

With Me.Controls
  For indx = 0 To .Count - 1
    If Me.Controls(indx).Tag = "1" Then
       If (TypeOf Me.Controls(indx) Is TextBox) Then
           Me.Controls(indx).Text = ""
       ElseIf (TypeOf Me.Controls(indx) Is MaskEdBox) Then
          tempMask = Me.Controls(indx).Mask
          Me.Controls(indx).Mask = ""
          Me.Controls(indx).Text = ""
          Me.Controls(indx).Mask = tempMask
      Else
         Me.Controls(indx).Caption = ""
      End If
    End If
  Next
End With
DoEvents

End Sub

Public Sub lockFields(bDoLock As Boolean)

Dim indx As Integer

For indx = 0 To Me.Controls.Count - 1
  If Me.Controls(indx).Tag = "1" Then
    If (TypeOf Me.Controls(indx) Is TextBox) Then
      If (bDoLock = True) Then
        Me.Controls(indx).Locked = True
        Me.Controls(indx).BackColor = vbWhite
      Else
        Me.Controls(indx).Locked = False
        Me.Controls(indx).BackColor = vbYellow
      End If
    ElseIf (TypeOf Me.Controls(indx) Is MaskEdBox) Then
      If (bDoLock = True) Then
        Me.Controls(indx).Enabled = False
        Me.Controls(indx).BackColor = vbWhite
      Else
        Me.Controls(indx).Enabled = True
        Me.Controls(indx).BackColor = vbYellow
      End If
   End If
 End If
Next
DoEvents

End Sub

Public Sub updateTree()

Dim indx As Integer
Dim rsAllNames As Recordset
Dim sqlNames As String
Dim sContactName As String
Dim currentAlpha As String

tvContact.Nodes.Clear

sqlNames = "SELECT ContactID, LastName, FirstName, MiddleInitial "
sqlNames = sqlNames & "FROM Contact ORDER BY"
sqlNames = sqlNames & " LastName, FirstName, MiddleInitial "

Set rsAllNames = dbContact.OpenRecordset(sqlNames)

If (rsAllNames.RecordCount > 0) Then
  rsAllNames.MoveFirst
End If

For indx = Asc("A") To Asc("Z")
 currentAlpha = Chr(indx)
    
 Set contactNode = tvContact.Nodes.Add _
    (, , currentAlpha, currentAlpha)
  
 If (Not rsAllNames.EOF) Then
  Do While UCase$(Left(rsAllNames!LastName, 1)) = currentAlpha
    With rsAllNames
      sContactName = !LastName & ", "
      sContactName = sContactName & !FirstName
      If (Not IsNull(!MiddleInitial)) Then
       sContactName = sContactName & " " & !MiddleInitial & "."
      End If
      End With

    DoEvents
    
    Set contactNode = tvContact.Nodes.Add(currentAlpha, _
    tvwChild, "ID" & CStr(rsAllNames!ContactID), sContactName)
    rsAllNames.MoveNext
    If (rsAllNames.EOF) Then
      Exit Do
    End If
  Loop
 End If
Next

sbStatus.Panels.Item(1).Text = "There are " & _
    rsAllNames.RecordCount & " contacts in the database."

rsAllNames.Close

DoEvents

End Sub

Public Sub updateForm()

Select Case iCurrentState
  Case NOW_ADDING, NOW_EDITING
    If (iCurrentState = NOW_ADDING) Then
      sbStatus.Panels.Item(2).Text = "Adding..."
      Call clearFields
    Else
      sbStatus.Panels.Item(2).Text = "Editing..."
    End If
    tbContact.Enabled = True
    tbContact.Tab = 0              '-- make the 1st tab current
    tbContact.TabEnabled(1) = False '-disable the 2nd and 3rd tabs
    tbContact.TabEnabled(2) = False
    tvContact.Enabled = False
    lockFields (False)        '-- unlock fields and set background
    txtFirstName.SetFocus     '-- set focus to first name field
    Toolbar1.Buttons(bAdd).Enabled = False
    Toolbar1.Buttons(bCancel).Enabled = True
    Toolbar1.Buttons(bSave).Enabled = True
    Toolbar1.Buttons(bDelete).Enabled = False
    Toolbar1.Buttons(bEdit).Enabled = False
    Toolbar1.Buttons(bQuit).Enabled = False
  Case NOW_IDLE
    sbStatus.Panels.Item(2).Text = "Ready."
    Toolbar1.Buttons(bAdd).Enabled = True
    Toolbar1.Buttons(bCancel).Enabled = False
    Toolbar1.Buttons(bSave).Enabled = False
    Toolbar1.Buttons(bQuit).Enabled = True
    If (Len(txtLastName)) Then
      Toolbar1.Buttons(bDelete).Enabled = True
      Toolbar1.Buttons(bEdit).Enabled = True
    Else
      Toolbar1.Buttons(bDelete).Enabled = False
      Toolbar1.Buttons(bEdit).Enabled = False
    End If
    tvContact.Enabled = True
    tbContact.TabEnabled(1) = True
    tbContact.TabEnabled(2) = True
  Case NOW_DELETING
    sbStatus.Panels.Item(2).Text = "Deleting...."
    Toolbar1.Buttons(bAdd).Enabled = False
    Toolbar1.Buttons(bCancel).Enabled = False
    Toolbar1.Buttons(bSave).Enabled = False
    Toolbar1.Buttons(bDelete).Enabled = False
    Toolbar1.Buttons(bEdit).Enabled = False
    Toolbar1.Buttons(bQuit).Enabled = False
  Case NOW_SAVING
    sbStatus.Panels.Item(2).Text = "Saving...."
    tvContact.Enabled = True
    Toolbar1.Buttons(bAdd).Enabled = False
    Toolbar1.Buttons(bCancel).Enabled = False
    Toolbar1.Buttons(bSave).Enabled = False
    Toolbar1.Buttons(bDelete).Enabled = False
    Toolbar1.Buttons(bEdit).Enabled = False
    Toolbar1.Buttons(bQuit).Enabled = False
    If (Len(mskBirthday)) Then
     lblBirthday = Format$(mskBirthday, "mmmm dd, yyyy")
    End If
End Select

DoEvents

End Sub


Public Sub setUpListView()

Dim clmHdr As ColumnHeader

Set clmHdr = lvCalls.ColumnHeaders. _
             Add(, , "Date / Time", lvCalls.Width \ 3)

Set clmHdr = lvCalls.ColumnHeaders. _
             Add(, , "Type of Call", lvCalls.Width \ 3)
             
Set clmHdr = lvCalls.ColumnHeaders. _
             Add(, , "Call Identifier", lvCalls.Width \ 3)
lvCalls.View = lvwReport

End Sub


Private Sub lvCalls_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Dim nSortCol As Integer
        
' When a ColumnHeader object is clicked, the list view
' control is sorted by the SubItems of that column.
' Set the SortKey to the index of the ColumnHeader - 1

nSortCol = ColumnHeader.Index - 1
    
If (lvCalls.SortKey = nSortCol) Then
   lvCalls.SortOrder = 1 - lvCalls.SortOrder
Else
   lvCalls.SortKey = nSortCol
   lvCalls.SortOrder = lvwAscending
End If
    
'-- Do the sort now
lvCalls.Sorted = True

End Sub

Private Sub lvCalls_ItemClick(ByVal Item As MSComctlLib.ListItem)
If (rsCallType.RecordCount > 0) Then
    rsCallType.MoveFirst
    '-- Find the record that has the ID --
    rsCallType.FindFirst "CallCounter = " & _
                 lvCalls.ListItems(Item.Index).SubItems(2)
     txtNotes = rsCallType!NotesOnPhoneCall
End If

End Sub


Private Sub lvCalls_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then

     If (rsCallType.RecordCount < 1) Then
       mnuDelete.Enabled = False
     Else
       mnuDelete.Enabled = True
     End If
     PopupMenu mnuPopup
End If

End Sub

Private Sub mnuAddNew_Click()
 frmCall.sContactName = sCurrentContactName
 frmCall.lContactNumber = lCurrentContactKey
 frmCall.Show vbModal
 Call populateListView

End Sub

Private Sub mnuDelete_Click()
Dim indx As Integer
Dim rsDeleteCall As Recordset
Dim sDeleteCall As String

indx = MsgBox("Are you sure you wish to delete this call from " & _
              lvCalls.ListItems(lvCalls.SelectedItem.Index) & "?", _
              vbYesNo + vbQuestion, progname)


If (indx <> vbYes) Then Exit Sub

sDeleteCall = "DELETE * FROM Notes WHERE CallCounter = " & _
              lvCalls.ListItems(lvCalls.SelectedItem.Index).SubItems(2)

dbContact.Execute (sDeleteCall)
Call populateListView

End Sub

Private Sub tbContact_Click(PreviousTab As Integer)
If (tbContact.Tab = 2) Then
   lblDateCreated = Format$(rsContactTable.DateCreated, _
    "dddd mmmm dd, yyyy hh:mm AMPM")
   lblLastUpdated = Format$(rsContactTable.LastUpdated, _
    "dddd mmmm dd, yyyy hh:mm AMPM")
   lblRecordCount = "Contacts in Database: " & _
    rsContactTable.RecordCount
   lblDiskReads = ISAMStats(0)
   lblDiskWrites = ISAMStats(1)
   lblReadCache = ISAMStats(2)
   lblReadAhead = ISAMStats(3)
   lblLocksPlaced = ISAMStats(4)
   lblReleaseLocks = ISAMStats(5)
End If

End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
 Case bAdd  '-- Add New

     iCurrentState = NOW_ADDING
     Call updateForm

 Case bCancel  '-Cancel

     If (bFieldsPopulated = True) Then
       Call populateFields
    End If
    Call lockFields(True)
    iCurrentState = NOW_IDLE
    Call updateForm

 Case bSave '-Saving

'-- Here we are saving either a new or edited entry --
    If (Not validateEntry()) Then
       Exit Sub
    End If
    postContact
 
Case bDelete  '-Deleting

    Dim indx As Integer
    Dim sMsg As String
    Dim sDeleteSQL As String
    sMsg = "Delete " & tvContact.SelectedItem & _
    " and all related call logs?"
    indx = MsgBox(sMsg, vbYesNo + vbCritical, progname)
    If (indx <> vbYes) Then Exit Sub
    sDeleteSQL = "DELETE * FROM Contact WHERE ContactID = " _
    & lCurrentContactKey
    dbContact.Execute (sDeleteSQL)

    Call initializeForm

Case bEdit  '-- Editing

    iCurrentState = NOW_EDITING
    updateForm

Case bQuit  '-- Quitting

   rsContactTable.Close
   dbContact.Close
   Set rsContactTable = Nothing
   Set dbContact = Nothing
   Unload Me

End Select

End Sub

Private Sub tvContact_NodeClick(ByVal Node As MSComctlLib.Node)
If (Len(Node.Key) = 1) Then Exit Sub

'-- Here we retrieve the contact the user clicked on --
lCurrentContactKey = CLng(Mid$(Node.Key, 3, Len(Node.Key)))
With rsContactTable
   .Index = "PrimaryKey"
   .Seek "=", lCurrentContactKey
   If Not .NoMatch Then
     bFieldsPopulated = True
     sCurrentContactName = tvContact.SelectedItem
     Call populateFields
     Call populateListView
     tbContact.Enabled = True
   Else
     MsgBox ("Ohhhh Nooo")
   End If
End With

End Sub


Public Sub populateFields()

Dim sBirthday As String

'-- Here we retrieve the fields from the database and --
'-- populate the fields in the user interface.        --

Call clearFields

With rsContactTable
  If (Not IsNull(!LastName)) Then txtLastName = !LastName
  If (Not IsNull(!MiddleInitial)) Then
    txtMiddleInitial = !MiddleInitial
  End If
  If (Not IsNull(!FirstName)) Then txtFirstName = !FirstName
  If (Not IsNull(!HomeStreet)) Then
    txtHomeStreet = !HomeStreet
  End If
  If (Not IsNull(!HomeCity)) Then
    txtHomeCity = !HomeCity
  End If
  If (Not IsNull(!HomeState)) Then
    txtHomeState = !HomeState
  End If
  If (Not IsNull(!HomeZip)) Then
    mskHomeZip = !HomeZip
  End If
  If (Not IsNull(!HomePhone)) Then
    mskHomePhone = !HomePhone
  End If
  If (Not IsNull(!HomeFax)) Then
    mskHomeFax = !HomeFax
  End If
  If (Not IsNull(!HomeEmail)) Then
    txtHomeEmail = !HomeEmail
  End If
  If (Not IsNull(!HomeCellPhone)) Then
    mskHomeCellPhone = !HomeCellPhone
  End If
  If (Not IsNull(!Birthday)) Then
    sBirthday = !Birthday
    convertDate sBirthday
    mskBirthday = sBirthday
    lblBirthday = Format$(!Birthday, "dddd mmmm dd, yyyy")
  End If
  DoEvents

 Call updateForm

End With

End Sub

Public Sub convertDate(sBirthday As String)

Dim sYear

Select Case Len(sBirthday)

Case 10 'needed to keep centuries correct prior to 1900 and after 2029.
Exit Sub

Case 9
If Mid$(sBirthday, 2, 1) = "/" Then
sBirthday = "0" & sBirthday
Else
sBirthday = Left(sBirthday, 3) & "0" & Mid$(sBirthday, 4, 6)
End If
Exit Sub

Case 8
Select Case Mid$(sBirthday, 2, 1)
Case "/"
sBirthday = "0" & Left(sBirthday, 2) & "0" & Right(sBirthday, 6)
Exit Sub
Case Else
End Select

Case 7
Select Case Mid$(sBirthday, 2, 1)
Case "/"
sBirthday = "0" & Left(sBirthday, 7)
Case Else
sBirthday = Left(sBirthday, 3) & "0" & Right(sBirthday, 4)
End Select

Case 6
Select Case Mid$(sBirthday, 2, 1)
Case Is = "/"
sBirthday = "0" & Left(sBirthday, 2) & "0" & Right(sBirthday, 4)
Case Else
End Select

Case Else
End Select

sYear = Right(sBirthday, 2)
If sYear >= 30 Then
sBirthday = Mid$(sBirthday, 1, 6) & "19" & sYear
Else
sBirthday = Mid$(sBirthday, 1, 6) & "20" & sYear
End If
End Sub
Public Sub populateListView()

Dim itemToAdd As ListItem
Dim noteSQL As String

lvCalls.ListItems.Clear
txtNotes = ""
txtNotes.Locked = True

noteSQL = "SELECT DISTINCTROW Notes.DateOfCall,"
noteSQL = noteSQL & "Notes.CallTypeID, Notes.NotesOnPhoneCall, "
noteSQL = noteSQL & " Notes.CallCounter, CallType.CallDescription,"
noteSQL = noteSQL & " Notes.ContactID "
noteSQL = noteSQL & " FROM Notes "
noteSQL = noteSQL & " INNER JOIN CallType ON Notes.CallTypeID ="
noteSQL = noteSQL & " CallType.CallTypeID "
noteSQL = noteSQL & " WHERE Notes.ContactID = " & _
    lCurrentContactKey
noteSQL = noteSQL & " ORDER BY Notes.DateOfCall DESC"

Set rsCallType = dbContact.OpenRecordset(noteSQL)

If (rsCallType.RecordCount > 0) Then
   rsCallType.MoveFirst
    While Not rsCallType.EOF
       Set itemToAdd = lvCalls.ListItems.Add(, , _
          Format$(rsCallType!DateOfCall, "dddd mmmm dd, yyyy"))
       itemToAdd.SubItems(1) = rsCallType!CallDescription
       itemToAdd.SubItems(2) = CStr(rsCallType!CallCounter)
       rsCallType.MoveNext
   Wend
   sbStatus.Panels.Item(1).Text = "There are " & _
    rsCallType.RecordCount & " calls logged for " & _
    sCurrentContactName
Else
   Set itemToAdd = lvCalls.ListItems.Add(, , "No calls logged")
   sbStatus.Panels.Item(1).Text = "No calls logged for " _
    & sCurrentContactName
End If

lvCalls.SelectedItem = lvCalls.ListItems(1)
Call lvCalls_ItemClick(lvCalls.SelectedItem)
DoEvents

End Sub

Public Function validateEntry() As Boolean

Dim indx As Integer

validateEntry = True
sbStatus.Panels.Item(2).Text = "Validating..."
If (Len(txtFirstName) < 1) Then
  tbContact.Tab = 0
  indx = MsgBox("Please enter the first name of the contact.", _
          vbInformation + vbOKOnly, progname)
  txtFirstName.SetFocus
  validateEntry = False
  Exit Function
End If

If (Len(txtLastName) < 1) Then
  tbContact.Tab = 0
  indx = MsgBox("Please enter the last name of the contact.", _
          vbInformation + vbOKOnly, progname)
  txtLastName.SetFocus
  validateEntry = False
  Exit Function
End If

mskBirthday.PromptInclude = False
If (Len(mskBirthday.Text) > 0) Then
  mskBirthday.PromptInclude = True
  If (Not IsDate(mskBirthday)) Then
    tbContact.Tab = 0
    indx = MsgBox("Please enter a valid birthdate mm/dd/yyyy.", _
          vbInformation + vbOKOnly, progname)
    mskBirthday.SetFocus
    validateEntry = False
    Exit Function
  End If
End If
mskBirthday.PromptInclude = False

End Function

Public Sub postContact()

Dim rsMaxIDNumber As Recordset
Dim sqlMaxID As String
Dim lNewContactID As Long

Screen.MousePointer = vbHourglass
sbStatus.Panels.Item(2).Text = "Posting Contact...."

If (iCurrentState = NOW_ADDING) Then
     rsContactTable.AddNew
Else
  With rsContactTable
     .MoveFirst
     .Index = "PrimaryKey"    'remember to change PrimaryKey to contactName
                              'if your Contacts database was created using
                              'VisData as demonstrated on pages 400 - 408
     .Seek "=", lCurrentContactKey
     If Not .NoMatch Then
       rsContactTable.Edit
     Else
      MsgBox ("Ohhhh Nooo")
     End If
   End With
End If

With rsContactTable
    If (Len(txtFirstName)) Then !FirstName = txtFirstName
    If (Len(txtMiddleInitial)) Then !MiddleInitial = _
    txtMiddleInitial
    If (Len(txtLastName)) Then !LastName = txtLastName
    If (Len(txtHomeStreet)) Then !HomeStreet = txtHomeStreet
    If (Len(txtHomeCity)) Then !HomeCity = txtHomeCity
    If (Len(txtHomeState)) Then !HomeState = txtHomeState
    If (Len(mskHomeZip)) Then !HomeZip = mskHomeZip
    If (Len(mskHomePhone)) Then !HomePhone = mskHomePhone
    If (Len(mskHomeFax)) Then !HomeFax = mskHomeFax
    If (Len(mskHomeCellPhone)) Then !HomeCellPhone = _
    mskHomeCellPhone
    If (Len(txtHomeEmail)) Then !HomeEmail = txtHomeEmail
    mskBirthday.PromptInclude = False
    If (Len(mskBirthday.Text) > 0) Then
      mskBirthday.PromptInclude = True
      !Birthday = mskBirthday
      lblBirthday = Format$(!Birthday, "dddd mmmm dd, yyyy")
    End If
    mskBirthday.PromptInclude = True
    .Update

End With

DoEvents

If (iCurrentState = NOW_ADDING) Then
  Call initializeForm
Else
  iCurrentState = NOW_IDLE
  Call lockFields(True)
  Call updateForm
End If

sbStatus.Panels.Item(2).Text = "Ready."
Screen.MousePointer = vbDefault

End Sub

