VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPublishers 
   Caption         =   "Publishers"
   ClientHeight    =   4740
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6840
   LinkTopic       =   "Form1"
   ScaleHeight     =   4740
   ScaleWidth      =   6840
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   20
      Top             =   4140
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
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   3120
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      DataField       =   "Fax"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   8
      Left            =   5040
      TabIndex        =   8
      Text            =   "XXXXXXXXXXXXXXX"
      Top             =   2520
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      DataField       =   "Telephone"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   7
      Left            =   3240
      TabIndex        =   7
      Text            =   "XXXXXXXXXXXXXXX"
      Top             =   2520
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      DataField       =   "Zip"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   6
      Left            =   1440
      TabIndex        =   6
      Text            =   "XXXXXXXXXXXXXXX"
      Top             =   2520
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      DataField       =   "State"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   5
      Left            =   120
      TabIndex        =   5
      Text            =   "XXXXXXXXXX"
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      DataField       =   "City"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   4
      Left            =   4440
      TabIndex        =   4
      Text            =   "XXXXXXXXXXXXXXXXXXXX"
      Top             =   1800
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      DataField       =   "Address"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Text            =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
      Top             =   1800
      Width           =   4095
   End
   Begin VB.TextBox Text1 
      DataField       =   "Company Name"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Text            =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
      Top             =   1080
      Width           =   6615
   End
   Begin VB.TextBox Text1 
      DataField       =   "Name"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   1
      Left            =   1320
      TabIndex        =   1
      Text            =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
      Top             =   360
      Width           =   5415
   End
   Begin VB.TextBox Text1 
      DataField       =   "PubID"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Text            =   "XXXXXX"
      Top             =   360
      Width           =   735
   End
   Begin VB.Data Data1 
      Align           =   2  'Align Bottom
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
      Top             =   4395
      Width           =   6840
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Comments"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   9
      Left            =   3480
      TabIndex        =   19
      Top             =   2880
      Width           =   3255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Fax"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   8
      Left            =   5040
      TabIndex        =   18
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Telephone"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   7
      Left            =   3240
      TabIndex        =   17
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Zip"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   6
      Left            =   1440
      TabIndex        =   16
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "State"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   15
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "City"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   4
      Left            =   4440
      TabIndex        =   14
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Address"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   13
      Top             =   1560
      Width           =   4095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Company Name"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   12
      Top             =   840
      Width           =   6615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Name"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   1
      Left            =   1320
      TabIndex        =   11
      Top             =   120
      Width           =   5415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Publisher's ID"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   120
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


Private Sub Data1_Reposition()
With Data1.Recordset
  Data1.Caption = "Publisher " & (.AbsolutePosition + 1) & _
                             " of " & lTotalRecords
  ProgressBar1.Value = .PercentPosition
End With
End Sub

Private Sub Form_Activate()

With Data1.Recordset
    .MoveLast
    lTotalRecords = .RecordCount
    .MoveFirst
End With

End Sub

