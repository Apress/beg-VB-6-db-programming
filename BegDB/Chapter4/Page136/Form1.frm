VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5085
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5475
   LinkTopic       =   "Form1"
   ScaleHeight     =   5085
   ScaleWidth      =   5475
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      DataField       =   "Comments"
      DataSource      =   "Data1"
      Height          =   375
      Index           =   9
      Left            =   3450
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   4080
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      DataField       =   "Fax"
      DataSource      =   "Data1"
      Height          =   375
      Index           =   8
      Left            =   3480
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   3120
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      DataField       =   "Telephone"
      DataSource      =   "Data1"
      Height          =   375
      Index           =   7
      Left            =   3450
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      DataField       =   "Zip"
      DataSource      =   "Data1"
      Height          =   375
      Index           =   6
      Left            =   3480
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   1200
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      DataField       =   "State"
      DataSource      =   "Data1"
      Height          =   375
      Index           =   5
      Left            =   3450
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   240
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      DataField       =   "City"
      DataSource      =   "Data1"
      Height          =   375
      Index           =   4
      Left            =   690
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   4080
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      DataField       =   "Address"
      DataSource      =   "Data1"
      Height          =   375
      Index           =   3
      Left            =   690
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   3120
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      DataField       =   "Company Name"
      DataSource      =   "Data1"
      Height          =   375
      Index           =   2
      Left            =   720
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      DataField       =   "Name"
      DataSource      =   "Data1"
      Height          =   375
      Index           =   1
      Left            =   720
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1200
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      DataField       =   "PubID"
      DataSource      =   "Data1"
      Height          =   375
      Index           =   0
      Left            =   690
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   240
      Width           =   1695
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
      Top             =   4740
      Width           =   5475
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Comments"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   9
      Left            =   3480
      TabIndex        =   19
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Fax"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   8
      Left            =   3480
      TabIndex        =   18
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Telephone"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   7
      Left            =   3480
      TabIndex        =   17
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Zip"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   6
      Left            =   3480
      TabIndex        =   16
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "State"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   5
      Left            =   3480
      TabIndex        =   15
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "City"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   4
      Left            =   720
      TabIndex        =   14
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Address"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   3
      Left            =   720
      TabIndex        =   13
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Company Name"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   2
      Left            =   720
      TabIndex        =   12
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Name"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   1
      Left            =   720
      TabIndex        =   11
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Publisher's ID"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   10
      Top             =   0
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
