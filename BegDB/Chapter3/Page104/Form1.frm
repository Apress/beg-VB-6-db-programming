VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3585
   LinkTopic       =   "Form1"
   ScaleHeight     =   1980
   ScaleWidth      =   3585
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\BegDB\Biblio.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Titles"
      Top             =   1440
      Width           =   3015
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      DataField       =   "ISBN"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   13
      Mask            =   "#-#######-#-#"
      PromptChar      =   "_"
   End
   Begin VB.Label Label1 
      Caption         =   "Formatted ISBN Numbers"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

