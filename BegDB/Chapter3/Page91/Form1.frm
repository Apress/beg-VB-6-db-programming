VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3225
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4485
   LinkTopic       =   "Form1"
   ScaleHeight     =   3225
   ScaleWidth      =   4485
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      DataField       =   "ISBN"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Text            =   "Text3"
      Top             =   1440
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      DataField       =   "Year Published"
      DataSource      =   "Data1"
      Height          =   405
      Left            =   1560
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   840
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      DataField       =   "Title"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   240
      Width           =   2775
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\BegDB\Biblio.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Titles"
      Top             =   2280
      Width           =   2340
   End
   Begin VB.Label Label3 
      Caption         =   "ISBN"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Year Published"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Title"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

