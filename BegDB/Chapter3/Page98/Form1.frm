VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4815
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7080
   LinkTopic       =   "Form1"
   ScaleHeight     =   4815
   ScaleWidth      =   7080
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data2 
      Caption         =   "Titles"
      Connect         =   "Access"
      DatabaseName    =   "C:\BegDB\Biblio.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4320
      Visible         =   0   'False
      Width           =   2580
   End
   Begin VB.Data Data1 
      Caption         =   "Publishers"
      Connect         =   "Access"
      DatabaseName    =   "C:\BegDB\Biblio.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Publishers"
      Top             =   4320
      Width           =   2580
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Form1.frx":0000
      Height          =   2415
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   4260
      _Version        =   393216
      AllowUserResizing=   3
   End
   Begin VB.TextBox Text1 
      DataField       =   "Name"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   480
      Width           =   4455
   End
   Begin VB.Label lblTitles 
      Caption         =   "Label2"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   "Publisher"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Data1_Reposition()
Dim lPubID As Long
lPubID = Data1.Recordset!PubID
Data2.RecordSource = "SELECT * FROM TITLES WHERE PubID = " _
    & lPubID
Data2.Refresh

End Sub

Private Sub Text1_Change()
lblTitles.Caption = "Titles published by  " & Text1
End Sub
