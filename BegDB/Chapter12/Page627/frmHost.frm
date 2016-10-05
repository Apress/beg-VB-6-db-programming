VERSION 5.00
Object = "*\AdataCtl.vbp"
Begin VB.Form frmHost 
   Caption         =   "VB Database Programming Data Control"
   ClientHeight    =   2250
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6075
   LinkTopic       =   "Form1"
   ScaleHeight     =   2250
   ScaleWidth      =   6075
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      DataField       =   "Title"
      DataSource      =   "dbCtl1"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   120
      Width           =   3255
   End
   Begin dataCtl.dbCtl dbCtl1 
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   6030
      _ExtentX        =   10636
      _ExtentY        =   2566
      RecordSource    =   "Titles"
      ConnectionString=   "Provider=Microsoft.Jet.OLEDB.3.51;Data Source=C:\BegDB\Biblio.mdb"
   End
End
Attribute VB_Name = "frmHost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

