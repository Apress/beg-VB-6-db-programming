VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Form1 
   Caption         =   "Bound Column Example"
   ClientHeight    =   4590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8025
   LinkTopic       =   "Form2"
   ScaleHeight     =   4590
   ScaleWidth      =   8025
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text7 
      DataField       =   "Comments"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   4080
      TabIndex        =   7
      Text            =   "Comments"
      Top             =   3240
      Width           =   3735
   End
   Begin VB.TextBox Text6 
      DataField       =   "Subject"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Text            =   "Subject"
      Top             =   3240
      Width           =   3615
   End
   Begin VB.TextBox Text5 
      DataField       =   "Notes"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   5640
      TabIndex        =   5
      Text            =   "Notes"
      Top             =   2280
      Width           =   2175
   End
   Begin VB.TextBox Text4 
      DataField       =   "Description"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Text            =   "Description"
      Top             =   2280
      Width           =   3135
   End
   Begin VB.TextBox Text3 
      DataField       =   "ISBN"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Text            =   "ISBN"
      Top             =   2280
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      DataField       =   "Year Published"
      DataSource      =   "Adodc1"
      Height          =   405
      Left            =   4440
      TabIndex        =   2
      Text            =   "Year Published"
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      DataField       =   "Title"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Text            =   "Title"
      Top             =   1320
      Width           =   3855
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   5280
      Top             =   3960
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=C:\BegDB\Biblio.mdb"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=C:\BegDB\Biblio.mdb"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Publishers"
      Caption         =   "Publishers Table"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   240
      Top             =   3960
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   2
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=C:\BegDB\Biblio.mdb"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=C:\BegDB\Biblio.mdb"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Titles"
      Caption         =   "Titles Table"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "Form1.frx":0000
      DataField       =   "PubID"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Name"
      BoundColumn     =   "PubID"
      Text            =   "DataCombo1"
   End
   Begin VB.Label Label7 
      Caption         =   "Comments"
      Height          =   255
      Left            =   4080
      TabIndex        =   14
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "Subject"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "Notes"
      Height          =   255
      Left            =   5640
      TabIndex        =   12
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "Description"
      Height          =   255
      Left            =   2280
      TabIndex        =   11
      Top             =   1920
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "ISBN"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Year Published"
      Height          =   255
      Left            =   4440
      TabIndex        =   9
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Book Title"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   960
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

