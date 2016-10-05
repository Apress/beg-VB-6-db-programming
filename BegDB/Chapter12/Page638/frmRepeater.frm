VERSION 5.00
Object = "{5C8CED40-8909-11D0-9483-00A0C91110ED}#1.0#0"; "MSDATREP.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmRepeater 
   Caption         =   "ADO Data Repeater Example"
   ClientHeight    =   2940
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8460
   LinkTopic       =   "Form1"
   ScaleHeight     =   2940
   ScaleWidth      =   8460
   StartUpPosition =   3  'Windows Default
   Begin MSDataRepeaterLib.DataRepeater DataRepeater1 
      Bindings        =   "frmRepeater.frx":0000
      Height          =   2085
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   3678
      _StreamID       =   -1412567295
      _Version        =   393216
      Caption         =   "DataRepeater1"
      BeginProperty RepeatedControlName {21FC0FC0-1E5C-11D1-A327-00AA00688B10} 
         _StreamID       =   -1412567295
         _Version        =   65536
         Name            =   "controlPrj.pubCtl"
      EndProperty
      RepeaterBindings=   3
      BeginProperty RepeaterBinding(0) {7D21A594-FC9B-11D0-A320-00AA00688B10} 
         _StreamID       =   -1412567295
         _Version        =   65536
         PropertyName    =   "company"
         DataField       =   "Company Name"
      EndProperty
      BeginProperty RepeaterBinding(1) {7D21A594-FC9B-11D0-A320-00AA00688B10} 
         _StreamID       =   -1412567295
         _Version        =   65536
         PropertyName    =   "id"
         DataField       =   "PubID"
      EndProperty
      BeginProperty RepeaterBinding(2) {7D21A594-FC9B-11D0-A320-00AA00688B10} 
         _StreamID       =   -1412567295
         _Version        =   65536
         PropertyName    =   "name"
         DataField       =   "Name"
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   2520
      Top             =   2520
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   582
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
      Caption         =   "Adodc1"
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
End
Attribute VB_Name = "frmRepeater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

