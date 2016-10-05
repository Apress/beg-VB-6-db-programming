VERSION 5.00
Begin VB.Form frmUpdate 
   Caption         =   "Modifying Biblio.mdb"
   ClientHeight    =   2820
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3765
   LinkTopic       =   "Form1"
   ScaleHeight     =   2820
   ScaleWidth      =   3765
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\BegDB\Biblio.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2160
      Width           =   1740
   End
   Begin VB.CommandButton cmdNewYork 
      Caption         =   "&Update to New York"
      Height          =   615
      Left            =   840
      TabIndex        =   1
      Top             =   1200
      Width           =   1695
   End
   Begin VB.CommandButton cmdIpswitch 
      Caption         =   "&Change to Ipswitch"
      Height          =   615
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "frmUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdIpswitch_Click()
With Data1
    .RecordSource = "SELECT * FROM Publishers WHERE City = 'New York'"
    .Refresh
    .Recordset.MoveFirst
    While (Not .Recordset.EOF)
        .Recordset.Edit
        .Recordset!City = "Ipswitch"
        .Recordset.Update
        .Recordset.MoveNext
    Wend
End With

End Sub

Private Sub cmdNewYork_Click()
Dim dbBiblio As Database

Set dbBiblio = OpenDatabase("C:\BegDB\Biblio.mdb")
dbBiblio.Execute "UPDATE Publishers SET City = " & _
            "'New York' WHERE City = 'Ipswitch'"
End Sub
