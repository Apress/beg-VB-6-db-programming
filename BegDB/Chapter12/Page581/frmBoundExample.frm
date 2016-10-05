VERSION 5.00
Begin VB.Form frmBoundExample 
   Caption         =   "Binding Collection Example"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMoveNext 
      Caption         =   "&Move Next"
      Height          =   495
      Left            =   2040
      TabIndex        =   3
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Show Message"
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   1455
   End
   Begin VB.TextBox txtName 
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Text            =   "txtName"
      Top             =   240
      Width           =   1935
   End
   Begin VB.TextBox txtPubID 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Text            =   "txtPubID"
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmBoundExample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private clsBoundClass As New myBoundClass
Private bndPublishers As New BindingCollection

Private Sub Check1_Click()
clsBoundClass.displayMoveComplete = Check1.Value
End Sub

Private Sub cmdMoveNext_Click()
clsBoundClass.MoveNext
End Sub

Private Sub Form_Load()
With bndPublishers
  .DataMember = "Publishers"
  Set .DataSource = clsBoundClass
  .Add txtPubID, "Text", "PubID"
  .Add txtName, "Text", "Name"
  MsgBox "Number of items bound:  " & .Count
End With

End Sub
