VERSION 5.00
Begin VB.UserControl pubCtl 
   ClientHeight    =   855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6465
   ScaleHeight     =   855
   ScaleWidth      =   6465
   Begin VB.TextBox txtPubID 
      Height          =   375
      Left            =   5400
      TabIndex        =   2
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox txtCompany 
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   360
      Width           =   3015
   End
   Begin VB.TextBox txtName 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "PubID"
      Height          =   255
      Left            =   5400
      TabIndex        =   5
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Company Name"
      Height          =   255
      Left            =   2280
      TabIndex        =   4
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "Name"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "pubCtl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Sub txtCompany_Change()
PropertyChanged "Company" 'automatically makes this persistent
End Sub

Private Sub txtName_Change()
PropertyChanged "Name"
End Sub

Private Sub txtPubID_Change()
PropertyChanged "PubID"
End Sub

Public Property Get company() As String
Attribute company.VB_MemberFlags = "14"
company = txtCompany
End Property

Public Property Let company(ByVal newCompanyName As String)
txtCompany = newCompanyName
End Property

Public Property Get name() As String
Attribute name.VB_MemberFlags = "14"
name = txtName
End Property

Public Property Let name(ByVal newName As String)
txtName = newName
End Property

Public Property Get id() As Long
Attribute id.VB_MemberFlags = "14"
id = txtPubID
End Property

Public Property Let id(ByVal newPubID As Long)
txtPubID = newPubID
End Property

