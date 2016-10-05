VERSION 5.00
Begin VB.Form AsciiCodes 
   Caption         =   "ASCII Converter"
   ClientHeight    =   5475
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3285
   LinkTopic       =   "Form1"
   ScaleHeight     =   5475
   ScaleWidth      =   3285
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   5325
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3255
   End
End
Attribute VB_Name = "AsciiCodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim iAscii As Integer

For iAscii = 65 To 90
    List1.AddItem "Ascii Value: " & iAscii & " = " & Chr(iAscii) & " and Character " & Chr(iAscii) & " =" & Asc(Chr(iAscii))
Next
    
    
End Sub

