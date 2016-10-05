VERSION 5.00
Begin VB.Form frmTables 
   Caption         =   "Form1"
   ClientHeight    =   3330
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2640
   LinkTopic       =   "Form1"
   ScaleHeight     =   3330
   ScaleWidth      =   2640
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "DAO Collection"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   2400
      Width           =   1695
   End
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmTables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim myDatabase As Database
Dim myTable As TableDef         'this will hold the tableDef(inition)
Dim myFields As Fields
Dim myField As Field

Set myDatabase = OpenDatabase("C:\BegDB\BIBLIO.MDB")

For Each myTable In myDatabase.TableDefs    'dots to get the tableDefs
     List1.AddItem myTable.Name & " has " & myTable.Fields.Count _
            & " fields"
Next

End Sub
