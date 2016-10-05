VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDBAnalyzer 
   Caption         =   "Microsoft Access Database Analyzer"
   ClientHeight    =   4500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6690
   Icon            =   "frmDBAnalyzer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4500
   ScaleWidth      =   6690
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAnalyze 
      Caption         =   "&Analyze Database"
      Height          =   495
      Left            =   3960
      TabIndex        =   5
      Top             =   3840
      Width           =   1815
   End
   Begin VB.CommandButton cmdLocate 
      Caption         =   "&Select Database"
      Height          =   495
      Left            =   840
      TabIndex        =   4
      Top             =   3840
      Width           =   1815
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3120
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   3240
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.TreeView dbTree 
      Height          =   2535
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   4471
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Scanning Tables in Database"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   3000
      Width           =   6375
   End
   Begin VB.Label lblTitle 
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
   End
End
Attribute VB_Name = "frmDBAnalyzer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dataBaseLocation As String


Private Sub cmdAnalyze_Click()
Dim dbase As Database
Dim tble As TableDef
Dim fld As Field
Dim prp As Property
Dim propString As String
Dim treeNode As Node
Dim currentFieldNumber As Integer
Dim currentFieldKey As String
Dim currentPropertyNumber As Integer
Dim currentPropertyKey As String
Dim currentTableName As String
Dim currentTableCount As Integer

Screen.MousePointer = vbHourglass

'-Don't permit the user to click when analyzing --
cmdAnalyze.Enabled = False
cmdLocate.Enabled = False

currentTableCount = 0
currentPropertyNumber = 1
currentFieldNumber = 1

'-Clear the TreeView control from any prior analysis --
dbTree.Nodes.Clear

'-Reasonability test ---
If (InStr(1, dataBaseLocation, ".MDB", 1) < 1) Then
  Screen.MousePointer = vbDefault
  MsgBox ("Please select an MDB file.")
  cmdAnalyze.Enabled = True
  cmdLocate.Enabled = True
  Exit Sub
End If

'-Set our reference variable to the selected database --
Set dbase = OpenDatabase(dataBaseLocation)

'-Reset the progress bar --
ProgressBar1.Value = 0
ProgressBar1.Max = dbase.TableDefs.Count

'-- The Root Node: The Name of the database --
Set treeNode = dbTree.Nodes.Add(, , "r", _
         "Database: " & dataBaseLocation)

'-- The next heirarchical structure is tables --
Set treeNode = dbTree.Nodes.Add("r", tvwChild, _
         "tble", "Tables")

'-- First, retrieve each table in the database --
For Each tble In dbase.TableDefs
    currentTableCount = currentTableCount + 1
    ProgressBar1.Value = currentTableCount
    currentTableName = "" & tble.Name & ""
    Set treeNode = dbTree.Nodes.Add("tble", _
        tvwChild, currentTableName, tble.Name)
  
'-- Now place the header 'Properties' under the table entry
    currentPropertyKey = "" & "Property" & _
        CStr(currentPropertyNumber) & ""
    currentPropertyNumber = currentPropertyNumber + 1
    Set treeNode = dbTree.Nodes.Add(currentTableName, tvwChild, _
        currentPropertyKey, "Properties")
  
    With tble
         For Each prp In tble.Properties
             propString = RetrieveProp(prp)
            Set treeNode = dbTree.Nodes.Add(currentPropertyKey, _
            tvwChild, , propString)
        Next
    
       currentFieldKey = "" & "Field " & CStr(currentFieldNumber) _
        & ""
       currentFieldNumber = currentFieldNumber + 1
       Set treeNode = dbTree.Nodes.Add(currentTableName, _
        tvwChild, currentFieldKey, "Fields")
    
       For Each fld In .Fields
          Set treeNode = dbTree.Nodes.Add(currentFieldKey, _
        tvwChild, , fld.Name)
     
      Next
  End With
Next

ProgressBar1.Value = 0
cmdAnalyze.Enabled = True
cmdLocate.Enabled = True
Screen.MousePointer = vbDefault

End Sub

Private Sub cmdLocate_Click()
On Error GoTo ErrHandler

With CommonDialog1

     .CancelError = True
     ' Set flags
     .Flags = cdlOFNHideReadOnly
     ' Set filters
     .Filter = "MS Access Files (*.MDB)|*.MDB"
     ' Specify default filter
     .FilterIndex = 1
     .DialogTitle = "Select table to analyze"
     ' Display the Open dialog box
     .ShowOpen
     'Display name of selected file
     dataBaseLocation = .FileName

End With

lblTitle = "Table Selected: " & dataBaseLocation
cmdAnalyze.Enabled = True

Exit Sub
    
ErrHandler:
  'User pressed the Cancel button
  dataBaseLocation = ""
  lblTitle = "Selection Canceled"
  Exit Sub

End Sub


Public Function RetrieveProp(myProperty As Property) As String

Dim tempProperty As Variant

Err = 0

On Error Resume Next
tempProperty = myProperty.Value

If (Err = 0) Then
   RetrieveProp = myProperty.Name
   If Len(tempProperty) Then
       RetrieveProp = RetrieveProp & "  -  Value: " & tempProperty
   End If
Else
      RetrieveProp = myProperty.Name
End If

End Function

