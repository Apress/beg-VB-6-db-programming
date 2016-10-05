Attribute VB_Name = "modCode"
Option Explicit

Public dbContact As Database

Public Const progname = "Address Book"

Public rsContactTable As Recordset

Public Enum currentState
  NOW_ADDING
  NOW_SAVING
  NOW_EDITING
  NOW_DELETING
  NOW_IDLE
End Enum
 
Public Enum bButton
  bAdd = 1
  bCancel = 3
  bSave = 5
  bDelete = 7
  bEdit = 9
  bQuit = 11
End Enum

Public Function openTheDatabase() As Boolean

Dim dbPath As String

On Error GoTo dbErrors

dbPath = App.Path & "\contacts.mdb"

Set dbContact = DBEngine.Workspaces(0).OpenDatabase(dbPath, False)

Set rsContactTable = dbContact.OpenRecordset("Contact", dbOpenTable)

openTheDatabase = True

Exit Function

dbErrors:

openTheDatabase = False
MsgBox (Err.Description)

End Function

