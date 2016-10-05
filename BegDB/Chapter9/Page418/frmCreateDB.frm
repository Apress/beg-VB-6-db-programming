VERSION 5.00
Begin VB.Form frmCreateDB 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCreateDB 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
End
Attribute VB_Name = "frmCreateDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim db As Database
Dim tblDef As TableDef
Dim fldDef As Field
Dim indx As Index
Dim dbName As String

Private Sub cmdCreateDB_Click()
dbName = "C:\BegDB\CONTACTS.MDB"
If (Len(Dir(dbName))) Then
  Kill dbName
End If

Set db = DBEngine.Workspaces(0).CreateDatabase(dbName, dbLangGeneral)

Call createContactTable

Call createNotesTable

Call createCallTypeTable

Call createRelationships

MsgBox ("Database successfully created.")

End Sub

Public Sub createContactTable()

'-- Create the Contact table here --
Set tblDef = db.CreateTableDef("Contact")

'-Create the fields in the contact table --
Set fldDef = tblDef.CreateField("ContactID", dbLong)
fldDef.Attributes = dbAutoIncrField
tblDef.Fields.Append fldDef
Set fldDef = tblDef.CreateField("LastName", dbText, 20)
tblDef.Fields.Append fldDef
Set fldDef = tblDef.CreateField("FirstName", dbText, 15)
tblDef.Fields.Append fldDef
Set fldDef = tblDef.CreateField("MiddleInitial", dbText, 1)
tblDef.Fields.Append fldDef
Set fldDef = tblDef.CreateField("Birthday", dbDate)
tblDef.Fields.Append fldDef
Set fldDef = tblDef.CreateField("HomeStreet", dbText, 20)
tblDef.Fields.Append fldDef
Set fldDef = tblDef.CreateField("HomeCity", dbText, 12)
tblDef.Fields.Append fldDef
Set fldDef = tblDef.CreateField("HomeState", dbText, 2)
tblDef.Fields.Append fldDef
Set fldDef = tblDef.CreateField("HomeZip", dbText, 10)
tblDef.Fields.Append fldDef
Set fldDef = tblDef.CreateField("HomePhone", dbText, 15)
tblDef.Fields.Append fldDef
Set fldDef = tblDef.CreateField("HomeFax", dbText, 15)
tblDef.Fields.Append fldDef
Set fldDef = tblDef.CreateField("HomeEmail", dbText, 30)
tblDef.Fields.Append fldDef
Set fldDef = tblDef.CreateField("HomeCellPhone", dbText, 15)
tblDef.Fields.Append fldDef
db.TableDefs.Append tblDef

'-- Add the Primary Key --
Set indx = tblDef.CreateIndex("PrimaryKey")
Set fldDef = indx.CreateField("ContactID")
indx.Fields.Append fldDef
indx.Primary = True
tblDef.Indexes.Append indx

End Sub

Public Sub createNotesTable()

'-- Create the Notes table here --
Set tblDef = db.CreateTableDef("Notes")

'-Create the fields in the Notes table --
Set fldDef = tblDef.CreateField("ContactID", dbLong)
tblDef.Fields.Append fldDef
Set fldDef = tblDef.CreateField("DateOfCall", dbDate)
tblDef.Fields.Append fldDef
Set fldDef = tblDef.CreateField("CallTypeID", dbLong)
tblDef.Fields.Append fldDef
Set fldDef = tblDef.CreateField("NotesOnPhoneCall", dbMemo)
tblDef.Fields.Append fldDef
Set fldDef = tblDef.CreateField("CallCounter", dbLong)
fldDef.Attributes = dbAutoIncrField
tblDef.Fields.Append fldDef
db.TableDefs.Append tblDef

End Sub

Public Sub createCallTypeTable()

'-- Create the CallType table here --
Set tblDef = db.CreateTableDef("CallType")

'-Create the fields in the CallType table --
Set fldDef = tblDef.CreateField("CallTypeID", dbLong)
fldDef.Attributes = dbAutoIncrField
tblDef.Fields.Append fldDef
Set fldDef = tblDef.CreateField("CallDescription", dbText, 20)
tblDef.Fields.Append fldDef
db.TableDefs.Append tblDef

'-- Add the Primary Key --
Set indx = tblDef.CreateIndex("PrimaryKey")
Set fldDef = indx.CreateField("CallTypeID")
indx.Fields.Append fldDef
indx.Primary = True
tblDef.Indexes.Append indx

End Sub

Public Sub createRelationships()

Dim makeRelation As Relation
Dim fld As Field

Set makeRelation = db.CreateRelation("MyRelationship")
makeRelation.Table = "Contact"
makeRelation.ForeignTable = "Notes"

Set fld = makeRelation.CreateField("ContactID")
fld.ForeignName = "ContactID"
makeRelation.Fields.Append fld
makeRelation.Attributes = dbRelationDeleteCascade
db.Relations.Append makeRelation

End Sub

