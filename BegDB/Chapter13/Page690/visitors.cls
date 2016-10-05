VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "visitors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function setVisitor(ByVal firstName As Variant, ByVal lastName As Variant) As Variant

Dim adoConnection As ADODB.Connection
Dim adoRecordset As ADODB.Recordset
Dim connectString As String

connectString = "Provider=Microsoft.Jet.OLEDB.3.51;" & _
                 "Data Source=C:\BegDB\visitors.mdb"

Set adoConnection = CreateObject("ADODB.Connection")
Set adoRecordset = CreateObject("ADODB.Recordset")

adoConnection.Open connectString
adoRecordset.Open "SELECT * FROM SiteVisitors", adoConnection, _
                               adOpenDynamic, adLockOptimistic

With adoRecordset
   .AddNew
   !firstName = firstName
   !lastName = lastName
   !previousVisit = Now()
   !totalVisits = 1
   .Update
End With

'-- Return a unique ID so we can set the cookie
setVisitor = adoRecordset!cookieID

'-- Close the recordset and the connection
adoRecordset.Close
adoConnection.Close

End Function

Public Function getVisitor(ByVal visitorNumber As Variant) As Variant

Dim adoConnection As ADODB.Connection
Dim adoRecordset As ADODB.Recordset
Dim connectString As String
Dim sqlString As String
Dim visitorArray() As Variant   'must be a variant

connectString = "Provider=Microsoft.Jet.OLEDB.3.51;" & _
                 "Data Source=C:\begdb\visitors.mdb"

Set adoConnection = CreateObject("ADODB.Connection")
Set adoRecordset = CreateObject("ADODB.Recordset")

adoConnection.Open connectString

sqlString = "Select * From SiteVisitors WHERE CookieID = " & visitorNumber

adoRecordset.Open sqlString, adoConnection, adOpenDynamic, adLockOptimistic

If (adoRecordset.BOF) And (adoRecordset.EOF) Then
  ReDim visitorArray(1)
  visitorArray(0) = "Not Found"
  visitorArray(1) = "Not Found"
  visitorArray(2) = "Not Found"
  visitorArray(3) = "Not Found"
Else
  ReDim visitorArray(4)
  visitorArray(0) = adoRecordset!firstName
  visitorArray(1) = adoRecordset!lastName
  visitorArray(2) = adoRecordset!previousVisit
  adoRecordset!previousVisit = Now()
  adoRecordset!totalVisits = adoRecordset!totalVisits + 1
  adoRecordset.Update
  visitorArray(3) = adoRecordset!totalVisits
End If

adoRecordset.Close
adoConnection.Close

getVisitor = visitorArray
End Function



