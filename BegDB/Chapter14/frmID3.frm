VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmID3 
   Caption         =   "Product Analysis"
   ClientHeight    =   5985
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7890
   Icon            =   "frmID3.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5985
   ScaleWidth      =   7890
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtAnalysis 
      Height          =   855
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Text            =   "frmID3.frx":0442
      Top             =   4440
      Width           =   6255
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   5400
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.StatusBar sbStatus 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   6
      Top             =   5685
      Width           =   7890
      _ExtentX        =   13917
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   5794
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "18:54"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "15/07/99"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdAnalyze 
      Caption         =   "&Analyze"
      Height          =   495
      Left            =   6600
      TabIndex        =   5
      Top             =   4800
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid resultsGrid 
      Height          =   1695
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   2990
      _Version        =   393216
      FixedCols       =   0
      AllowUserResizing=   1
   End
   Begin MSFlexGridLib.MSFlexGrid miningGrid 
      Height          =   1335
      Left            =   2760
      TabIndex        =   3
      Top             =   1200
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   2355
      _Version        =   393216
      FixedRows       =   0
      FixedCols       =   0
   End
   Begin VB.ListBox lstProduct 
      Height          =   840
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   2415
   End
   Begin VB.ListBox lstCategory 
      Height          =   840
      ItemData        =   "frmID3.frx":0450
      Left            =   120
      List            =   "frmID3.frx":0452
      TabIndex        =   0
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label Label4 
      Caption         =   "Region                                  Language                                     Country"
      Height          =   255
      Left            =   2760
      TabIndex        =   9
      Top             =   960
      Width           =   5055
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ID3 Product Management (P)roduct (A)nalysis (T)ool"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3360
      TabIndex        =   8
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Products"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Categories"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmID3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim adoConnection As ADODB.Connection

Private Sub cmdAnalyze_Click()
txtAnalysis = ""
Call setupID3
Call buildID3
Call determineEntropy

End Sub

Private Sub Form_Load()
'-- Show what we are doing here --
sbStatus.Panels.Item(2).Text = "Loading..."

Me.Show
DoEvents

'-------------------------------------
'-- Open the database with ADO --
'-------------------------------------
sbStatus.Panels.Item(1).Text = "Opening the database..."
If (Not openTheDatabase()) Then
  sbStatus.Panels.Item(1).Text = "Database failed..."
  sbStatus.Panels.Item(2).Text = "Error."
  Exit Sub
End If

Call setupID3

sbStatus.Panels.Item(1).Text = "Updating list boxes..."
Call updateListBoxes

sbStatus.Panels.Item(1).Text = ""
sbStatus.Panels.Item(2).Text = "Ready."

End Sub


Public Function openTheDatabase() As Boolean

'-- Here we want to open the database
Dim sConnectionString As String

On Error GoTo dbError

sbStatus.Panels.Item(1).Text = "Opening the database."

'-- Set reference to a new connection --
Set adoConnection = New ADODB.Connection

'-- Build the connection string
sConnectionString = "Provider=Microsoft.Jet.OLEDB.3.51;" _
                 & "Data Source=C:\BegDB\Nwind.mdb"

adoConnection.Open sConnectionString

openTheDatabase = True
Exit Function

dbError:
MsgBox (Err.Description)
openTheDatabase = False
sbStatus.Panels.Item(2).Text = "Could not open database."
End Function

Public Sub setupID3()

Dim adoTempCountry As ADODB.Recordset
'holds each unique country

Dim adoTempid3 As ADODB.Recordset
'holds a recordset for ID3

Dim sCountry As String
Dim sSql As String

'-----------------------------------------------------
'-- Get the unique countries and place in temp --
'-----------------------------------------------------
Set adoTempCountry = New ADODB.Recordset
sSql = "SELECT DISTINCT Country FROM Customers"
Set adoTempCountry = adoConnection.Execute(sSql)
'-----------------------------------------------

'-----------------------------------------------------
'-- Clean out all of the data from the ID3 table --
'-----------------------------------------------------
adoConnection.Execute ("DELETE * FROM ID3")

'----------------------------
'-- Now set up the ID3 --
'---------------------------
Set adoTempid3 = New ADODB.Recordset
adoTempid3.CursorType = adOpenKeyset
adoTempid3.LockType = adLockOptimistic
adoTempid3.Open "ID3", adoConnection

'-------------------------------------------------------
'-- Iterate through the country table and add each  --
'-- unique entry to the Country field of ID3.       --
'-------------------------------------------------------

While Not adoTempCountry.EOF

   sCountry = adoTempCountry!Country
   
   With adoTempid3
     .AddNew
     !Country = sCountry
   
   Select Case UCase(sCountry)
        Case "ARGENTINA"
          !CountryRegion = "South America"
          !CountryLanguage = "Spanish"
         Case "AUSTRIA"
          !CountryRegion = "Europe"
          !CountryLanguage = "German"
        Case "BELGIUM"
          !CountryRegion = "Europe"
          !CountryLanguage = "French"
        Case "BRAZIL"
          !CountryRegion = "South America"
          !CountryLanguage = "Portuguese"
        Case "CANADA"
          !CountryRegion = "North America"
          !CountryLanguage = "English"
        Case "DENMARK"
          !CountryRegion = "Scandinavia"
          !CountryLanguage = "Danish"
        Case "FINLAND"
          !CountryRegion = "Scandinavia"
          !CountryLanguage = "Finnish"
        Case "FRANCE"
          !CountryRegion = "Europe"
          !CountryLanguage = "French"
        Case "GERMANY"
          !CountryRegion = "Europe"
          !CountryLanguage = "German"
        Case "IRELAND"
          !CountryRegion = "British Isles"
          !CountryLanguage = "English"
        Case "ITALY"
          !CountryRegion = "Europe"
          !CountryLanguage = "Italian"
        Case "MEXICO"
          !CountryRegion = "North America"
          !CountryLanguage = "Spanish"
        Case "NORWAY"
          !CountryRegion = "Scandinavia"
          !CountryLanguage = "Norwegian"
        Case "POLAND"
          !CountryRegion = "Europe"
          !CountryLanguage = "Polish"
        Case "PORTUGAL"
          !CountryRegion = "Mediterranean"
          !CountryLanguage = "Portuguese"
        Case "SPAIN"
          !CountryRegion = "Mediterranean"
          !CountryLanguage = "Spanish"
        Case "SWEDEN"
          !CountryRegion = "Scandinavia"
          !CountryLanguage = "Swedish"
        Case "SWITZERLAND"
          !CountryRegion = "Europe"
          !CountryLanguage = "German"
        Case "UK"
          !CountryRegion = "British Isles"
          !CountryLanguage = "English"
        Case "USA"
          !CountryRegion = "North America"
          !CountryLanguage = "English"
        Case "VENEZUELA"
          !CountryRegion = "South America"
          !CountryLanguage = "Spanish"
        Case Else
          !CountryRegion = "Unknown"
          !CountryLanguage = "Unknown"
     End Select
     .Update
   End With
   adoTempCountry.MoveNext
 Wend
adoTempCountry.Close
adoTempid3.Close

Call gridID3

End Sub

Public Sub gridID3()

Dim adoID3 As ADODB.Recordset
Dim sSql As String
Dim iRows As Integer
Dim iCols As Integer
Dim iRowLoop As Integer
Dim iColLoop As Integer


sSql = "SELECT CountryRegion, CountryLanguage, Country"
sSql = sSql & " FROM ID3 GROUP BY "
sSql = sSql & "CountryRegion, CountryLanguage, Country"

Set adoID3 = New ADODB.Recordset
adoID3.CursorLocation = adUseClient

adoID3.Open sSql, adoConnection, , , adCmdText

adoID3.MoveFirst

iRows = adoID3.RecordCount
iCols = adoID3.Fields.Count

miningGrid.Rows = iRows
miningGrid.Cols = iCols

'----------------------------
'-- Set up the grid here --
'----------------------------
miningGrid.Row = 0
miningGrid.ColAlignment(0) = 7
    
For iColLoop = 0 To miningGrid.Cols - 1
  miningGrid.Col = iColLoop
  miningGrid.MergeCol(iColLoop) = True
  miningGrid.ColWidth(iColLoop) = 1500    'Set column's width
Next
miningGrid.MergeCells = flexMergeRestrictColumns

For iRowLoop = 0 To iRows - 1
  For iColLoop = 0 To iCols - 1
     miningGrid.Row = iRowLoop
     miningGrid.Col = iColLoop
     miningGrid.Text = adoID3.Fields(iColLoop)
  Next
  adoID3.MoveNext
Next

adoID3.Close

Set adoID3 = Nothing


End Sub


Public Sub updateListBoxes()

Dim adoTempRecordset As ADODB.Recordset
Dim sSql As String

'-- Let the user know what we are doing here ---
sbStatus.Panels.Item(1).Text = "Updating list boxes."

'--------------------------------------
'-- Set up the Categories list box --
'--------------------------------------
Set adoTempRecordset = New ADODB.Recordset
adoTempRecordset.CursorLocation = adUseClient
adoTempRecordset.Open _
   "SELECT * FROM Categories ORDER BY CategoryName", _
   adoConnection

ProgressBar1.Max = adoTempRecordset.RecordCount

lstCategory.Clear
With adoTempRecordset
  If .RecordCount > 0 Then .MoveFirst
  While Not .EOF
    ProgressBar1.Value = .AbsolutePosition
    lstCategory.AddItem !CategoryName
    lstCategory.ItemData(lstCategory.NewIndex) = !CategoryID
    .MoveNext
  Wend
End With
lstCategory.ListIndex = 0
ProgressBar1.Value = 0

adoTempRecordset.Close

'--------------------------
'-Set up the Products -
'--------------------------
sSql = "SELECT ProductName, ProductID FROM Products"
sSql = sSql & " WHERE CategoryID = " & lstCategory.ItemData(lstCategory.ListIndex)

adoTempRecordset.Open sSql, adoConnection

ProgressBar1.Max = adoTempRecordset.RecordCount

lstProduct.Clear
With adoTempRecordset
   If .RecordCount > 0 Then .MoveFirst
   While Not .EOF
     ProgressBar1.Value = .AbsolutePosition
     lstProduct.AddItem !ProductName
     lstProduct.ItemData(lstProduct.NewIndex) = !ProductID
     .MoveNext
   Wend
End With

lstProduct.ListIndex = 0
ProgressBar1.Value = 0

sbStatus.Panels.Item(1).Text = ""

End Sub


Private Sub lstCategory_Click()
Dim adoTempProducts As ADODB.Recordset
Dim adoTempPicture As ADODB.Recordset

Dim sSql As String

sSql = "SELECT ProductName, ProductID FROM Products"
sSql = sSql & " WHERE CategoryID = "
sSql = sSql & lstCategory.ItemData(lstCategory.ListIndex)

Set adoTempProducts = New ADODB.Recordset
adoTempProducts.CursorLocation = adUseClient

adoTempProducts.Open sSql, adoConnection


If adoTempProducts.RecordCount > 0 Then _
    adoTempProducts.MoveLast
ProgressBar1.Max = adoTempProducts.RecordCount
adoTempProducts.MoveFirst

lstProduct.Clear
With adoTempProducts
   If .RecordCount > 0 Then .MoveFirst
   While Not .EOF
     ProgressBar1.Value = .AbsolutePosition
     lstProduct.AddItem !ProductName
     lstProduct.ItemData(lstProduct.NewIndex) = !ProductID
     .MoveNext
   Wend
End With
lstProduct.ListIndex = 0
ProgressBar1.Value = 0
adoTempProducts.Close

sbStatus.Panels.Item(1).Text = ""

Call showAnalysis

End Sub

Private Sub lstProduct_Click()
Call showAnalysis
End Sub

Public Sub showAnalysis()

Dim sAnalysis As String

sAnalysis = "Search / Analysis Criteria:  " & vbCrLf
sAnalysis = sAnalysis & " Category:  " & lstCategory & vbCrLf
sAnalysis = sAnalysis & " Product:  " & lstProduct & vbCrLf
txtAnalysis = sAnalysis

End Sub
Private Sub buildID3()

Dim sSql As String
Dim sCriteria As String
Dim sCurrentCountry As String
Dim sCurrentLanguage As String
Dim sCurrentRegion As String
Dim adoRequestedData As ADODB.Recordset
Dim adoID3 As ADODB.Recordset

sbStatus.Panels.Item(1).Text = "Building Query."
sbStatus.Panels.Item(2).Text = "Working..."

'-------------------------------------------------------------
'-- This SQL will retrieve the Category Name, Total,
'-- Order Date, Country, and Product Name for the categories
'-- requested.
'-------------------------------------------------------------
sSql = "SELECT DISTINCTROW Categories.CategoryName, "
sSql = sSql & " SUM([Order Details].Quantity) AS Total,"
sSql = sSql & " Orders.OrderDate, Customers.Country,"
sSql = sSql & " Products.ProductName "
sSql = sSql & " FROM (Customers INNER JOIN Orders ON "
sSql = sSql & " Customers.CustomerID = Orders.CustomerID)"
sSql = sSql & " INNER JOIN ((Categories INNER JOIN Products"
sSql = sSql & " ON Categories.CategoryID ="
sSql = sSql & " Products.CategoryID) INNER JOIN "
sSql = sSql & " [Order Details] ON Products.ProductID = "
sSql = sSql & " [Order Details].ProductID) ON "
sSql = sSql & " Orders.OrderID = [Order Details].OrderID "

'----------------------------
'-- Now check the criteria --
'----------------------------
sCriteria = ""
sCriteria = sCriteria & " Categories.CategoryID = " & _
    lstCategory.ItemData(lstCategory.ListIndex)
sCriteria = sCriteria & " AND "
sCriteria = sCriteria & " Products.ProductID = " & _
    lstProduct.ItemData(lstProduct.ListIndex)
  
sSql = sSql & " WHERE " & sCriteria

sSql = sSql & " GROUP BY "
sSql = sSql & " Categories.CategoryName, "
sSql = sSql & "[Order Details].Quantity, Products.ProductName,"
sSql = sSql & " Orders.OrderDate, Customers.Country "
sSql = sSql & " ORDER BY Categories.CategoryName "

'MsgBox sSql

'----------------------------------------------------
'-- Open a recordset with the results of the query --
'----------------------------------------------------
Set adoRequestedData = New ADODB.Recordset
adoRequestedData.CursorLocation = adUseClient
adoRequestedData.CursorType = adOpenDynamic
adoRequestedData.Open sSql, adoConnection

'--------------------------------------------
'-- If there were no records, exit the sub --
'--------------------------------------------
With adoRequestedData
   If (.RecordCount < 1) Then
     MsgBox "No records"
     Exit Sub
   Else
     ProgressBar1.Max = .RecordCount
     .MoveFirst
   End If
End With

sbStatus.Panels.Item(1).Text = "Determining Sales Information"

'------------------------------------------------------
'-- Now, create the ID3 table and prepare to refresh --
'------------------------------------------------------
Set adoID3 = New ADODB.Recordset
sSql = "SELECT * FROM ID3"
adoID3.Open sSql, adoConnection, adOpenDynamic, _
   adLockOptimistic, adCmdText
'------------------------------------------------------------
'-- Loop through all of the records in the recordset
'-- returned by the query and update the ID3 table based
'-- on the results.
'------------------------------------------------------------
With adoRequestedData
 While Not .EOF
  
   ProgressBar1.Value = .AbsolutePosition
   DoEvents
   sCurrentCountry = !Country
   adoID3.Filter = "Country = '" & sCurrentCountry & "'"
   If ((Not adoID3.BOF) And (Not adoID3.EOF)) Then
     adoID3!Category = !CategoryName
     adoID3!Product = !ProductName
     If (Year(!OrderDate) = "1995") Then
             adoID3!OldQuantity = adoID3!OldQuantity + !Total
     ElseIf (Year(!OrderDate) = "1996") Then
             adoID3!NewQuantity = adoID3!NewQuantity + !Total
     End If
     If (adoID3!OldQuantity < adoID3!NewQuantity) Then
             adoID3!SalesUp = True
     Else
             adoID3!SalesUp = False
     End If
     If (adoID3!NewQuantity > 0) And (adoID3!OldQuantity > 0) _
          Then
        adoID3!UpByHowMuch = _
CSng(Format((adoID3!NewQuantity / adoID3!OldQuantity), "##.###"))
     End If
     '-- Were there were only sales in the current year --
     If (adoID3!NewQuantity > 0) And (adoID3!OldQuantity < 1) _
          Then
        adoID3!UpByHowMuch = CSng(Format(1, "##.###"))
     End If
     '-- Where no sales were only sales in the previous year --
     If (adoID3!NewQuantity < 1) And (adoID3!OldQuantity > 0) _
          Then
        adoID3!UpByHowMuch = CSng(Format(-1, "##.###"))
     End If
     adoID3.Update
   End If
   .MoveNext
 Wend
End With

adoConnection.Execute _
("DELETE * FROM ID3 WHERE ((OldQuantity = 0)" _
 & " AND (NewQuantity = 0))")
sbStatus.Panels.Item(1).Text = ""
sbStatus.Panels.Item(2).Text = "Ready."
ProgressBar1.Value = 0

Exit Sub
myError:
MsgBox (Err.Description)

End Sub
Private Sub determineEntropy()

Dim adoTemp As ADODB.Recordset
Dim sSql As String
Dim totalSamples As Integer
Dim entropyCountry As Single
Dim entropyCountryLanguage As Single
Dim entropyCountryRegion As Single
Dim Position(2, 1) As Variant 'holds the classifications

sbStatus.Panels.Item(1).Text = "Determining Entropy..."
sbStatus.Panels.Item(2).Text = "Working..."

adoConnection.Execute ("DELETE * FROM ID3 WHERE Category = NULL")
'-----------------------------------------------------
'-- Determine how many records are in the ID3 Table --
'-----------------------------------------------------
Set adoTemp = New ADODB.Recordset
sSql = "SELECT count(*) as HowMany from ID3"
adoTemp.Open sSql, adoConnection
totalSamples = adoTemp!HowMany
adoTemp.Close

'----------------------------------------------------------
'-- Determine the relative Entropy on each of the fields --
'----------------------------------------------------------
entropyCountryLanguage = getEntropy _
     ("CountryLanguage", totalSamples)
Position(0, 0) = entropyCountryLanguage
Position(0, 1) = "CountryLanguage"
entropyCountryRegion = getEntropy("CountryRegion", totalSamples)
Position(1, 0) = entropyCountryRegion
Position(1, 1) = "CountryRegion"
entropyCountry = getEntropy("Country", totalSamples)
Position(2, 0) = entropyCountry
Position(2, 1) = "Country"

Call qsort(Position, LBound(Position), UBound(Position))

txtAnalysis = "ID3 Analysis of the Category " & lstCategory & _
    " and the Product " & lstProduct & vbCrLf
txtAnalysis = txtAnalysis & "The lesser the entropy, the more "
txtAnalysis = txtAnalysis & "important is this Attribute "
txtAnalysis = txtAnalysis & "to overall Sales" & vbCrLf
txtAnalysis = txtAnalysis & Position(0, 1) & " Entropy: " & _
    Position(0, 0) & vbCrLf
txtAnalysis = txtAnalysis & Position(1, 1) & " Entropy: " & _
    Position(1, 0) & vbCrLf
txtAnalysis = txtAnalysis & Position(2, 1) & " Entropy: " & _
    Position(2, 0) & vbCrLf
txtAnalysis = txtAnalysis & "Product Manager - Review sales to: "
txtAnalysis = txtAnalysis & Position(0, 1) & vbCrLf

Call gridTheResults

sbStatus.Panels.Item(1).Text = ""
sbStatus.Panels.Item(2).Text = "Ready."

End Sub


Public Function getEntropy(ByVal sField As String, itotalRecords As Integer) As Single

'--------------------------------------------------------------
'-- First, determine how many unique values of this field are
'-- in the table. Store them and determine the entropy for each
'--------------------------------------------------------------

Dim sSql As String
Dim adoHoldGroup As ADODB.Recordset
Dim totalUp As Integer
Dim totalDown As Integer
Dim totalSamples As Integer
Dim fractionalEntropy As Single
Dim entropy As Single
Dim searchValue As String

fractionalEntropy = 0

Set adoHoldGroup = New ADODB.Recordset
adoHoldGroup.CursorLocation = adUseClient

'-- grab the info here--
sSql = "SELECT * FROM ID3 ORDER BY " & sField
adoHoldGroup.Open sSql, adoConnection
If (adoHoldGroup.BOF) And (adoHoldGroup.EOF) Then
  MsgBox "No records in the ID3 Table"
  Exit Function
End If

totalSamples = adoHoldGroup.RecordCount
adoHoldGroup.MoveFirst

searchValue = adoHoldGroup.Fields.Item(sField)

With adoHoldGroup
  
  While Not .EOF
       totalUp = 0
       totalDown = 0
       '---------------------------------------------------
       '-- Loop through all of the records with like values
       '---------------------------------------------------
       Do While (searchValue = .Fields.Item(sField))
          If !SalesUp = True Then
            totalUp = totalUp + 1
          Else
            totalDown = totalDown + 1
          End If
          .MoveNext
          If (.EOF) Then Exit Do
       Loop
       '---------------------------------------------------
       '-- Now determine the entropy for that group --
       '---------------------------------------------------
       If (totalUp = 0) Then
         fractionalEntropy = -1#
       ElseIf (totalDown = 0) Then
         fractionalEntropy = -1#
       Else
            fractionalEntropy = -(totalUp / totalSamples) * Log(totalUp / totalSamples) - (totalDown / totalSamples) * Log(totalDown / totalSamples)
       End If
       fractionalEntropy = (fractionalEntropy / itotalRecords)
       entropy = entropy + fractionalEntropy
       If (.EOF) Then
         getEntropy = entropy + fractionalEntropy
       Else
         searchValue = adoHoldGroup.Fields.Item(sField)
       End If
  Wend
End With

End Function


Public Sub gridTheResults()

'------------------------------------------------
'-- Now let's update the grid with the regions --
'------------------------------------------------
Dim adoID3 As ADODB.Recordset
Dim sSql As String
Dim iRows As Integer
Dim iCols As Integer
Dim iRowLoop As Integer
Dim iColLoop As Integer


sSql = "SELECT UpByHowMuch, OldQuantity, NewQuantity,"
sSql = sSql & " CountryRegion, CountryLanguage, Country FROM"
sSql = sSql & " ID3 ORDER BY UpByHowMuch DESC,"
sSql = sSql & " CountryRegion, CountryLanguage"


Set adoID3 = New ADODB.Recordset
adoID3.CursorLocation = adUseClient

adoID3.Open sSql, adoConnection, , , adCmdText

adoID3.MoveFirst

iRows = adoID3.RecordCount
iCols = adoID3.Fields.Count

resultsGrid.Rows = iRows
resultsGrid.Cols = iCols
'--------------------------
'-- Set up the grid here --
'--------------------------
resultsGrid.Row = 0

    
For iColLoop = 0 To resultsGrid.Cols - 1
  With resultsGrid
     .Col = iColLoop
     .ColWidth(iColLoop) = 1400
     .ColAlignment(iColLoop) = 7
     Select Case iColLoop
       Case 0
         .Text = "Growth Factor"
         .MergeCol(iColLoop) = True
       Case 1
         .Text = "Previous Qty"
         .MergeCol(iColLoop) = True
       Case 2
         .Text = "Recent Qty"
         .MergeCol(iColLoop) = True
       Case 3
         .Text = "Country Region"
       Case 4
          .Text = "Country Language"
       Case 5
          .Text = "Country"
     End Select
  End With
Next

resultsGrid.MergeCells = flexMergeFree

For iRowLoop = 1 To iRows - 1
  For iColLoop = 0 To iCols - 1
     resultsGrid.Row = iRowLoop
     resultsGrid.Col = iColLoop
     resultsGrid.Text = adoID3.Fields(iColLoop)
  Next
  adoID3.MoveNext
Next

adoID3.Close

Set adoID3 = Nothing

End Sub


Public Sub qsort(ByRef myArray() As Variant, ByVal iLowBound As Integer, ByVal iHighBound As Integer)


Dim intX As Integer
Dim intY As Integer
Dim intMiddle As Integer
Dim sHoldString As String
Dim varMidBound As Variant
Dim varTmp As Variant

If iHighBound > iLowBound Then
   intMiddle = ((iLowBound + iHighBound) \ 2)
   varMidBound = myArray(intMiddle, 0)
   
   intX = iLowBound
   intY = iHighBound
   
   Do While intX <= intY
      '-- if a value lower in the array is > than one --
      '-- higher in the array, swap them now          --
      If myArray(intX, 0) >= varMidBound _
           And myArray(intY, 0) <= varMidBound Then
             varTmp = myArray(intX, 0)
             sHoldString = myArray(intX, 1)
             myArray(intX, 0) = myArray(intY, 0)
             myArray(intX, 1) = myArray(intY, 1)
             myArray(intY, 0) = varTmp
             myArray(intY, 1) = sHoldString
             intX = intX + 1
             intY = intY - 1
      Else
        If myArray(intX, 0) < varMidBound Then
           intX = intX + 1
        End If
        If myArray(intY, 0) > varMidBound Then
          intY = intY - 1
        End If
      End If
   Loop
   Call qsort(myArray(), iLowBound, intY)
   Call qsort(myArray(), intX, iHighBound)
   
End If
End Sub


