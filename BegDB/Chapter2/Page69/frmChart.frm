VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmChart 
   Caption         =   "Order Details"
   ClientHeight    =   5190
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   6675
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   5190
   ScaleWidth      =   6675
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   300
      Left            =   60
      TabIndex        =   1
      Top             =   3960
      Width           =   1080
   End
   Begin MSChart20Lib.MSChart chtReport 
      Height          =   3840
      Left            =   60
      OleObjectBlob   =   "frmChart.frx":0000
      TabIndex        =   0
      Top             =   60
      Width           =   6360
   End
End
Attribute VB_Name = "frmChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Const MARGIN_SIZE = 60 'In Twips
Private Const SHAPE_COMMAND = "SHAPE {select UnitPrice,Quantity from [Order Details] Order by UnitPrice} AS ChildCommand COMPUTE ChildCommand, COUNT(ChildCommand.[Quantity]) AS [Quantity] BY [UnitPrice]"
Private Const CONNECT_STRING = "PROVIDER=MSDataShape;Data Source=C:\BegDB\Nwind.mdb;Data Provider=Microsoft.Jet.OLEDB.3.51"
Private Const FIELD_X = "UnitPrice"
Private Const FIELD_Y = "Quantity"
Private Const FIELD_Z = ""
Private Const VBERR_INVALID_PROCEDURE_CALL = 5
Private Const MARKERS_VISIBLE = 0
Private Const BRACKET_LEFT = "["
Private Const BRACKET_RIGHT = "]"
Private Const SPACE_CHAR = " "

Private Sub cmdClose_Click()
    Unload Me
End Sub

'-------------------------------------------------------------------------
'Purpose:   Display an error message to the user
'In:
'   [oError]
'           Error object containing error information
'-------------------------------------------------------------------------
Private Sub DisplayError(oError As ErrObject)
    MsgBox oError.Description, vbExclamation, App.Title
End Sub

Private Sub Form_Load()
    Dim conShape As ADODB.Connection
    Dim recShape As ADODB.Recordset
    
    On Error GoTo Form_Load_Error
    'Create and open connection to the Data Shape provider
    Set conShape = New ADODB.Connection
    conShape.Open CONNECT_STRING
    'Create and open a recordset
    Set recShape = New ADODB.Recordset
    recShape.Open SHAPE_COMMAND, conShape
    'Fill the chart with the recordset data
    ShowRecordsInChart recShape, FIELD_X, FIELD_Y, FIELD_Z
    'Show or hide markers
    ShowMarkers MARKERS_VISIBLE
    Exit Sub
Form_Load_Error:
    DisplayError Err
    Exit Sub
End Sub

Private Sub Form_Resize()
    Dim sngButtonTop As Single
    Dim sngScaleWidth As Single
    Dim sngScaleHeight As Single
    
    On Error GoTo Form_Resize_Error
    With Me
        sngScaleWidth = .ScaleWidth
        sngScaleHeight = .ScaleHeight
        'Move Close button to the lower right corner
        With .cmdClose
            sngButtonTop = sngScaleHeight - (.Height + MARGIN_SIZE)
            .Move sngScaleWidth - (.Width + MARGIN_SIZE), sngButtonTop
        End With
        .chtReport.Move MARGIN_SIZE, _
                        MARGIN_SIZE, _
                        sngScaleWidth - (2 * MARGIN_SIZE), _
                        sngButtonTop - (2 * MARGIN_SIZE)
    End With
    Exit Sub
Form_Resize_Error:
    'An error will occur if the user sizes
    'the form so small that negative heights
    'or widths are calculated
    Resume Next
End Sub

'-------------------------------------------------------------------------
'Purpose:   Determines if the passed key is being used in the
'           passed collection.
'In:
' [cCol]    The collection to check for key use in.
' [sKey]    The key to look for.
'Return:    If the key is being used by the collection, true
'           is returned.  Otherwise, false is returned.
'-------------------------------------------------------------------------
Private Function IsKeyInCollection(cCol As Collection, sKey As String) As Boolean
    Dim v As Variant
    On Error Resume Next
    v = cCol.Item(sKey)
    'It is important to check for error 5, rather than checking for
    'any error, because an error could occur even if the key is valid.
    'If the key existed, but it was associated with an element that
    'was an object, an error would occur because 'Set' wasn't used
    'to assign it to 'v'.
    IsKeyInCollection = (Err.Number <> VBERR_INVALID_PROCEDURE_CALL)
    Err.Clear
End Function

'----------------------------------------------------------
'Purpose:   Shows or Hides series markers, according to the
'           parameter.
'In:
' [bShow]   If true, all the series markers will be shown.
'           Otherwise, all the series markers will be hidden.
'----------------------------------------------------------
Private Sub ShowMarkers(bShow As Boolean)
    Dim i As Long
    On Error GoTo ShowMarkers_Click_Error
    With chtReport.Plot
        For i = 1 To .SeriesCollection.Count
            .SeriesCollection(i).SeriesMarker.Show = bShow
        Next
    End With
    Exit Sub
ShowMarkers_Click_Error:
    DisplayError Err
    Exit Sub
End Sub

'----------------------------------------------------------
'Purpose:   Displays the data summarized in the passed recordset
'           in the Chart.
'In:
' [recParent]
'           A recordset created using a Shape command, that
'           groups by one or two fields, and summarizes one.
' [sFldX]
'           The name of the field to group by on the X axis.
' [sFldY]
'           The name of the field to summarize on the Y axis.
' [sFldZ]
'           The name of the field to group by on the Z axis. This
'           field should be a zero length string, if the recordset
'           only groups by one field.
'----------------------------------------------------------
Private Sub ShowRecordsInChart(recParent As Recordset, _
                               sFldX As String, _
                               sFldY As String, _
                               sFldZ As String)
                                   
    Dim bUseZ As Boolean
    Dim cRows As Collection
    Dim cCols As Collection
    Dim lCol As Long
    Dim lRow As Long
    Dim lMaxCol As Long
    Dim lMaxRow As Long
    Dim sValue As String
    
    On Error GoTo ShowRecordsInChart_Error
    If Len(sFldZ) = 0 Then bUseZ = False Else bUseZ = True
    
    Set cRows = New Collection
    Set cCols = New Collection
    
    With Me.chtReport
        'Turn off chart painting
        .Repaint = False
        With .DataGrid
            'Clear the chart
            .DeleteRows 1, .RowCount
            .DeleteColumns 1, .ColumnCount
            .DeleteColumnLabels 1, .ColumnLabelCount
            .DeleteRowLabels 1, .RowLabelCount
            'Make sure there is one level of labels
            .InsertColumnLabels 1, 1
            .InsertRowLabels 1, 1
            'If the Z axis is not being used, make
            'sure there is one column
            If Not bUseZ Then .InsertColumns 1, 1
            recParent.MoveFirst
            Do Until recParent.EOF
                'Make sure a row is added for this X field
                sValue = FixNull(recParent.Fields(sFldX).Value, False)
                If Not IsKeyInCollection(cRows, sValue) Then
                    lMaxRow = lMaxRow + 1
                    lRow = lMaxRow
                    'Store the row index associated with
                    'the Row name
                    cRows.Add lRow, sValue
                    .InsertRows lRow, 1
                    .RowLabel(lRow, 1) = sValue
                Else
                    lRow = cRows.Item(sValue)
                End If
                
                'Make sure a column is added for this Z field
                If bUseZ Then
                    sValue = FixNull(recParent.Fields(sFldZ).Value, False)
                    If Not IsKeyInCollection(cCols, sValue) Then
                        lMaxCol = lMaxCol + 1
                        lCol = lMaxCol
                        'Store the column index associated with
                        'the column name
                        cCols.Add lCol, sValue
                        .InsertColumns lCol, 1
                        .ColumnLabel(lCol, 1) = sValue
                    Else
                        lCol = cCols.Item(sValue)
                    End If
                    'Set the datapoint value for this record's row and column
                    .SetData lRow, lCol, FixNull(recParent.Fields.Item(sFldY).Value, True), 0
                Else
                    'Set the datapoint value for this record's row
                    'There is only one column in this case
                    .SetData lRow, 1, FixNull(recParent.Fields.Item(sFldY).Value, True), 0
                End If
                'Move the recordset to the next record
                recParent.MoveNext
            Loop
        End With
        'Turn painting back on
        .Repaint = True
    End With
    Exit Sub
ShowRecordsInChart_Error:
    'Make sure the charts painting is turned back on
    Me.chtReport.Repaint = True
    DisplayError Err
    Exit Sub
End Sub

'-------------------------------------------------------------------------
'Purpose:   Checks a variant value for null.  If the value is null, returns
'           a vbNullString or a zero.
'In:
' [vField]
'           The variant to check for null.
' [bNumericRequired]
'           If true, return 0 if the variant is null.  Otherwise, return
'           vbNullString.
'-------------------------------------------------------------------------
Private Function FixNull(vField As Variant, _
                        bNumericRequired As Boolean) As Variant
    If IsNull(vField) Then
        If bNumericRequired Then
            FixNull = 0
        Else
            FixNull = vbNullString
        End If
    Else
        FixNull = vField
    End If
End Function



