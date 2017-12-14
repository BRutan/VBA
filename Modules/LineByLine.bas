Attribute VB_Name = "LineByLine"
Public Sub LineByLine(ByRef tableRange As Range, RVUColumn As Long, CPTColumn As Long, ProposedPriceColumn As Long, SuggestedPriceColumn As Long, ByRef CPTManifest As Range, Optional functionMode As Integer)
'' TODO:
'''''''
'' * Create CPT Manifest Workbook

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Description:
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' * Macro partially automates the Line-By-Line process.
'   -6_12_2017: Highlights the Proposed Price column in red where the RVU and Proposed Prices are not in order (ex: high RVU item has lower price than lower RVU item).
' * Parameter List:
'   - tableRange: table containing Line-By-Line data within a hospital's pricing file. (At the very least will contain CPT, Current Price, RVU, Proposed Price).
'   - RVUColumn: column containing the RVUs for each item (if selected range contains more than one column, takes column # of first column in selected range).
'   - CPTColumn: column containing the CPTs for each item (if selected range contains more than one column, takes column # of first column in selected range).
'   - ProposedPriceColumn: column containing the proposed prices (if selected range contains more than one column, takes column # of first column in selected range).
'   - SuggestedPricecolumn: column containing the suggested prices (if selected range contains more than one column, takes column # of first column in selected range).
'   - CPTManifest: Range containing mapping of CPTs to "groups" (grouping together related CPTs / have modifiers).
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Notes:
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' *
'
'
'

' Make backup of active sheet:
Call Misc.BackupActiveSheet(Application.ActiveSheet)

Call Macro_Utilities.CodeOptimizeSettings(True)

' Rescale the columns:
GroupColumn = CPTColumn - tableRange.Column + 2
RVUColumn = RVUColumn - tableRange.Column + 1
ProposedPriceColumn = ProposedPriceColumn - tableRange.Column + 1
SuggestedPriceColumn = SuggestedPriceColumn - tableRange.Column + 1

' Declare variables for use in macro:
Dim row As Integer
Dim currGroup As String
Dim groupRows As IntVector: Set groupRows = New IntVector
Dim ProposedPrices As IntVector: Set ProposedPrices = New IntVector

' Insert "Key" column in the tableRange to right of the "keyColumn" and make color yellow:
'' (Note: temporary, will ideally use the FY<Current Year> LBL Key column)
' To use, add keyColumn as parameter (Long type):
'tableRange.Cells(1, keyColumn).Offset(0, 1).EntireColumn.Insert
'tableRange.Cells(1, keyColumn).Value = "Key"
'tableRange.Cells(1, keyColumn).EntireColumn.Interior.Color = RGB(255, 255, 0)

' Determine if item should be an anchor, put "A" in keyColumn if true:
'' (Note: temporary, ideally will omit this step and just do pricing based upon rules
'' and update the LBL key).

''''' Variables for sorting:
Dim hasChanged As Boolean
Dim i, j As Integer
Dim tempRow As Integer

row = 1

Do While row <> tableRange.Rows.count + 1
    
    ' Skip all blank and non-numeric rows:
    Do While IsNumeric(tableRange.cells(row, GroupColumn).value) = False Or tableRange.cells(row, GroupColumn).value = vbNullString
        row = row + 1
    Loop
    
    groupRows.Clear
    
    currGroup = tableRange.cells(row, GroupColumn).value
    groupRows.Push (row)
    row = row + 1
    ' Accumulate Group Rows until the group number changes:
    Do While StrComp(tableRange.cells(row, GroupColumn).value, currGroup) = 0 And row <> tableRange.Rows.count + 1
        groupRows.Push (row)
        row = row + 1
    Loop
    ' Use rules to determine anchor, override proposed prices if needed:
    If groupRows.Size = 1 Then
        ' If only one group member in the table, proceed to next loop iteration:
        GoTo NextMain
    Else
''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Rules go HERE
''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '' (6.8.2017: For now just determine if proposed prices are appropriate given RVUs (i.e. well ordered). If not then flag
        '' by highlighting cells.)
        ' Order the rows by RVU:
        hasChanged = True
        counter = 0
        Do While hasChanged = True
            hasChanged = False
            For i = 1 To groupRows.Size
                For j = i To groupRows.Size
                    If val(tableRange.cells(groupRows.GetValue(CInt(i)), RVUColumn).value) < val(tableRange.cells(groupRows.GetValue(CInt(j)), RVUColumn).value) Then
                        tempRow = groupRows.GetValue(CInt(i))
                        groupRows.SetValue CInt(i), groupRows.GetValue(CInt(j))
                        groupRows.SetValue CInt(j), tempRow
                        hasChanged = True
                    End If
                Next j
            Next i
        Loop
        ' Now that the rows are ordered by RVU, check to see if order matches the order of Proposed Prices.
        ' If false, then highlight all of the cells in the group:
        For i = 1 To groupRows.Size
            ' Copy Proposed Prices:
            ProposedPrices.Push (val(tableRange.cells(groupRows.GetValue(CInt(i)), ProposedPriceColumn).value))
        Next i
        ' Determine if Proposed Prices are out of order:
        For i = 1 To ProposedPrices.Size
            If ProposedPrices.GetValue(CInt(i)) < ProposedPrices.GetValue(CInt(i + 1)) Then
                ' Highlight ProposedPrice cells for records in group:
                tableRange.Range(cells(groupRows.Min, ProposedPriceColumn), cells(groupRows.Max, ProposedPriceColumn)).Interior.Color = RGB(255, 150, 150)
            End If
        Next i
''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''
    End If
NextMain:
Loop

Set mainRange = tableRange

Call Macro_Utilities.CodeOptimizeSettings(False)

End Sub

Private Function CreateGroupColumn(tableRange As Range, CPTColumn As Long, CPTManifest As Range) As Long
    '' Description: Create group column in the pricing file table and return its column number in the table.
    '' TODO: Create CPT manifest workbook and macro to generate the group column
    '' For now assume that CPTManifest columns are / CPT / GroupNum /
    
    Dim currCPT As String
    Dim GroupColumn As Long
    Dim notMatched As Boolean
    
    ' Insert row beside CPT column to put group number:
    tableRange.cells(1, CPTColumn).Offset(0, 1).EntireRow.Insert
    GroupColumn = CPTColumn + 1
    
    Dim tableRow, manifestRow As Long
    For tableRow = 1 To tableRange.cells(1, CPTColumn).Rows.count
        currCPT = tableRange.cells(tableRow, GroupColumn).value
        For manifestRow = 1 To CPTManifest.Rows.count
          notMatched = True
          ' Search for CPT in the manifest:
          If (StrComp(currCPT, CPTManifest.cells(manifestRow, 1).value) = 0) Then
            ' If found then put group number into the GroupColumn cell:
            notMatched = False
            tableRange.cells(tableRow, GroupColumn).value = CPTManifest.cells(manifestRow, 2).value
            Exit For
          End If
        Next manifestRow
        ' If CPT is not in the manifest, put "#N/A":
        If notMatched = True Then
            tableRange.cells(tableRow, GroupColumn).value = "#N/A"
        End If
    Next tableRow
    
    ' Return the group column number:
    CreateGroupColumn = GroupColumn
    
End Function

Function PriceStepper(targetRVU As String, Price1 As String, Price2 As String, RVU1 As String, RVU2 As String, Optional firstIsIntercept As Boolean = False) As Double
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Description: Outputs "stepped" price that makes grouped prices consistent with RVUs (are monotonically increasing).
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo Exit_With_Error
Application.Volatile

'' Check input integrity (at least one passed RVU must not be 0):
If (StrComp(Trim(Price1), vbNullString) = 0 And StrComp(Trim(Price2), vbNullString) = 0) Or (StrComp(Trim(RVU1), vbNullString) = 0 And StrComp(Trim(RVU2), vbNullString) = 0) _
    Or (StrComp(Trim(targetRVU), vbNullString) = 0) Then
    GoTo Exit_With_Error
ElseIf StrComp(Trim(RVU1), vbNullString) = 0 And CDbl(RVU2) = 0 Then
    GoTo Exit_With_Error
ElseIf StrComp(Trim(RVU2), vbNullString) = 0 And CDbl(RVU1) = 0 Then
    GoTo Exit_With_Error
ElseIf IsNumeric(Price1) = False Or IsNumeric(Price2) = False Or IsNumeric(RVU1) = False Or IsNumeric(RVU2) = False Then
    GoTo Exit_With_Error
ElseIf CDbl(RVU1) - CDbl(RVU2) = 0 Then
    GoTo Exit_With_Error
End If

'' Set relevant variables for output:
Dim price_1, price_2, rvu_1, rvu_2, slope, intercept As Double

price_1 = CDbl(Price1)
price_2 = CDbl(Price2)
rvu_1 = CDbl(RVU1)
rvu_2 = CDbl(RVU2)

''' Compute the slope:
slope = (price_1 - price_2) / (rvu_1 - rvu_2)

''' Compute the intercept (actual price - "predicted price" determined by slope alone):
intercept = IIf(firstIsIntercept = True, price_1, price_2) - slope * IIf(firstIsIntercept = True, rvu_1, rvu_2)

''' Output the stepped price:
PriceStepper = CDbl(targetRVU) * slope + intercept

Exit Function

Exit_With_Error:
    PriceStepper = "#N/A"
    Exit Function

End Function
