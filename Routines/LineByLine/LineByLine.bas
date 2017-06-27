Attribute VB_Name = "LineByLine"

Public Sub Start_LineByLineForm()
'' Description:
'' *Use this subroutine to run the userform:
    Application.ActiveSheet.ScrollArea = ""
    Dim form As New LineByLineForm
    form.Show
End Sub

Public Sub LineByLine(ByRef tableRange As Range, RVUColumn As Long, CPTColumn As Long, ProposedPriceColumn As Long, SuggestedPriceColumn As Long, ByRef CPTManifest As Range, Optional functionMode As Integer)
'' TODO:
'' * Create CPT Manifest Workbook
'' * Finish Undo action.
'' *

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

Application.ScreenUpdating = False

' Rescale the columns:
GroupColumn = CPTColumn - tableRange.column + 2
RVUColumn = RVUColumn - tableRange.column + 1
ProposedPriceColumn = ProposedPriceColumn - tableRange.column + 1
SuggestedPriceColumn = SuggestedPriceColumn - tableRange.column + 1

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

Do While row <> tableRange.Rows.Count + 1
    
    ' Skip all blank and non-numeric rows:
    Do While IsNumeric(tableRange.Cells(row, GroupColumn).Value) = False Or tableRange.Cells(row, GroupColumn).Value = vbNullString
        row = row + 1
    Loop
    
    groupRows.Clear
    
    currGroup = tableRange.Cells(row, GroupColumn).Value
    groupRows.Push (row)
    row = row + 1
    ' Accumulate Group Rows until the group number changes:
    Do While StrComp(tableRange.Cells(row, GroupColumn).Value, currGroup) = 0 And row <> tableRange.Rows.Count + 1
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
                    If val(tableRange.Cells(groupRows.GetValue(CInt(i)), RVUColumn).Value) < val(tableRange.Cells(groupRows.GetValue(CInt(j)), RVUColumn).Value) Then
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
            ProposedPrices.Push (val(tableRange.Cells(groupRows.GetValue(CInt(i)), ProposedPriceColumn).Value))
        Next i
        ' Determine if Proposed Prices are out of order:
        For i = 1 To ProposedPrices.Size
            If ProposedPrices.GetValue(CInt(i)) < ProposedPrices.GetValue(CInt(i + 1)) Then
                ' Highlight ProposedPrice cells for records in group:
                tableRange.Range(Cells(groupRows.Min, ProposedPriceColumn), Cells(groupRows.Max, ProposedPriceColumn)).Interior.Color = RGB(255, 150, 150)
            End If
        Next i
''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''
    End If
NextMain:
Loop

Set mainRange = tableRange
Application.OnUndo "Undoing LineByLine...", "UndoLineByLine"

Application.ScreenUpdating = True

End Sub

Private Function CreateGroupColumn(tableRange As Range, CPTColumn As Long, CPTManifest As Range) As Long
    '' Description: Create group column in the pricing file table and return its column number in the table.
    '' TODO: Create CPT manifest workbook and macro to generate the group column
    '' For now assume that CPTManifest columns are / CPT / GroupNum /
    
    Dim currCPT As String
    Dim GroupColumn As Long
    Dim notMatched As Boolean
    
    ' Insert row beside CPT column to put group number:
    tableRange.Cells(1, CPTColumn).Offset(0, 1).EntireRow.Insert
    GroupColumn = CPTColumn + 1
    
    Dim tableRow, manifestRow As Long
    For tableRow = 1 To tableRange.Cells(1, CPTColumn).Rows.Count
        currCPT = tableRange.Cells(tableRow, GroupColumn).Value
        For manifestRow = 1 To CPTManifest.Rows.Count
          notMatched = True
          ' Search for CPT in the manifest:
          If (StrComp(currCPT, CPTManifest.Cells(manifestRow, 1).Value) = 0) Then
            ' If found then put group number into the GroupColumn cell:
            notMatched = False
            tableRange.Cells(tableRow, GroupColumn).Value = CPTManifest.Cells(manifestRow, 2).Value
            Exit For
          End If
        Next manifestRow
        ' If CPT is not in the manifest, put "#N/A":
        If notMatched = True Then
            tableRange.Cells(tableRow, GroupColumn).Value = "#N/A"
        End If
    Next tableRow
    
    ' Return the group column number:
    CreateGroupColumn = GroupColumn
    
End Function
