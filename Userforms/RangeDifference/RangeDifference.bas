Attribute VB_Name = "RangeDifference"
Public Sub Start_RangeDifferenceForm()
'' Description:
'' * Refer to this subroutine in a worksheet button/control to display the RangeDifferenceForm.
    RangeDifferenceForm.Show
End Sub
Public Sub RangeDifference(ByRef A As Range, ByRef B As Range, ByRef outputRange As Range, Optional hasHeaders As Boolean)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Description:
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' * Finds all "records" (row with values in each column) that are in worksheet range A that are NOT in
' * worksheet range B (A - B).
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Notes:
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' * # of columns in A and B must match.
' * If the hasHeaders option is true, then A & B's column titles must match.

Application.ScreenUpdating = False

Dim numBlanks, outputRangeRow As Integer: outputRangeRow = 1
Dim origBRow As Integer: origBRow = 1
Dim aRow, bRow, col As Integer: aRow = 1: bRow = 1

' If option to look at headers included, then add headers to outputRange and skip
If hasHeaders = True Then
    For col = 1 To A.Columns.Count
        outputRange.Cells(outputRangeRow, col).Value = A.Cells(1, col).Value
    Next col
    outputRangeRow = outputRangeRow + 1
    aRow = aRow + 1
    bRow = bRow + 1
    origBRow = bRow
End If

' Perform the difference algorithm:
Do While aRow <> A.Rows.Count + 1
    ' Determine if record is blank. If true then proceed to next record in A:
    numBlanks = 0
    For col = 1 To A.Columns.Count
        If A.Cells(aRow, col).Value = vbNullString Then
            numBlanks = numBlanks + 1
        End If
    Next col
    ' If record is blank (all cells are blank) then skip to next record in A:
    If numBlanks = A.Columns.Count Then
        GoTo SkipA
    End If
    bRow = origBRow
    Do While bRow <> B.Rows.Count + 1
        ' Compare record in A to record in B. If match found, then break loop and proceed to next record.
        ' Otherwise continue searching through B:
        matchedColumns = 0
        For col = 1 To B.Columns.Count
            If StrComp(A.Cells(aRow, col).Value, B.Cells(bRow, col).Value) = 0 Then
                GoTo SkipA
            ElseIf bRow <> B.Rows.Count Then
                GoTo SkipB
            End If
        Next col
        ' If routine reaches this point then record in A must have been matched.
        ' Put to output range:
        Set outputRange = outputRange.Resize(RowSize:=outputRangeRow)
        ' Add to outPut:
        For col = 1 To A.Columns.Count
            outputRange.Cells(outputRangeRow, col).Value = A.Cells(aRow, col).Value
        Next col
        outputRangeRow = outputRangeRow + 1
SkipB:
    bRow = bRow + 1
    Loop
SkipA:
aRow = aRow + 1
Loop

Application.ScreenUpdating = True

End Sub
