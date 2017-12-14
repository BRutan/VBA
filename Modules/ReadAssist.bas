Attribute VB_Name = "ReadAssist"
Option Explicit
''''''''''''''
'' IN PROGRESS:
'' 1. Figure out how to overwrite worksheet functions from Personal Workbook.
'' 2. Add some type of warning

Private isOn As Boolean

Sub Worksheet_SelectionChange(ByVal Target As Excel.Range)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Description: Used to highlight row and column for easier row reference.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If isOn = True Then
    
    ' Use static variables to store previous row and column (note: sometimes fails to be added):
    Static prevRow As Range
    Static prevColumn As Range
    
    Dim thisWS As Worksheet: Set thisWS = ActiveSheet
    Dim selectedCell As Range: Set selectedCell = Application.Selection
    Dim formatCond1, formatCond2 As FormatCondition
    
    ' Delete conditional formatting for previously selected row/column:
    If Not (prevRow Is Nothing) Then
        prevRow.FormatConditions.Delete
    End If
    If Not (prevColumn Is Nothing) Then
        prevColumn.FormatConditions.Delete
    End If
    Set prevRow = selectedCell.EntireRow
    Set prevColumn = selectedCell.EntireColumn
    
    ' Add conditional formatting to whole column and row of selected cell:
    Set formatCond1 = selectedCell.EntireRow.FormatConditions.Add(Type:=xlExpression, Formula1:="=1=1")
    Set formatCond2 = selectedCell.EntireColumn.FormatConditions.Add(Type:=xlExpression, Formula1:="=1=1")
    
    With formatCond1
        .Interior.Color = vbYellow
        .Interior.Pattern = xlSolid
        .Font.Color = vbBlack
    End With
    With formatCond2
        .Interior.Color = vbYellow
        .Interior.Pattern = xlSolid
        .Font.Color = vbBlack
    End With
End If

End Sub

Sub TurnOnReadAssist()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Description: Used to turn read assist on or off.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

If isOn = False Then
    isOn = True
Else
    isOn = False
End If

End Sub
