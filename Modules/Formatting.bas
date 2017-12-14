Attribute VB_Name = "Formatting"
Option Explicit

Sub CDMSummarySizing()

Dim cdmRecSummaryRange As Range: Set cdmRecSummaryRange = Application.Selection
Dim row, col As Integer

' Size the columns:
For row = 1 To cdmRecSummaryRange.Rows.count
    For col = 1 To cdmRecSummaryRange.Columns.count
        If InStr(cdmRecSummaryRange.cells(row, col).value, "Subtotal") <> 0 Then
            ' Grey out the row:
            
        End If
    Next col
Next row


End Sub

Sub SpawnCDMRecSummary()
'
' SpawnCDMRecSummary Macro
' Spawns the CDM Reconciliation Summary, still needs to have columns positioned.
'
' FORMAT: Line Count and # Lines w/ Usage have column width of 12
' Units/Charges have column width of 13
' Area of Review has column width of 26
Range("B1:K29").Clear
Application.ScreenUpdating = False
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "Area Of Review"
    Range("C2").Select
    ActiveCell.FormulaR1C1 = "Line Count"
    Range("D2").Select
    ActiveCell.FormulaR1C1 = "Line Count w/ Usage"
    Range("E2").Select
    ActiveCell.FormulaR1C1 = "IP"
    Range("F2").Select
    ActiveCell.FormulaR1C1 = "OP"
    Range("G2").Select
    ActiveCell.FormulaR1C1 = "Total"
    Range("E2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Range("H2").Select
    ActiveSheet.Paste
    Range("E1:G1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    ActiveCell.FormulaR1C1 = "Units"
    Range("E1:G1").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("H1").Select
    ActiveSheet.Paste
    ActiveCell.FormulaR1C1 = "Charges"
    Range("K1:K2").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    ActiveCell.FormulaR1C1 = "% of Total"
    Range("D1:D2").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("C1:C2").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("B1:B2").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("B29").Select
    ActiveCell.FormulaR1C1 = "Totals"
    Range("C29").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[-26]C:R[-1]C)"
    Range("C29:J29").Select
    Selection.FillRight
    Range("K29").Select
    Selection.End(xlUp).Select
    Range("K3").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]/R29C10"
    Range("K3:K28").Select
    Selection.FillDown
    Range("B1:B2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range("B1:K2").Select
    With Selection
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With
    Range("B1:B2").Select
    Application.CutCopyMode = False
    Range("B1:D2").Select
    Application.Run "PERSONAL.XLSB!Buonopane_Blue"
    Range("E1:J1").Select
    Application.Run "PERSONAL.XLSB!Buonopane_Blue"
    Range("K1:K2").Select
    Application.Run "PERSONAL.XLSB!Buonopane_Blue"
    Range("J2").Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Application.Run "PERSONAL.XLSB!Buonopane_Gray"
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("B1:B2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range("B1:K2").Select
    Selection.Font.Bold = True
    Range("B3").Select
    Selection.End(xlDown).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range("B29:K29").Select
    Application.Run "PERSONAL.XLSB!Buonopane_Blue"
    Selection.Font.Bold = True
    Range("K29").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]/R29C10"
    Range("K28").Select
    Selection.End(xlUp).Select
    Rows("1:2").Select
    Range("B1").Activate
    Selection.RowHeight = 30.75
    Columns("B:B").ColumnWidth = 15
    Range("C3:D28").Select
    Application.Run "PERSONAL.XLSB!Buonopane_Yellow"
    Range("K3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("K3:K28").Select
    Application.Run "PERSONAL.XLSB!Buonopane_Yellow"
    Range("B1:K6").Select
    Range("K1").Activate
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
With Range("B1:K29")
    .Font.Name = "Arial"
    .Font.Size = 11
End With
Range("B1:K1").Font.Bold = True
Range("B29:K29").Font.Bold = True

Application.ScreenUpdating = True
End Sub

Option Explicit

Sub SpawnBenchmarkStats()
''''''''''''''''''''''''''''''''''''''''''''''''''''
'' IN PROGRESS:
''''''''''''''''''''''''''''''''''''''''''''''''''''
' BenchmarkStatsSpawn Macro:
' Spawns the Benchmark Stats template used with pricing files.
' Row height is 33 for two merged rows
Application.ScreenUpdating = False

If StrComp(Range("B1").value, "Area of Review") <> 0 Then
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Area of Review"
    Range("C2").Select
    ActiveCell.FormulaR1C1 = "Charges w/" & Chr(10) & "BM's"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = _
        "Current Price (CP) Compared to Local Hospital Benchmarks"
    Range("D2").Select
    ActiveCell.FormulaR1C1 = "<MIN"
    Range("E2").Select
    ActiveCell.FormulaR1C1 = ">= MIN" & Chr(10) & "<= 10th"
    Range("F2").Select
    ActiveCell.FormulaR1C1 = "> 10th" & Chr(10) & "<= 20th"
    Range("G2").Select
    ActiveCell.FormulaR1C1 = ">20th"
    Range("G2").Select
    ActiveCell.FormulaR1C1 = "> 20th" & Chr(10) & "<= 30th"
    Range("H2").Select
    ActiveCell.FormulaR1C1 = "> 30th" & Chr(10) & "<= 40th"
    Range("I2").Select
    ActiveCell.FormulaR1C1 = ">50th"
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "> 50th" & Chr(10) & "<= 60th"
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "> 40th" & Chr(10) & "<= 50th"
    Range("J2").Select
    ActiveCell.FormulaR1C1 = "> 50th" & Chr(10) & "<= 60th"
    Range("K2").Select
    ActiveCell.FormulaR1C1 = ">60th" & Chr(10) & "<= 70th"
    Range("L2").Select
    ActiveCell.FormulaR1C1 = ">70th" & Chr(10) & "<= 80th"
    Range("M2").Select
    ActiveCell.FormulaR1C1 = ">80th" & Chr(10) & "<= 90th"
    Range("N2").Select
    ActiveCell.FormulaR1C1 = "> 90th" & Chr(10) & "<= MAX"
    Range("O2").Select
    ActiveCell.FormulaR1C1 = ">MAX"
    Range("F2").Select
    ActiveCell.FormulaR1C1 = "> 10th" & Chr(10) & "<= 20th"
    Range("H2").Select
    ActiveCell.FormulaR1C1 = "> 30th" & Chr(10) & "<= 40th"
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "> 40th" & Chr(10) & "<= 50th"
    Range("K2").Select
    ActiveCell.FormulaR1C1 = "> 60th" & Chr(10) & "<= 70th"
    Range("L2").Select
    ActiveCell.FormulaR1C1 = "> 70th" & Chr(10) & "<= 80th"
    Range("M2").Select
    ActiveCell.FormulaR1C1 = "> 80th" & Chr(10) & "<= 90th"
    Range("O2").Select
    ActiveCell.FormulaR1C1 = "> MAX"
    Range("O3").Select
    ActiveCell.FormulaR1C1 = ""
    Range("O2").Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Range("M2").Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Range("B1:B2").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("C1:C2").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("D1:O1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Application.Run "PERSONAL.XLSB!Buonopane_Blue"
    Range("B1:C2").Select
    Range("C1").Activate
    Application.Run "PERSONAL.XLSB!Buonopane_Blue"
    Range("D2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Application.Run "PERSONAL.XLSB!Buonopane_Gray"
    Range("B1:O2").Select
    Selection.Font.Bold = True
    Columns("B:B").Select
    Selection.ColumnWidth = 17
    Columns("C:C").ColumnWidth = 11.29
    Columns("B:O").Select
    With Selection
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With
    Range("B3:O4").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Rows("5:5").Select
    Range("B5").Activate
    Selection.RowHeight = 3
    Range("B5:O5").Select
    With Selection.Font
        .Color = -16777024
        .TintAndShade = 0
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 192
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("B6").Select
    Range("A5").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range("A5:XFC5").Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Range("A5:N5").Select
    Range("B5:O5").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 192
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("B3:O4").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.UnMerge
    Range("B3:B4").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("C3:C4").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("D4").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C/R3C[-1]"
    Range("D4:O4").Select
    Selection.FillRight
    Range("C3:C4").Select
    Range("B3:B4").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 16250343
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("B3:O4").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("C3:C4").Select
    ActiveWindow.LargeScroll Down:=1
    Range("B39:B40").Select
    ActiveWindow.LargeScroll Down:=1
    Range("B81:B82").Select
    Range("B6:B7").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("C6:C7").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("B6:C7").Select
    Application.Run "PERSONAL.XLSB!Buonopane_Blue"
    Range("D6:O6").Select
    Application.Run "PERSONAL.XLSB!Buonopane_Blue"
    Range("B6:O7").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("B6:B7").Select
    ActiveCell.FormulaR1C1 = "Total"
    Range("C6:C7").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[-3]C)"
    Range("B6:C7").Select
    Selection.Font.Bold = True
    Range("B6:C7").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With
    Range("D6:O6").Select
    Selection.Font.Bold = True
    Range("D3:O3,D6:O6").Select
    Range("D6").Activate
    Selection.NumberFormat = "#,##0"
    Range("D7").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C/R6C3"
    Range("D7:O7").Select
    Selection.FillRight
    Selection.Style = "Percent"
    Range("D4:O4").Select
    Selection.Style = "Percent"
    ActiveWindow.DisplayGridlines = False
End If
    '' TODO: Add loop that formats all charges and percent cells where percentages > 10% to yellow.
''''''''''''''''''''''''''''''''
'' Paint all >10% charges and percentages with Buonopane Yellow
''''''''''''''''''''''''''''''''
' Get final row:
    Dim rowNum As Integer: rowNum = 1
    Do While InStr(1, Range("B" + CStr(rowNum)).value, "Grand Total") = 0
        rowNum = rowNum + 1
    Loop
    Dim templateRange As Range: Set templateRange = Range("B1:O" + CStr(rowNum))
    Dim currTwoCellGroup As Range
    Dim colNum As Integer: colNum = 1
    
    
    
    
    
''''''''''''''''''''''''''''''''
'' Add Chart:
''''''''''''''''''''''''''''''''
    ' Add Chart:
    ActiveSheet.Shapes.AddChart2(227, xlLine).Select
    With ActiveSheet.Shapes(1)
        .IncrementLeft -325.5
        .IncrementTop -48
        .ScaleWidth 1.6020833333, msoFalse, _
        msoScaleFromTopLeft
        .ScaleHeight 1.0833333333, msoFalse, _
        msoScaleFromTopLeft
        .IncrementLeft -2.25
        .IncrementTop -13.5
    End With
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(1).values = "=Sheet3!$D$" & CStr(rowNum) & ":$O$" & CStr(rowNum)
    ActiveChart.FullSeriesCollection(1).Name = "=""Current Price"""
    ActiveChart.FullSeriesCollection(1).XValues = "=Sheet3!$D$2:$O$2"
    ActiveChart.ChartTitle.Select
    Selection.Delete
    ActiveChart.SetElement (msoElementLegendRight)
    ActiveSheet.ChartObjects(1).Activate
    ActiveChart.Legend.Select
    Selection.Left = 35.539
    Selection.Top = 13.562
    ActiveChart.ChartArea.Select
    ActiveSheet.Shapes("Chart 2").ScaleWidth 0.9908974057, msoFalse, _
        msoScaleFromTopLeft
    ActiveChart.PlotArea.Select
    Selection.width = 528.325
    ActiveChart.Axes(xlValue).MajorGridlines.Select
    ActiveChart.PlotArea.Select
    ActiveChart.ChartArea.Select
    ActiveChart.PlotArea.Select
    ActiveChart.Axes(xlValue).MajorGridlines.Select
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorAccent1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
    End With
    Application.CommandBars("Format Object").Visible = False
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
    End With

Application.ScreenUpdating = True

End Sub

