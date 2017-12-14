Attribute VB_Name = "Misc"
Option Explicit
Public storedWB As Workbook
Function SortRobustReference(cells As Range) As String
''' IN PROGRESS: Determine if will actually work (does indirect work via containing a static variable?)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Description: Use instead of linking cells directly to allow references to be sorting robust (won't point
'' to wrong cell if sort a range).
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Application.Volatile





End Function

Sub BorderAllAround(ByRef wb_in As Workbook, wsName As String, rangeAddress As String)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Description: Puts borders all around passed range at rangeAddress contained in worksheet wsName in workbook wb_in.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo Exit_Sub
Dim range_in As Range: Set range_in = wb_in.Worksheets(wsName).Range(rangeAddress)

range_in.Borders(xlDiagonalDown).LineStyle = xlNone
range_in.Borders(xlDiagonalUp).LineStyle = xlNone
With range_in.Borders(xlEdgeLeft)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With range_in.Borders(xlEdgeTop)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With range_in.Borders(xlEdgeBottom)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With range_in.Borders(xlEdgeRight)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With range_in.Borders(xlInsideVertical)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With
With range_in.Borders(xlInsideHorizontal)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
End With

Exit Sub

Exit_Sub:
    Exit Sub

End Sub
Sub PrependZerosToDRGs()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Description: Prepends 0s to DRG codes that do not have length of 3. Ex: converts "1" to "001".
'' Note: Assumes that selected range contains DRG codes.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim thisWS As Worksheet: Set thisWS = Application.ActiveSheet
Call Macro_Utilities.BackupActiveSheet(thisWS)

Dim selectedRange As Range: Set selectedRange = Application.Selection
Dim currDRGCode As String
Dim row As Long

For row = 2 To selectedRange.Rows.count
    currDRGCode = CStr(Trim(selectedRange.cells(row, 1).value))
    ' Skip blanks:
    If StrComp(currDRGCode, vbNullString) <> 0 Then
        ' Prepend zeros until has length of 3:
        Do While Len(currDRGCode) < 3
            currDRGCode = "0" & currDRGCode
        Loop
        selectedRange.cells(row, 1).value = "'" & currDRGCode
    End If
Next row

End Sub


Sub TEST()
''' Test VBA object functionality here.

' Workbooks.Open "C:\Users\brutan\Desktop\Projects\ICD-10\ICD9 to 10 Proc and DX Crosswalk FINAL 10.20.17.xlsx"
 
' Misc.RemoveBlankColumns

Macro_Utilities.CodeOptimizeSettings (False)

' Dim temp As Workbook: Set temp = Application.ActiveWorkbook

End Sub

Sub ExportAllVBA(workbook_in As String)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Description: Export all forms and macros to specified folder on computer's filesystem.
'' Best to call from the immediate window, passing workbook name string to the subroutine.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Declare constants:
    Const Module = 1
    Const ClassModule = 2
    Const Form = 3
    Const Document = 100
    Const Padding = 24
        
    Dim VBComponent As Object
    Dim count As Integer
    Dim path As String
    Dim directory As String
    Dim extension As String
    Dim fso As New FileSystemObject
    Dim currWB, targetWB As Workbook
    Dim hasWB As Boolean
    Dim errorString As String: errorString = "Error: "
    Dim hasError As Boolean
    
    On Error GoTo Exit_With_Error
    If InStr(1, workbook_in, "\") <> 0 Then
        errorString = errorString & vbCr & vbTab & "Please pass workbook name only. Do not pass workbook path."
        hasError = True
        GoTo Exit_With_Error
    Else
        ' Remove extension from passed workbook name string if provided:
        If InStr(1, workbook_in, ".x") <> 0 Then
            workbook_in = Mid(workbook_in, 1, InStr(1, workbook_in, ".x") - 1)
        ElseIf InStr(1, workbook_in, ".X") <> 0 Then
            workbook_in = Mid(workbook_in, 1, InStr(1, workbook_in, ".X") - 1)
        End If
    End If
    For Each currWB In Workbooks
        If InStr(1, currWB.Name, Trim(workbook_in)) <> 0 Then
            Set targetWB = currWB
            hasWB = True
            Exit For
        End If
    Next currWB
    
    ' If workbook not found then set error message and display on sub exit:
    If hasWB = False Then
        errorString = errorString & vbCr & vbTab & "Workbook " & Chr(34) & workbook_in & Chr(34) & " could not be found."
        hasError = True
        GoTo Exit_With_Error
    End If
    
    ' Get the output folder path:
    directory = InputBox("Please enter the macro output directory for workbook " & targetWB.Name & ":")
    
    If StrComp(directory, vbNullString) = 0 Then
        errorString = errorString & vbCr & vbTab & "Please enter the output directory."
        hasError = True
        GoTo Exit_With_Error
    End If
    
    ''''''''''''
    '''' TODO: Replace with folder for each type, put into main loop.
    If Not fso.FolderExists(directory) Then
        Call fso.CreateFolder(directory)
    End If
    
    Dim classFolder, formFolder, moduleFolder As String: classFolder = "Classes": formFolder = "Userforms": moduleFolder = "Modules"
    Dim classFolderExists, formFolderExists, moduleFolderExists As Boolean
    
    Dim currFolder As String
    ''' TODO: add folder creation and adding to folder
    For Each VBComponent In targetWB.VBProject.VBComponents
        Select Case VBComponent.Type
            Case ClassModule, Document
                extension = ".cls"
                currFolder = classFolder
            Case Form
                extension = ".frm"
                currFolder = formFolder
            Case Module
                extension = ".bas"
                currFolder = moduleFolder
            Case Else
                extension = ".txt"
                currFolder = "Misc"
        End Select
        
        '' Create folder if hasn't been created already:
        If Not fso.FolderExists(directory & "\" & currFolder) Then
            fso.CreateFolder (directory & "\" & currFolder)
        End If
                        
        On Error Resume Next
        Err.Clear
        
        path = directory & "\" & currFolder & "\" & VBComponent.Name & extension
        On Error Resume Next
        Call VBComponent.Export(path)

        If Err.Number <> 0 Then
            errorString = errorString & vbCr & vbTab & "Failed to export " & VBComponent.Name & " to " & path
            hasError = True
        Else
            count = count + 1
            Debug.Print "Exported " & Left$(VBComponent.Name & ":" & Space(Padding), Padding) & path
        End If
    Next
    
    Exit Sub
    
Exit_With_Error:
    ' Display custom error if occurred, otherwise display system error:
    If hasError = True Then
        msgBox errorString
    Else
        msgBox Err.Description
    End If
    Exit Sub
    
End Sub
Function ThompsonTau(ByVal array_in As Variant, ByVal percentile As Double, ByVal geometric As Boolean) As ArrayList
''''' IN PROGRESS:
'' TODO:
'' 1. Implement geometric method.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Description: Runs the "Thompson Tau" outlier detection method (http://www.mne.psu.edu/cimbala/me345/Lectures/Outliers.pdf),
'' returns an array of detected outliers in the past data set (array_in).
'' Note: Any non-numeric data in array_in is ignored.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''' For now, simply return empty arraylist and exit the function if geometric is true:
If geometric = True Then
    GoTo Exit_With_Error
End If

Dim startIndex, endIndex As Long: startIndex = LBound(array_in): endIndex = UBound(array_in)

'' For debugging purposes:
'If endIndex > 2 Then Stop

' Return array with null string if < 3 elements:
If endIndex <= 2 Or percentile <= 0 Or percentile >= 1 Then
    GoTo Exit_With_Error
End If

' Convert passed variant array into an arraylist (for easier manipulation). Only accept data that are numeric:
Dim index As Long
Dim dataList As New ArrayList
' Dim growthDict As New Dictionary

For index = startIndex To endIndex
    If IsNumeric(array_in(index)) = True Then
'        growthDict.Add array_in(index), 0
        dataList.Add array_in(index)
    End If
Next index

On Error GoTo Exit_With_Error
' Ensure ArrayList reference has been set:
Macro_Utilities.AddRelevantReferences

Dim outliers As New ArrayList
Dim mean, stdDev As Double
Dim currCount As Long: currCount = dataList.count
Dim dev1, dev2, testOutlierDataPoint, testDev As Double
Dim tStat As Double
Dim tau As Double
Dim numElements As Long: numElements = dataList.count
Dim hasOutliers As Boolean: hasOutliers = True

''' Search for outliers using iterative method:
Do While hasOutliers = True And currCount > 2
'' Continue to search for outliers until only two elements remain or no outliers were detected in current batch:
    hasOutliers = False
    currCount = dataList.count
    dataList.Sort
    ' Calculate mean and standard deviation depending on specification:
    If geometric = True Then
    ' Geometric mean and standard deviation:
        ' Calculate each price's growth rate between current price and next highest price:
        For index = 0 To currCount - 2
'            growthDict.Item(dataList.Item(index)) = dataList.Item(index + 1) / dataList.Item(index)
        Next index
        ' Calculate geometric mean:
        mean = 1
        For index = 0 To currCount - 2
'            mean = mean * growthDict.Item(dataList.Item(index))
        Next index
        mean = mean ^ (1 / currCount)
        stdDev = 0
        For index = 0 To currCount - 1
            stdDev = stdDev + Math.Log(dataList.Item(index) / mean) ^ 2
        Next index
        stdDev = stdDev / currCount
        stdDev = Sqr(stdDev)
    Else
    ' Arithmetic mean and standard deviation method:
        mean = Application.WorksheetFunction.Average(dataList.ToArray)
        stdDev = Application.WorksheetFunction.StDev_P(dataList.ToArray)
        If stdDev = 0 Then
            ' Exit function since all data points are the same:
            GoTo Exit_With_Error
        End If
    End If
    tStat = Abs(Application.WorksheetFunction.T_Inv_2T(percentile, currCount - 2))
    tau = tStat * (numElements - 1)
    tau = tau / (Sqr(numElements) * Sqr(numElements - 2 + tStat * tStat))
    '' Check which element, largest or smallest, has the greatest absolute deviation:
    
    dev1 = Abs(dataList.Item(0) - mean)
    dev2 = Abs(dataList.Item(currCount - 1) - mean)
    If dev1 > dev2 Then
    ' If min value had greater deviation, set as the test data point:
        testOutlierDataPoint = dataList.Item(0)
        testDev = dev1
    ElseIf dev1 < dev2 Then
    ' If max value had greatest deviation, set as the test data point:
        testOutlierDataPoint = dataList.Item(currCount - 1)
        testDev = dev2 / stdDev
    Else
    ' If deviations were the same, just set to the max data point:
        testOutlierDataPoint = dataList.Item(currCount - 1)
        testDev = dev2 / stdDev
    End If
    ' Determine if an outlier:
    If testDev > tau * stdDev Then
        hasOutliers = True
        ' Add data point to outliers list, remove data point from data set then proceed to next iteration:
        outliers.Add testOutlierDataPoint
        dataList.Remove testOutlierDataPoint
    Else
        ' Exit iterative process since no outliers:
        Exit Do
    End If
Loop

' Return detected outliers and exit function:
Set ThompsonTau = outliers
Exit Function

Exit_With_Error:
    ' If an error occurs, return an empty arraylist and exit:
    Set ThompsonTau = New ArrayList
    Exit Function
    
End Function

Public Function CreateCustomMsgBox(headerString As String, titleString As String, messageString As String, width_in As Double, height_in As Double) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Description: Create custom message box with width and height. Returns true if ok button was clicked, false if not.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''' TODO:
'' 1. Determine why okbutton is to right of cancel button (figure out positioning rules).

Dim msgBox As New CustomMsgBox
On Error GoTo Exit_With_Error

' Set message box features:
msgBox.width = width_in
msgBox.height = height_in
msgBox.TitleLabel.Caption = titleString
msgBox.MessageLabel.Caption = messageString
msgBox.Caption = headerString
' Position the buttons to the bottom and center of the form:
msgBox.OkButton.Left = msgBox.width / 2 + msgBox.OkButton.width
msgBox.OkButton.Top = msgBox.height - 60 - msgBox.OkButton.height
msgBox.CancelButton.Left = msgBox.width / 2 - 12 ' 12 is the distance between the buttons.
msgBox.CancelButton.Top = msgBox.height - 60 - msgBox.CancelButton.height

msgBox.Show

If msgBox.result = True Then
    CreateCustomMsgBox = True
    GoTo Exit_Sub
Else
    CreateCustomMsgBox = False
    GoTo Exit_Sub
End If

Exit_Sub:
    Unload msgBox
    Exit Function
Exit_With_Error:
    CreateCustomMsgBox = False
    Unload msgBox

End Function

Public Function CodeSplit(InputRange As Range) As String
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' IN PROGRESS (Functioning but add more flexibility)
'''' TODOs:
'''' 1. Add handling for ICD9s/10s (check for '.'s, then increment by .01 instead of 1)
'''' 2. Add handling for HCPCS (start with letter).
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Description: convert range of numbers (ex: 180-190) into literal numbers (180, 181,..., 190).
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Check that only one cell is selected:
If InputRange.Rows.count <> 1 And InputRange.Columns.count <> 1 Then
    GoTo Exit_With_Error
End If

Dim outputString As String: outputString = vbNullString
Dim inputCodeString, outputCodeString As String: inputCodeString = InputRange.value: outputCodeString = vbNullString
Dim currClause, startNum, endNum As String
Dim startIndex, endIndex, currNum, dashIndex As Long
Dim hasDash As Boolean

' Optimize code settings:
Call Macro_Utilities.CodeOptimizeSettings(True)

On Error GoTo Exit_With_Error

startIndex = 1

Do While startIndex < Len(inputCodeString)
    ' Get the current 'clause' (in between commas):
    If InStr(startIndex, inputCodeString, ",") <> 0 Then
        endIndex = InStr(startIndex, inputCodeString, ",")
    Else
        endIndex = Len(inputCodeString) + 1
    End If
    currClause = Trim(Mid(inputCodeString, startIndex, endIndex - startIndex))
    ' If code has dash, get the two numbers and put all numbers in between dash.
    ' Else put current clause into output string and continue.
    If InStr(1, currClause, "-") <> 0 Then
        If IsNumeric(Trim(Mid(currClause, 1, InStr(1, currClause, "-") - 1))) Then
            startNum = Trim(Mid(currClause, 1, InStr(1, currClause, "-") - 1))
        Else
            ' Go to next clause:
            GoTo NextClause
        End If
        If IsNumeric(Trim(Mid(currClause, InStr(1, currClause, "-") + 1, Len(currClause) - InStr(1, currClause, "-")))) Then
            endNum = Trim(Mid(currClause, InStr(1, currClause, "-") + 1, Len(currClause) - InStr(1, currClause, "-")))
        Else
            ' Go to next clause:
            GoTo NextClause
        End If
        ' Check if endNum greater than startNum. If false then continue to next clause:
        If CLng(startNum) >= CLng(endNum) Then
            GoTo NextClause
        End If
        ' Append all numbers in range to the outputCodeString:
        currNum = CLng(startNum)
        Do While currNum < CLng(endNum) + 1
            If StrComp(Trim(outputCodeString), vbNullString) = 0 Then
                outputCodeString = CStr(currNum)
            Else
                outputCodeString = outputCodeString & ", " & CStr(currNum)
            End If
            currNum = currNum + 1
        Loop
    Else
NextClause:
        If StrComp(Trim(outputCodeString), vbNullString) = 0 Then
            outputCodeString = currClause
        Else
            outputCodeString = outputCodeString & ", " & currClause
        End If
    End If
    startIndex = endIndex + 1
Loop
    
CodeSplit = outputCodeString

Exit Function

Exit_With_Error:
    CodeSplit = "#N/A"
    Call Macro_Utilities.CodeOptimizeSettings(False)
    Exit Function

End Function
Public Sub MultiSheetCopy()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Description:
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' - Create multiple copies of active sheet (helpful for breaking out Rate Matrix into tables).

Call Macro_Utilities.CodeOptimizeSettings(True)


Dim numCopies As Integer
Dim numCopiesString As String: numCopiesString = InputBox("Enter number of desired copies: ", "Copy This Sheet Multiple Times")

If StrComp(numCopiesString, vbNullString) = 0 Then
    Exit Sub
ElseIf IsNumeric(numCopiesString) = False Then
    msgBox "Error: " + numCopiesString + " must be a numeric value > 0"
    Exit Sub
ElseIf CInt(numCopiesString) > 20 Then
    msgBox "Error: number of copies must be less than or equal to 20"
    Exit Sub
ElseIf CInt(numCopiesString) <= 0 Then
    msgBox "Error: number of copies must be a positive number."
    Exit Sub
End If

numCopies = CInt(numCopiesString)

' Create the copies:
Dim activeWS As Worksheet: Set activeWS = Application.ActiveSheet
Dim i As Integer

For i = 1 To numCopies
    activeWS.Copy After:=ActiveWorkbook.Sheets(Application.ActiveWorkbook.Worksheets.count)
Next i

Exit_Sub:
    Call Macro_Utilities.CodeOptimizeSettings(False)
    Exit Sub
    
Exit_With_Error:
    Call Macro_Utilities.CodeOptimizeSettings(False)
    msgBox ("Error: Other macro error.")
    Exit Sub

End Sub
Function FindCostBand(ChargeFormula As Range, OriginalCost As Range, binRange As Range, MinColumn As Integer) As String
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Description:
''' *Finds the cost band for a line item in the Pharmacy Pricing File.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Parameters:
''' *ChargeFormula (Range): The Charge Formula for the line item.
''' *OriginalCost (Range): The Lookup Cost.
''' *binRange (Range): The range in the Pricing Toggle containing Cost Bands (Note: first column MUST be the Description, last column MUST be the MAX cost band.)
''' *MinColumn (Integer): the number for the MIN column in the binRange.

If ChargeFormula.Columns.count <> 1 Or ChargeFormula.Rows.count <> 1 Or OriginalCost.Columns.count <> 1 Or OriginalCost.Rows.count <> 1 Then
    msgBox "Invalid Category."
    Exit Function
ElseIf binRange.Columns.count <> 10 Then
    msgBox "Invalid Bin Range."
    Exit Function
End If

Call Macro_Utilities.CodeOptimizeSettings(True)

Dim ChargeFormulaString As String: ChargeFormulaString = UCase(ChargeFormula.cells(1, 1).value)
Dim rowIter, MaxColumn As Integer: rowIter = 1: MaxColumn = MinColumn + 1
Dim CostVal As Double: CostVal = OriginalCost.cells(1, 1).value

Do Until StrComp(binRange.cells(rowIter, 1).value, ChargeFormulaString) = 0
    rowIter = rowIter + 1
    ' If category name reaches NO COST INFO (i.e. the charge formula name does not match anything), then return "NO COST INFO" and exit function.
    If StrComp(binRange.cells(rowIter, 1).value, "NO COST INFO") = 0 Then
        FindCostBand = "NO COST INFO"
        Exit Function
    End If
Loop

' Find the Cost Band.
Do Until (CostVal >= binRange.cells(rowIter, MinColumn).value _
And CostVal <= binRange.cells(rowIter, MaxColumn).value) _
Or rowIter = binRange.Columns.count
    rowIter = rowIter + 1
    If StrComp(binRange.cells(rowIter, 1).value, "NO COST INFO") = 0 Then
        FindCostBand = "NO COST INFO"
        Exit Function
    End If
Loop

' Return the Cost Band.
FindCostBand = binRange.cells(rowIter, 2).value

Exit_Sub:
    Call Macro_Utilities.CodeOptimizeSettings(False)
    Exit Sub
Exit_With_Error:
    Call Macro_Utilities.CodeOptimizeSettings(False)
    msgBox ("Error: Other macro error.")
    Exit Sub
    
End Function
Function AddCriteriaToSortAndFilter(ByRef criteriaContents As Range, ByRef tableRange As Range, _
criteriaColumnNum As Integer) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' IN PROGRESS
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Description:
''' * Sort and filter using contents of highlighted range as criteria.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Parameters:
'''
'''
''' TODO: get starting point of tableRange,

'' Input Validation:
Dim errorString As String: errorString = "Error: "
Dim hasErrors As Boolean: hasErrors = False
If criteriaColumnNum < 0 Then
    errorString = errorString + vbCr + "Criteria column number must be positive."
    hasErrors = True
End If
If criteriaColumnNum > tableRange.Columns.count Then
    errorString = errorString + vbCr + "Criteria column falls outside of table range."
    hasErrors = True
End If
If criteriaContents.Columns.count > 1 Then
    errorString = errorString + vbCr + "Criteria content range can only be one column."
    hasErrors = True
End If

If hasErrors = True Then
    msgBox errorString
    AddCriteriaToSortAndFilter = False
    Exit Function
End If

Call Macro_Utilities.CodeOptimizeSettings(True)

' Declare array to hold contents:
Dim contentArray() As String
ReDim contentArray(criteriaContents.Rows.count - 1)
Dim i As Integer

For i = 1 To criteriaContents.Rows.count
    contentArray(i - 1) = criteriaContents.cells(i, 1).value
Next i

' Perform the auto-filtering:
tableRange.AutoFilter Field:=criteriaColumnNum, Criteria1:=contentArray, Operator:=xlFilterValues

Call Macro_Utilities.CodeOptimizeSettings(False)

AddCriteriaToSortAndFilter = True

End Function
Sub AddCriteriaToSortAndFilter_SUBROUTINE(ByRef criteriaContents As Range, ByRef tableRange As Range, _
criteriaColumnNum As Integer)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' IN PROGRESS
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Description:
''' * Sort and filter using contents of highlighted range as criteria.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Parameters:
'''
'''
''' TODO: get starting point of tableRange,

'' Input Validation:
Dim errorString As String: errorString = "Error: "
Dim hasErrors As Boolean: hasErrors = False
If criteriaColumnNum < 0 Then
    errorString = errorString + vbCr + "Criteria column number must be positive."
    hasErrors = True
End If
If criteriaColumnNum > tableRange.Columns.count Then
    errorString = errorString + vbCr + "Criteria column falls outside of table range."
    hasErrors = True
End If
If criteriaContents.Columns.count > 1 Then
    errorString = errorString + vbCr + "Criteria content range can only be one column."
    hasErrors = True
End If

If hasErrors = True Then
    msgBox errorString
    AddCriteriaToSortAndFilter = False
    Exit Sub
End If

Call Macro_Utilities.CodeOptimizeSettings(True)

' Declare array to hold contents:
Dim contentArray() As String
ReDim contentArray(criteriaContents.Rows.count - 1)
Dim i As Integer

For i = 1 To criteriaContents.Rows.count
    contentArray(i - 1) = criteriaContents.cells(i, 1).value
Next i

' Perform the auto-filtering:
tableRange.AutoFilter Field:=criteriaColumnNum, Criteria1:=contentArray, Operator:=xlFilterValues

Call Macro_Utilities.CodeOptimizeSettings(False)

AddCriteriaToSortAndFilter = True

End Sub
Sub NALineBreakout_Deprecated()

Call Macro_Utilities.BackupActiveSheet(Application.ActiveSheet)

Dim row As Long, lastRow As Long, s

Call Macro_Utilities.CodeOptimizeSettings(True)

lastRow = cells(Rows.count, 1).End(xlUp).row
For row = lastRow To 2 Step -1
    If InStr(cells(row, 4), ", ") Then
        s = Split(cells(row, 4), ", ")
        Rows(row + 1).Resize(UBound(s)).Insert
        cells(row + 1, 1).Resize(UBound(s), 3).value = cells(row, 1).Resize(, 3).value
        cells(row + 1, 5).Resize(UBound(s)).value = cells(row, 5).value
        cells(row + 1, 6).Resize(UBound(s)).value = cells(row, 6).value
        cells(row + 1, 7).Resize(UBound(s)).value = cells(row, 7).value
        cells(row + 1, 8).Resize(UBound(s)).value = cells(row, 8).value
        cells(row + 1, 9).Resize(UBound(s)).value = cells(row, 9).value
        cells(row + 1, 10).Resize(UBound(s)).value = cells(row, 10).value
        cells(row + 1, 11).Resize(UBound(s)).value = cells(row, 11).value
        cells(row + 1, 12).Resize(UBound(s)).value = cells(row, 12).value
        cells(row + 1, 13).Resize(UBound(s)).value = cells(row, 13).value
        cells(row + 1, 14).Resize(UBound(s)).value = cells(row, 14).value
        cells(row + 1, 15).Resize(UBound(s)).value = cells(row, 15).value
        cells(row + 1, 16).Resize(UBound(s)).value = cells(row, 16).value
        cells(row + 1, 17).Resize(UBound(s)).value = cells(row, 17).value
        cells(row + 1, 18).Resize(UBound(s)).value = cells(row, 18).value
        cells(row + 1, 19).Resize(UBound(s)).value = cells(row, 19).value
        cells(row + 1, 20).Resize(UBound(s)).value = cells(row, 20).value
        cells(row + 1, 21).Resize(UBound(s)).value = cells(row, 21).value
        cells(row + 1, 22).Resize(UBound(s)).value = cells(row, 22).value
        cells(row + 1, 23).Resize(UBound(s)).value = cells(row, 23).value
        cells(row + 1, 24).Resize(UBound(s)).value = cells(row, 24).value
        cells(row + 1, 25).Resize(UBound(s)).value = cells(row, 25).value
        cells(row + 1, 26).Resize(UBound(s)).value = cells(row, 26).value
        cells(row + 1, 27).Resize(UBound(s)).value = cells(row, 27).value
        cells(row + 1, 28).Resize(UBound(s)).value = cells(row, 28).value
        cells(row, 4).Resize(UBound(s) + 1).value = Application.Transpose(s)
    End If
Next row

Call Macro_Utilities.CodeOptimizeSettings(False)

End Sub
Function CountColorIf(rSample As Range, rArea As Range) As Long
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Description: Counts the number of times a cell with color matching rSample's color appears in selected range (rArea).
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Application.Volatile
    Dim rAreaCell As Range
    Dim lMatchColor As Long
    Dim lCounter As Long

    lMatchColor = rSample.Interior.Color
    For Each rAreaCell In rArea
        If rAreaCell.Interior.Color = lMatchColor Then
            lCounter = lCounter + 1
        End If
    Next rAreaCell
    CountColorIf = lCounter
End Function
Sub RemoveNewLines()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Description: remove all newline characters (CHR(10)) commonly used in column headers in pricing files/rate matrix/etc.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim i, j As Integer
    Dim selectedRange As Range: Set selectedRange = Selection
    For i = 1 To selectedRange.Columns.count
        For j = 1 To selectedRange.Rows.count
            selectedRange.cells(j, i).value = Replace(selectedRange.cells(j, i).value, Chr(10), " ")
        Next j
    Next i
End Sub
Sub RemoveNewLines_wParams(Optional passedRange As Range)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Description: remove all newline characters (CHR(10)) commonly used in column headers in pricing files/rate matrix/etc.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim i, j As Integer
    Dim selectedRange As Range
    If passedRange Is Nothing Then
        Set selectedRange = Selection
    Else
        Set selectedRange = passedRange
    End If
    For i = 1 To selectedRange.Columns.count
        For j = 1 To selectedRange.Rows.count
            selectedRange.cells(j, i).value = Replace(selectedRange.cells(j, i).value, Chr(10), " ")
        Next j
    Next i
End Sub
Function GetDRGs(codeString As String) As String
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Description: Get DRGs associated with code (very specific function).
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim sIndex, fIndex As Integer: sIndex = -99: fIndex = -99
Dim char As Integer: char = 1

Dim tempChar As String

Do While char <> Len(codeString)
    tempChar = Mid(codeString, char, 1)
    If IsNumeric(Mid(codeString, char, 1)) = True And sIndex = -99 Then
        sIndex = char
    ElseIf char <> 1 Then
        If StrComp(Trim(Mid(codeString, char, 1)), vbNullString) = 0 And _
        IsNumeric(Mid(codeString, char - 1, 1)) = True Then
            fIndex = char
            Exit Do
        End If
    End If
    char = char + 1
Loop

If fIndex = -99 Then
    fIndex = InStr(1, codeString, ",")
End If

GetDRGs = Mid(codeString, sIndex, fIndex - sIndex)

End Function
Public Function PriorityOutput(selectedRange As Range) As String
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Description: Outputs the first non-blank cell, from left to right.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' NOTE: the selected range can only be one row.

On Error GoTo Exit_With_Error

If selectedRange.Rows.count > 1 Then
    GoTo Exit_With_Error
End If

Dim col As Integer

For col = 1 To selectedRange.Columns.count
    If StrComp(Trim(selectedRange.cells(1, col).value), vbNullString) <> 0 Then
        PriorityOutput = selectedRange.cells(1, col).value
        Exit Function
    End If
Next col

PriorityOutput = vbNullString
Exit Function

Exit_With_Error:
    PriorityOutput = "#N/A"
    Exit Function

End Function
Sub OutputGroupings()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Description: Group all codes by price in selected list (for fee schedules/groupers).
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''\
Dim i, j, outputRow As Integer
Dim selectedRange As Range: Set selectedRange = Application.Selection
Dim OutputRange As Range: Set OutputRange = Range("E1")
Dim currPrice, accumCodes As String

outputRow = 1
i = 1

Do While i <> selectedRange.Rows.count + 1
    currPrice = selectedRange.cells(i, 2).value
    accumCodes = selectedRange.cells(i, 1).value
    i = i + 1
    Do While StrComp(currPrice, selectedRange.cells(i, 2).value) = 0 And _
        i <> selectedRange.Rows.count + 1
        accumCodes = accumCodes & ", " & selectedRange.cells(i, 1).value
        i = i + 1
    Loop
    OutputRange.cells(outputRow, 2).value = currPrice
    OutputRange.cells(outputRow, 1).value = accumCodes
    outputRow = outputRow + 1
Loop

End Sub
Function CountIf_Array(matchVal As Variant, ByRef arr As Variant) As Long
''''''''''''''''''''''''''''''''''''''
'' IN PROGRESS:
''''''''''''''''''''''''''''''''''''''
'' TODO:
'' 1. Add more valid types.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Description:
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'' Description: CountIf on an array.
    
    If UBound(arr) - LBound(arr) = 0 Then
        CountIf_Array = 0
        Exit Function
    End If
    
    On Error GoTo Skip
    Dim i, count As Long
    For i = LBound(arr) To UBound(arr)
        If VarType(arr) = vbString Then
            If StrComp(matchVal, arr(i)) = 0 Then
                count = count + 1
            End If
        Else
            If matchVal = arr(i) Then
                count = count + 1
            End If
        End If
Skip:
    Next i
    
    CountIf_Array = count

End Function
Sub RemoveBlankColumns()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Description:
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim selectedRange As Range: Set selectedRange = Application.Selection
Dim ws As Worksheet: Set ws = Application.ActiveSheet
Dim col As Integer


' Ask if want to back up sheet:
Macro_Utilities.BackupActiveSheet ws
Macro_Utilities.CodeOptimizeSettings True

' On Error GoTo Exit_With_Error

Dim totColumns As Integer: totColumns = selectedRange.Columns.count
Dim removedColCount As Integer: removedColCount = 0
col = 1

Do While col < totColumns - removedColCount + 1
    ' If value is blank then delete column.
    If StrComp(Trim(selectedRange.cells(1, col).value), vbNullString) = 0 Then
        Range(selectedRange.cells(1, col).Address(RowAbsolute:=False, ColumnAbsolute:=False)).Delete
        removedColCount = removedColCount + 1
    End If
    col = col + 1
Loop

Exit_Sub:
    Macro_Utilities.CodeOptimizeSettings False
    Exit Sub

Exit_With_Error:
    Macro_Utilities.CodeOptimizeSettings True
    msgBox "Other Macro Error."
    Exit Sub

End Sub

Sub WorkbookOpenClose(workbookPath As String, Open_Workbook As Boolean, ByRef success As Boolean)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Description:
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Opens or closes workbook with passed name. Workbook reference is stored in storedWB global variable, for use in
'' functions that need to open workbooks (functions cannot open workbooks in VBA).

'' Set the global storedWB workbook reference if seek to open, close otherwise:
Dim currWB As Workbook
Dim wbName As String: wbName = workbookPath
Dim hasChanged As Boolean
'' Get the workbook name from passed Open_Workbook (contains full path name);
Do While InStr(1, wbName, "\") <> 0
    ' Remove all backslashes until get final (assumed) workbook name:
    wbName = Mid(wbName, InStr(1, wbName, "\") + 1, Len(wbName))
    hasChanged = True
Loop

If hasChanged = False Then
    GoTo Exit_With_Error
End If

' Remove extension:
If InStr(1, wbName, ".x") <> 0 Then
    wbName = Mid(wbName, 1, InStr(1, wbName, ".x") - 1)
ElseIf InStr(1, wbName, ".X") <> 0 Then
    wbName = Mid(wbName, 1, InStr(1, wbName, ".X") - 1)
End If

On Error GoTo Exit_With_Error
hasChanged = False
If Open_Workbook = True Then
    ' If workbook has not been de-referenced then close and null out:
    If Not (storedWB Is Nothing) Then
        storedWB.Close
        Set storedWB = Nothing
    End If
    Workbooks.Open (workbookPath)
    ' Assign opened workbook to global reference:
    For Each currWB In Workbooks
        If InStr(1, currWB.Name, wbName) <> 0 Then
            Set storedWB = currWB
            hasChanged = True
            Exit For
        End If
    Next currWB
Else
    ' Close workbook and null out:
    If Not (storedWB Is Nothing) Then
        storedWB.Close
        Set storedWB = Nothing
        hasChanged = True
    End If
End If

' If failed to find workbook, set success to False and exit:
If hasChanged = False Then
    GoTo Exit_With_Error
End If

success = True
Exit Sub

Exit_With_Error:
    success = False
    Exit Sub
    
End Sub

Sub ManuallyGrabICD9sAnd10s()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Description:
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''

Dim selectedRange As Range: Set selectedRange = Application.Selection
' Assume 1st column contains ICD9s, second contains ICD10s:
Dim ICD9outputString, ICD10outputString As String
Dim row As Long: row = 2
Dim outputRow As Long: outputRow = 1

Do While row <= selectedRange.Rows.count
    ICD10outputString = vbNullString
    ICD9outputString = vbNullString
'    If outputRow = 3 Then Stop
    Do While Len(ICD10outputString) < (29900 - 10762) And row <= selectedRange.Rows.count
        If Len(ICD10outputString) = 0 Then
            ICD10outputString = Chr(34) & CStr(Trim(selectedRange.cells(row, 2).value)) & Chr(34)
            ICD9outputString = Chr(34) & CStr(Trim(selectedRange.cells(row, 1).value)) & Chr(34)
        Else
            ICD10outputString = ICD10outputString & "," & Chr(34) & CStr(Trim(selectedRange.cells(row, 2).value)) & Chr(34)
            ICD9outputString = ICD9outputString & "," & CStr(Trim(selectedRange.cells(row, 1).value)) & Chr(34)
        End If
        row = row + 1
    Loop
    '' Output each corresponding string into 3rd and 4th column:
    selectedRange.cells(outputRow, 3).value = CStr(ICD9outputString)
    selectedRange.cells(outputRow, 4).value = CStr(ICD10outputString)
    outputRow = outputRow + 1
Loop

End Sub

Function IsError(input_val As String) As Boolean

On Error GoTo Exit_With_Error
If StrComp(input_val, "#VALUE!") = 0 Or StrComp(input_val, "#N/A") = 0 Or StrComp(input_val, "#DIV/0!") = 0 _
    Or StrComp(input_val, "#REF!") = 0 Or StrComp(input_val, "#NULL!") = 0 Or StrComp(input_val, "#NUM!") = 0 Then
        IsError = True
Else
    IsError = False
End If

Exit Function

Exit_With_Error:
    IsError = False

End Function

Function BlankOnError(input_val As String) As String

On Error GoTo Exit_With_Error
If StrComp(input_val, "#VALUE!") = 0 Or StrComp(input_val, "#N/A") = 0 Or StrComp(input_val, "#DIV/0!") = 0 _
    Or StrComp(input_val, "#REF!") = 0 Or StrComp(input_val, "#NULL!") = 0 Or StrComp(input_val, "#NUM!") = 0 Then
        BlankOnError = vbNullString
Else
    BlankOnError = input_val
End If

Exit Function

Exit_With_Error:
    BlankOnError = False

End Function


End Function
