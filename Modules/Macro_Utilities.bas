Attribute VB_Name = "Macro_Utilities"
Option Explicit
Function AddRelevantReferences() As String
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Description: Add all relevant references to this workbook.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Note: Returns errorMessage if any references fail to add.

On Error GoTo Exit_With_Error
Dim ref As Object
Dim currElem As Integer
Dim addRefs, addRefsPath As Variant
Dim hasRef As Boolean

addRefs = Array("stdole", "RefEdit", "VBScript_RegExp_55", "Scripting", "mscorlib")
addRefsPath = Array("C:\Windows\System32\stdole2.tlb", "C:\Program Files\Microsoft Office\Root\Office16\REFEDIT.DLL", _
    "C:\Windows\System32\vbscript.dll", "C:\Windows\System32\scrrun.dll", "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\mscorlib.dll")
Dim errorMessage As String: errorMessage = "Error: Some references failed to add. " & vbCr & vbTab _
            & "Did you enable VBA access to the project object model?" & vbCr & vbTab _
            & "(https://www.ibm.com/support/knowledgecenter/en/SSD29G_2.0.0/com.ibm.swg.ba.cognos.ug_cxr.2.0.0.doc/t_ug_cxr_enable_developer_macro_settings.html)"

For currElem = 0 To UBound(addRefs)
    hasRef = False
    For Each ref In ThisWorkbook.VBProject.References
        If StrComp(Trim(addRefs(currElem)), Trim(ref.Name)) = 0 Then
            hasRef = True
            Exit For
        End If
    Next ref
    If hasRef = False Then
        ' Add the reference:
        ThisWorkbook.VBProject.References.AddFromFile addRefsPath(currElem)
    End If
Next currElem

AddRelevantReferences = vbNullString
Exit Function

Exit_With_Error:
    AddRelevantReferences = errorMessage
    Exit Function

End Function

Function BackupActiveSheet(ws As Worksheet)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Description: Makes a backup of the active sheet. Should be done before every significant macro since cannot undo.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    On Error GoTo On_Error
    BackupPrompt.Show
    If BackupPrompt.result = True Then
        Dim newWsName As String: newWsName = ws.Name + "_Backup"
        Dim wsNameExists As Boolean
        Dim searchSheet As Worksheet
        Dim iterNum As Integer: iterNum = 1
        Dim tempNum As Integer
        ' Check if sheet name already exists:
        Do
            wsNameExists = False
            For Each searchSheet In ActiveWorkbook.Worksheets
                If StrComp(searchSheet.Name, newWsName) = 0 Then
                    tempNum = InStr(1, newWsName, "(")
                    If InStr(1, newWsName, "(") = 0 Then
                    ' First time appending number:
                        If Len(newWsName) > 30 Then
                            newWsName = Mid(newWsName, 1, 30 - Len("_Backup (" & CStr(iterNum) & ")") - 1)
                            newWsName = newWsName & "_Backup (" & CStr(iterNum) & ")"
                        End If
                        newWsName = newWsName & " (" & CStr(iterNum) & ")"
                        iterNum = iterNum + 1
                        wsNameExists = True
                    Else
                        newWsName = Mid(newWsName, 1, InStr(1, newWsName, " (") - 1)
                        newWsName = newWsName & " (" & CStr(iterNum) & ")"
                        If Len(newWsName) > 30 Then
                            newWsName = Mid(newWsName, 1, 30 - Len("_Backup (" & CStr(iterNum) & ")") - 1)
                            newWsName = newWsName & "_Backup (" & CStr(iterNum) & ")"
                        End If
                        iterNum = iterNum + 1
                        wsNameExists = True
                    End If
                End If
            Next searchSheet
        Loop While wsNameExists = True
        ' Copy the provided worksheet then denote as backup:
        ws.Copy Before:=Worksheets(ws.Name)
        Worksheets(ws.index - 1).Name = newWsName
        ws.Activate
    End If
    ' Save the current workbook if specified:
    If BackupPrompt.SaveWorkbookOption = True Then
        ActiveWorkbook.Save
    End If
    ' Clear out the prompt:
    Unload BackupPrompt
    Exit Function
    
On_Error:
    Unload BackupPrompt
    msgBox "Error: BackupActiveSheet macro error."
    Exit Function
    
End Function

Function ColumnIntersect(ByVal mainRange As Range, ByVal colRange As Range) As Integer
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Description: Determines if colRange is on the mainRange. Function returns the column number for colRange relative
'' to the mainRange's address or returns 0 if colRange does not intersect with the mainRange.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Assume first column is the check column if colRange has more than one column
Set colRange = colRange.Resize(1, 1)
Set mainRange = mainRange.Resize(1, mainRange.Columns.count)

Dim colAddress As String: colAddress = colRange.Address(ReferenceStyle:=xlR1C1)
Dim mainAddress As String: mainAddress = mainRange.Address(ReferenceStyle:=xlR1C1)
Dim mainColStart, mainColEnd As String

colAddress = Mid(colAddress, InStr(1, colAddress, "C") + 1, Len(colAddress))
If mainRange.Columns.count > 1 Then
    mainColStart = Mid(mainAddress, InStr(1, mainAddress, "C") + 1, InStr(1, mainAddress, ":") - InStr(1, mainAddress, "C") - 1)
    mainAddress = Mid(mainAddress, InStr(1, mainAddress, ":") + 1, Len(mainAddress))
    mainColEnd = Mid(mainAddress, InStr(1, mainAddress, "C") + 1, Len(mainAddress))
Else
    ' If mainrange only has one column, just grab single column number:
    mainColStart = Mid(mainAddress, InStr(1, mainAddress, "C") + 1, Len(mainAddress))
End If

If mainRange.Columns.count > 1 Then
    ' If more than one column, check if column number is in between, inclusive, of the mainRange's column numbers:
    If CInt(colAddress) >= mainColStart And CInt(colAddress) <= mainColEnd Then
        ColumnIntersect = CInt(colAddress) - mainColStart + 1
        Exit Function
    Else
        ColumnIntersect = 0
        Exit Function
    End If
Else
    ' If only one column, just check if column number is the same as the mainRange column number.
    If CInt(colAddress) = CInt(mainColStart) Then
        ColumnIntersect = 1
        Exit Function
    Else
        ColumnIntersect = 0
        Exit Function
    End If
End If

End Function

Function CheckIfRangeBlank(ByVal InputRange As Range) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Description: returns true if range is entirely blank.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim i, j As LongLong
Dim currVal As String
Dim resetOpt As Boolean

If Macro_Utilities.OptimizationIsOn() = False Then
    Call Macro_Utilities.CodeOptimizeSettings(True)
    resetOpt = True
End If

For i = 1 To InputRange.Rows.count
    For j = 1 To InputRange.Columns.count
        currVal = CStr(InputRange.cells(i, j).value)
        If StrComp(Trim(currVal), vbNullString) <> 0 Then
            CheckIfRangeBlank = False
            GoTo Exit_Function
        End If
    Next j
Next i

CheckIfRangeBlank = True

GoTo Exit_Function

Exit_Function:
    If resetOpt = True Then
        Call Macro_Utilities.CodeOptimizeSettings(False)
    End If

End Function

Sub CodeOptimizeSettings(turnOff As Boolean)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Description: Turns off all unnecessary settings on worksheet/workbook to speed up code.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If turnOff = False Then
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    Application.DisplayStatusBar = True
    Application.EnableEvents = True
    Application.ActiveWorkbook.ActiveSheet.DisplayPageBreaks = True
    Application.ScreenUpdating = True
Else
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    Application.DisplayStatusBar = False
    Application.EnableEvents = False
    Application.ActiveWorkbook.ActiveSheet.DisplayPageBreaks = False
    Application.ScreenUpdating = False
End If

End Sub

Function RangeCheck(rangeAddress As String) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''
'' IN PROGRESS:
''''''''''''''''''''''''''''''''''''''''''''''''''
'' TODO: Add ability to check if range is on different worksheet (the same worksheet as the range).
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Description: returns false if passed string is not a valid range address.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

On Error GoTo Exit_False
Dim isR1C1 As Boolean
Dim filter1, filter2 As String
Dim output As Boolean

' Remove potential unnecessary spaces:
rangeAddress = Trim(rangeAddress)

' Check if absolute address, filter out worksheet name and exclamation if necessary:
If InStr(1, rangeAddress, "!") <> 0 Then
    rangeAddress = Mid(rangeAddress, InStr(1, rangeAddress, "!") + 1, Len(rangeAddress))
End If

'' Check if R1C1:
If InStr(1, rangeAddress, "R") <> 0 Then
    filter1 = Mid(rangeAddress, InStr(1, rangeAddress, "R") + 1, Len(rangeAddress))
    If InStr(1, filter1, "C") <> 0 Then
        '' Check if integer exists between the R and the C:
        filter1 = Mid(rangeAddress, InStr(1, rangeAddress, "R") + 1, InStr(1, rangeAddress, "C") - InStr(1, rangeAddress, "R") - 1)
        filter2 = Mid(rangeAddress, InStr(1, rangeAddress, "C") + 1, Len(rangeAddress))
        If IsNumeric(filter1) = True And InStr(1, filter1, ".") = 0 And IsNumeric(filter2) = True And InStr(1, filter2, ".") = 0 Then
            output = True
            GoTo Exit_Function
        Else
            isR1C1 = False
        End If
    End If
End If

'' Check if A1A1 style if necessary:
Dim currChar, columnFull, rowNumFull As String
Dim index, strLen As Long

If isR1C1 = False Then
    ' Remove dollar symbols used in range locking if necessary:
    rangeAddress = Replace(rangeAddress, "$", "")
    strLen = Len(rangeAddress)
    index = 1
    Do
    ' Get full first column letters (max of two):
        columnFull = vbNullString
        currChar = Mid(rangeAddress, index, 1)
        Do While Macro_Utilities.IsChar(currChar) = True And index < strLen + 1
            columnFull = columnFull & currChar
            index = index + 1
            currChar = Mid(rangeAddress, index, 1)
        Loop
        ' If reached end of address then return false:
        If index >= strLen + 1 Then
            output = False
            GoTo Exit_Function
        ElseIf Len(columnFull) > 3 Then
            ' If more than three letters in the column then cannot represent a column, thus return false:
            output = False
            GoTo Exit_Function
        '''' TODO: determine maximum possible letter in column name.
        End If
        ' Get first row number:
        rowNumFull = vbNullString
        Do While Macro_Utilities.IsChar(currChar) = False And StrComp(currChar, ":") <> 0 And index < strLen + 1
            rowNumFull = rowNumFull & currChar
            index = index + 1
            currChar = Mid(rangeAddress, index, 1)
        Loop
        ' Check the row number:
        If IsNumeric(rowNumFull) = False Then
            ' If not a number then return false:
            output = False
            GoTo Exit_Function
        ElseIf InStr(1, rowNumFull, ".") <> 0 Then
            ' If is a number but contains a decimal then return false:
            output = False
            GoTo Exit_Function
        ElseIf CLng(rowNumFull) > 65536 Then
            ' If integer is greater than 65536, the maximum row number, then return false:
            output = False
            GoTo Exit_Function
        Else
            ' If reached the end of the string then return true:
            If index >= strLen + 1 Then
                output = True
                GoTo Exit_Function
            Else
                index = index + 1
            End If
        End If
    ' Repeat loop if address contains a ":":
    Loop While StrComp(currChar, ":") = 0
    ' If reaches this point then assume string represents an address:
    output = True
End If
    
Exit_Function:
    RangeCheck = output
    Exit Function

Exit_False:
    RangeCheck = False
    Exit Function

End Function
Function FileExists(fileName As String, localFolderPath As String) As Boolean
'''''''''''' IN PROGRESS:

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Description: Checks if file with passed name exists in the local folder.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim exists As Boolean

If FilenameHasInvalidCharacters(fileName) = True Then
    GoTo Exit_Func
End If



FileExists:
    FileExists = exists


End Function

Sub WarningMessage(message As String, seconds As Integer)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Description: Generates warning message passed by caller that remains for passed amount of time (in seconds).
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
CreateObject("WScript.Shell").PopUp message, time, "FYI: "

End Sub
Function RemoveNewLineChars(inputString As String) As String
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Description: Returns string without newline type characters.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo Exit_With_Error
If InStr(1, inputString, vbNewLine) <> 0 Then
    inputString = Replace(inputString, vbNewLine, "")
End If
If InStr(1, inputString, vbCrLf) <> 0 Then
    inputString = Replace(inputString, vbCrLf, "")
End If
If InStr(1, inputString, vbCr) <> 0 Then
    inputString = Replace(inputString, vbCr, "")
End If
If InStr(1, inputString, vbLf) <> 0 Then
    inputString = Replace(inputString, vbLf, "")
End If
If InStr(1, inputString, vbLf) <> 0 Then
    inputString = Replace(inputString, vbLf, "")
End If

RemoveNewLineChars = inputString

Exit Function

Exit_With_Error:
    RemoveNewLineChars = vbNullString
    Exit Function

End Function
Function FilenameHasInvalidCharacters(fileName As String) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Description: Macro determines if passed potential Windows file name has invalid characters.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Macro_Utilities.CodeOptimizeSettings (True)

On Error GoTo Exit_With_Error

Dim invalidChars As Variant
invalidChars = Array(">", "<", ":", Chr(34), ":", "/", "\", "|", "?", "*")
Dim char, numChars As Long: numChars = Len(fileName)
Dim invalidCharIndex, numInvalidChars As Integer: numInvalidChars = UBound(invalidChars)
For char = 1 To numChars
    For invalidCharIndex = 0 To numInvalidChars
        If StrComp(invalidChars(invalidCharIndex), Mid(fileName, char, 1)) = 0 Then
            Macro_Utilities.CodeOptimizeSettings (False)
            FilenameHasInvalidCharacters = True
            Exit Function
        End If
    Next invalidCharIndex
Next char

Macro_Utilities.CodeOptimizeSettings (False)

FilenameHasInvalidCharacters = False
Exit Function

Exit_With_Error:
    Macro_Utilities.CodeOptimizeSettings (False)
    FilenameHasInvalidCharacters = False
    Exit Function

End Function

Function OptimizationIsOn() As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Description: Returns true if the optimization settings are running. Useful preventing turning optimization settings back on
' in functions/subs within functions/subs that already have optimization settings on.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If Application.Calculation = xlCalculationManual And Application.EnableEvents = False And _
    Application.ActiveWorkbook.ActiveSheet.DisplayPageBreaks = False And Application.DisplayStatusBar = False And Application.ScreenUpdating = False Then
    OptimizationIsOn = True
Else
    OptimizationIsOn = False
End If

End Function

Sub DisplayAlert(message As String, title As String, time As Long)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Description: Display message with title that auto-closes after certain amount of time.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If time < 0 Then
    Exit Sub
End If

CreateObject("WScript.Shell").PopUp message, time, title

End Sub

Function WorkbookNameIsOpen(workbookName As String) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Description: Returns true if workbook with passed name is open. Should run if trying to save a workbook since cannot save
' a workbook with same name as another open workbook.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim currWB As Workbook
Dim currWBName As String
' Potential file extensions include
On Error GoTo Exit_With_Error

For Each currWB In Workbooks
    ' Remove file type extension:
    If InStr(1, currWB.Name, ".x") <> 0 Then
        currWBName = Mid(currWB.Name, 1, InStr(1, currWB.Name, ".x") - 1)
    ElseIf InStr(1, currWB.Name, ".X") <> 0 Then
        currWBName = Mid(currWB.Name, 1, InStr(1, currWB.Name, ".X") - 1)
    End If
    If StrComp(currWBName, workbookName) = 0 Then
        WorkbookNameIsOpen = True
        Exit Function
    End If
Next currWB

WorkbookNameIsOpen = False

Exit Function

Exit_With_Error:
    WorkbookNameIsOpen = False
    Exit Function
    
End Function

Function WorkbookExistsLocally(workbookNamePath As String) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Description: Returns true if workbook name exists at path. workbookNamePath is of form ex: 'C:\Program Files\...\File_Name.xlsm'
' or similar.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

On Error GoTo Exit_With_Error
If StrComp(Dir(workbookNamePath), vbNullString) = 0 Then
    WorkbookExistsLocally = False
Else
    WorkbookExistsLocally = True
End If

Exit Function

Exit_With_Error:
    WorkbookExistsLocally = False

End Function

Function WorksheetExists(wbName As String, wsName As String) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Description: Returns true if worksheet with name exists in passed workbook. Returns false if workbook name is invalid or
' worksheet name does not exist in valid workbook.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo Exit_False
Dim thisWB As Workbook: Set thisWB = Workbooks(wbName)
Dim ws As Worksheet: Set ws = thisWB.Worksheets(wsName)

WorksheetExists = True

Exit Function

Exit_False:
    WorksheetExists = False
    
End Function

Function R1C1converter(Address As String, isR1C1 As Boolean) As String
'''''''''' IN PROGRESS:
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Description: Converts R1C1 address into an A1 address and vice versa.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo Exit_With_Error
Dim output As Variant
Dim functionInput As String: functionInput = "=SUM(" & Address & ")"

If isR1C1 = True Then
    'Convert A1 to R1C1:
    output = Application.ConvertFormula(functionInput, xlA1, xlR1C1, RelativeTo:=Range("D1"))
Else
    'Convert R1C1 to A1:
    output = Application.ConvertFormula(functionInput, xlR1C1, xlA1, RelativeTo:=Range("D1"))
End If

output = Mid(functionInput, InStr(1, functionInput, "=SUM(") + 1, InStr(1, functionInput, ")") - 1)
R1C1converter = CStr(output)

Exit Function

Exit_With_Error:
    R1C1converter = "#N/A"

End Function

Function IsChar(ByVal value As String) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Description: Returns true if character (or first character of passed string if len > 1) is
'' an alphabetical character.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo Exit_With_Error
If Len(value) = 0 Then
    GoTo Exit_With_Error
End If
value = Mid(value, 1, 1)

If (Asc(value) >= 65 And Asc(value) <= 90) Or (Asc(value) >= 97 And Asc(value) <= 122) Then
    IsChar = True
    Exit Function
Else
    IsChar = False
    Exit Function
End If

Exit_With_Error:
    IsChar = False
    Exit Function

End Function

Function EqualizeRangeRows(targetAddress As String, convertAddress As String) As String
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Description: Converts convertAddress' rows into targetAddress' rows while maintaining its columns.
'' Note: If either input is blank or is not an address then returns a blank string.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Dependencies: Need the RangeCheck function (present in Macro_Utilities module).
If StrComp(Trim(targetAddress), vbNullString) = 0 Or StrComp(Trim(convertAddress), vbNullString) = 0 Or _
    RangeCheck(targetAddress) = False Or RangeCheck(convertAddress) = False Then
    GoTo Exit_With_Error
End If

targetAddress = Trim(targetAddress)
convertAddress = Trim(convertAddress)

Dim targetAddressSheet, convertAddressSheet As String
'' Remove $ symbols if necessary:
targetAddress = Replace(targetAddress, "$", "")
convertAddress = Replace(convertAddress, "$", "")
'' Remove and store sheet names from both inputs if necessary:
If InStr(1, targetAddress, "!") <> 0 Then
    targetAddressSheet = Mid(targetAddress, 1, InStr(1, targetAddress, "!"))
    targetAddress = Mid(targetAddress, InStr(1, targetAddress, "!") + 1, Len(targetAddress))
End If
If InStr(1, convertAddress, "!") <> 0 Then
    convertAddressSheet = Mid(convertAddress, 1, InStr(1, convertAddress, "!"))
    convertAddress = Mid(convertAddress, InStr(1, convertAddress, "!") + 1, Len(convertAddress))
End If

'' Get targetAddress' rows:
Dim firstRow, lastRow As String
Dim index As Long
'' Extract first cell's row number:
For index = 1 To IIf(InStr(1, targetAddress, ":") <> 0, InStr(1, targetAddress, ":") - 1, Len(targetAddress))
    If IsNumeric(Mid(targetAddress, index, 1)) = True Then
        firstRow = firstRow & Mid(targetAddress, index, 1)
    End If
Next index
'' Check if address spans multiple cells:
If InStr(1, targetAddress, ":") <> 0 Then
    ' Address spans multiple cells. Extract row number of second cell:
    For index = InStr(1, targetAddress, ":") + 1 To Len(targetAddress)
        If IsNumeric(Mid(targetAddress, index, 1)) = True Then
            lastRow = lastRow & Mid(targetAddress, index, 1)
        End If
    Next index
End If
'' Get convertAddress' columns:
Dim column1, column2 As String
' Get first column:
For index = 1 To IIf(InStr(1, convertAddress, ":") <> 0, InStr(1, convertAddress, ":") - 1, Len(convertAddress))
    If IsNumeric(Mid(convertAddress, index, 1)) = False Then
        column1 = column1 & Mid(convertAddress, index, 1)
    End If
Next index
' Get second column if necessary:
If InStr(1, convertAddress, ":") <> 0 Then
    For index = InStr(1, convertAddress, ":") + 1 To Len(convertAddress)
        If IsNumeric(Mid(convertAddress, index, 1)) = False Then
            column2 = column2 & Mid(convertAddress, index, 1)
        End If
    Next index
Else
    column2 = column1
End If

'' Reconstitute the convertAddress and output:
convertAddress = convertAddressSheet & "$" & column1 & "$" & firstRow
If StrComp(lastRow, vbNullString) <> 0 Then
    convertAddress = convertAddress & ":" & "$" & column2 & "$" & lastRow
End If

EqualizeRangeRows = convertAddress

Exit Function

Exit_With_Error:
    EqualizeRangeRows = vbNullString
    
End Function

