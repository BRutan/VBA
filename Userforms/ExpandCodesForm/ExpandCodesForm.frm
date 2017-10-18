VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ExpandCodesForm 
   Caption         =   "Expand Codes Macro"
   ClientHeight    =   3576
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4296
   OleObjectBlob   =   "ExpandCodesForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ExpandCodesForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub ExitButton_Click()
    Me.Hide
    Unload Me
End Sub
Private Sub RunButton_Click()
''''''''''''''''''''''''
'' IN PROGRESS:
''''''''''''''''''''''''
'' TODOs:
'' 1. Make output formatting consistent.
    
'' Description: Puts one code in comma separated string of codes per line, copying all other contents per line. Useful
'' for breaking out lines in Rate Matrix for Net Analysis.
'' Notes:
'' 1. Must select the entire table that you want to expand codes into. Does not check if selection is valid.
'' 2. Must run for each code column when performing rate matrix line break-out when procedure is multi-coded.

' Ensure proper inputs:
Dim errorMessage As String: errorMessage = "Error: "
Dim hasError As Boolean
If StrComp(Trim(ExpandCodesForm.CodeColumnBox.value), vbNullString) = 0 Then
    errorMessage = errorMessage & vbCr & "Please enter the code column."
    hasError = True
End If
If StrComp(Trim(ExpandCodesForm.TableRangeBox.value), vbNullString) = 0 Then
    errorMessage = errorMessage & vbCr & "Please enter the table range."
    hasError = True
End If

If hasError = True Then
    GoTo Exit_With_Error
End If

' Get the table range:
On Error GoTo Exit_With_Error
errorMessage = "Error: provided table range is invalid."
Dim tableRange As Range: Set tableRange = Range(ExpandCodesForm.TableRangeBox.value)

' Get the column range:
Dim colRange As Range: Set colRange = Range(ExpandCodesForm.CodeColumnBox.value)
' Set column to first column in range if passed column range has more than one column:
If colRange.Columns.count <> 1 Then
    colRange = colRange.Resize(ColumnSize:=1)
End If
' Check if column does not intersect with table or is invalid:
errorMessage = "Error: column range is invalid."
Dim colString As String: colString = colRange.Address(ReferenceStyle:=xlR1C1)
Dim TEMP As String: TEMP = Mid(colString, InStr(1, colString, "C") + 1, Len(colString) - InStr(1, colString, "C"))
If InStr(1, TEMP, ":") <> 0 Then
    TEMP = Mid(TEMP, 1, InStr(1, TEMP, ":") - 1)
End If

Dim colNum As Integer: colNum = CInt(TEMP)

If Intersect(tableRange, Range(ExpandCodesForm.CodeColumnBox.value)) Is Nothing Then
    errorMessage = "Error: Code column is not on the selected table."
    GoTo Exit_With_Error
End If

' Get the relative column number on the table:
Dim currCol As Integer
Dim currColString As String
If tableRange.Columns.count <> 1 Then
    For currCol = 1 To tableRange.Columns.count
        currColString = tableRange.Cells(1, currCol).Address(ReferenceStyle:=xlR1C1)
        ' Extract the column number:
        TEMP = Mid(currColString, InStr(1, currColString, "C") + 1, Len(currColString) - InStr(1, currColString, "C"))
        If colNum = CInt(TEMP) Then
            colNum = currCol
            Exit For
        End If
    Next currCol
End If

' Paste all rows as values if specified:
If ExpandCodesForm.PasteAsValuesOption = True Then
    tableRange.Copy
    tableRange.PasteSpecial xlPasteValues
End If

' Ask to backup active sheet:
Call Macro_Utilities.BackupActiveSheet(tableRange.Worksheet)

' Display message to warn against closing macro:
CreateObject("WScript.Shell").PopUp "If macro hangs, do not exit. Is running in background.", 3, "FYI: "

Call Macro_Utilities.CodeOptimizeSettings(True)

' Perform the main routine:
Dim codes As New ArrayList
Dim currCodes As String
Dim currCell As Range
Dim row, tempRow, tempCol As Long: row = 1

' If range has column headers, start on next row:
If ExpandCodesForm.HasHeadersOption = True Then
    row = 2
End If

Do While row < CLng(tableRange.Rows.count) + CLng(1)
    currCodes = tableRange.Cells(row, colNum).value
    codes.Clear
    ' Load all of the codes in the current cell.
    Do While Len(currCodes) <> 0
        If InStr(1, currCodes, ",") = 0 Then
            ' If no commas left, assume whatever remains is a code and push onto codes vector.
            If Len(currCodes) <> 0 And StrComp(Trim(currCodes), vbNullString) <> 0 Then
                codes.Add Trim(currCodes)
            End If
            Exit Do
        Else
            TEMP = Trim(Mid(currCodes, 1, InStr(1, currCodes, ",") - 1))
            codes.Add Trim(Mid(currCodes, 1, InStr(1, currCodes, ",") - 1))
            currCodes = Mid(currCodes, InStr(1, currCodes, ",") + 1, Len(currCodes) - InStr(1, currCodes, ",") + 1)
        End If
    Loop
    If codes.count > 1 Then
    ' Insert appropriate number of lines, copying content of other columns:
        tempRow = row
        TEMP = CStr(row + 1) & ":" & CStr(row + codes.count - 1)
        tableRange.Rows(TEMP).EntireRow.Insert Shift:=xlDown
        ' Copy contents of current row for each new row:
        For currCol = 1 To tableRange.Columns.count
            ' Skip column containing long code string:
            If currCol <> colNum Then
                tableRange.Columns(currCol).Rows(TEMP).value = tableRange.Cells(row, currCol).value
            End If
        Next currCol
        ' Fill in each code:
        For currCol = 0 To codes.count - 1
            tableRange.Cells(row, colNum).value = codes.Item(currCol)
            row = row + 1
        Next currCol
    Else
SkipRow:
    ' Go to next row since no codes or only one code:
        row = row + 1
    End If
Loop

GoTo Macro_Exit

' Sub goes here on error:
Exit_With_Error:
    Me.Hide
    Unload ExpandCodesForm
' Print error message if logical error occurred:
    If hasError = True Then
        MsgBox errorMessage
    Else
        MsgBox "Error: Other macro error."
    End If
    Call Macro_Utilities.CodeOptimizeSettings(False)
    Exit Sub
' Sub goes here on input error (from the ExpandCodesForm Userform):
Input_Error:
    Me.Hide
    Unload ExpandCodesForm
    MsgBox "Error: Input is invalid."
    Call Macro_Utilities.CodeOptimizeSettings(False)
    Exit Sub
    
Macro_Exit:
    Me.Hide
    Unload ExpandCodesForm
    Call Macro_Utilities.CodeOptimizeSettings(False)
    Exit Sub

End Sub
