VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LineByLineForm 
   Caption         =   "Line By Line"
   ClientHeight    =   4944
   ClientLeft      =   120
   ClientTop       =   456
   ClientWidth     =   5616
   OleObjectBlob   =   "LineByLineForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "LineByLineForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim originalContents() As String
Dim mainRange As Range
Dim originalSheet As Worksheet

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' LineByLineForm
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Description:
'' * Used to run the LineByLine macro.
'' Instructions:
'' * Place Button or any other control onto the sheet you want to run the macro on. When the Button/control is clicked the
'' Line By Line userform will pop up and remain until you either successfully run the macro or click cancel.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' ***************************** Main Control *****************************
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RunButton_Click()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' errorMessage1 indicates which inputs are empty.
Dim errorMessage1 As String: errorMessage1 = "Error: "
Dim hasError As Boolean

' Ensure input integrity:
If tableRangeInput.value = vbNullString Then
    errorMessage1 = errorMessage1 + vbCr + "Please enter the table range."
    hasError = True
End If
If CPTManifestInput.value = vbNullString Then
    errorMessage1 = errorMessage1 + vbCr + "Please enter CPT Manifest range in table."
    hasError = True
End If
If CPTColumnInput.value = vbNullString Then
    errorMessage1 = errorMessage1 + vbCr + "Please enter CPT column range in table."
    hasError = True
End If
If RVUColumnInput.value = vbNullString Then
    errorMessage1 = errorMessage1 + vbCr + "Please enter RVU column in table."
    hasError = True
End If
If ProposedPriceInput.value = vbNullString Then
    errorMessage1 = errorMessage1 + vbCr + "Please enter Proposed Price column in table."
    hasError = True
End If
If SuggestedPriceInput.value = vbNullString Then
    errorMessage1 = errorMessage1 + vbCr + "Please enter Suggested Price column in table."
    hasError = True
End If

If hasError = True Then
    ' If any of the fields are empty then show error message and quit:
    MsgBox errorMessage1
    Exit Sub
End If

' Make copy of original table:
Dim currWs As Worksheet: Set currWs = Application.ActiveSheet
Dim newWS As Worksheet
Set newWS = Application.ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
newWS.Name = "Original Table pg." + CStr(ThisWorkbook.Sheets.count)

Dim rangeCopy As Range
Set rangeCopy = newWS.Range(Cells(1, 1), Cells(Range(tableRangeInput.value).Rows.count, Range(tableRangeInput.value).Columns.count))
Dim row, col As Integer
For row = 1 To rangeCopy.Rows.count
    For col = 1 To rangeCopy.Columns.count
    ' Copy contents
        rangeCopy.Cells(row, col).value = Range(tableRangeInput.value).Cells(row, col).value
    Next col
Next row

'' Make copy of worksheet state for use in Undo:
'Dim tableRange As Range: Set tableRange = Range(tableRangeInput.Text)
'ReDim originalContents(tableRange.Rows.Count, tableRange.Columns.Count)
'Dim i, j As Integer
'For i = 1 To UBound(originalContents, 1)
'    For j = 1 To UBound(originalContents, 2)
'        originalContents(i, j) = tableRange.Cells(i, j).Value
'    Next j
'Next i

Set originalSheet = Application.ActiveSheet

'Application.OnUndo "Undoing LineByLine...", "UndoLineByLine"
'' Run the LineByLine sub:

LineByLine.LineByLine Range(tableRangeInput.value), Range(RVUColumnInput.value).Column, Range(CPTColumnInput.value).Column, _
Range(ProposedPriceInput.value).Column, Range(SuggestedPriceInput.value).Column, Range(CPTManifestInput.value), 0

MsgBox "Done"

Me.Hide
Unload Me

End Sub
Private Sub CancelButton_Click()
    Me.Hide
    Unload Me
End Sub

End Sub
