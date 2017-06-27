VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LineByLineForm 
   Caption         =   "Line By Line"
   ClientHeight    =   4710
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5640
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
Dim errorMessage1 As String
Dim baseError As String: baseError = "Please enter the following inputs: "

' Ensure input integrity:
If tableRangeInput.Value = vbNullString Then
    errorMessage1 = baseError + " Table Range"
End If
If CPTManifestInput.Value = vbNullString Then
    If errorMessage1 = vbNullString Then
        errorMessage1 = baseError + " CPT Manifest Table"
    Else
        errorMessage1 = errorMessage1 + vbCr + "CPT Manifest Table"
    End If
End If
If CPTColumnInput.Value = vbNullString Then
    If errorMessage1 = vbNullString Then
        errorMessage1 = baseError + " CPT Column"
    Else
        errorMessage1 = errorMessage1 + vbCr + "CPT Column"
    End If
    ' Append
End If
If RVUColumnInput.Value = vbNullString Then
    If errorMessage1 = vbNullString Then
        errorMessage1 = baseError + " RVU Column"
    Else
        errorMessage1 = errorMessage1 + vbCr + "RVU Column"
    End If
End If
If ProposedPriceInput.Value = vbNullString Then
    If errorMessage1 = vbNullString Then
        errorMessage1 = baseError + " Proposed Price"
    Else
        errorMessage1 = errorMessage1 + vbCr + "Proposed Price"
    End If
End If
If SuggestedPriceInput.Value = vbNullString Then
    If errorMessage1 = vbNullString Then
        errorMessage1 = baseError + " Suggested Price"
    Else
        errorMessage1 = errorMessage1 + vbCr + "Suggested Price"
    End If
End If

If errorMessage1 <> vbNullString Then
    ' If any of the fields are empty then show error message and quit:
    MsgBox errorMessage1
    Exit Sub
End If

' Make copy of worksheet state for use in Undo:
Dim tableRange As Range: Set tableRange = Range(tableRangeInput.Text)
ReDim originalContents(tableRange.Rows.Count, tableRange.Columns.Count)
Dim i, j As Integer
For i = 1 To UBound(originalContents, 1)
    For j = 1 To UBound(originalContents, 2)
        originalContents(i, j) = tableRange.Cells(i, j).Value
    Next j
Next i

Set originalSheet = Application.ActiveSheet

Application.OnUndo "Undoing LineByLine...", "UndoLineByLine"
' Run the LineByLine sub:

LineByLine.LineByLine Range(tableRangeInput.Value), Range(RVUColumnInput.Value).column, Range(CPTColumnInput.Value).column, _
Range(ProposedPriceInput.Value).column, Range(SuggestedPriceInput.Value).column, Range(CPTManifestInput.Value), 0

MsgBox "Done"

Me.Hide
Unload Me

End Sub

Private Sub CancelButton_Click()
    Me.Hide
    Unload Me
End Sub




''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Form Control Event Definitions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Private Sub CPTColumnInput_BeforeDragOver(Cancel As Boolean, ByVal Data As MSForms.DataObject, ByVal x As stdole.OLE_XPOS_CONTAINER, ByVal y As stdole.OLE_YPOS_CONTAINER, ByVal DragState As MSForms.fmDragState, Effect As MSForms.fmDropEffect, ByVal Shift As Integer)
'
'End Sub
'Private Sub CPTManifestInput_BeforeDragOver(Cancel As Boolean, ByVal Data As MSForms.DataObject, ByVal x As stdole.OLE_XPOS_CONTAINER, ByVal y As stdole.OLE_YPOS_CONTAINER, ByVal DragState As MSForms.fmDragState, Effect As MSForms.fmDropEffect, ByVal Shift As Integer)
'
'End Sub
'Private Sub ProposedPriceInput_BeforeDragOver(Cancel As Boolean, ByVal Data As MSForms.DataObject, ByVal x As stdole.OLE_XPOS_CONTAINER, ByVal y As stdole.OLE_YPOS_CONTAINER, ByVal DragState As MSForms.fmDragState, Effect As MSForms.fmDropEffect, ByVal Shift As Integer)
'
'End Sub
'Private Sub RVUColumnInput_BeforeDragOver(Cancel As Boolean, ByVal Data As MSForms.DataObject, ByVal x As stdole.OLE_XPOS_CONTAINER, ByVal y As stdole.OLE_YPOS_CONTAINER, ByVal DragState As MSForms.fmDragState, Effect As MSForms.fmDropEffect, ByVal Shift As Integer)
'
'End Sub
'Private Sub SuggestedPriceInput_BeforeDragOver(Cancel As Boolean, ByVal Data As MSForms.DataObject, ByVal x As stdole.OLE_XPOS_CONTAINER, ByVal y As stdole.OLE_YPOS_CONTAINER, ByVal DragState As MSForms.fmDragState, Effect As MSForms.fmDropEffect, ByVal Shift As Integer)
'
'End Sub
'Private Sub tableRangeInput_BeforeDragOver(Cancel As Boolean, ByVal Data As MSForms.DataObject, ByVal x As stdole.OLE_XPOS_CONTAINER, ByVal y As stdole.OLE_YPOS_CONTAINER, ByVal DragState As MSForms.fmDragState, Effect As MSForms.fmDropEffect, ByVal Shift As Integer)
'
'End Sub


Public Sub UndoLineByLine()
    On Error GoTo Message
    
    If originalSheet Is Nothing Then
        Exit Sub
    Else
        originalSheet.Activate
        ' Replace contents of mainRange with original contents:
        Dim i, j As Integer
        For i = 1 To originalRange.Rows.Count
            For j = 1 To originalRange.Columns.Count
                mainRange.Cells(i, j).Value = originalRange.Cells(i, j).Value
            Next j
        Next i
    End If

Message:
    MsgBox "Error: Can't Undo."

End Sub
