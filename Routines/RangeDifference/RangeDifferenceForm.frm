VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RangeDifferenceForm 
   Caption         =   "Range Difference"
   ClientHeight    =   3555
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4605
   OleObjectBlob   =   "RangeDifferenceForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "RangeDifferenceForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' RangeDifferenceForm Userform
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Description:
'' * Userform used to get ranges for use with RangeDifference subroutine.
'' * RangeDifference finds all "records" (row with values in each column) that are in worksheet range A that are NOT in
'' * worksheet range B (A - B), and places them in the specified output worksheet range.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Notes:
'' * # of columns must match for ranges A and B.
'' * If the Has Headers radiobutton is clicked, the subroutine will determine if the column names for each range
'' * match. If they don't, then the subroutine will display a message and not run.
Private Sub RunButton_Click()

Dim baseMessage As String: baseMessage = "(Error) Enter the following inputs:"
Dim hasChanged As Boolean: hasChanged = False

If ARangeInput.Value = vbNullString Then
    baseMessage = baseMessage + "A Range" + vbCr
    hasChanged = True
End If
If BRangeInput.Value = vbNullString Then
    baseMessage = baseMessage + "B Range" + vbCr
    hasChanged = True
End If
If OutputRangeInput.Value = vbNullString Then
    baseMessage = baseMessage + "Output Range" + vbCr
    hasChanged = True
End If

' If any fields are omitted, show message box and do nothing:
If hasChanged = True Then
    MsgBox baseMessage, vbOKOnly
    Exit Sub
End If

Dim A, B As Range: Set A = Range(ARangeInput.Value): Set B = Range(BRangeInput.Value)
Dim outputRange As Range: Set outputRange = Range(OutputRangeInput.Value)

' Ensure integrity of inputs:
If A.Columns.Count <> B.Columns.Count Then
    MsgBox "Error: # of columns do not match", vbOKOnly
End If

' Ensure that headers match (if required):
If HeaderOptionRadio.Value = True Then
    Dim i As Integer
    For i = 1 To A.Columns.Count
        If StrComp(A.Cells(1, i).Value, B.Cells(1, i).Value) Then
            MsgBox "Error: Headers do not match."
            Exit Sub
        End If
    Next i
End If

' Resize the outputRange to match the sets' # of columns:
Set outputRange = outputRange.Resize(ColumnSize:=A.Columns.Count)

' If all fields filled, run the RangeDifference macro:

RangeDifference.RangeDifference Range(ARangeInput.Value), Range(BRangeInput.Value), outputRange, HeaderOptionRadio.Value

End Sub
Private Sub ExitButton_Click()
    Me.Hide
    Unload Me
End Sub
