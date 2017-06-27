VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringVector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''''''''''''''''''''''''''''''''''''''''''
'' StringVector Class:
''''''''''''''''''''''''''''''''''''''''''
'' Description:
'' *Vector (1 x m) of Strings with push/pop functionality.
'' Member Functions:
'' *Class_Initialize(): Class constructor. Set size to 0.
'' *Class_Terminate(): Class destructor.
'' // Accessors:
'' *Size(): Return number of elements in vector.
'' *Value(index): Return value at index. Returns nothing if out of bounds.
'' // Mutators:
'' *Push(value): Append a new value into the vector.
'' *Pop(): Return value at end of vector and remove.
'' *CoreArray(array): Erase current vector and convert Integer array into a vector object.
'' *Value(index, value): Alter element's value at index with passed value.
'' *Clear(): Erase entire vector, resize to 1x1.
'' ***** Notes: Indices start at 1.
''''''''''''''''''''''''''''''''''''''''''
'' Class Variables/Objects
''''''''''''''''''''''''''''''''''''''''''
Private pCoreArr() As String
Private pSize As Integer
''''''''''''''''''''''''''''''''''''''''''
'' Constructors/Destructor
''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    pSize = 0
    ReDim pCoreArr(pSize + 1)
End Sub
Private Sub Class_Terminate()

End Sub
''''''''''''''''''''''''''''''''''''''''''
'' Accessors
''''''''''''''''''''''''''''''''''''''''''
Public Property Get Size() As Integer
    Size = pSize
End Property
Public Function GetValue(index As Integer) As String
    If (index > pSize Or index < 1) Then
        Exit Function
    End If
    GetValue = pCoreArr(CInt(index))
End Function
''''''''''''''''''''''''''''''''''''''''''
'' Mutators
''''''''''''''''''''''''''''''''''''''''''
Public Sub Clear()
'' Erase all contents of the core array and redim to be 1 x 1:
    Erase pCoreArr
    pSize = 0
End Sub
Public Sub Push(val As String)

    If pSize = 0 Then
        pSize = pSize + 1
        ReDim pCoreArr(pSize)
        pCoreArr(pSize) = val
    Else
        pSize = pSize + 1
        ReDim Preserve pCoreArr(pSize)
        pCoreArr(pSize) = val
    End If
    
End Sub
Public Function Pop() As String
    If pSize = 0 Then
        Exit Function
    End If
    
    Dim outVal As Integer
    outVal = pCoreArr(pSize)
    pSize = pSize - 1
    If pSize <> 0 Then
        ReDim Preserve pCoreArr(pSize)
    End If
    Pop = outVal
    
End Function
Public Sub Sort(isDescending As Boolean)
    If pSize = 0 Or pSize = 1 Then
        Exit Sub
    End If
    ' Use simple bubblesort:
    Dim hasChanged As Boolean: hasChanged = True
    Dim i, j As Integer
    Dim tempVal As String
    If isDescending = True Then
    ' Sort in descending order:
        While hasChanged = True
            hasChanged = False
            For i = 1 To pSize
                For j = 1 To pSize
                    If StrComp(pCoreArr(i), pCoreArr(j)) < 0 Then
                        ' Swap the elements:
                        tempVal = pCoreArr(i)
                        pCoreArr(i) = pCoreArr(j)
                        pCoreArr(j) = tempVal
                        hasChanged = True
                    End If
                Next j
            Next i
        Loop
    Else
    ' Sort in ascending order:
    While hasChanged = True
            hasChanged = False
            For i = 1 To pSize
                For j = 1 To pSize
                    If StrComp(pCoreArr(i), pCoreArr(j)) > 0 Then
                        ' Swap the elements:
                        tempVal = pCoreArr(i)
                        pCoreArr(i) = pCoreArr(j)
                        pCoreArr(j) = tempVal
                        hasChanged = True
                    End If
                Next j
            Next i
        Loop
    
    End If
End Sub
Public Property Let CoreArray(ByRef arr_in() As String)
'' Erase and copy contents of passed array into class core array, make first index 0:
    Erase pCoreArr
    ReDim pCoreArr(UBound(arr_in) - LBound(arr_in), UBound(arr_in, 1) - LBound(arr_in, 1))
    Dim i, j As Integer
    For i = LBound(arr_in) To UBound(arr_in) + 1
        For j = LBound(arr_in, 1) To UBound(arr_in, 1) + 1
            pCoreArr(i - LBound(arr_in), j - LBound(arr_in, 1)) = arr_in(i, j)
        Next j
    Next i
End Property
Public Sub SetValue(index As Integer, val As String)
    If (index > pSize Or index < 1) Then
        Exit Sub
    End If
    If val > pMax Then
        pMax = val
    ElseIf val < pMin Then
        pMin = val
    End If
    pCoreArr(index) = val
End Sub


