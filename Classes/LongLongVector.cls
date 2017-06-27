VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LongLongVector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''''''''''''''''''''''''''''''''''''''''''
'' LongLongVector Class:
''''''''''''''''''''''''''''''''''''''''''
'' Description:
'' *Vector (1 x m) of LongLongs with push/pop functionality.
'' Member Functions:
'' *Class_Initialize(): Class constructor. Set Min, Max and Size to 0.
'' *Class_Terminate(): Class destructor.
'' // Accessors:
'' *Size(): Return number of elements in vector.
'' *Value(index): Return value at index. Returns nothing if out of bounds.
'' // Mutators:
'' *Push(value): Append a new value into the vector.
'' *Pop(): Return value at end of vector and remove.
'' *CoreArray(array): Erase current vector and convert LongLong array into a vector object.
'' *Value(index, value): Alter element's value at index with passed value.
'' *Clear(): Erase entire vector, resize to 1x1.
'' ***** Notes: Indices start at 1.
''''''''''''''''''''''''''''''''''''''''''
'' Class Variables/Objects
''''''''''''''''''''''''''''''''''''''''''
Private pCoreArr() As LongLong
Private pSize As Integer
Private pMax, pMin As LongLong
''''''''''''''''''''''''''''''''''''''''''
'' Constructors/Destructor
''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    pSize = pMax = pMin = 0
    ReDim pCoreArr(pSize + 1)
End Sub
Private Sub Class_Terminate()

End Sub
''''''''''''''''''''''''''''''''''''''''''
'' Accessors
''''''''''''''''''''''''''''''''''''''''''
Public Property Get Size() As LongLong
    Size = pSize
End Property
Public Function GetValue(index As LongLong) As LongLong
    If (index > pSize Or index < 1) Then
        Exit Function
    End If
    GetValue = pCoreArr(CInt(index))
End Function
Public Property Get Max() As LongLong
    Max = pMax
End Property
Public Property Get Min() As LongLong
    Min = pMin
End Property
''''''''''''''''''''''''''''''''''''''''''
'' Mutators
''''''''''''''''''''''''''''''''''''''''''
Public Sub Clear()
'' Erase all contents of the core array and redim to be 1 x 1:
    Erase pCoreArr
    pSize = pMin = pMax = 0
End Sub
Public Sub Push(val As LongLong)
    
    If val > pMax Then
        pMax = val
        If pMin = 0 Then
            pMin = pMax
        End If
    End If
    If val < pMin Then
        pMin = val
        If pMin = 0 Then
            pMin = pMax
        End If
    End If
    
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
Public Function Pop() As LongLong
    If pSize = 0 Then
        Exit Function
    End If
    
    Dim outVal As LongLong
    outVal = pCoreArr(pSize)
    If pSize <> 0 Then
        pSize = pSize - 1
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
    Dim tempVal As LongLong
    If isDescending = True Then
    ' Sort in descending order:
        While hasChanged = True
            hasChanged = False
            For i = 1 To pSize
                For j = 1 To pSize
                    If pCoreArr(i) < pCoreArr(j) Then
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
                    If pCoreArr(i) > pCoreArr(j) Then
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
Public Property Let CoreArray(ByRef arr_in() As LongLong)
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
Public Sub SetValue(index As Integer, val As LongLong)
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
