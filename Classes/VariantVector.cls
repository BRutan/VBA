VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VariantVector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''''''''''''''''''''''''''''''''''''''''''
'' VariantVector Class:
''''''''''''''''''''''''''''''''''''''''''
'' Description:
'' *Vector (1 x m) of Variants with push/pop functionality.
'' Member Functions:
'' *Class_Initialize(): Class constructor. Initialize pCoreArr to be 1 x 1 and set size to 0.
'' *Class_Terminate(): Class destructor.
'' // Accessors:
'' *Size(): Return number of elements in vector.
'' *Value(index): Return value at index. Returns nothing if out of bounds.
'' // Mutators:
'' *Push(value): Append a new value into the vector.
'' *Pop(): Return value at end of vector and remove.
'' *CoreArray(array): Erase current vector and convert variant array into a vector object.
'' *Value(index, value): Alter element's value at index with passed value.
'' *Clear(): Erase entire vector, resize to 1x1.
''''''''''''''''''''''''''''''''''''''''''
'' Class Variables/Objects
''''''''''''''''''''''''''''''''''''''''''
Private pCoreArr() As Variant
Dim nSize As LongLong
''''''''''''''''''''''''''''''''''''''''''
'' Constructors/Destructor
''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    nSize = 0
    ReDim CoreArr(nRows)
End Sub
Private Sub Class_Terminate()

End Sub
''''''''''''''''''''''''''''''''''''''''''
'' Accessors
''''''''''''''''''''''''''''''''''''''''''
Public Property Get Size() As LongLong
    Size = nSize
End Property
Public Function GetValue(index As LongLong) As Variant
    If (index >= nSize Or index < 0) Then
        Exit Function
    End If
    Value = pCoreArr(index)
End Function
''''''''''''''''''''''''''''''''''''''''''
'' Mutators
''''''''''''''''''''''''''''''''''''''''''
Public Sub Push(val As Variant)
    nSize = nSize + 1
    ReDim Preserve pCoreArr(nSize)
    pCoreArr(nSize - 1) = val
End Sub
Public Function Pop() As Variant
    Dim outVal As Variant
    outVal = pCoreArr(nSize - 1)
    nSize = nSize - 1
    ReDim Preserve pCoreArr(nSize - 1)
    
    Pop = outVal
End Function
Public Property Let CoreArray(ByRef arr_in() As Variant)
'' Erase and copy contents of passed array into class core array:
    Erase pCoreArr
    ReDim pCoreArr(UBound(arr_in) - LBound(arr_in), UBound(arr_in, 1) - LBound(arr_in, 1))
    Dim i, j As Long
    For i = LBound(arr_in) To UBound(arr_in)
        For j = LBound(arr_in, 1) To UBound(arr_in, 1)
            pCoreArr(i - LBound(arr_in), j - LBound(arr_in, 1)) = arr_in(i, j)
        Next j
    Next i
End Property
Public Sub SetValue(index As LongLong, val As Variant)
    If (index >= nRows Or index < 0) Then
        Exit Sub
    End If
    pCoreArr(index) = val
End Sub
Public Sub Clear()
'' Erase all contents of the core array and redim to be 1 x 1:
    Erase pCoreArr
    nSize = 0
    ReDim pCoreArr(nSize)
End Sub
