' Ensures that lower bound of all
' the arrays will starts from 0
Option Base 0

' Declaration of static (fixed-size) array
Dim Array(4) As String ' Now capacity is 5

' Declaration of dynamic array
Dim Array() As String ' Now no array allocated

' Allocation of dynamic array space
ReDim Array(0) ' Now capacity is 1

' Increasing array capacity with keeping
' existing elements
ReDim Preserve Array(UBound(Array) + 1) ' Now capacity is 2

' Deleting of the array from memory
Erase Array

'
' Iterating
'

Dim i As Integer

For i = 0 To UBound(Array)
    MsgBox Array(i)
Next i

'
' Emptiness Checking
'

' The most reliable way for any array type.
' Returns True if the array is unallocated (empty),
' and returns False otherwise
Function IsArrayEmpty(arr As Variant) As Boolean
    Dim ub As Long
    On Error Resume Next
    ub = UBound(arr)
    IsArrayEmpty = (Err.Number = 0)
    On Error GoTo 0
End Function

' The quickest and unreliable way
If Not Not myArr Then
    ' Array is initialized
Else
    ' Array is empty/uninitialized
End If

' Checking of plain Variant (without parentheses).
' This function does not work on arrays; it
' only works on single Variant variables that
' have not yet been assigned a value
Dim myArr As Variant

If IsEmpty(myArr) Then
    MsgBox "Empty"
End If
