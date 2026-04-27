' Declaration
Dim myColl As Collection

' Initialization
Set myColl = New Collection

' Adding value-key pairs
myColl.Add "value1", "key1"
myColl.Add "value2", "key2"

' Getting keys
MsgBox myColl("key1") ' Returns "value1"
MsgBox myColl(2)      ' Returns "value2"

' Iterating
Dim item As Variant

For Each item In myColl
    MsgBox item
Next item

' Key checking (non-native but the fastest way)
Function IsKeyExists(ByRef col As Collection, ByVal key As String) As Boolean
    On Error Resume Next
    Dim temp As Variant
    temp = col(key)
    IsKeyExists = (Err.Number = 0)
    On Error GoTo 0
End Function

' Value checking
Function IsValueExists(ByRef col As Collection, ByVal val As Variant) As Boolean
    Dim item As Variant
    
    For Each item In col
        If item = val Then
            IsValueExists = True
            
            Exit Function
        End If
    Next item
    
    IsValueExists = False
End Function
