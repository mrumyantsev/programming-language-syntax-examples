' Forces the explicit declaration
' of all variables
Option Explicit

' Global array variable
Dim List() As String

' Global object variable
Dim Obj As Object

Sub ClearGlobalVariables()
    Erase List

    Set Obj = Nothing
End Sub
