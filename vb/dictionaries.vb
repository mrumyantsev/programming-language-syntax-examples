' Declaration 1
Dim dict As Object

' Declaration 2
Dim dict As Scripting.Dictionary

' Initialization 1
Set dict = CreateObject("Scripting.Dictionary")

' Initialization 2
Set dict = New Scripting.Dictionary

' Adding key-value pairs
dict.Add "key1", "value1"

' Setting value by key
dict("key1") = "value1"

' Checking key existence (native method)
If dict.Exists("key1") Then
    ' Key is present
End If
