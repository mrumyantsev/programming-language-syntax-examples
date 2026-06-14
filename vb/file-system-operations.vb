'
' Saving string to a file, using
' FileSystemObject (TextStream)
'

Dim fso As Object

Dim ts As Object
Dim filePath As String
Dim data As String

filePath = "C:\test.txt"
data = "123"

Set fso = CreateObject("Scripting.FileSystemObject")
' CreateTextFile parameters: (Filename, Overwrite, Unicode)
Set ts = fso.CreateTextFile(filePath, True)

ts.WriteLine data
ts.Close

Set ts = Nothing
Set fso = Nothing
