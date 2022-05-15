Attribute VB_Name = "modClient"
'@Folder("TDD")
Option Explicit

Public Sub test()
    Dim tCase As TestCase
    
    With New TestEngine
        Set tCase = TestCase.Create("Test of Function Add")
        tCase.IsEqual Add(1, 2), 3
        .AddTestCase tCase
        .RunTests
    End With
End Sub

Public Function Add(ByVal x As Integer, y As Integer) As Integer
    Add = x + y
End Function
