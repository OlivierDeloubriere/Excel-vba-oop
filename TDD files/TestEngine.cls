VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("TDD")
'@PredeclaredId
Option Explicit

Private Type TTestEngine
    TestCases As Collection
End Type

Private this As TTestEngine

Private Sub class_Initialize()
    Set this.TestCases = New Collection
End Sub

Public Sub AddTestCase(ByVal myTestCase As TestCase)
    this.TestCases.Add myTestCase
End Sub

Public Sub RunTests()
    Dim testCaseItem As TestCase
    For Each testCaseItem In this.TestCases
        Debug.Print Chr(13) & testCaseItem.Description & " -- "; testCaseItem.OutputMessage
    Next testCaseItem
End Sub
