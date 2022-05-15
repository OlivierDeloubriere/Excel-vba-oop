VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestCase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("TDD")
'@PredeclaredId

Option Explicit

Private Type TTestCase
    Description As String
    TestResult As Boolean
    OutputMessage As String
End Type

Private this As TTestCase
Public Function Create(ByVal myDesc As String) As TestCase
    Dim result As TestCase
    Set result = New TestCase
    
    result.Description = myDesc
    result.TestResult = False
    result.OutputMessage = vbNullString
    
    Set Create = result
End Function
Public Property Let TestResult(ByVal value As Boolean)
    this.TestResult = value
End Property
Public Property Get TestResult() As Boolean
    TestResult = this.TestResult
End Property
Public Property Let OutputMessage(ByVal value As String)
    this.OutputMessage = value
End Property
Public Property Get OutputMessage() As String
    OutputMessage = this.OutputMessage
End Property
Public Property Let Description(ByVal value As String)
    this.Description = value
End Property
Public Property Get Description() As String
    Description = this.Description
End Property

Public Sub IsEqual(ByVal value1 As Variant, value2 As Variant)
    If value1 = value2 Then
        TestResult = True
        OutputMessage = "Test Passed --"
    Else
        TestResult = False
        OutputMessage = value1 & " is not equal to " & value2 & " -- Test Failed --"
    End If
End Sub
