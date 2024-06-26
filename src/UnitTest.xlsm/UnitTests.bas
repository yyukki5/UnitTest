Attribute VB_Name = "UnitTests"
Option Explicit

Sub CreateTests()
    UnitTest.CreateRunTests
End Sub

Sub RunTests()
   Dim test As New UnitTest

    test.RegisterTest "Test_Test"
    test.RegisterTest "Add_Scenario_ExpectedBehavior", 4, 17, 1
    test.RegisterTest "Add_Scenario_ExpectedBehavior", 10, 16, 26
    test.RegisterTest "Add_Scenario_ExpectedBehavior", 1, 2, 3
    test.RegisterTest "Add_Scenario_ExpectedBehavior", 2, 3, 4

    test.RunTests UnitTest
End Sub

'[Fact]
Sub Test_Test()
    Dim col As Collection
    Set col = New Collection
    
    With UnitTest.NameOf("Test for sample")
        .AssertTrue True
        .AssertTrue False
        .AssertTrue "Hello"
        .AssertTrue col
        
        .AssertFalse False
        .AssertFalse True
        .AssertEqual 1, 1
        .AssertEqual 1, 2
        .AssertNotEqual 1, 2
        .AssertNotEqual 1, 1
    
        On Error Resume Next '<--- Need for .AssertHasError(), .AssertHasNoError()
        Err.Raise 9001
        .AssertHasError
        Err.Raise 9001, "Sample", "This is sample Error."
        .AssertHasNoError
    
        Err.Clear
        .AssertHasError
        Err.Clear
        .AssertHasNoError
        On Error GoTo 0
        
    End With
End Sub


'[Theory]
'[InlineData(4, 17, 1)]
'[InlineData(10, 16, 26)]
'[MemberData(GetMembers)]
Sub Add_Scenario_ExpectedBehavior(a, b, res As Double)
    ' Arrange
    Dim result As Double
    ' Act
    result = Add(a, b)
    ' Assert
    UnitTest.AssertEqual res, result
End Sub

Private Function Add(a, b) As Double
    Add = a + b
End Function

Public Function GetMembers() As Collection
    Dim c As New Collection, i As Long
    For i = 1 To 2
       c.Add Array(i, i + 1, i + 2)
    Next i
    Set GetMembers = c
End Function

