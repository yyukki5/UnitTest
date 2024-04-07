Attribute VB_Name = "UnitTests"
Option Explicit

Private test_ As New UnitTest

Sub RunTests()
    Dim test As New UnitTest

    test.RegisterTest "Test_Test"
    test.RegisterTest "Test_Test1"
    test.RegisterTest "Add_Scenario_ExpectedBehavior", 4, 17, 1
    test.RegisterTest "Add_Scenario_ExpectedBehavior", 10, 16, 26
    
    test.RunTests test_
End Sub


Sub Test_Test()
    test_.AssertTrue True
    test_.AssertTrue False
    test_.AssertFalse False
    test_.AssertFalse True
    test_.AssertEqual 1, 1
    test_.AssertEqual 1, 2
    test_.AssertNotEqual 1, 2
    test_.AssertNotEqual 1, 1

    On Error Resume Next '<--- Need for .AssertHasError(), .AssertHasNoError()
    Err.Raise 9001
    test_.AssertHasError
    Err.Raise 9001, "Sample", "This is sample Error."
    test_.AssertHasNoError

    Err.Clear
    test_.AssertHasError
    Err.Clear
    test_.AssertHasNoError
    On Error GoTo 0
    
End Sub

Sub Add_Scenario_ExpectedBehavior(a, b, res)
    ' Arrange
    Dim result As Double
    ' Act
    result = Add(a, b)
    ' Assert
    test_.AssertEqual res, result
End Sub

Private Function Add(a, b) As Double
    Add = a + b
End Function
