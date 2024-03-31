Attribute VB_Name = "UnitTests"
Option Explicit

Private test_ As New UnitTest

Sub RunTests()
    Dim test As UnitTest
    Set test = New UnitTest

    test.RegisterTest "Test_Test"
    
    test.RunTests
End Sub


Sub Test_Test()
    Set test_ = New UnitTest
    test_.AssertTrue True
    test_.AssertTrue False
    test_.AssertFalse False
    test_.AssertFalse True
    test_.AssertEqual 1, 1
    test_.AssertEqual 1, 2
    test_.AssertNotEqual 1, 2
    test_.AssertNotEqual 1, 1
    
    On Error Resume Next
    Err.Raise 9001
    test_.AssertHasError
    Err.Raise 9001
    test_.AssertHasNoError
    
    Err.Clear
    test_.AssertHasError
    Err.Clear
    test_.AssertHasNoError
    
End Sub
