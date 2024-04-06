VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UnitTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ------------------------------------------------------------------------------------------------------------
' Project           : https://github.com/yyukki5/UnitTest
' Version           : 1.0.0
' Dependencies      :
' Features
'  - Assert*        : Asserting for test result.
'  - RegisterTest   : Register Test for RunTests()
'  - RunTests       : Run registered tests
' - ...
' ------------------------------------------------------------------------------------------------------------

Private hasError_ As Boolean
Private errNumber_ As Long
Private errDescription_ As String
Private errSource_ As String
Private errResult_ As String

Private testName_ As String
Private testNames_ As Collection

Private Sub Class_Initialize()
    Set testNames_ = New Collection
End Sub

Property Get HasError() As Boolean
    HasError = hasError_
End Property

Property Get ErrorNumber() As Long
    ErrorNumber = errNumber_
End Property

Property Get ErrorResult() As String
    ErrorResult = errResult_
End Property

Private Property Get isSingleTest() As Boolean
    isSingleTest = True
    If testNames_ Is Nothing Then Exit Property
    If testNames_.Count = 0 Then Exit Property
    isSingleTest = False
End Property

' ---------------------------------------------------------------------------------------------------------------------------
' Before using, switch error check in VBE > Tools > Options > General > Error Trapping > Break on Unhandled Errors
Public Function RunTests(uTest As UnitTest)
    Dim testName
    uTest.ClearErr
    uTest.RegisterTest "dummy_for_not_single_test"
    Debug.Print vbCrLf & "--- Start tests (" & Format(Now) & Right(Format(Timer, "0.000"), 4) & ") ---"
    For Each testName In testNames_
        RunTest (testName), uTest
    Next
    Debug.Print "--- Finish tests (" & Format(Now) & Right(Format(Timer, "0.000"), 4) & ") ---" & vbCrLf
    Set testNames_ = Nothing
    uTest.ClearRegisteredTests
End Function

Public Function RegisterTest(testName As String)
    testNames_.Add testName
End Function

Public Function ClearRegisteredTests()
    Do While testNames_.Count > 0
    testNames_.Remove (testNames_.Count)
    Loop
End Function

Public Function AssertTrue(condition)
    On Error GoTo errCondition
    If condition = True Then Else Call Err.Raise(9001, "", "NG: Should be True.")
    Exit Function
errCondition:
    CopyErr
    If isSingleTest Then DebugPrintErr
End Function

Public Function AssertFalse(condition)
    On Error GoTo errCondition
    If condition = False Then Else Call Err.Raise(9002, "", "NG: Should be False.")
    Exit Function
errCondition:
    CopyErr
     If isSingleTest Then DebugPrintErr
End Function

Public Function AssertEqual(expected, actual)
    On Error GoTo errCondition
    If VarType(expected) = vbObject And VarType(actual) = vbObject Then If expected.Equals((actual)) Then GoTo pass Else Call Err.Raise(9003, "", "NG: Should be equal.")
    If expected = actual Then Else Call Err.Raise(9003, "", "NG: Should be Equal. expected is " & expected & ", actual is " & actual)
pass:
    Exit Function
errCondition:
    CopyErr
    If isSingleTest Then DebugPrintErr
End Function

Public Function AssertNotEqual(expected, actual)
    On Error GoTo errCondition
    If VarType(expected) = vbObject And VarType(actual) = vbObject Then If expected.Equals((actual)) = False Then GoTo pass Else Call Err.Raise(9004, "", "NG: Should NOT be equal.")
    If expected <> actual Then Else Call Err.Raise(9004, "", "NG: Should NOT be Equal. expected is " & expected & ", actual is " & actual)
pass:
    Exit Function
errCondition:
    CopyErr
    If isSingleTest Then DebugPrintErr
End Function

Public Function AssertHasError()
     If Err.Number <> 0 Then
    Else
        On Error GoTo errCondition
        Call Err.Raise(9005, "", "NG: Should have error, but has NO error.")
    End If
    Err.Clear
    Exit Function
errCondition:
    CopyErr
    If isSingleTest Then DebugPrintErr: Err.Clear
End Function

Public Function AssertHasNoError()
    If Err.Number = 0 Then
    Else
        Dim num As Long:    num = Err.Number
        Dim desc As String:  desc = Err.Description
        Dim src As String:     src = Err.Source
        On Error GoTo errCondition
        Call Err.Raise(9006, "", "NG: Should have NO error, but has error." & vbCrLf & _
            "      - " & "Number: " & num & vbCrLf & _
            "      - " & "Description: " & desc & vbCrLf & _
            "      - " & "Source: " & src)
    End If
    Exit Function
errCondition:
    CopyErr
    If isSingleTest Then DebugPrintErr:     Err.Clear
End Function

Public Function ClearErr()
    hasError_ = False
    errNumber_ = 0
    errDescription_ = ""
    errSource_ = ""
    errResult_ = ""
End Function

' ------------------------------------------------------------------------
Private Function RunTest(testName As String, uTest As UnitTest)
On Error Resume Next
    Application.Run (testName)
    CopyErr  ' When application run has error, catch
    If hasError_ = False Then
        hasError_ = uTest.HasError
        errNumber_ = uTest.ErrorNumber
        errResult_ = uTest.ErrorResult
    End If
    Debug.Print (IIf(hasError_, "NG", "OK") & ": " & testName & IIf(hasError_, errResult_, ""))
    ClearErr
    uTest.ClearErr
End Function

Private Function DebugPrintErr()
    Debug.Print errDescription_
End Function

Private Function CopyErr()
    hasError_ = IIf(Err.Number > 0, True, hasError_)
    errNumber_ = IIf(Err.Number = 0, errNumber_, Err.Number)
    errDescription_ = IIf(Err.Number = 0, errDescription_, Err.Description)
    errSource_ = IIf(Err.Number = 0, errSource_, Err.Source)
    errResult_ = IIf(Err.Number = 0, errResult_, errResult_ & vbCrLf & "  - " & errSource_ & errDescription_)
End Function

