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
' Version           : 1.1.0
' ------------------------------------------------------------------------------------------------------------

Private hasError_ As Boolean
Private errNumber_ As Long
Private errDescription_ As String
Private errSource_ As String
Private errResult_ As String

Private testFunctions_ As Collection

Private Const KEY_NAME As String = "TestName"
Private Const KEY_PARAMETERS As String = "Parameters"

Private Sub Class_Initialize()
    Set testFunctions_ = New Collection
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
    If testFunctions_ Is Nothing Then Exit Property
    If testFunctions_.Count = 0 Then Exit Property
    isSingleTest = False
End Property

' ---------------------------------------------------------------------------------------------------------------------------
' Before using, switch error check in VBE > Tools > Options > General > Error Trapping > Break on Unhandled Errors
Public Function RunTests(uTest As UnitTest)
    Dim testFunction As Collection, testFunc As Collection
    
    uTest.ClearErr
    uTest.RegisterTest "dummy_for_not_single_test"
    
    Debug.Print vbCrLf & "--- Start tests (" & Format(Now) & Right(Format(Timer, "0.000"), 4) & ") ---"
    For Each testFunction In testFunctions_
        RunTest testFunction, uTest
    Next
    Debug.Print "--- Finish tests (" & Format(Now) & Right(Format(Timer, "0.000"), 4) & ") ---" & vbCrLf
    
    Set testFunctions_ = Nothing
    uTest.ClearRegisteredTests
End Function

Public Function RegisterTest(testName As String, ParamArray parameters())
    Dim testFunc As Collection
    Set testFunc = New Collection
    testFunc.Add testName, KEY_NAME
    If UBound(parameters()) > 0 Then testFunc.Add parameters(), KEY_PARAMETERS Else testFunc.Add Array(), KEY_PARAMETERS
    testFunctions_.Add testFunc
End Function

Public Function ClearRegisteredTests()
    Do While testFunctions_.Count > 0
    testFunctions_.Remove (testFunctions_.Count)
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
Private Function RunTest(testFunc As Collection, uTest As UnitTest)
On Error Resume Next
    ApplicationRun testFunc
    CopyErr  ' When application run has error, catch
    If hasError_ = False Then
        hasError_ = uTest.HasError
        errNumber_ = uTest.ErrorNumber
        errResult_ = uTest.ErrorResult
    End If
    Debug.Print (IIf(hasError_, "NG", "OK") & ": " & GetFunctionStringWithArgument(testFunc) & IIf(hasError_, errResult_, ""))
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

Private Function ApplicationRun(testFunc As Collection) As Variant
    Dim p, funcName As String, res
    
    funcName = testFunc(KEY_NAME)
    p = testFunc(KEY_PARAMETERS)
    Select Case UBound(p) + 1
        Case 0: res = Run(funcName)
        Case 1: res = Run(funcName, p(0))
        Case 2: res = Run(funcName, p(0), p(1))
        Case 3: res = Run(funcName, p(0), p(1), p(2))
        Case 4: res = Run(funcName, p(0), p(1), p(2), p(3))
        Case 5: res = Run(funcName, p(0), p(1), p(2), p(3), p(4))
        Case 6: res = Run(funcName, p(0), p(1), p(2), p(3), p(4), p(5))
        Case 7: res = Run(funcName, p(0), p(1), p(2), p(3), p(4), p(5), p(6))
        Case 8: res = Run(funcName, p(0), p(1), p(2), p(3), p(4), p(5), p(6), p(7))
        Case 9: res = Run(funcName, p(0), p(1), p(2), p(3), p(4), p(5), p(6), p(7), p(8))
        Case 10: res = Run(funcName, p(0), p(1), p(2), p(3), p(4), p(5), p(6), p(7), p(8), p(9))
        Case 11: res = Run(funcName, p(0), p(1), p(2), p(3), p(4), p(5), p(6), p(7), p(8), p(9), p(10))
        Case 12: res = Run(funcName, p(0), p(1), p(2), p(3), p(4), p(5), p(6), p(7), p(8), p(9), p(10), p(11))
        Case 13: res = Run(funcName, p(0), p(1), p(2), p(3), p(4), p(5), p(6), p(7), p(8), p(9), p(10), p(11), p(12))
        Case 14: res = Run(funcName, p(0), p(1), p(2), p(3), p(4), p(5), p(6), p(7), p(8), p(9), p(10), p(11), p(12), p(13))
        Case 15: res = Run(funcName, p(0), p(1), p(2), p(3), p(4), p(5), p(6), p(7), p(8), p(9), p(10), p(11), p(12), p(13), p(14))
        Case 16: res = Run(funcName, p(0), p(1), p(2), p(3), p(4), p(5), p(6), p(7), p(8), p(9), p(10), p(11), p(12), p(13), p(14), p(15))
        Case 17: res = Run(funcName, p(0), p(1), p(2), p(3), p(4), p(5), p(6), p(7), p(8), p(9), p(10), p(11), p(12), p(13), p(14), p(15), p(16))
        Case 18: res = Run(funcName, p(0), p(1), p(2), p(3), p(4), p(5), p(6), p(7), p(8), p(9), p(10), p(11), p(12), p(13), p(14), p(15), p(16), p(17))
        Case 19: res = Run(funcName, p(0), p(1), p(2), p(3), p(4), p(5), p(6), p(7), p(8), p(9), p(10), p(11), p(12), p(13), p(14), p(15), p(16), p(17), p(18))
        Case 20: res = Run(funcName, p(0), p(1), p(2), p(3), p(4), p(5), p(6), p(7), p(8), p(9), p(10), p(11), p(12), p(13), p(14), p(15), p(16), p(17), p(18), p(19))
        Case 21: res = Run(funcName, p(0), p(1), p(2), p(3), p(4), p(5), p(6), p(7), p(8), p(9), p(10), p(11), p(12), p(13), p(14), p(15), p(16), p(17), p(18), p(19), p(20))
        Case 22: res = Run(funcName, p(0), p(1), p(2), p(3), p(4), p(5), p(6), p(7), p(8), p(9), p(10), p(11), p(12), p(13), p(14), p(15), p(16), p(17), p(18), p(19), p(20), p(21))
        Case 23: res = Run(funcName, p(0), p(1), p(2), p(3), p(4), p(5), p(6), p(7), p(8), p(9), p(10), p(11), p(12), p(13), p(14), p(15), p(16), p(17), p(18), p(19), p(20), p(21), p(22))
        Case 24: res = Run(funcName, p(0), p(1), p(2), p(3), p(4), p(5), p(6), p(7), p(8), p(9), p(10), p(11), p(12), p(13), p(14), p(15), p(16), p(17), p(18), p(19), p(20), p(21), p(22), p(23))
        Case 25: res = Run(funcName, p(0), p(1), p(2), p(3), p(4), p(5), p(6), p(7), p(8), p(9), p(10), p(11), p(12), p(13), p(14), p(15), p(16), p(17), p(18), p(19), p(20), p(21), p(22), p(23), p(24))
        Case 26: res = Run(funcName, p(0), p(1), p(2), p(3), p(4), p(5), p(6), p(7), p(8), p(9), p(10), p(11), p(12), p(13), p(14), p(15), p(16), p(17), p(18), p(19), p(20), p(21), p(22), p(23), p(24), p(25))
        Case 27: res = Run(funcName, p(0), p(1), p(2), p(3), p(4), p(5), p(6), p(7), p(8), p(9), p(10), p(11), p(12), p(13), p(14), p(15), p(16), p(17), p(18), p(19), p(20), p(21), p(22), p(23), p(24), p(25), p(26))
        Case 28: res = Run(funcName, p(0), p(1), p(2), p(3), p(4), p(5), p(6), p(7), p(8), p(9), p(10), p(11), p(12), p(13), p(14), p(15), p(16), p(17), p(18), p(19), p(20), p(21), p(22), p(23), p(24), p(25), p(26), p(27))
        Case 29: res = Run(funcName, p(0), p(1), p(2), p(3), p(4), p(5), p(6), p(7), p(8), p(9), p(10), p(11), p(12), p(13), p(14), p(15), p(16), p(17), p(18), p(19), p(20), p(21), p(22), p(23), p(24), p(25), p(26), p(27), p(28))
        Case 30: res = Run(funcName, p(0), p(1), p(2), p(3), p(4), p(5), p(6), p(7), p(8), p(9), p(10), p(11), p(12), p(13), p(14), p(15), p(16), p(17), p(18), p(19), p(20), p(21), p(22), p(23), p(24), p(25), p(26), p(27), p(28), p(29))
        Case Else: Err.Raise 9001, "", "(UnitTest) Registered Test has Error, Arguments are over 30."
    End Select
    ApplicationRun = res
End Function

Private Function GetFunctionStringWithArgument(testFunc As Collection) As String
    Dim str As String, args, arg
    
    str = testFunc(KEY_NAME)
    args = testFunc(KEY_PARAMETERS)
    If UBound(args) < 0 Then GetFunctionStringWithArgument = str: Exit Function
    
    str = str & "("
    For Each arg In args: str = str & arg & ",": Next
    str = Left(str, Len(str) - 1) & ")"
    GetFunctionStringWithArgument = str
End Function

