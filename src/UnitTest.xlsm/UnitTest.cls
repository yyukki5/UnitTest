VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UnitTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ------------------------------------------------------------------------------------------------------------
' Project           : https://github.com/yyukki5/UnitTest
' Version           : 1.3.0
' ------------------------------------------------------------------------------------------------------------

Private hasError_ As Boolean
Private errNumber_ As Long
Private errDescription_ As String
Private errSource_ As String
Private errResult_ As String

Private testFunctions_ As Collection
Private resultTexts_ As Collection

Private Const KEY_NAME As String = "TestName"
Private Const KEY_PARAMETERS As String = "Parameters"
Private Const RUN_TESTS_NAME As String = "RunTests"

Private Sub Class_Initialize()
    Set testFunctions_ = New Collection
    Set resultTexts_ = New Collection
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

Property Get Results() As Collection
    Set Results = resultTexts_
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
    If condition = True Then Else Call Err.Raise(9001, "", "AssertTrue(); Expected: True; Actual: " & condition)
    PrePassed "AssertTrue()"
    Exit Function
errCondition:
    ErrorProcess
End Function

Public Function AssertFalse(condition)
    On Error GoTo errCondition
    If condition = False Then Else Call Err.Raise(9002, "", "AssertFalse(); Expected: False; Actual: " & condition)
    PrePassed "AssertFalse()"
    Exit Function
errCondition:
    ErrorProcess
End Function

Public Function AssertEqual(Expected, Actual)
    On Error GoTo errCondition
    If Equals(Expected, Actual) = False Then Call Err.Raise(9003, "", "AssertEqual() : Values differ; Expected: " & ToString(Expected) & "; Actual: " & ToString(Actual))
    PrePassed "AssertEqual(); Expected is " & ToString(Expected) & "; Actual is " & ToString(Actual)
    Exit Function
errCondition:
    ErrorProcess
End Function

Public Function AssertNotEqual(Expected, Actual)
    On Error GoTo errCondition
    If Equals(Expected, Actual) Then Call Err.Raise(9004, "", "AssertNotEqual() : Values are equal; Expected: Not " & ToString(Expected) & "; Actual: " & ToString(Actual))
    PrePassed "AssertNotEqual(); Expected is " & ToString(Expected) & ", Actual is " & ToString(Actual)
    Exit Function
errCondition:
    ErrorProcess
End Function

Public Function AssertSame(Expected, Actual)
    On Error GoTo errCondition
    If Expected Is Actual Then Else Call Err.Raise(9007, "", "AssertSame() : Values are not the same instance; ")
    PrePassed "AssertSame()"
    Exit Function
errCondition:
    ErrorProcess
End Function

Public Function AssertNotSame(Expected, Actual)
    On Error GoTo errCondition
    If Expected Is Actual Then Call Err.Raise(9008, "", "AssertNotSame() : Values are the same instance")
    PrePassed "AssertNotSame()"
    Exit Function
errCondition:
    ErrorProcess
End Function

Public Function AssertContains(Expected, collect)
    On Error GoTo errCondition
    Dim src
    For Each src In collect
        If Equals(Expected, src) Then PrePassed "AssertContains()":  Exit Function
    Next
    Call Err.Raise(9009, "", "AssertContains() : Item not found in collection; Collection:" & ToString(collect) & "; Not found: " & ToString(Expected))
errCondition:
    ErrorProcess
End Function

Public Function AssertDoesNotContains(Expected, collect)
    On Error GoTo errCondition
    Dim src
    For Each src In collect
        If Equals(Expected, src) Then Call Err.Raise(9010, "", "AssertDoesNotContain() : Item found in collection; Collection: " & ToString(collect) & "; Found: " & ToString(Expected))
    Next
    PrePassed "AssertDoesNotContain()"
    Exit Function
errCondition:
    ErrorProcess
End Function

Public Function AssertInRange(Actual, low, high)
    On Error GoTo errCondition
    If low <= Actual And Actual <= high Then Else Call Err.Raise(9011, "", "AssertInRange() : Value not in range; Range: (" & low & "-" & high & "); Actual: " & Actual)
    PrePassed "AssertInRange() : " & Actual & " is in Range(" & low & "-" & high & ")"
    Exit Function
errCondition:
    ErrorProcess
End Function

Public Function AssertNotInRange(Actual, low, high)
    On Error GoTo errCondition
    If low <= Actual And Actual <= high Then Call Err.Raise(9012, "", "AssertNotInRange() : Value in range; Range: (" & low & "-" & high & "); Actual; " & Actual)
    PrePassed "AssertNotInRange() : " & Actual & " is out of Range(" & low & "-" & high & ")"
    Exit Function
errCondition:
    ErrorProcess
End Function

Public Function AssertStartsWith(expectedStartString, actualString)
    On Error GoTo errCondition
    If Left(actualString, Len(expectedStartString)) = expectedStartString Then Else Call Err.Raise(9013, "", "AssertStartsWith() : String start does not match; String: """ & actualString & """; Expected start: " & expectedStartString & """")
    PrePassed "AssertStartsWith() : Value is started with """ & expectedStartString & """"
    Exit Function
errCondition:
    ErrorProcess
End Function

Public Function AssertEndsWith(expectedEndString, actualString)
    On Error GoTo errCondition
    If Right(actualString, Len(expectedEndString)) = expectedEndString Then Else Call Err.Raise(9014, "", "AssertEndsWith() : String end does not match; String: """ & actualString & """; Expected end: " & expectedEndString & """")
    PrePassed "AssertEndsWith() : Value is not started with """ & expectedEndString & """"
    Exit Function
errCondition:
    ErrorProcess
End Function

Public Function AssertEmpty(collect)
    On Error GoTo errCondition
    Dim src
    For Each src In collect
        Call Err.Raise(9015, "", "AssertEmpty() : Collection was not empty; Collection: " & ToString(collect))
    Next
    PrePassed "AssertEmpty()"
    Exit Function
errCondition:
    ErrorProcess
End Function

Public Function AssertNotEmpty(collect)
    On Error GoTo errCondition
    Dim src
    For Each src In collect
        PrePassed "AssertNotEmpty()"
        Exit Function
    Next
    Call Err.Raise(9016, "", "AssertNotEmpty() : Collection was empty")
errCondition:
    ErrorProcess
End Function

Public Function AssertNull(obj)
    On Error GoTo errCondition
    If IsNull(obj) Then Else Call Err.Raise(9017, "", "AssertNull() : Value is not null; Expected: Null; Actual: " & ToString(obj))
    PrePassed "AssertNull()"
    Exit Function
errCondition:
    ErrorProcess
End Function

Public Function AssertNotNull(obj)
    On Error GoTo errCondition
    If IsNull(obj) Then Call Err.Raise(9018, "", "AssertNotNull() : Value does not have a value")
    PrePassed "AssertNotNull()"
    Exit Function
errCondition:
    ErrorProcess
End Function

Public Function AssertNothing(obj)
    On Error GoTo errCondition
    If VarType(obj) <> vbObject Then Call Err.Raise(9019, "", "AssertNothing() : Value is not object;")
    If obj Is Nothing Then Else Call Err.Raise(9019, "", "AssertNothing() : Value is not Nothing; Expected: Nothing; Actual: " & ToString(obj))
    PrePassed "AssertNothing()"
    Exit Function
errCondition:
    ErrorProcess
End Function

Public Function AssertNotNothing(obj)
    On Error GoTo errCondition
    If VarType(obj) <> vbObject Then Call Err.Raise(9020, "", "AssertNotNothing() : Value is not object;")
    If obj Is Nothing Then Call Err.Raise(9020, "", "AssertNotNothing() : Value does not have a value")
    PrePassed "AssertNotNothing() : Value is NOT Nothing"
    Exit Function
errCondition:
    ErrorProcess
End Function

Public Function AssertIsType(expectedType As String, obj)
    On Error GoTo errCondition
    If TypeName(obj) = expectedType Then Else Call Err.Raise(9021, "", "AssertIsType() : Value is not the exact type; Expected: """ & expectedType & """; Actual: """ & TypeName(obj) & """")
    PrePassed "AssertIsType() : Value type is """ & expectedType & """"
    Exit Function
errCondition:
    ErrorProcess
End Function

Public Function AssertIsNotType(expectedType As String, obj)
    On Error GoTo errCondition
    If TypeName(obj) = expectedType Then Call Err.Raise(9022, "", "AssertIsNotType() : Value is the exact type; Expected: """ & expectedType & """; Actual: """ & TypeName(obj) & """")
    PrePassed "AssertIsNotType() : Value type is NOT  """ & expectedType & """"
    Exit Function
errCondition:
    ErrorProcess
End Function

Public Function AssertSingle(collect)
    On Error GoTo errCondition
    If collect.Count <> 1 Then Call Err.Raise(9023, "", "AssertSingle() : The collection contained some items; Collection: " & ToString(collect))
    PrePassed "AssertSingle()"
    Exit Function
errCondition:
    ErrorProcess
End Function

Public Function AssertDistinct(collect)
    On Error GoTo errCondition
    Dim col1, col2, cnt1 As Long, cnt2 As Long
    For Each col1 In collect
        cnt1 = cnt1 + 1
        For Each col2 In collect
            cnt2 = cnt2 + 1
            If cnt1 <> cnt2 Then
                If Equals(col1, col2) Then Call Err.Raise(9023, "", "AssertDistinct() : Duplicate item found; Collection: " & ToString(collect))
            End If
        Next
        cnt2 = 0
    Next
    PrePassed "AssertDistinct()"
    Exit Function
errCondition:
    ErrorProcess
End Function

Public Function AssertHasError()
     If Err.Number <> 0 Then
    Else
        On Error GoTo errCondition
        Call Err.Raise(9005, "", "AssertHasError() : No error raised")
    End If
    Err.Clear
    PrePassed "AssertHasError()"
    Exit Function
errCondition:
    ErrorProcess
End Function

Public Function AssertHasNoError()
    If Err.Number = 0 Then
    Else
        Dim num As Long:    num = Err.Number
        Dim desc As String:  desc = Err.Description
        Dim src As String:     src = Err.Source
        On Error GoTo errCondition
        Call Err.Raise(9006, "", "AssertHasNoError() : Error raised." & _
            " (" & "Number: " & num & _
            ", " & "Description: " & desc & _
            ", " & "Source: " & src & ")")
    End If
    PrePassed "AssertHasNoError()"
    Exit Function
errCondition:
    ErrorProcess
End Function

Public Function ClearErr()
    hasError_ = False
    errNumber_ = 0
    errDescription_ = ""
    errSource_ = ""
    errResult_ = ""
End Function

' Before Using this function: Need to check in option of application, "Trust access to the VBA project object model"
Public Sub CreateRunTests(Optional moduleName As String = "UnitTests", Optional privateUnitTestName As String = "UnitTest")
    Dim i As Long, txt As String, txts, code As String, testFuncs As New Collection, testFuncInfos As Collection, runTestsLine As Long, codeModuleObject As Object
          
    Set codeModuleObject = Application.VBE.ActiveVBProject.VBComponents(moduleName).CodeModule
    txt = codeModuleObject.Lines(1, codeModuleObject.CountOfLines)
    txts = Split(txt, vbCrLf)
    For i = LBound(txts) To UBound(txts)
        If InStr(txts(i), "'[Fact") > 0 Or InStr(txts(i), "'[Theory") > 0 Then
            If InStr(txts(i), "Skip") > 0 Then GoTo nextloop
            Set testFuncInfos = New Collection
            If runTestsLine = 0 Then runTestsLine = i
            testFuncInfos.Add GetFunctionName(txts, i), "Name"
'            If InStr(txts(i), "'[Fact]") > 0 Then
            If InStr(txts(i), "'[Theory]") > 0 Then testFuncInfos.Add GetTheoryData(txts, i), "Data"
            testFuncs.Add testFuncInfos
        End If
nextloop:
    Next i
    code = GetRunTestsCode(testFuncs, privateUnitTestName)
    
    With codeModuleObject
        On Error GoTo nonRunTests
        runTestsLine = .ProcStartLine(RUN_TESTS_NAME, 0)
        .DeleteLines runTestsLine, .ProcCountLines(RUN_TESTS_NAME, 0)
        .InsertLines runTestsLine, code
        Exit Sub
nonRunTests:
        .InsertLines runTestsLine, code
        Exit Sub
    End With
End Sub

Public Function NameOf(testUnitName As String) As UnitTest
    If isSingleTest Then Debug.Print vbCrLf & "--- " & testUnitName & " -------------------------"
    Set NameOf = Me
End Function

' Private functions ------------------------------------------------------------------------
Private Function RunTest(testFunc As Collection, uTest As UnitTest)
On Error Resume Next
    Dim txt As String, startTime As Double, testTimeMs As Long
    startTime = Now + Timer
    ApplicationRun testFunc
    testTimeMs = (Now + Timer - startTime) * 1000
    CopyErr  ' When application run has error, catch
    If hasError_ = False Then
        hasError_ = uTest.HasError
        errNumber_ = uTest.ErrorNumber
        errResult_ = uTest.ErrorResult
    End If
    
    Debug.Print GetResultText(AddResultInfo(testFunc, testTimeMs))
    ClearErr
    uTest.ClearErr
End Function

Private Function AddResultInfo(testFunc As Collection, testTimeMs) As Collection
    Dim resInfos As New Collection
    resInfos.Add GetFunctionStringWithArgument(testFunc), "testName"
    resInfos.Add hasError_ = False, "result"
    resInfos.Add IIf(hasError_, Mid(errResult_, Len(vbCrLf) + 1), ""), "errResult"
    resInfos.Add testTimeMs, "Time"
    resultTexts_.Add resInfos
    Set AddResultInfo = resInfos
End Function

Private Function GetResultText(resInfos As Collection) As String
    GetResultText = IIf(resInfos("result"), "Success", "Failure") & " : " & resInfos("testName") & "  (" & resInfos("Time") & "[ms])" & IIf(resInfos("result"), "", vbCrLf & resInfos("errResult"))
End Function

Private Function PrePassed(msg)
'    errResult_ = errResult_ & vbCrLf & "  - : " & msg 'hiding message at RunTests
    If isSingleTest Then Debug.Print "- Success : " & msg
End Function

Private Function ErrorProcess()
    CopyErr
    If isSingleTest Then Debug.Print "! Failure : " & errDescription_
End Function

Private Function CopyErr()
    hasError_ = IIf(Err.Number > 0, True, hasError_)
    errNumber_ = IIf(Err.Number = 0, errNumber_, Err.Number)
    errDescription_ = IIf(Err.Number = 0, errDescription_, Err.Description)
    errSource_ = IIf(Err.Number = 0, errSource_, Err.Source)
    errResult_ = IIf(Err.Number = 0, errResult_, errResult_ & vbCrLf & "  ! Failure : " & errDescription_ & errSource_)
End Function

Private Function ApplicationRun(testFunc As Collection) As Variant
    Dim p, funcName As String, res
    
    funcName = testFunc(KEY_NAME)
    p = testFunc(KEY_PARAMETERS)
    With Application
    Select Case UBound(p) + 1
        Case 0: res = .Run(funcName)
        Case 1: res = .Run(funcName, p(0))
        Case 2: res = .Run(funcName, p(0), p(1))
        Case 3: res = .Run(funcName, p(0), p(1), p(2))
        Case 4: res = .Run(funcName, p(0), p(1), p(2), p(3))
        Case 5: res = .Run(funcName, p(0), p(1), p(2), p(3), p(4))
        Case 6: res = .Run(funcName, p(0), p(1), p(2), p(3), p(4), p(5))
        Case 7: res = .Run(funcName, p(0), p(1), p(2), p(3), p(4), p(5), p(6))
        Case 8: res = .Run(funcName, p(0), p(1), p(2), p(3), p(4), p(5), p(6), p(7))
        Case 9: res = .Run(funcName, p(0), p(1), p(2), p(3), p(4), p(5), p(6), p(7), p(8))
        Case 10: res = .Run(funcName, p(0), p(1), p(2), p(3), p(4), p(5), p(6), p(7), p(8), p(9))
        Case 11: res = .Run(funcName, p(0), p(1), p(2), p(3), p(4), p(5), p(6), p(7), p(8), p(9), p(10))
        Case 12: res = .Run(funcName, p(0), p(1), p(2), p(3), p(4), p(5), p(6), p(7), p(8), p(9), p(10), p(11))
        Case 13: res = .Run(funcName, p(0), p(1), p(2), p(3), p(4), p(5), p(6), p(7), p(8), p(9), p(10), p(11), p(12))
        Case 14: res = .Run(funcName, p(0), p(1), p(2), p(3), p(4), p(5), p(6), p(7), p(8), p(9), p(10), p(11), p(12), p(13))
        Case 15: res = .Run(funcName, p(0), p(1), p(2), p(3), p(4), p(5), p(6), p(7), p(8), p(9), p(10), p(11), p(12), p(13), p(14))
        Case 16: res = .Run(funcName, p(0), p(1), p(2), p(3), p(4), p(5), p(6), p(7), p(8), p(9), p(10), p(11), p(12), p(13), p(14), p(15))
        Case 17: res = .Run(funcName, p(0), p(1), p(2), p(3), p(4), p(5), p(6), p(7), p(8), p(9), p(10), p(11), p(12), p(13), p(14), p(15), p(16))
        Case 18: res = .Run(funcName, p(0), p(1), p(2), p(3), p(4), p(5), p(6), p(7), p(8), p(9), p(10), p(11), p(12), p(13), p(14), p(15), p(16), p(17))
        Case 19: res = .Run(funcName, p(0), p(1), p(2), p(3), p(4), p(5), p(6), p(7), p(8), p(9), p(10), p(11), p(12), p(13), p(14), p(15), p(16), p(17), p(18))
        Case 20: res = .Run(funcName, p(0), p(1), p(2), p(3), p(4), p(5), p(6), p(7), p(8), p(9), p(10), p(11), p(12), p(13), p(14), p(15), p(16), p(17), p(18), p(19))
        Case 21: res = .Run(funcName, p(0), p(1), p(2), p(3), p(4), p(5), p(6), p(7), p(8), p(9), p(10), p(11), p(12), p(13), p(14), p(15), p(16), p(17), p(18), p(19), p(20))
        Case 22: res = .Run(funcName, p(0), p(1), p(2), p(3), p(4), p(5), p(6), p(7), p(8), p(9), p(10), p(11), p(12), p(13), p(14), p(15), p(16), p(17), p(18), p(19), p(20), p(21))
        Case 23: res = .Run(funcName, p(0), p(1), p(2), p(3), p(4), p(5), p(6), p(7), p(8), p(9), p(10), p(11), p(12), p(13), p(14), p(15), p(16), p(17), p(18), p(19), p(20), p(21), p(22))
        Case 24: res = .Run(funcName, p(0), p(1), p(2), p(3), p(4), p(5), p(6), p(7), p(8), p(9), p(10), p(11), p(12), p(13), p(14), p(15), p(16), p(17), p(18), p(19), p(20), p(21), p(22), p(23))
        Case 25: res = .Run(funcName, p(0), p(1), p(2), p(3), p(4), p(5), p(6), p(7), p(8), p(9), p(10), p(11), p(12), p(13), p(14), p(15), p(16), p(17), p(18), p(19), p(20), p(21), p(22), p(23), p(24))
        Case 26: res = .Run(funcName, p(0), p(1), p(2), p(3), p(4), p(5), p(6), p(7), p(8), p(9), p(10), p(11), p(12), p(13), p(14), p(15), p(16), p(17), p(18), p(19), p(20), p(21), p(22), p(23), p(24), p(25))
        Case 27: res = .Run(funcName, p(0), p(1), p(2), p(3), p(4), p(5), p(6), p(7), p(8), p(9), p(10), p(11), p(12), p(13), p(14), p(15), p(16), p(17), p(18), p(19), p(20), p(21), p(22), p(23), p(24), p(25), p(26))
        Case 28: res = .Run(funcName, p(0), p(1), p(2), p(3), p(4), p(5), p(6), p(7), p(8), p(9), p(10), p(11), p(12), p(13), p(14), p(15), p(16), p(17), p(18), p(19), p(20), p(21), p(22), p(23), p(24), p(25), p(26), p(27))
        Case 29: res = .Run(funcName, p(0), p(1), p(2), p(3), p(4), p(5), p(6), p(7), p(8), p(9), p(10), p(11), p(12), p(13), p(14), p(15), p(16), p(17), p(18), p(19), p(20), p(21), p(22), p(23), p(24), p(25), p(26), p(27), p(28))
        Case 30: res = .Run(funcName, p(0), p(1), p(2), p(3), p(4), p(5), p(6), p(7), p(8), p(9), p(10), p(11), p(12), p(13), p(14), p(15), p(16), p(17), p(18), p(19), p(20), p(21), p(22), p(23), p(24), p(25), p(26), p(27), p(28), p(29))
        Case Else: Err.Raise 9001, "", "(UnitTest) Registered Test has Error, Arguments are over 30."
    End Select
    End With
    ApplicationRun = res
End Function

Private Function Equals(Expected, Actual) As Boolean
    Dim result
    Equals = False
    If VarType(Expected) <> VarType(Actual) Then
        If IsNumeric(Expected) And IsNumeric(Actual) Then Else _
            Err.Raise 9101, "", "Var type is not equal. expected is " & TypeName(Expected) & ", actual is " & TypeName(Actual)
    End If
    If VarType(Expected) < vbArray Then
        If VarType(Expected) = vbObject Then
            If TryEqual(Expected, Actual, result) Then
                If result Then Equals = True
            Else
                Err.Raise 9102, "", "Cannot evaluate Equal(). type name  """ & TypeName(Expected) & """"
            End If
        Else
            If Expected = Actual Then Equals = True
        End If
    Else
        
    End If
End Function

Private Function TryEqual(a, b, result) As Boolean
    On Error GoTo errEqual
    result = a.Equals((b))
    TryEqual = True
    Exit Function
errEqual:
    TryEqual = False
End Function

Private Function ToString(val) As String
    Dim txt As String
    If VarType(val) = vbObject Then
        On Error GoTo errToString
        If TypeName(val) = "Collection" Then
            Dim col As Collection: Set col = val
            Dim c
            txt = "["
            For Each c In col: txt = txt & CStr(c) & ",": Next
            txt = Left(txt, Len(txt) - 1) & "]"
            ToString = IIf(col.Count > 0, txt, "")
            Exit Function
        Else
            ToString = val.ToString()
        End If
errToString:
    Else
        If VarType(val) < vbArray Then
            ToString = CStr(val)
        Else
            Dim ar
            txt = "["
            For Each ar In val: txt = txt & CStr(ar) & ",": Next
            txt = Left(txt, Len(txt) - 1) & "]"
            ToString = txt
        End If
    End If
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

Private Function GetFunctionName(texts, startLine As Long) As String
    Dim i As Long, funcName As String

    For i = startLine To UBound(texts)
        If InStr(LCase(texts(i)), "sub ") > 0 Then funcName = texts(i): Exit For
        If InStr(LCase(texts(i)), "function ") > 0 Then funcName = texts(i): Exit For
    Next i
    funcName = Replace(Replace(Replace(Replace(funcName, "Sub ", ""), "Function ", ""), "Property Get ", ""), "Property Set ", "")
    funcName = Left(funcName, InStr(funcName, "(") - 1)
    funcName = Replace(funcName, " ", "")
    GetFunctionName = funcName
End Function

Private Function GetTheoryData(texts, startLine As Long) As Collection
    Dim i As Long, funcName As String, theoryDatas As New Collection, dataName, memberData, memberDatas As Collection

    For i = startLine To UBound(texts)
        If InStr(LCase(texts(i)), "[inlinedata(") > 0 Then
            theoryDatas.Add Replace(Replace(Mid(texts(i), Len("'[inlinedata(") + 1), " ", ""), ")]", "")
        End If
        If InStr(LCase(texts(i)), "[memberdata(") > 0 Then
            dataName = Replace(Replace(Mid(texts(i), Len("'[memberdata(") + 1), " ", ""), ")]", "")
            Set memberDatas = Application.Run(dataName)
            For Each memberData In memberDatas
                theoryDatas.Add Join(memberData, ",")
            Next
        End If
        If InStr(LCase(texts(i)), "sub ") > 0 Then funcName = texts(i): Exit For
        If InStr(LCase(texts(i)), "function ") > 0 Then funcName = texts(i): Exit For
    Next i
    
    Set GetTheoryData = theoryDatas
End Function

Private Function GetRunTestsCode(testFuncs As Collection, privateTestUnitName As String) As String
    Dim s, code As String, codes As New Collection, testFuncInfos As Collection

    With codes
        .Add "Sub " & RUN_TESTS_NAME & "()"
        .Add "   Dim test As New UnitTest"
        .Add ""
        For Each testFuncInfos In testFuncs
            If testFuncInfos.Count = 1 Then
                .Add "    test.RegisterTest """ & testFuncInfos.Item("Name") & """"
            Else
                For Each s In testFuncInfos.Item("Data")
                    .Add "    test.RegisterTest """ & testFuncInfos.Item("Name") & """, " & s
                Next
            End If
        Next
        .Add ""
        .Add "    test.RunTests " & privateTestUnitName
        .Add "End Sub"
    End With
    For Each s In codes: code = code & vbCrLf & s:       Next
    
    GetRunTestsCode = code
End Function
