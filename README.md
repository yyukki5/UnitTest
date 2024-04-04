# UnitTest
For the unit test in VBA.  
You only import UnitTest.cls, You can use simply unit tests.

- UnitTest : src\UnitTest.xlsm
    - UnitTest.cls  
    - UnitTests.bas <-- Sample code for tests.

## Feature
 - Assert*() functions validate test result.
    - AssertTrue(), AssertFalse(), AssertEqual(), AssertNotEqual(), AssertHasError(), AssertHasNoError()
 - RunTests() function run registered test by RegisterTest().
    - Registered test name is called from UnitTest.cls by Application.Run().
 - Run a test method, you can validate assert*() functions.

Before using, switch error trapping in VBE > Tools > Options > General > Error Trapping > "Break on Unhandled Errors".

## Sample Code
Write in Modules.
~~~
Private test_ As New UnitTest '<-- (1)Need 

Sub Sample_RunTests()
    Dim test As New UnitTest
    test.RegisterTest "Test_Test" ' Register test method name by string
    test.RunTests test_  ' <--- Need argument (1)
End Sub

Sub Test_Test()
    test_.AssertEqual 1, 1
    test_.AssertEqual 1, 2
    test_.AssertNotEqual 1, 2
    test_.AssertNotEqual 1, 1

    On Error Resume Next '<-- (2)Need for AssertHasError, AssertHasNoError
    Err.Raise 9001
    test_.AssertHasError
    Err.Raise 9001, "Sample", "This is sample Error."
    test_.AssertHasNoError
    On Error GoTo 0    
End Sub
~~~
When Sample_RunTests() has run, Show below in immediate window.
~~~
--- Start tests (2024/04/05 1:23:45.678) ---
NG: Test_Test
  - NG: Should be Equal. expected is 1, actual is 2
  - NG: Should be NOT Equal. expected is 1, actual is 1
  - NG: Should have NO error, but has error.
      - Number: 9001
      - Description: This is sample Error.
      - Source: Sample
  - NG: Should have error, but has NO error.
--- Finish tests (2024/04/05 1:23:45.876) ---
~~~
And also you can run Test_Test(), get a like result.



## Caution
Character code is CRLF, not LF.  
Windows user can use by download and directly import. Non Windows users, please confirm character code.


## Japanese Note
ArrayExを作った時の副産物レポジトリ  
場当たり的に作ったのですが、便利だったので誰かの役に立つといいなと思っています  

 - UnitTest.cls をプロジェクトに入れるだけで使えます。クラス一つだけで実装しているのが気に入っています  
 - テストコードは標準モジュールに書きます  
 （Private field に変数を準備したり(1), On Error Resume Nextを準備したり(2)と書き方に少しクセがあります）  
 - RegisterTest() で、テストメソッドを登録して、RunTests()で実行します
    - (1) のPrivate field を引数として与えてください
 - テストメソッド単体で実行することもできます。
    - メソッドを作っているときはテストメソッドをF5キーで実行して確かめて修正してを繰り返して、全体をテストしたいときはRunTests()のメソッドを実行して使っていました
 - 使用するときは、ツール>オプション>全般>エラー トラップ>"エラー処理対象外のエラーで中断" に設定してご使用ください

 文字コードが CRLFになるように保存しています。Windowsユーザーはダウンロードしてインポートすればそのまま使用できます  
 Windows以外のOSを利用されている方は文字コードをご確認ください。  
 (Fetchすれば大丈夫だと思いますが...念のため)



