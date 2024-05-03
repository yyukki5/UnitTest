# UnitTest
For the unit test in VBA.  
Only import UnitTest.cls, You can use simply unit tests.
src
 - UnitTest.cls  <-- Import this file.
 - UnitTests.bas <-- Sample code.

## Sample
![img](https://github.com/yyukki5/UnitTest/assets/136491951/a21f2b28-3da7-4683-8973-899eff5d83e4)

## Feature
 - Assert*() functions validate test result.
    - AssertTrue(), AssertFalse(), AssertEqual(), AssertNotEqual(), AssertHasError(), AssertHasNoError()
 - Run single test method by Run macro(F5), you can validate by Assert*() functions.
 - RunTests() function run registered test by RegisterTest().
    - Registered test name is called from UnitTest.cls by Application.Run().
 - CreateRunTests() function can create RunTests().
    - RunTests() method has been created, based on header ( comment as "[Fact]", "[Theory]") of test function.
    - [Fact], [Theory] can use "Skip".
    - [Theory] can use "InlineData()", "MemberData".
 

Before using, switch error trapping in VBE > Tools > Options > General > Error Trapping > "Break on Unhandled Errors".

### Caution
Character code is CRLF, not LF.  
Windows user can use by download and directly import. Non Windows users, please confirm character code.


## Japanese Note
ArrayExを作った時の副産物レポジトリ  
場当たり的に作ったのですが、便利だったので誰かの役に立つといいなと思っています  

 - UnitTest.cls をプロジェクトに入れるだけで使えます。クラス一つだけで実装しているのが気に入っています  
 - テストコードは標準モジュールに書きます  
 （Private field に変数を準備したり(1), On Error Resume Nextを準備したり(2)と書き方に少しクセがあります）  
 - テストメソッド単体で実行するできます
 - 複数のテストメソッドを一度に実行する場合は、RegisterTest() でテストメソッドを登録してRunTests()で実行します
    - (1) のPrivate field を引数として与えてください
    - メソッドを作っているときはテストメソッドをF5キーで実行して確かめて修正してを繰り返して、全体をテストしたいときはRunTests()のメソッドを実行して使っていました
 - 使用するときは、ツール>オプション>全般>エラー トラップ>"エラー処理対象外のエラーで中断" に設定してご使用ください

 文字コードが CRLFになるように保存しています。Windowsユーザーはダウンロードしてインポートすればそのまま使用できます  
 Windows以外のOSを利用されている方は文字コードをご確認ください。  
 (Fetchすれば大丈夫だと思いますが...念のため)
