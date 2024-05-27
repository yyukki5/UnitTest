# UnitTest
For the unit test in VBA.  
Only import UnitTest.cls, You can use simply unit tests.
src
 - UnitTest.cls  <-- Import this file.
 - UnitTests.bas <-- Sample code.

## Sample
![img](https://github.com/yyukki5/UnitTest/assets/136491951/6301d0ff-57b7-452e-8c91-f5adb23a2ffb)

## Feature
 - Assert*() functions validate test result.
    - AssertTrue(), AssertFalse(), AssertEqual(), AssertNotEqual(), AssertSame(), AssertNotSame(), AssertContains(), AssertStartsWith(), AssertEmpty(), AssertNothing(), AssertIsType(), AssertSingle(), AssertDistinct(), AssertHasError(), AssertHasNoError(), ...
 - Run single test method by Run macro(F5), you can validate by Assert*() functions.
 - RunTests() function run registered test by RegisterTest().
    - Registered test name is called from UnitTest.cls by Application.Run().
 - CreateRunTests() function can create RunTests().
    - RunTests() method has been created, based on header ( comment as "[Fact]", "[Theory]") of test function.
    - [Fact], [Theory] can use "Skip".
    - [Theory] can use "InlineData()", "MemberData()".
    - Need to check VBA Project Object Model Access.
 - Results property has test result (Test name, result, description, time).
 
Before using, switch error trapping in VBE > Tools > Options > General > Error Trapping > "Break on Unhandled Errors".

### Caution
Character code is CRLF, not LF.  
Windows user can use by download and directly import. Non Windows users, please confirm character code.


## Japanese Note
ArrayExを作った時の副産物レポジトリ  
場当たり的に作ったのですが、便利だったので誰かの役に立つといいなと思っています  

主な仕様
 - UnitTest.cls をプロジェクトに入れるだけで使えます。クラス一つだけで実装しているのが気に入っています  
 - テストコードは標準モジュールに書きます 
 - テストメソッドをデバッグ実行することで、テスト単体で実行することができます
 - 複数のテストメソッドを一度に実行する場合は、RegisterTest() でテストメソッドを登録してRunTests()で実行します
   - CreateRunTests()を用いてRunTests()を作成することができます
    - テストメソッドとして識別する場合は、テストメソッド名の直前にコメントで直前に[Fact]や [Theory] と付けて下さい
    - [Theory] を用いた場合は "InlineData()", "MemberData()" を用いてテストメソッドの引数を指定できます
    - アプリケーションの"VBA プロジェクトオブジェクトモデルへのアクセスを信頼する"にチェックを入れた状態で使用してください
 - テストの結果はDebug.Printされます。また Results プロパティでも取得できます
 - 使用するときは、ツール>オプション>全般>エラー トラップ>"エラー処理対象外のエラーで中断" に設定してご使用ください

<br>
 文字コードが CRLFになるように保存しています。Windowsユーザーはダウンロードしてインポートすればそのまま使用できます  
 Windows以外のOSを利用されている方は文字コードをご確認ください。  
 (Fetchすれば大丈夫だと思いますが...念のため)
