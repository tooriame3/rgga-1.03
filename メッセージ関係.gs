// アクティブなスプレッドシートにトーストを表示
function PopupStartMain(){
  SpreadsheetApp.getActiveSpreadsheet().toast(
    '[ 定期実行される処理 ]' +'　　　　' +
    '「 DO NOT OPEN 」のメールを確認中です。'
    ,'スクリプト実行中', -1);
}
function PopupStartMain2(){
  SpreadsheetApp.getActiveSpreadsheet().toast(
    '[ 定期実行される処理 ]' +'　　　　' +
    '「 終値データ 」の更新中です。'
    ,'スクリプト実行中', -1);
}
function PopupStartA(){
  SpreadsheetApp.getActiveSpreadsheet().toast(
    '[ CSVデータによる全置換 ]' +'　　' +
    '終了までお待ちください。'
    ,'スクリプト実行中', -1);
}

function PopupStartB(){
  SpreadsheetApp.getActiveSpreadsheet().toast(
    '[ CSVデータによる修正 ]' +'　　　' +
    '終了までお待ちください。'
    ,'スクリプト実行中', -1);
}

//使用していない
function PopupStartC(){
  SpreadsheetApp.getActiveSpreadsheet().toast(
    '[ CSVデータによる修復 ]' +'　　　' +
    '終了までお待ちください。'
    ,'スクリプト実行中', -1);
}

//使用するCSVの期間の表示
function PopupStartC1(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh2 = ss.getSheetByName('修復作業用');
  var t1 = sh2.getRange('U2').getValue();
  t1=Utilities.formatDate(t1, "JST","yyyy/MM/dd HH:mm");
  var t2 = sh2.getRange('U3').getValue();
  t2=Utilities.formatDate(t2, "JST","yyyy/MM/dd HH:mm");
  SpreadsheetApp.getActiveSpreadsheet().toast(
    '[ CSVデータによる修正 ]' +'　　　' +
    '（期間）'+
    t1+
    '      ～'+t2
    ,'スクリプト実行中', -1);
}

function PopupStartD(){
  SpreadsheetApp.getActiveSpreadsheet().toast(
    '[ 全取引履歴の保存 ]' +'　　　　　' +
    'ダウンロードの準備中です。'
    ,'スクリプト実行中', -1);
}

function PopupStartSummaryMail(){
  SpreadsheetApp.getActiveSpreadsheet().toast(
    '[ サマリーメールの送信 ]' +'　　　' +
    'メールを作成しています。'
    ,'スクリプト実行中', -1);
}

function PopupStartSetMyID(){
  SpreadsheetApp.getActiveSpreadsheet().toast(
    '[ リンクの変更 ]' +'　　　　　　　' +
    'マニュアルに従ってください。'
    ,'スクリプト実行中', -1);
}

function PopupStartG(){
  SpreadsheetApp.getActiveSpreadsheet().toast(
    '[ データ追加の修復 ]' +'　　　　　' +
    '終了までお待ちください。'
    ,'スクリプト実行中', -1);
}

//共通
function PopupEnd(){
  SpreadsheetApp.getActiveSpreadsheet().toast(
    '処理が終わりました。' +'　　  　　' +
    '再計算を実行します。集計表の再作成には時間がかかります。'
    ,'スクリプトの終了', 10);
}

function PopupEndMain(){
  SpreadsheetApp.getActiveSpreadsheet().toast(
    '[ 定期実行される処理 ]' +'　　　　' +
    '次の処理の準備中です。'
    ,'スクリプトの終了', 20);
}
function PopupEndMain2(){
  SpreadsheetApp.getActiveSpreadsheet().toast(
    '[ 定期実行される処理 ]' +'　　　　' +
    'ユーザー処理があれば、数分後に実行します。'
    ,'スクリプトの終了', -1);
}
function PopupEndMain3(){
  SpreadsheetApp.getActiveSpreadsheet().toast(
    '[ 定期実行される処理 ]' +'　　　　' +
    '全ての処理が終わりました。'
    ,'スクリプトの終了', 5);
}


function PopupEndD(){
  SpreadsheetApp.getActiveSpreadsheet().toast(
    '[ 全取引履歴の保存 ]' +'　　　　　' +
    'ダウンロードの準備中です。'
    ,'スクリプト実行中', -1);
}

function PopupEndSummaryMail(){
  SpreadsheetApp.getActiveSpreadsheet().toast(
    '[ サマリーメールの送信 ]' +'　　　' +
    'メールを送信しました。'
    ,'スクリプトの終了', 5);
}

function PopupEndSetMyID(){
  SpreadsheetApp.getActiveSpreadsheet().toast(
    '[ リンクの変更 ]' +'　　　　　　　' +
    'リンクが正常にできました。'
    ,'スクリプトの終了', 10);
}

function PopupEndSetMyID2(){
  SpreadsheetApp.getActiveSpreadsheet().toast(
    '[ リンクの変更 ]' +'　　　　　　　' +
    'アクセスが未許可または再計算中です。'
    ,'スクリプトの終了', 10);
}

function PopupEndmainTest(){
  SpreadsheetApp.getActiveSpreadsheet().toast(
    '[ サマリーメールの作成 ]' +'　　　' +
    'サマリーメールが受信できたら動作テスト完了です。数分かかります。'
    ,'スクリプトの継続', 90);
}
//デバッグ用---------------------
function Popup1(){
  SpreadsheetApp.getActiveSpreadsheet().toast(
    '--1---　' +
    ''
    ,'スクリプト実行中', -1);
}
function Popup2(){
  SpreadsheetApp.getActiveSpreadsheet().toast(
    '--2---　' +
    ''
    ,'スクリプト実行中', -1);
}
function Popup3(){
  SpreadsheetApp.getActiveSpreadsheet().toast(
    '--3---　' +
    ''
    ,'スクリプト実行中', -1);
}
function Popup4(){
  SpreadsheetApp.getActiveSpreadsheet().toast(
    '--4---　' +
    ''
    ,'スクリプト実行中', -1);
}
function Popup5(){
  SpreadsheetApp.getActiveSpreadsheet().toast(
    '--5---　' +
    ''
    ,'スクリプト実行中', -1);
}

//********大部分はPopupに変更

function startMsg() {
  Browser.msgBox('処理が始まりました。計算終了までお待ちください。');
}
function startMsgA() {
  Browser.msgBox('CSVデータで置き換えます。計算終了までお待ちください。');
}
function startMsgB() {
  Browser.msgBox('CSVデータを追加します。計算終了までお待ちください。');
}
function startMsgC() {
  Browser.msgBox('修復処理が始まりました。CSVの期間外は修復しません。');
}
function startMsgD() {
  Browser.msgBox('CSVファイルへの変換が終わるまで、しばらくお待ちください。');
}
function endMsg() {
  Browser.msgBox('処理が終わりました。シートを再計算します。');
}
function endMsg2() {
  Browser.msgBox('ファイルが空でした。');
}
function msg2(){
  Browser.msgBox(
    "修復処理が終わりました。シートの再計算が終わったら（右上の計算バーが消えます） \\n\\n" +
    "岡三のＨＰにある「年間取引損益照会」と、シートの本年合計とを比較し、違っている場合は、 \\n" +
    "「使用方法」に記された作者まで、ご連絡くだされば幸いです。"
  );
}
function endMsgG2(){
  Browser.msgBox(
    "今の状態でデータ修復の再実行を行っても、同じ結果になってしまいます。 \\n\\n" +
    "シートにあるデータをすべて含むような期間を指定したCSVを岡三ＨＰからダウンロードして、 \\n" +
    "「１：CSVファイルで処理」からもう一度、修復を実行してみてください。"
  );
}


/*インプットボックスを表示させてテキスト入力する*/
function daysInput() {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheetName1 ='終値データ'; 
  var sheet1 = spreadsheet.getSheetByName(sheetName1);
 
  //inputboxを表示して名前を入力・daysへ格納
  var days = Browser.inputBox('描画日数を1～5(半角)を入力してください');
  
  //セルに日数を入力
  if ((days==1)+(days==2)+(days==3)+(days==4)+(days==5)){
    var rng = sheet1.getRange("Q14").setValue(days);
  };
  //
  SpreadsheetApp.getActiveSpreadsheet().toast(
    '上部に表示される「スクリプトを実行中しています・・・」が消えるまで、お待ちください。 '+
    '------------------------------------------------'
    ,'スクリプト実行中', 30);
  //
  changeAxes();
}

