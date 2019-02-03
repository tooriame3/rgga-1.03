function onOpen() {
  "use strict";
  add_menu();
  add_menuVUP();
  sheetHide();
}

function add_menu() {
  "use strict";
  // 現在アクティブなスプレッドシートを取得
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  //var spreadsheetObj = SpreadsheetApp.getActiveSpreadsheet();
  var menuList       = [];
  menuList.push({
    name : "１：CSVファイルで処理",
    functionName : "openDialog3"
  });
  menuList.push({
    name : "２：リンクの変更",
    functionName : "setMyID"
  });
    menuList.push({
    name : "３：サマリーメールの送信",
    functionName : "summaryMail"
  });
  menuList.push({
    name : "４：自動更新の動作テスト",
    functionName : "mainTest"
  });
  menuList.push({
    name : "５：全取引履歴の保存",
    functionName : "exportcsv"
  });
  menuList.push({
    name : "６：データ修正を修復",
    functionName : "correctSheetData2"
  });
  menuList.push({
    name : "７：作業用シートの表示",
    functionName : "sheetShow"
  });
  menuList.push({
    name : "８：作業用シートの非表示",
    functionName : "sheetHide"
  });
  menuList.push({
    name : "９：シート全データの削除",
    functionName : "clearMyData"
  });
  menuList.push({
    name : "＊：ＨＰから成績を取得",
    functionName : "getIkinariData2"
  });
  menuList.push({
    name : "Ｕ：ユーザーの独自処理",
    functionName : "syori1"
  });
  menuList.push({
    name : "Ｇ：取引グラフの日数変更",
    functionName : "daysInput"
  });
  menuList.push({
    name : "Ｅ：インポートエラーチエック",
    functionName : "errCheck"
  });
  ss.addMenu("ユーティリテイ", menuList);
}

function syori1(){
}

function setMyID(){
  //createTrigger2(); //このとき、ついでにやっておく
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  //スプレッドシートのidを取得し
  var getid = ss.getId();
  var sh1 = ss.getSheetByName('取引履歴');
  sh1.showSheet();
  ss.setActiveSheet(sh1, true);
  var strformula1 ="=IMPORTRANGE(\"${getid}\",\"入力と訂正!B2:P\")".replace("${getid}",getid);
  sh1.getRange('B2').setFormula(strformula1);
  //
  var sh2 = ss.getSheetByName('修復作業用');
  var strformula2 ="=IMPORTRANGE(\"${getid}\",\"履歴の修復!C2:L\")".replace("${getid}",getid);
  sh2.getRange('C2').setFormula(strformula2);
  var strformula3 ="=IMPORTRANGE(\"${getid}\",\"履歴の修復!O2:P\")".replace("${getid}",getid);
  sh2.getRange('P2').setFormula(strformula3);
  //
  PopupStartSetMyID();
  var ok = false;
  while (!ok){    
    try{
      var a=sh1.getRange('A2').getValue();
      ok = true;
      PopupEndSetMyID();
    }catch(e){
      PopupEndSetMyID2();
    }
  }
}

function errCheck(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('取引履歴');
  try {
    var p=sh.getRange('B2').getValue();
    var str = String(p);
    Browser.msgBox('インポートエラーはありません。');  
  } catch (e) {
    var error = e;
    //「インポートエラー」が含まれている場合
    if ( error.message.match(/インポートエラー/)) {
      Browser.msgBox('インポートエラーです。');
    } else {
      Browser.msgBox('多分、計算途中です。');
    }
  }
}

function summaryMail(){
  PopupStartSummaryMail();
  sendMail3();
  PopupEndSummaryMail();
}

function mainTest(){
  createTrigger();
  main();
  var triggerKey = "triggerNext";    //トリガーIDを保存するときに使用するkey
  var funcName="runFromTriger_Next";
  setTrigger(triggerKey, funcName)
  PopupEnd();
}

function onLinkA(){ 
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh1 = ss.getSheetByName('取引履歴');
  var getid = ss.getId();
  var strformula1 ="=IMPORTRANGE(\"${getid}\",\"入力と訂正!B2:P\")".replace("${getid}",getid);
  sh1.getRange('B2').setFormula(strformula1);
}

function offLinkA(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh1 = ss.getSheetByName('取引履歴');
  sh1.getRange('B2').clearContent();
}

function onLinkB(){ 
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var getid = ss.getId();
  var sh2 = ss.getSheetByName('修復作業用');
  var strformula2 ="=IMPORTRANGE(\"${getid}\",\"履歴の修復!C2:L\")".replace("${getid}",getid);
  sh2.getRange('C2').setFormula(strformula2);
  var strformula3 ="=IMPORTRANGE(\"${getid}\",\"履歴の修復!O2:P\")".replace("${getid}",getid);
  sh2.getRange('P2').setFormula(strformula3);
}

function offLinkB(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh2 = ss.getSheetByName('修復作業用');
  sh2.getRange('C2').clearContent();
  sh2.getRange('P2').clearContent();
}

function openDialog3() {
  "use strict";
  var html = HtmlService.createHtmlOutputFromFile('index');
  SpreadsheetApp.getUi() // DocumentAppやFormApp等目的に合わせたオブジェクトを使う
  .showModelessDialog(html, '岡三ＨＰやメール添付のCSVファイルで処理');
}

//ユーティリティの置換
function uploadA(e) {
  var file = e.file;
  if (file.length == 0){
    endMsg2();
    return;
  }
  else {
    PopupStartA();
  };
  
  //offLinkA();
  var csvData = Utilities.parseCsv(file.getBlob().getDataAsString("Shift_JIS"));
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  var sheetName ="入力と訂正"; 
  var sheet = ss.getSheetByName(sheetName);
  
  clearMySheet(sheetName);
  //決済のみ抽出
  data = myFilter2(csvData);
  sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
  //
  sortSheet(sheetName);
  //タイトル行が２行目にくるので削除
  if (sheet.getRange('A2').getValue()=="約定日"){
    sheet.deleteRow(2);
  };
  
  //onLinkA();
  PopupEnd();
  var sh1 = ss.getSheetByName('チエック１');
  sh1.showSheet();
  ss.setActiveSheet(sh1, true);
}


//ユーティリティの修正
function uploadB(e) {
  var file = e.file;
  if (file.length == 0){
    endMsg2();
    return;
  }
  else {
    PopupStartB();
  };
  memoValueCopy(); //追加
  
  var csvData = Utilities.parseCsv(file.getBlob().getDataAsString("Shift_JIS"));
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  var sheetName ="履歴の修復"; 
  var sheet = ss.getSheetByName(sheetName);

  //offLinkA();
  //offLinkB();
  clearMySheet(sheetName);  
  //決済のみ抽出
  data = myFilter2(csvData);
  sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
  //タイトル行が２行目にくるので削除
  if (sheet.getRange('A2').getValue()=="約定日"){
    sheet.deleteRow(2);
  };

  //onLinkB();
  //onLinkA();
  correctSheetData();//修復.gs
  PopupEnd();
  var sh1 = ss.getSheetByName('チエック１');
  sh1.showSheet();
  ss.setActiveSheet(sh1, true);
}

//メモと戦略の値コピー
function memoValueCopy(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh1 = ss.getSheetByName("入力と訂正");
  var sh2 = ss.getSheetByName("メモと戦略の値");
  clearMySheet("メモと戦略の値")
  var last_row = getLastRowNumber_ColumnA("入力と訂正");
  if (last_row ==1){
    last_row = 2;
  };
  var data = [[]];  
  data = sh1.getRange(2, 16, last_row, 19).getValues();//１行目なしP:S
  sh2.getRange(2, 1, data.length, data[0].length).setValues(data);
  ss.setNamedRange('メモと戦略の値_処理', sh2.getRange(2, 1, data.length, data[0].length));
}

//データ修復の再実行
function  correctSheetData2() {
  //var ss = SpreadsheetApp.getActiveSpreadsheet();
  //var sh = ss.getSheetByName("メモと戦略の値");
  //sh.showSheet();
  //ss.setActiveSheet(sh, true);
  if (true){
    correctSheetData();
    PopupEnd();
  }else{
    endMsgG2();
  }
}

/*/取り引き列col=5が決済であるものの配列を返す
function myFilter2(data) {
  var result = data.filter( function( value ) {
    return value[4] == '決済';
  })
  return result;
  Logger.log( result );
}
/*/
function myFilter2(data) {
  var result = data;
  return result;
}

function sheetHide() {
  var sheetName1 = "取引履歴";
  var sheetName2 = "修復作業用";
  var sheetName3 = "履歴の修復";
  var sheetName4 = "対応表";
  var sheetName5 = "本家成績";
  var sheetName6 = "メモと戦略の値";
  var sheetName7 = "セッションラベル";
  var sheetName8 = "終値データ";

  var sheet1 = SpreadsheetApp.getActive().getSheetByName(sheetName1);
  sheet1.hideSheet();
  var sheet2 = SpreadsheetApp.getActive().getSheetByName(sheetName2);
  sheet2.hideSheet();
  var sheet3 = SpreadsheetApp.getActive().getSheetByName(sheetName3);
  sheet3.hideSheet();
  var sheet4 = SpreadsheetApp.getActive().getSheetByName(sheetName4);
  sheet4.hideSheet();
  var sheet5 = SpreadsheetApp.getActive().getSheetByName(sheetName5);
  sheet5.hideSheet();
  var sheet6 = SpreadsheetApp.getActive().getSheetByName(sheetName6);
  sheet6.hideSheet();
  var sheet7 = SpreadsheetApp.getActive().getSheetByName(sheetName7);
  sheet7.hideSheet();
  var sheet8 = SpreadsheetApp.getActive().getSheetByName(sheetName8);
  sheet8.hideSheet();  
}

function sheetShow() {
  var sheetName1 = "取引履歴";
  var sheetName2 = "修復作業用";
  var sheetName3 = "履歴の修復";
  var sheetName4 = "対応表";
  var sheetName5 = "本家成績";
  var sheetName6 = "メモと戦略の値";
  var sheetName7 = "セッションラベル";
  var sheetName8 = "終値データ";
    
  var sheet1 = SpreadsheetApp.getActive().getSheetByName(sheetName1);
  sheet1.showSheet();
  var sheet2 = SpreadsheetApp.getActive().getSheetByName(sheetName2);
  sheet2.showSheet();
  var sheet3 = SpreadsheetApp.getActive().getSheetByName(sheetName3);
  sheet3.showSheet();
  var sheet4 = SpreadsheetApp.getActive().getSheetByName(sheetName4);
  sheet4.showSheet();
  var sheet5 = SpreadsheetApp.getActive().getSheetByName(sheetName5);
  sheet5.showSheet();
  var sheet6 = SpreadsheetApp.getActive().getSheetByName(sheetName6);
  sheet6.showSheet();
  var sheet7 = SpreadsheetApp.getActive().getSheetByName(sheetName7);
  sheet7.showSheet();
  var sheet8 = SpreadsheetApp.getActive().getSheetByName(sheetName8);
  sheet8.showSheet();  
}

//main用
function uploadMain(file) {
  var csvData = Utilities.parseCsv(file.getBlob().getDataAsString("Shift_JIS"));
  // 現在アクティブなスプレッドシートを取得
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  var sheetName ="履歴の修復";
  // そのスプレッドシートにある"sheetName"という名前のシートを取得  
  var sheet = ss.getSheetByName(sheetName);

  //offLinkA();
  //offLinkB();
  clearMySheet(sheetName);  
  //決済のみ抽出はしていない
  data = myFilter2(csvData);
  //data=csvData
  try{
    sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
      //タイトル行が２行目にくるので削除
      if (sheet.getRange('A2').getValue()=="約定日"){
        sheet.deleteRow(2);
      };
    //onLinkB();
    //onLinkA();
    correctSheetData();//修復.gs
  }catch(e){
    //onLinkB();
    //onLinkA();
  }
}

function　clearMyData(){
  clearMySheet('入力と訂正');
}
