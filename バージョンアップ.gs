function add_menuVUP() {
  "use strict";
  // 現在アクティブなスプレッドシートを取得
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  //var spreadsheetObj = SpreadsheetApp.getActiveSpreadsheet();
  var menuList       = [];
  /*/
  menuList.push({
    name : "●●●　まだありません",
    functionName : "vup0"
  });
  /*/
  menuList.push({
    name : "■ NEW ■　1.00 ⇒ 1.01",
    functionName : "vup1"
  });
  ss.addMenu("バージョンアップ", menuList);
}
function vup0(){ 
}
function vup1(){ 
  var verNo = '1.01';
  var kousin = '1.01　自動売買の結果のページのタグの一部変更に対応。取引グラフの日数変更を可能にした。';
    SpreadsheetApp.getActiveSpreadsheet().toast(
    '[ 1.00 ⇒ 1.01 ]' +'　　　　　　　' +
    '終了までお待ちください。'
    ,'バージョンアップ実行中', -1);
  vup(verNo,kousin);
}

function vup(verNo,kousin) {
  //var verNo = '1.01';
  //var kousin = '1.01 自動売買の結果のページのタグの一部変更に対応。取引グラフの日数変更を可能にした。';
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('マニュアル');
  sh.showSheet();
  //var rng = sh.getRange("A1:A50");
  //sh.unhideRow(rng);
  if (sh.getRange('A1').getValue()!= verNo){
  sh.insertRowBefore(3);
  sh.getRange('B3').setValue(kousin);
  sh.getRange('A1').setValue(verNo);
    SpreadsheetApp.getActiveSpreadsheet().toast(
    '[ バージョンアップ ]' +'      　　　 　　　　' +
    '終了しました。'
    ,'バージョンアップの終了', 3);
  }else{
    SpreadsheetApp.getActiveSpreadsheet().toast(
    '[ バージョンアップ ]' +'  　　　　　 ' +
    '既にバージョンアップ済です。'
    ,'バージョンアップの終了', 3); 
  }
  var row2=findRow(sh,'使用方法',1);
  var rng = sh.getRange(2,1,row2-2,1);
  sh.hideRow(rng);
  //sh.unhideRow(rng);
}

function findRow(sheet,val,col){
  var dat = sheet.getDataRange().getValues(); //受け取ったシートのデータを二次元配列に取得
  for(var i=1;i<dat.length;i++){
    if(dat[i][col-1] === val){
      return i+1;
    }
  }
  return 0;
}

function setNameLookup() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('対応表');
  ss.setNamedRange('戦略名', sh.getRange('F2:F23'));
  ss.setNamedRange('番号から戦略名', sh.getRange('D2:F23'));
  ss.setNamedRange('メモから戦略名', sh.getRange('H2:K43'));
};
function sortLookup() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('対応表');
  sh.getRange('D2:F23').sort([{column: 4, ascending: true}]);
  sh.getRange('H2:K43').sort([{column: 8, ascending: true}]);
};
