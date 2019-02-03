///////////////////////////////////////
https://kido0617.github.io/js/2017-02-13-gas-6-minutes/

//指定したkeyに保存されているトリガーIDを使って、トリガーを削除する
function deleteTrigger(triggerKey) {
  var triggerId = PropertiesService.getScriptProperties().getProperty(triggerKey);
  
  if(!triggerId) return;
  
  ScriptApp.getProjectTriggers().filter(function(trigger){
    return trigger.getUniqueId() == triggerId;
  })
  .forEach(function(trigger) {
    ScriptApp.deleteTrigger(trigger);
  });
  PropertiesService.getScriptProperties().deleteProperty(triggerKey);
}
 
//トリガーを発行
function setTrigger(triggerKey, funcName){
  deleteTrigger(triggerKey);   //保存しているトリガーがあったら削除
  var dt = new Date();
  dt.setMinutes(dt.getMinutes() + 1);  //１分後に再実行
  var triggerId = ScriptApp.newTrigger(funcName).timeBased().at(dt).create().getUniqueId();
  //あとでトリガーを削除するためにトリガーIDを保存しておく
  PropertiesService.getScriptProperties().setProperty(triggerKey, triggerId);
}

/////////////////////////////////////////////////

//トリガーで呼ばれる関数
function runFromTriger() {
  var file = getFile("okasan", '取引集計（イキナリＡＵＴＯ版）');
  var ss = SpreadsheetApp.open(file);
  if (ss !==undefined){
    main();
  }
  //トリガーを発行
  var triggerKey = "triggerNext";    //トリガーIDを保存するときに使用するkey
  var funcName="runFromTriger_Next";
  setTrigger(triggerKey, funcName)
}

function runFromTriger_Next() {

  var file = getFile("okasan", '取引集計（イキナリＡＵＴＯ版）');
  var ss = SpreadsheetApp.open(file);
  if (ss !==undefined){
    main2();
  }
  //全て実行終えたらトリガーを削除する
  var triggerKey = "triggerNext";    //トリガーIDを保存するときに使用するkey
  deleteTrigger(triggerKey);
  //トリガーを発行
  var triggerKey = "triggerUser";    //トリガーIDを保存するときに使用するkey
  var funcName="runFromTriger_User";
  setTrigger(triggerKey, funcName)
  //
}
//
function runFromTriger_User() {
  var file = getFile("okasan", '取引集計（イキナリＡＵＴＯ版）');
  var ss = SpreadsheetApp.open(file);
  if (ss !==undefined){
    syori2();
  }
  //全て実行終えたらトリガーを削除する
  var triggerKey = "triggerUser";    //トリガーIDを保存するときに使用するkey
  deleteTrigger(triggerKey);
}
//
function runFromTriger2() {
  var file = getFile("okasan", '取引集計（イキナリＡＵＴＯ版）');
  var ss = SpreadsheetApp.open(file);
  if (ss !==undefined){
    getIkinariData2();
  }
}

//ユーザーの独自処理
function syori2(){
}
// Creates a trigger
function createTrigger(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var triggers = ScriptApp.getUserTriggers(ss);
  if (triggers.length ==1){
  ScriptApp.newTrigger("runFromTriger")
    .timeBased()
    .everyMinutes(30)
    .create();
  }
}
// Creates a trigger2
function createTrigger2(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var triggers = ScriptApp.getUserTriggers(ss);
  if (triggers.length ==0){
  ScriptApp.newTrigger("runFromTriger2")
    .timeBased()
    .atHour(11)
    .nearMinute(30)//スクリプトのタイムゾーンの午前11時半頃に実行
    .everyDays(1)// atHour（）を使用している場合は必要
    .create();
  }
}


//フォルダとファイルを名前から指定してファイルを所得
function getFile(dirname, filename) {
  var dirs = DriveApp.getFoldersByName(dirname);
  if(!dirs.hasNext()) {return undefined;}
  var dir = dirs.next();
  var files = dir.getFilesByName(filename);
  if(!files.hasNext()) {return undefined;}
  return files.next();
}

function main(){
  PopupStartMain();
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  var folderName = "okasan";
  var myfolder=DriveApp.getRootFolder().getFoldersByName(folderName).next();
  var fileName = "最新日毎データ.csv";
  onLinkA();
  onLinkB();
  getMail();

  if (getFile("okasan", "最新日毎データ.csv") != undefined) { 
    mycsv = myfolder.getFilesByName(fileName).next();
    importNew();//"最新日毎データ.csv";
    sortSheet("入力と訂正");
    trashMycsv();
  };  
  PopupEndMain();
}

function main2(){
  PopupStartMain2();
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  if (getFile("okasan", "最新終値データ.csv") != undefined) { 
    addClosingPriceData();
    summaryMail();
  };
  
  PopupEndMain2();
}

function trashMycsv(){
  mycsv=getFile("okasan", "最新日毎データ.csv");
  if (mycsv != undefined){
    mycsv.setTrashed(true);
  }; 
}

function importNew() {
  var folderName = "okasan";
  var myfolder=DriveApp.getRootFolder().getFoldersByName(folderName).next();
  var fileName = "最新日毎データ.csv";
  var folders = DriveApp.getFoldersByName(folderName);
  
  //指定フォルダを検索
  while (folders.hasNext()) {
    var folder = folders.next();
    if (folder.getName() == folderName) {
      myfolder = folder;
      var files = DriveApp.getFilesByName(fileName);
      //指定したCSVファイルを検索
      while (files.hasNext()) {
        var file = files.next();
        if (file.getName() == fileName) {
          //var data = file.getBlob().getDataAsString("Shift_JIS"); 
          //var csv = Utilities.parseCsv(data);            
          //uploadMain(csv);//修復用にデータを貼り付ける
          uploadMain(file);
          folder.removeFile(file);
          return;
        }
      }
    }
  }
}

//ソート入力と訂正
function sort1(){
  sortSheet("入力と訂正");
}
//ソート履歴の修復
function sort2(){
  sortSheet("履歴の修復");
}

function sortSheet(sheetName){
  //sheetName="入力と訂正"or"履歴の修復";
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var  sh = ss.getSheetByName(sheetName);
  var last_row = getLastRowNumber_ColumnA(sheetName);
  if (last_row>2){
    rng = sh.getRange(2, 1, last_row-1, 16);  // <--対象範囲  
    //そーと P=16で降順（重複キー降順）
    sh.getRange(2, 1, last_row-1, 16).sort([{column: 16, ascending: false},
                                            {column: 14, ascending: false},{column: 15, ascending: false}]);
  }
}




