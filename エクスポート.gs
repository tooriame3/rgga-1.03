//https://officeforest.org/wp/2018/11/25/google-apps-scriptでcsvファイルを取り扱う/
//ダイアログ用のグローバル変数
var url = "";
 
//CSVエクスポートするルーチン
function exportcsv(){
  PopupStartD();
  var folderName = "okasan";
  var targetFolder = DriveApp.getRootFolder().getFoldersByName(folderName).next().getId();
  var sheetName="入力と訂正";
  var col = 14; //戦略入力
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var  sh = ss.getSheetByName(sheetName);
  var last_row = getLastRowNumber_ColumnA(sheetName);
  var data = sh.getRange(1, 1, last_row, col).getValues();
  // 2次元配列になっているデータをcsvに変換
  //csvデータをshif-jisに変換
  var csvData = csvchange(data);
　var today = Utilities.formatDate(new Date(), "JST", "_MMdd");
  var filename = "全取引履歴"+today+".csv";
  var blob = Utilities.newBlob("", "text/comma-separated-values", filename).setDataFromString(csvData, "Shift_JIS");
  //blobデータをcsvファイルとしてドライブに保存
  var fileid = DriveApp.getFolderById(targetFolder).createFile(blob).getId();
  //  
  PopupEndD();  
  //ダウンロードリンクを生成
  url = "https://drive.google.com/uc?export=download&id=" + fileid;
  //ダウンロードリンクのダイアログを生成
  var output = HtmlService.createTemplateFromFile('download');
  var html = output.evaluate();
  var ui = SpreadsheetApp.getUi();
  ui.showModalDialog(html, "全取引履歴を保存"); 
}
 
//CSVデータ形式に整える関数
function csvchange(data){
  var rowlength = data.length;
  var columnlength = data[0].length;
  var csvdata = "";
  var csv = "";
  for(var i = 0;i<rowlength;i++){
    var tmp=data[i];
    //
    if (i>0){
      var date=formatDate(tmp[0]);
      tmp.shift();
      tmp.unshift(date);
    };
    //
    if (i < rowlength-1) {
      csvdata += tmp.join(",") + "\r\n";
      //csvdata += data[i].join(",") + "\r\n";
    }else{
      csvdata += tmp;
      //csvdata += data[i];
    }
  }
  return csvdata;
}

//日付
function formatDate(date){
  if(Object.prototype.toString.call(date) !== '[object Date]') return '';
  return Utilities.formatDate(date, "JST","yyyy/MM/dd HH:mm:ss");
}