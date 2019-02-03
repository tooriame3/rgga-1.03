function addClosingPriceData() {
  var folderName = "okasan";
  var myfolder=DriveApp.getRootFolder().getFoldersByName(folderName).next();
  var fileName = "最新終値データ.csv";
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
          uploadClosingPriceData(file);
          folder.removeFile(file);
          return;
        }
      }
    }
  }
}

function uploadClosingPriceData(file){
  // 現在アクティブなスプレッドシートを取得
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  var sheetName ='終値データ'; 
  var sh = ss.getSheetByName(sheetName);
  //sh.activate;
  // フォーマット
  var numFormats1 = 'yyyy/MM/dd';
  var numFormats2 = 'hh:mm:ss';
  //

    var csvData = Utilities.parseCsv(file.getBlob().getDataAsString("Shift_JIS"));
    data = csvData;
    if (data.length>=79){
      　var range = sh.getRange("A2:C");
        range.clearContent();
        var range = sh.getRange(1, 1, data.length, data[0].length);
        range.setValues(data); 
    }
    // 指定したセル範囲にフォーマットを適用
      sh.getRange("A:A").setNumberFormat(numFormats1);  
      sh.getRange("B:B").setNumberFormat(numFormats2);
    /*/sort
    last_row = getLastRowNumber_ColumnA(sheetName);
    if (last_row>2){
      rng = sh.getRange(2, 1, last_row-1, 3);  // <--対象範囲  
    　　sh.getRange(2, 1, last_row-1, 3).sort([{column: 1, ascending: true},{column: 2, ascending: true}]);
  　　}
    /*/
    //軸の最大値・最小値の変更
      changeAxes();
}

//最終行を返す
function getLastRowNumber_ColumnNo(sheetName,col){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var  sh = ss.getSheetByName(sheetName);
  var last_row = sh.getLastRow();
  var data = [[]];
  data = sh.getRange(1, col, last_row, col).getValues();//
  i = data.filter(String).length;
  return i;
}

function changeAxes() {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheetName1 ='終値データ'; 
  var sheet1 = spreadsheet.getSheetByName(sheetName1);
  var y_min =sheet1.getRange("R17").getValue();
  var y_max =sheet1.getRange("Q17").getValue();
  
  var sheetName2 ='取引グラフ'; 
  var sheet2 = spreadsheet.getSheetByName(sheetName2);
  var charts = sheet2.getCharts();
  var chart = charts[0];
  var newchart = chart;
  newchart = newchart.modify()
  .setOption('vAxes.0.viewWindowMode', 'explicit')
  .setOption ('vAxes.0.viewWindow.min', y_min)
  .setOption ('vAxes.0.viewWindow.max', y_max)
  .setOption('vAxis.format', 'short')
  .setOption('hAxis.showTextEvery',13)
  .setOption('hAxis.format', 'short')

  .build();
  sheet2.updateChart(newchart);
};
