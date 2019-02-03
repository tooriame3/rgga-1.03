//データの修正  
function  correctSheetData() {
  //計測
  var start = new Date();
  
  var sheetName0 = "入力と訂正";
  var sheetName1 = "取引履歴";
  var sheetName2 = "修復作業用";
  var sheetName3 = "履歴の修復";
  //onLinkA();
  //onLinkB();
  sortSheet(sheetName0);//念のためにソート
  PopupStartC1();
  var data1 = myFilter(sheetName1);//"取引履歴"保持対象がtrue
  var data2 = myFilter(sheetName2);//"修復作業用"追加対象がtrue  

  clearMySheet(sheetName0);
  appendData(sheetName0, data2);//ソートされている（メモあり）
  appendData(sheetName0, data1);//ソートされている
  sortSheet(sheetName0);
  
  //計測
  var end = new Date();
  var time_past = String((end - start) / 1000);
  Logger.log(time_past);   
}

//最終行を返す
function getLastRowNumber_ColumnA(sheetName){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var  sh = ss.getSheetByName(sheetName);
  var last_row = sh.getLastRow();
  var data = [[]];
  data = sh.getRange(1, 1, last_row, 1).getValues();//
  i = data.filter(String).length;
  return i;
}


//列col=17がtrueであるものの配列を返す
function myFilter(sheetName) {
  //sheetName="取引履歴";
  var col = 17; //判定列Q
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var  sh = ss.getSheetByName(sheetName);
  var last_row = getLastRowNumber_ColumnA(sheetName);
  var data = [[]];
  var data1 = [[]];
  
  data = sh.getRange(2, 1, last_row, col).getValues();//１行目なし
  data1 = data.filter(function(e){
    return e[col-1]
  }).map(
    function(e){return e.slice(0, 14)}
  );　//14＝戦略入力N
  return　data1;
}

//配列dataを"sheetName"にはりつけ
function appendData(sheetName, data) {
  if (data.length > 0){
    try{
      var check = data[0].length;
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var  sh = ss.getSheetByName(sheetName);
      var last_row = getLastRowNumber_ColumnA(sheetName);
      var range = sh.getRange(last_row+1, 1, data.length, data[0].length);
      range.setValues(data);     
    }catch(e){
      return;
    }
  }
}

//シートを２行目からクリア
function clearMySheet(sheetName) {
  //sheetName="入力と訂正";
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var  sh = ss.getSheetByName(sheetName);
  var last_col = 14; //N列
  var range = sh.getRange("A2:N");
  range.clearContent();
}




//https://www.sejuku.net/blog/21812
//array1と同じものがarray2に何個あるかを配列で返す
function countArray(array1,array2){
  var count = function(key) {
    return [array2.filter( function( value ) {
    //抽出
    return value[0] == key[0];
    }).length];
  };
  return array1.map(count);
}

  
