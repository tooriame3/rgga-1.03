//未使用
function getIkinariData(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var getid = ss.getId();
  var sh1 = ss.getSheetByName('ＨＰ検証用');
  sh1.showSheet();
  ss.setActiveSheet(sh1, true);
  
  PopupStartGetIkinariData();
  clearMySheet2();

  //ikinariData( strategyName,  memo       ,session,col)
  ikinariData("last15-2"     ,"LAST15"   ,"Day"  ,6);
  clearMySheet3();//LAST15の昨年以前をクリア
  
  ikinariData("analyzerd20-2","D20b"     ,""     ,6);
  ikinariData("ir30"         ,"IR30"     ,"Day"  ,6);
  ikinariData("finger"       ,"finger1"  ,"Day"  ,5);
  ikinariData("f30"          ,"F30"      ,""     ,6);
  ikinariData("ir_night"     ,"IR_night1","Night",6);
  ikinariData("y_break"      ,"Ybreak"   ,"Day"  ,5);
  ikinariData("analyzerd9"   ,"D9"       ,""     ,6);
  ikinariData("y_night"      ,"Ynight"   ,"Night",5);  
  ikinariData("inthebus"     ,"Bus"      ,"Day"  ,6);
  
  sortSheet2();
  PopupEndGetIkinariData();
}
/*/
function PopupStartGetIkinariData(){
  SpreadsheetApp.getActiveSpreadsheet().toast(
    '[ ＨＰから成績の取込み ]' +'　　　' +
    '処理には時間がかかります。'
    ,'スクリプト実行中', -1);
}
/*/
function PopupEndGetIkinariData(){
  SpreadsheetApp.getActiveSpreadsheet().toast(
    '[ ＨＰから成績の取込み ]' +'　　　' +
    '終了したので再計算します。'
    ,'スクリプトの終了', 10);
}
function PopupStartGetIkinariData(){
  SpreadsheetApp.getActiveSpreadsheet().toast(
    '[ ＨＰから成績の取込み ]' +'　　　' +
    'このシートのデータがすべて、入替わるまでお待ちください。'
    ,'スクリプト実行中', -1);
}

//そーと I=9, F=6で降順（約定日、セッション）
function sortSheet2(){
  sheetName="本家成績";
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var  sh = ss.getSheetByName(sheetName);
  var last_row = getLastRowNumber_ColumnA(sheetName);
  //rng = sh.getRange(2, 1, last_row-1, 10);  // <--対象範囲  
  sh.getRange(2, 1, last_row-1, 10).sort([{column: 10, ascending: false},
         {column: 6, ascending: false}]);
}

//シートを２行目からクリア
function clearMySheet2() {
  sheetName="本家成績";
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var  sh = ss.getSheetByName(sheetName);
  var range = sh.getRange("A2:G");
  range.clearContent();
}

//LAST15の１昨年以前をクリア
function clearMySheet3() {
  sheetName="本家成績";
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var  sh = ss.getSheetByName(sheetName);
  var last_row = getLastRowNumber_ColumnA(sheetName);
  rng = sh.getRange(last_row-30, 1, last_row, 7);  // <--対象範囲 
  rng.clearContent();
}


function ikinariData(strategyName,memo,session,col) {
  var str = "";
  //var strategyName = "ir30";
  //var memo = "IR30";
  //var session = "Day";
  var url = "https://gijuku.online/"+ strategyName + "/";
  try{
    var response = UrlFetchApp.fetch(url);
  }catch(e){
    return str;
  };   
  //スクレイピング
  var source = response.getContentText('UTF-8');
  var data = getContentOfTagName2(source, 'td class');
  var low = data.length/col;
  
  var data2 =[];
  var data3 =[];
  var tmp=[];
  
  for (var i=1; i<=low; i++){
    tmp=[];
    tmp = data.slice(0, col);
    //fingerは５列
    if (col==6){
      if (tmp[5]==""){
        tmp.pop();
        tmp.push(session);
      }
    }else{
      tmp.push(session);
    };
    tmp.push(memo);
    if (tmp[5]=="Day" || tmp[5]=="Night"){
      data2.push(tmp);
    };
    data.splice(0, col);
  };
  
  // 現在アクティブなスプレッドシートを取得
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  var sheetName ="本家成績";
  var sheet = ss.getSheetByName(sheetName);
  var last_row = getLastRowNumber_ColumnA(sheetName)
  sheet.getRange(last_row+1, 1, data2.length, data2[0].length).setValues(data2);
  var chck1=data2[data2.length-1];
  var chck2=last_row+1;
  var chck3=data2.length;
  var chck4=data2[0].length;
  var chck1=data2[data2.length-1];
}

function getContentOfTagName2(html, tagName) {
  var i = 0;
  var j = 0;
  var startOfTag;
  var endOfTag;
  var str = [ ];
  while(html.indexOf('<td class',j)!=-1){
    //"td class"の開始位置を取得
    j = html.indexOf('<td class',j);
    //次の">"位置 + 1を文字列の始めとする
    startOfStr = html.indexOf('>',j)+1;
    //次の"</td>"位置を文字列の終わりとする
    endOfStr = html.indexOf('</td>',j);
    //タグの間の文字列を配列に追加
    str[i] = html.substring(startOfStr, endOfStr);
    j = endOfStr + 1;
    i++;
  }
  return str;
}

//-------------------------------------------
//配列dataと同じサイズの、年dataの配列を返す
function makeYearData(data,n1){
  var now = new Date();
  var year =now.getYear();
  var year1=[String(year-1)+"年"];
  var year2=[String(year)+"年"];
  //var year1=["2018年"];//n1個
  //var year2=["2019年"];//残り(n2)は上に来る
  
  var len=data.length;
  var n2;
  //n1=-1 なら、すべてyear1
  if (n1==-1){
    n1=len;
    n2=0;
  }else{
    n2=len-n1;
  };
  var newdata=[];
  for (var i=1;i<=n2;i++){
    newdata.push(year2);
  };
  for (var i=1;i<=n1;i++){
    newdata.push(year1);
  };

  return (newdata); 
}

function testMYD(){
  data = [[1,2], [3,4],[5,6],[7,8]];
  newdata=makeYearData(data,-1);
  Logger.log(newdata);
}

function getIkinariData2(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var getid = ss.getId();
  var sh1 = ss.getSheetByName('チエック２');
  sh1.showSheet();
  ss.setActiveSheet(sh1, true);
  
  PopupStartGetIkinariData();
  clearMySheet2();

  //ikinariData( strategyName,  memo       ,session,col)
  ikinariData2("last15-2"     ,"LAST15"   ,"Day"  ,6);
  clearMySheet3();//LAST15の昨年以前をクリア
  
  ikinariData2("analyzerd20-2","D20b"     ,""     ,6);
  ikinariData2("ir30"         ,"IR30"     ,"Day"  ,6);
  ikinariData2("finger"       ,"finger1"  ,"Day"  ,5);
  ikinariData2("f30"          ,"F30"      ,""     ,6);
  ikinariData2("ir_night"     ,"IR_night1","Night",6);
  ikinariData2("y_break"      ,"Ybreak"   ,"Day"  ,5);
  ikinariData2("analyzerd9"   ,"D9"       ,""     ,6);
  ikinariData2("y_night"      ,"Ynight"   ,"Night",5);  
  ikinariData2("inthebus"     ,"Bus"      ,"Day"  ,6);
  
  sortSheet2();
  PopupEndGetIkinariData();
}

function ikinariData2(strategyName,memo,session,col) {
  var str = "";
  //var strategyName = "ir30";
  //var memo = "IR30";
  //var session = "Day";
  var url = "https://gijuku.online/"+ strategyName + "/";
  try{
    var response = UrlFetchApp.fetch(url);
  }catch(e){
    return str;
  };   
  //スクレイピング
  var source = response.getContentText('UTF-8');
  var data = getContentOfTagName2(source, 'td class');
  var low = data.length/col;
  
  var data2 =[];
  var data3 =[];
  var tmp=[];
  var n2=0;
  var setn2=false;
  var beforDay="9999/01/01";
  
  for (var i=1; i<=low; i++){
    tmp=[];
    tmp = data.slice(0, col);
    //fingerは５列
    if (col==6){
      if (tmp[5]==""){
        tmp.pop();
        tmp.push(session);
      }
    }else{
      tmp.push(session);
    };
    
    
    tmp.push(memo);
    if (tmp[5]=="Day" || tmp[5]=="Night"){
      data2.push(tmp);
      if (!setn2){
        if (formatJDate(tmp[0])>beforDay){
          setn2=true;          
        }else{
          n2++;
          beforDay=formatJDate(tmp[0]);
        };
      };
      data.splice(0, col);
    };
  };
  var n1;
  if (data2.length==n2){
    n1=-1;
  }else{
    n1=data2.length-n2;
  };
  
  //memo == "LAST15"　の暫定的な設定（Last15の１月のデータが掲載されるまでは2018,2017年のデータがあるから）
  if (memo == "LAST15"){
    if (data2[0][0].substr(0,2)=="12"){
      n1=-1;
    }
  }  
  
  // 現在アクティブなスプレッドシートを取得
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  var sheetName ="本家成績";
  var sheet = ss.getSheetByName(sheetName);
  var last_row = getLastRowNumber_ColumnA(sheetName)
  sheet.getRange(last_row+1, 1, data2.length, data2[0].length).setValues(data2);
  
  var yeardata=makeYearData(data2,n1);
  sheet.getRange(last_row+1, data2[0].length+1, yeardata.length, 1).setValues(yeardata);
  
}

function formatJDate(str) {
  //var str = '3月14日';
  var reg1 = /^([0-9]月)/;
  var reg2 = /月([0-9]日)/;
  str = str.replace(reg1,"0"+'$1');
  str = str.replace(reg2,"月0"+'$1');
  return str;
  Logger.log(str);
}


