function getMail() {
  Logger.log("start");
  //var ss = SpreadsheetApp.getActiveSpreadsheet();
  //アクティブなシートを取得
  //var sht = ss.getActiveSheet(); 
  //スプレッドシートのidを取得する
  //var idSS = ss.getId();
  //var fileSS = DriveApp.getFileById(idSS);
  //
  var Threads = GmailApp.getInboxThreads(); 
  for(var i=0;i<50;i++){ 
    var status = Threads[i].isUnread();
    if(status==true){ 
      var subject = Threads[i].getMessages()[0].getSubject();
      //抽出するメールを件名で絞り込む
      if(subject=="DO NOT OPEN!(by GAS)"){
        var msgs = Threads[i].getMessages();
        //スレッドのメッセージを回して最新のメールの添付ファイルの配列を取得する
        var attachments
        for(m in msgs) {
          var msg = msgs[m];          
          //そのメールが未読かチェックする
          if(msg.isUnread()){
            //未読だった場合、添付ファイルオブジェクトを取得
            attachments = msg.getAttachments();
            //添付ファイルオブジェクトの配列からファイルを全てコピーする
            for(var j = 0;j<attachments.length;j++) {
              var folder_name = 'okasan';
              var folder = DriveApp.getFoldersByName(folder_name).next();
              for(var j = 0;j<attachments.length;j++){
                var data = DriveApp.createFile(attachments[j]);
                folder.addFile(data);
                //ルートディレクトリに作られたファイルを削除する
                DriveApp.getRootFolder().removeFile(data);                  
                Logger.log("copy name: " + data.getName());
              }
              //メールを既読にする
              msg.markRead()
            }
          }
        }
        
        //ごみ箱に捨てる前にスレッドを念の為、既読にする
        Threads[i].markRead();        
        Threads[i].moveToTrash(); //保存終了したらゴミ箱に移動
      }
    }
  }  
  Logger.log("end");
}
//
function getChtImage(){
  var spreadsheet = SpreadsheetApp.getActive();
  var sheetName ='取引グラフ2'; 
  var sheet = spreadsheet.getSheetByName(sheetName);
  var charts = sheet.getCharts();
  var chart = charts[0];
  //グラフの画像を取得
  // 画像を作成
  var img = DriveApp.createFile(chart.getBlob());
 
  // 画像のURLを取得してimgタグ作成 
  var url = img.getDownloadUrl().replace('?e=download&gd=true','');
  var html = '<img src=\"'+url+'\">';

  img.setTrashed(true);
  var html2 = HtmlService.createHtmlOutput(html);//グラフの画像を取得;
  spreadsheet.show(html2);
 // var folderName = "okasan";
 // var myfolder=DriveApp.getRootFolder().getFoldersByName(folderName).next();
  //myfolder.createFile("取引グラフ.png", img2, 'image/png')
}

 


function getChtImage2(){
  var spreadsheet = SpreadsheetApp.getActive();
  var sheetName ='取引グラフ2'; 
  var sheet = spreadsheet.getSheetByName(sheetName);
  var charts = sheet.getCharts();
  var chart = charts[0];
  //グラフの画像を取得
  var imageBlob = chart.getBlob().getAs('application/pdf').setName("取引グラフ.pdf");//グラフの画像を取得
  
}
//*****************************************************************************
//最新バージョン　
function sendMail3() {
  //refreshSheet2(); //"サマリー"の更新を待つ
  var to_address = Session.getActiveUser().getEmail();
  var d = new Date();
  var date = Utilities.formatDate( d, 'Asia/Tokyo', 'MM/dd HH:mm');
  var mail_subject = "最新の取引集計サマリー_"+date ;
  
  var str1=ikinari2(); //自動売買義塾のスクレイピング
  var link1= "<a href=\"https://gijuku.online/#home\">gijyuku.online</a>";
  
  var str2=tecnical(); //テクニカルのスクレイピング
  //var str2="";
  var link2= "<a href=\"https://jp.investing.com/indices/japan-225-futures-technical\">Investing.com</a> JP225(CFD) "+date+ " 現在";
  //var link2="";
  
  var str="";
  str = "サマリーを添付しました。" +  "<br>"  + 
     str1 + "<br>" +  str2 + link2;
  //  "（参考）<br>" + "[WEBページより抜粋] "+ link1 + str1 + "<br>" +  str2 + link2;
  //var html = createHtmlOutput(str)
  //var mail_body = 'htmlメールが表示できませんでした';
  
  var mail_body;
  mail_body = "サマリー.csvを添付しました。" ;
  var summary;
  createCsv();
  summary = getFile("okasan", "サマリー.csv");
 // var chtIm = getFile("okasan", "取引グラフ.png");
 MailApp.sendEmail(to_address,mail_subject,mail_body,{htmlBody: str , attachments: [summary]});
  //MailApp.sendEmail(to_address,mail_subject,mail_body,{htmlBody: str , attachments: [summary,chtIm]});
}

function createCsv() {
  var csvData = loadData();
  writeDrive(csvData);
}

function writeDrive(csv) {
  //CSVファイルが置かれているGoogleDriveのフォルダー名を指定
  var folderName = "okasan";
  var myfolder=DriveApp.getRootFolder().getFoldersByName(folderName).next();
  var fileName = "サマリー.csv";
  var file = getFile("okasan", "サマリー.csv");
  if (file!=undefined){
    file.setTrashed(true);
  };
  var contentType = 'text/csv';
  var charset = "Shift_JIS";
  //var charset = 'utf-8';
  var blob = Utilities.newBlob('', contentType, fileName).setDataFromString(csv, charset);
  //var blob = Utilities.newBlob("", "text/comma-separated-values", filename).setDataFromString(csvData, "Shift_JIS");
  myfolder.createFile(blob);
}

//'サマリー'
function loadData() {
  // 現在アクティブなスプレッドシートを取得
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  // そのスプレッドシートにある"サマリー"という名前のシートを取得
  var sh = ss.getSheetByName('サマリー');
  ss.setActiveSheet(sh, true);
  var last_row = getLastRowNumber_ColumnA('サマリー');
  var last_col = sh.getLastColumn();
  var data = sh.getRange(1, 1, last_row, last_col).getValues();
  var data = sh.getDataRange().getValues();
  var csv = '';
  for(var i = 0; i < data.length; i++) {
    csv += data[i].join(',') + "\r\n";};
  Logger.log(csv);
  return csv;
}

function getContentOfTagName(html, tagName) {
  var i = 0;
  var j = 0;
  var startOfTag;
  var endOfTag;
  var str = [ ];
  while(html.indexOf('<' + tagName,j)!=-1){
    //"<タグ名"の開始位置を取得
    j = html.indexOf('<' + tagName,j);
    //次の">"位置 + 1を文字列の始めとする
    startOfStr = html.indexOf('>',j)+1;
    //次の"</タグ名>"位置を文字列の終わりとする
    endOfStr = html.indexOf('</' + tagName + '>',j);
    //タグの間の文字列を配列に追加
    str[i] = html.substring(startOfStr, endOfStr);
    j = endOfStr + 1;
    i++;
  }
  return str;
}

function tecnical() {
  var html = "";
  var url="https://jp.investing.com/indices/japan-225-futures";
  try{
    var response = UrlFetchApp.fetch(url);
  }catch(e){
    return str;
  }
  var response = UrlFetchApp.fetch(url);
  if (response != undefined){
    var source = response.getContentText('UTF-8');
    var kiji = getContentOfTagName(source, 'tbody');
    var title = getContentOfTagName(source, 'thead');
    var body = [];
    body[0] = "<table frame=\"border\" rules=\"all\">";
    body[1] = title[2];
    body[2] = kiji[2];
    body[3] = "<\/table>";
    html = "<h4>テクニカル<\/h4>"+ body[0]+ body[1]+ body[2]+ body[3];
  };
  return html;
}

//使用していない
function ikinari1() {

  var url="https://gijuku.online"
  var response = UrlFetchApp.fetch(url);
  var source = response.getContentText('UTF-8');
  var link = getContentOfTagName10(source, 'li');
  var i;
  for (i=0;link.length-1;i++){
    if ( link[i].match(/\/%e8%87%aa%e5%8b%95%e5%a3%b2%e8%b2%b7%e3%81%ae%e7%b5%90%e6%9e%9c/)) {
      break;
    }
  };  
  return link[i];
}


//使用していない
function getContentOfTagName10(html, tagName) {
  var i = 0;
  var j = 0;
  var startOfTag;
  var endOfTag;
  var str = [ ];
  while(html.indexOf('<' + tagName,j)!=-1){
    //"<タグ名"の開始位置を取得
    j = html.indexOf('<' + tagName,j);
    //
    startOfStr = html.indexOf('href=',j)+6;
    //
    endOfStr = html.indexOf('\/"',j)+1;
    //タグの間の文字列を配列に追加
    str[i] = html.substring(startOfStr, endOfStr);
    j = endOfStr + 1;
    i++;
  }
  return str;
}



//使用中
function ikinari2() {
  var str = "";
  var url = ikinari1();
  try{
    var response = UrlFetchApp.fetch(url);
  }catch(e){
      return str;
  };   
  //スクレイピング
  var source = response.getContentText('UTF-8');
  var kiji = getContentOfTagName(source, 'p');
  var title = getContentOfTagName(source, 'title');
  var title2 = getContentOfTagName(source, 'h2');

  //追加
   var lists = Parser.data(source).from('デイセッション============</p>').to('////////////').build();
  lists="<p>■===デイセッション============</p>"+lists
  //
  var i;
  for (i=0;title2.length-1;i++){
    if ( title2[i].match(/自動/)) {
      break;
    }
  };
  //
  var kiji_date = i;
  var kiji_first = 0;
  var kiji_end =0 ;
  for (i=0;title2.length-1;i++){
    if ( kiji[i].match(/■===デイ/)) {
      break;
    }
  };
  kiji_first = i ;
  
  for (i=0;title2.length-1;i++){
    if ( kiji[i].match(/→ <a/)) {
      break;
    }
  };
  kiji_end = i-1 ;
  //
  var body = "";
  add1 = "<p>";
  add2 = "<\/p>";
  body = "<h4>" + title2[kiji_date] + "<\/h4>";  
  /*/
  for (i = kiji_first; i<=kiji_end; i++){
    body = body + add1 + kiji[i] + add2;
  }
  /*/
  //これで上を置き換え
  body=body+lists
  //
  return body;
}

//ベーシック認証
/*
var url = "http://example.com/admin/";
var user = "hoge";
var pass = "huga";
var options = {
"headers" : {"Authorization" : " Basic " + Utilities.base64Encode(user + ":" + pass)}
};
var response = UrlFetchApp.fetch(url, options);
*/

//文字列strから最初の文字列preではじまり、最後の文字列sufで終わる
function fetchData(str, pre, suf) {
  var reg = new RegExp('.*?');
  var data = str.match(pre+reg+suf,'g')[0];
  return data;  
}
