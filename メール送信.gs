function sendMailcsvSheet(){
  var mail_subject = "test";
  var mail_body = "test";
  var sheetName = "使用方法";
  sendMail4(mail_subject,mail_body,sheetName);
}
//*****************************************************************************
//ユーザー用
function sendMail4(mail_subject,mail_body,sheetName) {

  var to_address = Session.getActiveUser().getEmail();
  var d = new Date();
  var date = Utilities.formatDate( d, 'Asia/Tokyo', 'MM/dd HH:mm');
    
  var csvData = loadData4(sheetName);
  writeDrive4(csvData,sheetName);
  
  var csvName = sheetName+".csv";
  attachFile = getFile("okasan", csvName);
  
  MailApp.sendEmail(to_address,mail_subject,mail_body,{attachments: [attachFile]});
}


function writeDrive4(csvData,sheetName) {
  //CSVファイルが置かれているGoogleDriveのフォルダー名を指定
  var folderName = "okasan";
  var myfolder=DriveApp.getRootFolder().getFoldersByName(folderName).next();
  var csvName = sheetName+".csv";
  var file = getFile("okasan", csvName);
  if (file!=undefined){
    file.setTrashed(true);
  };
  var contentType = 'text/csv';
  var charset = "Shift_JIS";
  //var charset = 'utf-8';
  var blob = Utilities.newBlob('', contentType, csvName).setDataFromString(csvData, charset);
  //var blob = Utilities.newBlob("", "text/comma-separated-values", filename).setDataFromString(csvData, "Shift_JIS");
  myfolder.createFile(blob);
}

//'sheetName'
function loadData4(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  var sh = ss.getSheetByName(sheetName);
  ss.setActiveSheet(sh, true);
  var last_row = getLastRowNumber_ColumnA(sheetName);
  var last_col = sh.getLastColumn();
  var data = sh.getRange(1, 1, last_row, last_col).getValues();
  var data = sh.getDataRange().getValues();
  var csv = '';
  for(var i = 0; i < data.length; i++) {
    csv += data[i].join(',') + "\r\n";};
  Logger.log(csv);
  return csv;
}