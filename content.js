var mailApp=MailApp;
var app=SpreadsheetApp;
var spreadsheet=app.getActiveSpreadsheet();
var sheet=spreadsheet.getSheetByName("listEmail");
var sheet1=spreadsheet.getSheetByName("Messege");
var values1=sheet1.getDataRange().getValues();

let drive=DriveApp;
let folderID=values1[2][1];

function sendMails() {

  var values=sheet.getDataRange().getValues();
  var last=values.length;
  let pdf=drive.getFolderById(folderID).getFilesByName(values1[3][1]).next().getAs('application/pdf');

  for (var row=0; row < last; row++){
    if(row > 0) {

      mailApp.sendEmail(
        values[row][0],
        values1[0][1],
        values1[1][1],
        {attachments: [pdf]
                       })
    }
  }
}


 
