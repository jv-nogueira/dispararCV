var mailApp=MailApp;
var app=SpreadsheetApp;
var spreadsheet=app.getActiveSpreadsheet();
var sheet=spreadsheet.getSheetByName("listEmail");
var sheet1=spreadsheet.getSheetByName("Messege");
var values1=sheet1.getDataRange().getValues();

let drive=DriveApp;
let folderID=values1[2][1]; // Coluna B, linha 2

function sendMails() {

  var values=sheet.getDataRange().getValues();
  var last=values.length;
  let pdf=drive.getFolderById(folderID).getFilesByName(values1[3][1]).next().getAs('application/pdf'); // Coluna B, linha 4

  for (var row=0; row < last; row++){
    if(row > 0) {

      mailApp.sendEmail(
        values[row][0], // Coluna A, linha 2 em diante (laço de repetição)
        values1[0][1], // Coluna B, linha 1
        values1[1][1], // Coluna B, linha 2
        {attachments: [pdf]
                       })
    }
  }
}


 
