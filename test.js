function confirmation() {
  let ui = SpreadsheetApp.getUi();
  let title = 'メールを送信しますか?';
  let response = ui.alert(title, ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES) {
    sendMail()
  } else {
    return false
  }
}

function sendMail() {
  let spreadSheet = SpreadsheetApp.openById('17xqYvL50zH6PjRIYukP2Rt52rMVMQXHVnJVmEhmQQg4');
  let sheet = spreadSheet.getSheetByName('クライアント一覧');
  let lastRow = sheet.getLastRow();
  let clients = sheet.getRange(1, 1, lastRow, 5).getValues();
  
  let sheet_2 = spreadSheet.getSheetByName('送信非対称クライアント一覧');
  let sheet_2_lastRow = sheet_2.getLastRow();
  let notSendClients = sheet_2.getRange(2, 2, sheet_2_lastRow -1).getValues().flat();
  
  let DOC_URL = 'https://docs.google.com/document/d/1UF1KdQ5MgrxEE9jNr2laP5uVIMi6i-EzSON6xPman4g/edit';
  let doc = DocumentApp.openByUrl(DOC_URL);
  let docText = doc.getBody().getText();
  
  let subject = sheet.getRange("H2").getValue(); //件名
  let kenmei = sheet.getRange("G2").getValue(); //From
  let options = {name: `${kenmei}`}; //From
  
  for(let i = 1; i < lastRow; i++) {
    if (!notSendClients.includes(clients[i][1])) {
      if (!sendedCheck(i, clients)) {
        let range = sheet.getRange(i + 1, 5);
        range.check();
        
        let company = clients[i][1]; //会社名
        let name = clients[i][2];  //名前
        let recipient = clients[i][3]; //宛先
        
        let body = docText
        .replace('{社名}', company)
        .replace('{担当者名}', name);
        
        //GmailApp.sendEmail(recipient, subject, body, options);
      }
    }
  }
}

function sendedCheck(i, values) {
  if (values[i][4] === true) {
    return true
  } else {
    return false
  }
}
