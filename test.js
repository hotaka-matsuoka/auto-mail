function sendMail() {
  let spreadSheet = SpreadsheetApp.openById('17xqYvL50zH6PjRIYukP2Rt52rMVMQXHVnJVmEhmQQg4')
  let sheet = spreadSheet.getSheetByName('シート1');
  let lastRow = sheet.getLastRow();
  let values = sheet.getRange(1, 1, lastRow, 5).getValues();
 
  let DOC_URL = 'https://docs.google.com/document/d/1UF1KdQ5MgrxEE9jNr2laP5uVIMi6i-EzSON6xPman4g/edit';
  let doc = DocumentApp.openByUrl(DOC_URL);
  let docText = doc.getBody().getText();
  
  let subject = 'メール送信テスト'; //題名
  let options = {name: 'これはメール自動送信のテストです'}; //件名
  
  for(let i = 1; i < lastRow; i++) {
    if (!sendedCheck(i, values)) {
      console.log(`${values[i][0]} ${values[i][1]}に送信しました`)
      let range = sheet.getRange(i + 1, 5);
      range.check();
      
      let company = values[i][1]; //会社名
      let name = values[i][2];  //名前
      let recipient = values[i][3]; //宛先
      
      let body = docText
      .replace('{社名}', company)
      .replace('{担当者名}', name);
      
      // GmailApp.sendEmail(recipient, subject, body, options);
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