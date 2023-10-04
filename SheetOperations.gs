function isMessageProcessed(messageId) {
  var sheet = SpreadsheetApp.openById('1xMM6EnNyO3r0BI_RYtC4Ntqjtkwyi80g9ZQEhdY9UXY').getSheetByName('lista_de_ID_mensagem');
  var ids = sheet.getRange("A:A").getValues();
  for (var i = 0; i < ids.length; i++) {
    if (ids[i][0] === messageId) {
      return true;
    }
  }
  return false;
}

function addProcessedMessageId(messageId) {
  var sheet = SpreadsheetApp.openById('1xMM6EnNyO3r0BI_RYtC4Ntqjtkwyi80g9ZQEhdY9UXY').getSheetByName('lista_de_ID_mensagem');
  sheet.appendRow([messageId]);
}

function recordPayment(data) {
  var spreadsheet = SpreadsheetApp.openById('1xMM6EnNyO3r0BI_RYtC4Ntqjtkwyi80g9ZQEhdY9UXY');
  var nameOfSheet = "dados_email_comprovante";
  var sheet = spreadsheet.getSheetByName(nameOfSheet);
  if(!isDuplicate(data.idDoPIX,sheet)){
    sheet.appendRow([data.idDoPIX, data.nomeDoEmail, data.dataDoEmail, data.horaDoEmail, data.urlDoComprovante, data.dataDoComprovante, data.valorDoComprovante, data.nomeDoComprovante]);
  }
  
}


function addDataToSheet(data) {
    var spreadsheet = SpreadsheetApp.openById('1xMM6EnNyO3r0BI_RYtC4Ntqjtkwyi80g9ZQEhdY9UXY');
    var sheet = spreadsheet.getSheetByName("dados_do_extrato");

    if (!isDuplicate(data.idOfTransaction, sheet)) {
        var nextRow = sheet.getLastRow() + 1;
        sheet.getRange(nextRow, 1, 1, 4).setValues([[data.idOfTransaction, data.amountTransaction, data.nameOfSender, data.dateTransaction]]);
    }
}

function isDuplicate(data, sheet) {
    var existingData = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues();
    return existingData.some(row => row[0] === data);
}

// function isDuplicatePIX(data, sheet) {
//     var existingData = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues();
//     return existingData.some(row => row[0] === data.idOfTransaction);
// }
