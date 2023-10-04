//pesquisa se o ID da mensagem já está na lista emails processados
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

//adiciona o ID da mensagem de email à lista de mensagens processadas
function addProcessedMessageId(messageId) {
  var sheet = SpreadsheetApp.openById('1xMM6EnNyO3r0BI_RYtC4Ntqjtkwyi80g9ZQEhdY9UXY').getSheetByName('lista_de_ID_mensagem');
  sheet.appendRow([messageId]);
}
