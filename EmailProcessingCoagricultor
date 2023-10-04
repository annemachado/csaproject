//Processa os emails dos coagricultores, extrai os dados dos comprovantes
function processEmailCoagricultor() {
  var myEmail = Session.getActiveUser().getEmail();  // Obter o e-mail do usuário autenticado
  var emailList = getEmailListFromSheet();
  var rootFolderName = 'comprovantes_de_pagamentos_CSA';
  var rootFolder = getOrCreateFolder(rootFolderName);
  for (var i =  0; i < emailList.length; i++){
    var searchQuery = "from:" + emailList[i];
    var threads = GmailApp.search(searchQuery);
    for (var j = 0; j < threads.length; j++) {
      var messages = threads[j].getMessages();
      for (var k = 0; k < messages.length; k++){
        var message = messages[k];
        var messageId = message.getId();
        if(!isMessageProcessed(messageId)){
          var remetente = message.getFrom();
          var emailRemetente = extractEmail(remetente);
          if(emailRemetente !== myEmail){  
            var onlyName = extraiNomeDoRemetente(remetente);
            var dateRecieved = message.getDate();
            var monthYearFolderName = Utilities.formatDate(dateRecieved, Session.getScriptTimeZone(), 'MMMM yyyy');
            var monthYearFolder = getOrCreateFolder(monthYearFolderName, rootFolder);
            var formattedDate = Utilities.formatDate(dateRecieved, "GMT-3", "dd/MM/yyyy");
            var formattedTime = Utilities.formatDate(dateRecieved, "GMT-3", "HH:mm:ss");
            var dadosEmailCoagricultor = {
              nomeDoEmail: onlyName,
              // corpoDoEmail: body,
              dataDoEmail: formattedDate,
              horaDoEmail: formattedTime,
              urlDoComprovante: null,
              dataDoComprovante: null,
              valorDoComprovante: null,
              nomeDoComprovante: null,
              idDoPIX: null
            };

            if (message.getAttachments().length > 0) {
              var attachments = message.getAttachments();
              for (var l = 0; l < attachments.length; l++) {
                var attachmentBlob = attachments[l].copyBlob(); // Retorna um Blob representando o anexo
                var arquivoComprovante = createFileInDrive(attachmentBlob,monthYearFolder);
                var textoDoComprovante = getTextFromPdf(arquivoComprovante); 
                var pixDetalhes = extractPixDetails(textoDoComprovante);
                Logger.log(textoDoComprovante); // Log the entire text for manual verification
                Logger.log(pixDetalhes);
                dadosEmailCoagricultor.urlDoComprovante = arquivoComprovante.getUrl();
                dadosEmailCoagricultor.dataDoComprovante = pixDetalhes.dateOfTransaction;
                dadosEmailCoagricultor.valorDoComprovante = pixDetalhes.value;
                dadosEmailCoagricultor.nomeDoComprovante = pixDetalhes.nameOfPayer;
                dadosEmailCoagricultor.idDoPIX = pixDetalhes.idOfPIX;
                recordPayment(dadosEmailCoagricultor);  
              }
            }
          addProcessedMessageId(messageId);
          }         
        }
      }
    }
  }  
}

//pesquisa os emails cadastrados dos coagricultores
function getEmailListFromSheet() {
  var spreadsheet = SpreadsheetApp.openById('1xMM6EnNyO3r0BI_RYtC4Ntqjtkwyi80g9ZQEhdY9UXY');
  var sheet = spreadsheet.getSheetByName('lista_de_emails'); //nome da aba desejada
  var range = sheet.getRange("A:A"); // Pegando todos os valores da coluna A
  var emails = range.getValues();

  // Filtrar emails não vazios
  var filteredEmails = [];
  for (var m = 0; m < emails.length; m++) {
      if (emails[m][0]) { // Se a célula não estiver vazia
          filteredEmails.push(emails[m][0]);
      }
  }
  return filteredEmails;
}

//Extrai texto dos arquivos (comprovantes) convertendo-os para doc
function getTextFromPdf(fileArquivo) {
  var pdfBlob = fileArquivo.getBlob();
  var docId = Drive.Files.insert({title: 'Temp', mimeType: MimeType.GOOGLE_DOCS}, pdfBlob).id;

  // Obtém o texto do Google Document
  var text = DocumentApp.openById(docId).getBody().getText();

  // Opcional: Exclui o Google Document temporário após a extração do texto
  DriveApp.getFileById(docId).setTrashed(true);

  return text;//Retorna o texto extraído do comprovante
}

//Extrai os dados do PIX
function extractPixDetails(textoDoComprovante) {
  var pixDetails = {
    value: null,
    dateOfTransaction: null,
    nameOfPayer: null,
    idOfPIX: null
  };

  // Extrair valor
  var valueMatch = textoDoComprovante.match(/(\d+,\d+)/);
  if (valueMatch) {
    var valueString = valueMatch[1];
    var valueFloat = parseFloat(valueString.replace(',', '.'));
    pixDetails.value = valueFloat;
  }

  // Extrair a data
  pixDetails.dateOfTransaction = extractDateFromText(textoDoComprovante);

  //Extrair o nome
  pixDetails.nameOfPayer = extractPayerName(textoDoComprovante);

  //Extrair o ID do PIX
  pixDetails.idOfPIX = extractPIXId(textoDoComprovante);

  return pixDetails;
}

function extractDateFromText(textoDoComprovante) {
      var regex = /(\d{2}\/\d{2}\/\d{4})/; 
    var match = textoDoComprovante.match(regex);
    Logger.log(match);
    if (match) {
        return match[1];
    }
    return null;
}

function extractPayerName(textoDoComprovante) {
  // Lista de possíveis padrões baseados nos exemplos fornecidos
    var patterns = [
        /Dados do pagador\s+Nome\s+([^\n]+)/,         // Padrão para Caixa Econômica Federal
        /CLIENTE:\s*([^\n]+)/,                       // Padrão para Banco do Brasil
        /Origem\s+Nome\s+([^\n]+)/,                    // Padrão para Nubank
        /Dados de quem pagou\s+Nome:\s+([^\n]+)/
    ];

    for (var i = 0; i < patterns.length; i++) {
        var match = textoDoComprovante.match(patterns[i]);
        if (match && match[1]) {
            return match[1].trim(); // Retorna o nome encontrado
        }
    }

    return null; // Retorna null se nenhum nome for encontrado
}

function extractPIXId(textoDoComprovante) {
    // Lista de possíveis padrões baseados nos exemplos fornecidos
    var patterns = [
        /ID:\s*([\w\d-]+)/,                                     // Padrão para Banco do Brasil
        /ID transação\s+([\w\d-]+)/,                            // Padrão para Caixa Econômica Federal
        /ID da transação:\s+([\w\d-]+)/,                        // Padrão para Nubank
        /Número de Controle:\s+([\w\d-]+)/                      
    ];

    for (var i = 0; i < patterns.length; i++) {
        var match = textoDoComprovante.match(patterns[i]);
        if (match && match[1]) {
            return match[1].trim(); // Retorna o ID encontrado
        }
    }

    return null; // Retorna null se nenhum ID for encontrado
}


function extraiNomeDoRemetente(remetente) {
  Logger.log(remetente);
  var match = remetente.match(/"?(.+?)"?\s+<.+>/);
  if (match) {
    return match[1];
  } else {
    return remetente;  // Se não conseguirmos extrair um nome, retorne o valor original (que é o endereço de e-mail)
  }
}


function extractEmail(remetente) {
  var emailRegex = /[a-zA-Z0-9._-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,6}/;
  var match = remetente.match(emailRegex);
  return match ? match[0] : null;  // Retorna o primeiro e-mail encontrado ou null se nenhum for encontrado
}


