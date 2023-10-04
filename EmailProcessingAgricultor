//processa os emails do agricultor que contem o extrato OFX da conta bancária
function processEmailsAndOFX() {
  var searchQuery = "from:annek.borges@hotmail.com";
  var threads = GmailApp.search(searchQuery);
  var extratosFolder = DriveApp.getFolderById('1JXNxVDYFV5iKrfJPyRB1xI73am9CrZkA');
  var extratosProcessedFolder = DriveApp.getFolderById('1HEdGw98_IZX85mWva8neAg56tV-Jni49');
    
  threads.forEach(thread => {
      thread.getMessages().forEach(message => {
          var messageId = message.getId();
          if(!isMessageProcessed(messageId)){
            message.getAttachments().forEach(attachment => {
              if (isValidOFXAttachment(attachment)) {
                var ofxFile = extratosFolder.createFile(attachment);
                processOFXFile(ofxFile); 
                ofxFile.moveTo(extratosProcessedFolder);
                addProcessedMessageId(messageId);  // Registra o ID da mensagem como processado
              };
            });  
          };
      });
  });
}

function processOFXFile(fileOFX) {
    var xmlContent = extractXmlContentFromOFX(fileOFX);
    if (!xmlContent) return;

    var transactions = getTransactionsFromXml(xmlContent);
    transactions.forEach(transaction => {
        var type = transaction.getChildText('TRNTYPE');
        if (type === 'CREDIT') {
            var data = formatTransactionData(transaction);
            addDataToSheet(data); // Esta função é do SheetProcessing.gs
        }
    });
}


function extractXmlContentFromOFX(fileOFX) {
    var content = fileOFX.getBlob().getDataAsString();
    var startPos = content.indexOf('<OFX>');
    if (startPos === -1) {
        Logger.log('Tag <OFX> não encontrada.');
        return null;
    }
    return content.substring(startPos);
}

function getTransactionsFromXml(xmlContent) {
    var xml;
    try {
        xml = XmlService.parse(xmlContent);
    } catch (e) {
        Logger.log('Erro ao analisar o conteúdo XML: ' + e.toString());
        return [];
    }

    var root = xml.getRootElement();
    return root.getChild('BANKMSGSRSV1').getChild('STMTTRNRS').getChild('STMTRS').getChild('BANKTRANLIST').getChildren('STMTTRN');
}

function formatTransactionData(transaction) {
    var valorString = transaction.getChildText('TRNAMT');
    var valorFloat = parseFloat(valorString.replace(',', '.'));
    return {
        dateTransaction: formatOFXDate(transaction.getChildText('DTPOSTED')),
        amountTransaction: valorFloat,
        nameOfSender: formatDescription(transaction.getChildText('MEMO')),
        idOfTransaction: transaction.getChildText(`FITID`)
    };
}


//Extrai o nome de quem fez a transação a partir do extrato
function formatDescription(rawDescription) {
    var match = rawDescription.match(/-(.*?)-/);
    return match && match[1] ? match[1].trim() : rawDescription;
}

//extrai a data e formata para dd/mm/aaa do extrato
function formatOFXDate(rawDate) {
  Logger.log(rawDate);
    return `${rawDate.substr(6, 2)}/${rawDate.substr(4, 2)}/${rawDate.substr(0, 4)}`;
}

function isValidOFXAttachment(attachment) {
    return attachment.getContentType() === 'application/ofx' && attachment.getName().endsWith('.ofx');
}
