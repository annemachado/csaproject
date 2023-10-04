function matchPayments() {
  var comprovantesSheet = SpreadsheetApp.openById('1xMM6EnNyO3r0BI_RYtC4Ntqjtkwyi80g9ZQEhdY9UXY').getSheetByName('dados_email_comprovante');
  var extratosSheet = SpreadsheetApp.openById('1xMM6EnNyO3r0BI_RYtC4Ntqjtkwyi80g9ZQEhdY9UXY').getSheetByName('dados_do_extrato');
  let matchedPayments = [];
  let unmatchedPayments = [];

  let extratos = extratosSheet.getRange("A1:E" + extratosSheet.getLastRow()).getValues(); 
  let comprovantes = comprovantesSheet.getRange("A1:I" + comprovantesSheet.getLastRow()).getValues();
  Logger.log(extratos);
  Logger.log(comprovantes);

  extratos.forEach((extratoRow, extratoIndex) => {
    if (extratoRow[4] !== "Verificado") { // Checando se o status é diferente de "Verificado"
      let matched = false;

      for (let i = 0; i < comprovantes.length; i++) {
        let comprovanteRow = comprovantes[i];
        
        if (comprovanteRow[8] !== "Verificado") { // Checando se o status é diferente de "Verificado"
          Logger.log(extratoRow[1]);
          Logger.log(comprovanteRow[6]);
          Logger.log(extratoRow[3]);
          Logger.log(comprovanteRow[5]);
          Logger.log(extratoRow[2].trim().toLowerCase());
          Logger.log(comprovanteRow[7].trim().toLowerCase());
          if (extratoRow[2].trim().toLowerCase() === comprovanteRow[7].trim().toLowerCase() 
            && extratoRow[1] == comprovanteRow[6]
            && extratoRow[3].getTime() === comprovanteRow[5].getTime()) {

            matchedPayments.push({
              extrato: extratoRow,
              comprovante: comprovanteRow
            });
            Logger.log(matchPayments);

            // Atualizar o status para "Verificado" nas planilhas
            extratosSheet.getRange(extratoIndex + 2, 5).setValue("Verificado");
            comprovantesSheet.getRange(i + 2, 9).setValue("Verificado");

            matched = true;
            break;
          }
        }
      }

      if (!matched) {
        unmatchedPayments.push(extratoRow);
        extratosSheet.getRange(extratoIndex + 2, 5).setValue("Comprovante Pendente");
        Logger.log(unmatchedPayments);
      }
    }
  });

    comprovantes.forEach((comprovanteRow, comprovanteIndex) => {
    if (comprovanteRow[8] !== "Verificado") { 
      let matched = false;

      for (let i = 0; i < extratos.length; i++) {
        let extratoRow = extratos[i];

        if (extratoRow[4] !== "Verificado") { 
          if (comprovanteRow[7].trim().toLowerCase() === extratoRow[2].trim().toLowerCase()
            && comprovanteRow[6] == extratoRow[1]
            && comprovanteRow[5].getTime() === extratoRow[3].getTime()) {

            matched = true;
            break;
          }
        }
      }

      if (!matched) {
        // Atualizando a coluna de status para "Extrato Pendente"
        comprovantesSheet.getRange(comprovanteIndex + 2, 9).setValue("Transação não identificada");
      }
    }
  });

  return {
    matchedPayments: matchedPayments,
    unmatchedPayments: unmatchedPayments
  };
}
