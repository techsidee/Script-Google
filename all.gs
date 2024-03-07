// Função para verificar se um e-mail já existe na planilha
function verificarEmailExistente(planilha, data, assunto) {
  var range = planilha.getDataRange();
  var valores = range.getValues();
  for (var i = 0; i < valores.length; i++) {
    var rowData = valores[i];
    // Convertendo a string de data de volta para objeto Date
    var rowDataDate = new Date(rowData[0]);
    // Verificando se a data e o assunto correspondem
    if (rowDataDate.getTime() === data.getTime() && rowData[1] === assunto) {
      return true;
    }
  }
  return false;
}

function importarEmailsParaPlanilha() {
  var planilha = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Testando");
  var etiqueta = "Allmail2"; 
  var label = GmailApp.getUserLabelByName(etiqueta);

  if (label !== null) {
    var threads = label.getThreads();
    for (var i = 0; i < threads.length; i++) {
      var mensagens = threads[i].getMessages();
      for (var j = 0; j < mensagens.length; j++) {
        var data = mensagens[j].getDate();
        var assunto = mensagens[j].getSubject();
        var corpo = mensagens[j].getPlainBody();

        var tipo = "Padrão";
        var numeroContrato = "";
        var valorAnterior = "";
        var novoValor = "";

        if (corpo.includes("Parcela foi modificado de")) {
          tipo = "Alteração de parcela";
          numeroContrato = corpo.match(/\(-\d+\.\d+\)/) ? corpo.match(/\(-\d+\.\d+\)/)[0].match(/\d+/)[0] : "Não encontrado";
          valorAnterior = corpo.match(/Parcela foi modificado de (\d+[\.,]\d+)/) ? corpo.match(/Parcela foi modificado de (\d+[\.,]\d+)/)[1] : "Não encontrado";
          novoValor = corpo.match(/para (\d+[\.,]\d+)/) ? corpo.match(/para (\d+[\.,]\d+)/)[1] : "Não encontrado";
        } else if (corpo.includes("Valor de Liberacao foi modificado de")) {
          tipo = "Alteração de troco";
          numeroContrato = corpo.match(/\(-\d+\.\d+\)/) ? corpo.match(/\(-\d+\.\d+\)/)[0].match(/\d+/)[0] : "Não encontrado";
          valorAnterior = corpo.match(/foi modificado de (\d+[\.,]\d+)/) ? corpo.match(/foi modificado de (\d+[\.,]\d+)/)[1] : "Não encontrado";
          novoValor = corpo.match(/para (\d+[\.,]\d+)/) ? corpo.match(/para (\d+[\.,]\d+)/)[1] : "Não encontrado";
        } else if (corpo.includes("Tabela foi modificado de")) {
          tipo = "Alteração de tabela";
          numeroContrato = corpo.match(/\(-\d+\.\d+\)/) ? corpo.match(/\(-\d+\.\d+\)/)[0].match(/\d+/)[0] : "Não encontrado";
          valorAnterior = corpo.match(/Tabela foi modificado de (.+) para/) ? corpo.match(/Tabela foi modificado de (.+) para/)[1] : "Não encontrado";
          novoValor = corpo.match(/para (.+)\./) ? corpo.match(/para (.+)\./)[1] : "Não encontrado";
        }

        // Verifica se o e-mail já existe na planilha
        var existeEmail = verificarEmailExistente(planilha, data, assunto);

        if (!existeEmail) {
          // Adiciona o e-mail à planilha apenas se ele não existir
          if (planilha) {
            planilha.appendRow([data, assunto, numeroContrato, tipo, valorAnterior, novoValor]);
          } else {
            Logger.log("Planilha não encontrada.");
          }
        }
      }
    }
  } else {
    Logger.log("A etiqueta '" + etiqueta + "' não foi encontrada.");
  }
}
