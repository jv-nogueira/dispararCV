function realizarBackup() {
  // Defina as informações da planilha de origem ("listEmail")
  var planilhaOrigem = SpreadsheetApp.getActiveSpreadsheet();
  var paginaOrigem = planilhaOrigem.getSheetByName("listEmail");
  var dadosOrigem = paginaOrigem.getRange("A2:A").getValues().filter(String); // Obtém os dados da coluna A

  // Defina as informações da planilha de origem ("Message")
  var planilhaOrigem2 = SpreadsheetApp.getActiveSpreadsheet();
  var paginaOrigem2 = planilhaOrigem2.getSheetByName("Messege");
  var valuesOrigem2 = paginaOrigem2.getDataRange().getValues();

  // Crie um array para armazenar os dados finais
  var dadosFinais = [];

  // Obtém a primeira folha na planilha de destino
  var idPlanilhaDestino = "1inF8gfmVOUeJqyJFjVe7ID1hh7aNKJCCyL7P7yDvzZE"; // Substitua pelo ID da sua planilha de destino
  var planilhaDestino = SpreadsheetApp.openById(idPlanilhaDestino);
  var backup = planilhaDestino.getSheets()[0];

  // Obtém os dados da coluna A na planilha de destino
  var dadosDestino = backup.getRange("A:A").getValues().flat().filter(String);

  // Encontra a próxima linha vazia na coluna A da planilha de destino
  var proximaLinhaVazia = dadosDestino.length + 1;

  // Itere sobre os dados de origem da "listEmail"
  dadosOrigem.forEach(function(valor, index) {
    // Adicione o valor da "listEmail"
    var linhaFinal = [valor];

    // Adicione os valores fixos da "Message"
    linhaFinal.push(
      valuesOrigem2[0][1], // Coluna B, linha 1
      valuesOrigem2[1][1], // Coluna B, linha 2
      valuesOrigem2[2][1], // Coluna B, linha 1
      valuesOrigem2[3][1], // Coluna B, linha 2
      // Adicione mais valores conforme necessário
    );

    // Adicione a data e hora atuais
    var dataHoraAtual = new Date();
    linhaFinal.push(dataHoraAtual);

    // Adicione a linha final ao array de dados finais
    dadosFinais.push(linhaFinal);
  });

  // Obtém a faixa de destino começando da próxima linha vazia
  var faixaDestino = backup.getRange(proximaLinhaVazia, 1, dadosFinais.length, dadosFinais[0].length);

  // Insere os dados na planilha de destino
  faixaDestino.setValues(dadosFinais);
}
