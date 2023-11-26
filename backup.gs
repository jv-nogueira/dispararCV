function realizarBackup() {
    // Defina as informações da planilha de origem
    var planilhaOrigem = SpreadsheetApp.getActiveSpreadsheet();
    var paginaOrigem = planilhaOrigem.getSheetByName("listEmail");
    var dadosOrigem = paginaOrigem.getRange("A2:A").getValues().filter(String); // Obtém os dados da coluna A
  
    // Defina as informações da planilha de destino
    var idPlanilhaDestino = "1inF8gfmVOUeJqyJFjVe7ID1hh7aNKJCCyL7P7yDvzZE"; // Substitua pelo ID da sua planilha de destino
  
    // Abre a planilha de destino diretamente
    var planilhaDestino = SpreadsheetApp.openById(idPlanilhaDestino);
  
    // Obtém a primeira folha na planilha de destino
    var backup = planilhaDestino.getSheets()[0];
  
    // Obtém os dados da coluna A na planilha de destino
    var dadosDestino = backup.getRange("A:A").getValues().flat().filter(String);
  
    // Adiciona os dados de origem nas linhas vazias da coluna A da planilha de destino
    dadosDestino = dadosDestino.concat(dadosOrigem);
  
    // Obtém a faixa de destino começando da primeira linha
    var faixaDestino = backup.getRange(1, 1, dadosDestino.length, 1);
  
    // Insere os dados na coluna de destino
    faixaDestino.setValues(dadosDestino.map(function(value) { return [value]; }));
  
  }
  