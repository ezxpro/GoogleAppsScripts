function obterMapaEstoque()
{
  const idPlanilha = "1J5twCP4L4Smy7FP4JbARD3qVB_mv8rnQ1fRN5y6LvaQ";
  const planilha = SpreadsheetApp.openById(idPlanilha);
  const pagina = planilha.getSheetByName("Stock");
  var intervalo = pagina.getRange('B1:D8');
  var dados = intervalo.getValues();
  var mapaEstoque = new Map();
  for (var i = 0; i < dados.length; i++)
  {
    var produto = dados[i][0];
    var quantidade = dados[i][2];

    mapaEstoque.set(produto, quantidade);
  }
  for (p of mapaEstoque)
  {
    Logger.log(p);
  }
  return mapaEstoque;
}