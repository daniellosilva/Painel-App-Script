function contarLinhasPreenchidas() {
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Relatórios');
  const dados = aba.getDataRange().getValues();
  let count = 0;

  for (let i = 1; i < dados.length; i++) {
    if (dados[i].join("") !== "") count++;
  }

  return count;
}

function buscarOportunidadesPorNome(termo) {
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Relatórios');
  const dados = aba.getDataRange().getValues();

  const resultados = {};
  let maxValor = 0;
  let melhorValorFormatado = "";

  for (let i = 1; i < dados.length; i++) {
    const nome = String(dados[i][2]).toLowerCase();
    const fase = dados[i][3];
    const valorStr = dados[i][7]; // Coluna H, índice 7

    if (nome.includes(termo.toLowerCase())) {
      if (!resultados[fase]) resultados[fase] = 1;
      else resultados[fase]++;

      const valorNumerico = Number(String(valorStr).replace(/[^\d.-]/g, ''));
      if (!isNaN(valorNumerico) && valorNumerico > maxValor) {
        maxValor = valorNumerico;
        melhorValorFormatado = valorStr;
      }
    }
  }

  return {
    quantidade: Object.values(resultados).reduce((a, b) => a + b, 0),
    fases: Object.keys(resultados),
    valorTotal: melhorValorFormatado || "R$ 0,00"
  };
}
