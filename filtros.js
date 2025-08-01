function filtrarPorSemanaNaMesmaAba(dataTexto) {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const abaDestino = planilha.getSheetByName("Relatórios");

  if (!dataTexto || typeof dataTexto !== "string") {
    showAlert("Data inválida.");
    return;
  }

  const partes = dataTexto.split("-");
  if (partes.length !== 3) {
    showAlert("Formato inválido. Use AAAA-MM-DD.");
    return;
  }

  const ano = parseInt(partes[0], 10);
  const mes = parseInt(partes[1], 10) - 1;
  const dia = parseInt(partes[2], 10);
  const data = new Date(ano, mes, dia);

  if (isNaN(data.getTime())) {
    showAlert("Data inválida.");
    return;
  }

  const inicioSemana = getSegundaFeiraISO(data);
  const fimSemana = new Date(inicioSemana);
  fimSemana.setDate(inicioSemana.getDate() + 6);
  fimSemana.setHours(23, 59, 59, 999);

  const intervalo = abaDestino.getDataRange();
  const valores = intervalo.getValues();
  const indiceDataAdicao = valores[0].length - 1;

  abaDestino.showRows(1, abaDestino.getMaxRows());

  for (let i = 1; i < valores.length; i++) {
    const valorData = valores[i][indiceDataAdicao];
    const linhaNumero = i + 1;
    if (valorData instanceof Date) {
      if (valorData < inicioSemana || valorData > fimSemana) {
        abaDestino.hideRows(linhaNumero);
      }
    } else {
      abaDestino.hideRows(linhaNumero);
    }
  }

  showAlert(`Filtro aplicado para a semana de ${formatarData(inicioSemana)} a ${formatarData(fimSemana)}.`);
}

function filtrarOportunidadesGanhasPorIntervalo(dataInicialStr, dataFinalStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Relatórios");

  // Converte strings para objetos Date
  const dataInicial = parseDataISO(dataInicialStr);
  const dataFinal = parseDataISO(dataFinalStr);

  if (!dataInicial || !dataFinal || isNaN(dataInicial.getTime()) || isNaN(dataFinal.getTime())) {
    showAlert("Datas inválidas. Certifique-se de que ambas as datas foram selecionadas.");
    return;
  }

  // Garante que a data final seja no fim do dia
  dataFinal.setHours(23, 59, 59, 999);

  sheet.showRows(1, sheet.getMaxRows());

  const dados = sheet.getDataRange().getValues();
  const COL_STATUS = 3;       // Coluna D
  const COL_DATA_ADICAO = 24; // Coluna Y
  let totalOportunidades = 0;

  for (let i = 1; i < dados.length; i++) {
    const status = dados[i][COL_STATUS];
    const dataAdicao = dados[i][COL_DATA_ADICAO];
    const linha = i + 1;

    const manter = (
      status === "Fechado e ganho" &&
      dataAdicao instanceof Date &&
      dataAdicao >= dataInicial &&
      dataAdicao <= dataFinal
    );

    if (manter) {
      totalOportunidades++;
    } else {
      sheet.hideRows(linha);
    }
  }

  const msg = `Oportunidades ganhas entre ${formatarData(dataInicial)} e ${formatarData(dataFinal)}:\n\n` +
              `• Total de oportunidades: ${totalOportunidades}`;
  showAlert("Filtro Aplicado", msg);
}

function parseDataISO(str) {
  const partes = str.split("-");
  if (partes.length !== 3) return null;

  const ano = parseInt(partes[0], 10);
  const mes = parseInt(partes[1], 10) - 1;
  const dia = parseInt(partes[2], 10);

  return new Date(ano, mes, dia);
}

function filtrarOportunidadesCriadas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  // 1ª entrada: Data inicial
  const respostaInicial = ui.prompt('Filtrar oportunidades criadas', 
                                    'Digite a DATA INICIAL (DD/MM/AAAA):', 
                                    ui.ButtonSet.OK_CANCEL);
  
  if (respostaInicial.getSelectedButton() !== ui.Button.OK) return;
  
  const dataInicialTexto = respostaInicial.getResponseText();
  const dataInicial = parseData(dataInicialTexto);
  
  if (!dataInicial || isNaN(dataInicial.getTime())) {
    showAlert('Data inicial inválida! Use o formato DD/MM/AAAA');
    return;
  }
  
  // 2ª entrada: Data final (opcional)
  const respostaFinal = ui.prompt('Filtrar oportunidades criadas', 
                                  'Digite a DATA FINAL (DD/MM/AAAA) ou deixe em branco para usar a data de hoje:', 
                                  ui.ButtonSet.OK_CANCEL);
  
  if (respostaFinal.getSelectedButton() !== ui.Button.OK) return;
  
  const dataFinalTexto = respostaFinal.getResponseText().trim();
  const dataFinal = dataFinalTexto ? parseData(dataFinalTexto) : new Date();
  
  if (!dataFinal || isNaN(dataFinal.getTime())) {
    showAlert('Data final inválida! Use o formato DD/MM/AAAA ou deixe em branco.');
    return;
  }
  
  const sheet = ss.getSheetByName("Relatórios");
  sheet.showRows(1, sheet.getMaxRows());
  
  const dados = sheet.getDataRange().getValues();
  
  const COL_DATA_CRIACAO = 13; // Coluna N (índice 0-based)

  let totalOportunidades = 0;
  
  for (let i = 1; i < dados.length; i++) {
    const dataCriacao = dados[i][COL_DATA_CRIACAO];

    const manterVisivel = (
      dataCriacao instanceof Date &&
      dataCriacao >= dataInicial &&
      dataCriacao <= dataFinal
    );

    if (manterVisivel) {
      totalOportunidades++;
    } else {
      sheet.hideRows(i + 1);
    }
  }

  let msg = `Oportunidades criadas entre ${formatarData(dataInicial)} e ${formatarData(dataFinal)}:\n\n`;
  msg += `• Total de oportunidades: ${totalOportunidades}`;
  
  showAlert('Filtro Aplicado', msg, ui.ButtonSet.OK);
}

function getSegundaFeiraISO(data) {
  const diaSemana = data.getDay();
  const diff = (diaSemana === 0 ? -6 : 1) - diaSemana; 
  const segunda = new Date(data);
  segunda.setDate(data.getDate() + diff);
  segunda.setHours(0, 0, 0, 0); 
  return segunda;
}