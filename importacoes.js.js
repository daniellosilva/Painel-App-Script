const COLUNA_DATA_REFERENCIA = 12; 
const COLUNA_M_FORMATAR = 13;     
const COLUNA_VALOR_REFERENCIA = 6; 
const ULTIMA_COLUNA_PARA_DUPLICATAS = 24; 

function executarImportacaoComBarra() {
  Logger.log("a barra abre")
  limparProgresso(); // sempre resetar antes de abrir o modal
  const template = HtmlService.createTemplateFromFile("barraProgresso");
  const html = template.evaluate()
    .setWidth(400)
    .setHeight(130);
  SpreadsheetApp.getUi().showModalDialog(html, "Importando...");
}

function mostrarTodasAsLinhas() {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const aba = planilha.getSheetByName("Relatórios");
  const totalLinhas = aba.getMaxRows();

  aba.showRows(1, totalLinhas);
  showAlert(`Todas as linhas foram exibidas na aba "${aba.getName()}".`);
}

function importarSemDuplicatas() {
  Logger.log("a importação começa")
  atualizarProgresso(0, "Iniciando importação...");

  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const abaOrigem = planilha.getSheetByName("Relatório Acompanhamento (Volátil)");
  const abaDestino = planilha.getSheetByName("Relatórios");

  atualizarProgresso(10, "Lendo dados da origem...");
  const dadosOrigemComCabecalhoCompleto = abaOrigem.getRange(1, 1, abaOrigem.getLastRow(), abaOrigem.getLastColumn()).getValues();

  if (dadosOrigemComCabecalhoCompleto.length < 2) {
    atualizarProgresso(100, "Aba origem vazia.");
    showAlert("A aba origem está vazia ou sem dados para copiar.");
    return;
  }

  atualizarProgresso(25, "Analisando duplicatas...");
  const cabecalhoOrigem = dadosOrigemComCabecalhoCompleto[0];
  const dadosOrigemParaComparacao = dadosOrigemComCabecalhoCompleto.slice(1).map(row => row.slice(0, ULTIMA_COLUNA_PARA_DUPLICATAS));
  const ultimaLinhaDestino = abaDestino.getLastRow();

  let dadosDestinoParaComparacao = [];
  if (ultimaLinhaDestino >= 2) {
    dadosDestinoParaComparacao = abaDestino.getRange(2, 1, ultimaLinhaDestino - 1, ULTIMA_COLUNA_PARA_DUPLICATAS).getValues();
  }

  const chavesDestino = new Set(dadosDestinoParaComparacao.map(linha => createRowKey(linha)));

  atualizarProgresso(40, "Filtrando novas linhas...");
  const novasLinhas = dadosOrigemParaComparacao.filter(linha => {
    const chave = createRowKey(linha);
    return !chavesDestino.has(chave);
  });

  if (novasLinhas.length > 0) {
    atualizarProgresso(55, "Preparando novas linhas...");
    const dataAtual = new Date();
    const novasLinhasComDataCompletas = dadosOrigemComCabecalhoCompleto.slice(1).filter(linhaCompleta => {
      const linhaParaComparacao = linhaCompleta.slice(0, ULTIMA_COLUNA_PARA_DUPLICATAS);
      const chave = createRowKey(linhaParaComparacao);
      return !chavesDestino.has(chave);
    }).map(linha => [...linha, dataAtual]);

    atualizarProgresso(70, "Inserindo dados no destino...");
    if (ultimaLinhaDestino === 0) {
      abaDestino.getRange(1, 1, 1, cabecalhoOrigem.length + 1).setValues([[...cabecalhoOrigem, "Data Adição"]]);
      abaDestino.getRange(2, 1, novasLinhasComDataCompletas.length, novasLinhasComDataCompletas[0].length).setValues(novasLinhasComDataCompletas);
    } else {
      abaDestino.getRange(ultimaLinhaDestino + 1, 1, novasLinhasComDataCompletas.length, novasLinhasComDataCompletas[0].length).setValues(novasLinhasComDataCompletas);
    }

    atualizarProgresso(80, "Executando pós-processamento...");
    formatarColunaMComoData();
    preencherPrevisaoFaturamentoMes();
    preencherPrevisaoFaturamento();
    calcularValorTotalAteFimAno();
    preencherMesesComValor();

    const ultimaLinha = abaDestino.getLastRow();
    if (ultimaLinha >= 2) {
      abaDestino.getRange(2, 38, ultimaLinha - 1).setNumberFormat("#,##0.00");
    }

    atualizarProgresso(100, "Importação concluída.");
    showAlert(`${novasLinhasComDataCompletas.length} nova(s) linha(s) foram adicionadas ao relatório.`);
  } else {
    atualizarProgresso(100, "Nada novo para importar.");
    showAlert("Nenhuma nova linha foi adicionada. Todos os dados já estavam presentes.");
  }
}

function calcularValorTotalAteFimAno() {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const aba = planilha.getSheetByName("Relatórios");
  const dados = aba.getDataRange().getValues();

  const COLUNA_G = 6; 
  const COLUNA_H = 7; 
  const cabecalho = dados[0];

  let colPrevisao = cabecalho.indexOf("Previsão de faturamento");
  if (colPrevisao === -1) {
    showAlert("Coluna 'Previsão de faturamento' não encontrada.");
    return;
  }

  let colTotal = cabecalho.indexOf("Valor Total");
  if (colTotal === -1) {
    colTotal = cabecalho.length;
    aba.getRange(1, colTotal + 1).setValue("Valor Total");
  }

  for (let i = 1; i < dados.length; i++) {
    const linha = dados[i];

    const valorG = Number(
      String(linha[COLUNA_G]).replace(/\./g, "").replace(",", ".")
    );
    const valorH = linha[COLUNA_H];
    const previsao = linha[colPrevisao];

    let valorTotal = "";

    if (!valorG || isNaN(valorG) || valorG === 0) { 
      valorTotal = valorH === true || valorH === "true" ? true : valorH;
    } else {
      let data;
      if (previsao instanceof Date) {
        data = previsao;
      } else {
        try {
          data = Utilities.parseDate(previsao, planilha.getSpreadsheetTimeZone(), "dd/MM/yyyy");
        } catch (e) {
          data = new Date(NaN);
        }
      }

      if (!isNaN(data.getTime())) {
        const mes = data.getMonth(); // 0 a 11
        const mesesRestantes = 11 - mes + 1;
        valorTotal = valorG * mesesRestantes;
      } else {
        valorTotal = valorG;
      }
    }

    aba.getRange(i + 1, colTotal + 1).setValue(valorTotal);
  }

  aba.getRange(2, colTotal + 1, aba.getLastRow() - 1).setNumberFormat("#,##0.00");
  showAlert("Coluna 'Valor Total' atualizada com base no fim do ano.");
}

function preencherPrevisaoFaturamento() {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const aba = planilha.getSheetByName("Relatórios");
  const dados = aba.getDataRange().getValues();

  const colDataReferencia = COLUNA_DATA_REFERENCIA; 
  const cabecalho = dados[0];

  let colPrevisao = cabecalho.indexOf("Previsão de faturamento");

  if (colPrevisao === -1) {
    colPrevisao = cabecalho.length;
    aba.getRange(1, colPrevisao + 1).setValue("Previsão de faturamento");
  }

  for (let i = 1; i < dados.length; i++) {
    const linha = dados[i];
    const dataValorDaCelula = linha[colDataReferencia]; 

    let data;
    if (dataValorDaCelula instanceof Date) {
        data = dataValorDaCelula;
    } else {
        try {
            data = Utilities.parseDate(dataValorDaCelula, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "dd/MM/yyyy");
        } catch (e) {
            data = new Date(NaN); 
        }
    }

    if (isNaN(data.getTime())) {
      continue;
    }

    const previsao = new Date(data);
    previsao.setMonth(previsao.getMonth() + 1); 
    previsao.setDate(1); 
    previsao.setHours(0, 0, 0, 0); 
    
    aba.getRange(i + 1, colPrevisao + 1).setValue(previsao);
  }

  const ultimaLinha = aba.getLastRow();
  if (ultimaLinha > 1) {
    aba.getRange(2, colPrevisao + 1, ultimaLinha - 1).setNumberFormat("dd/MM/yyyy");
  }
  showAlert("Coluna 'Previsão de faturamento' preenchida e formatada.");
}

function preencherPrevisaoFaturamentoMes() {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const aba = planilha.getSheetByName("Relatórios");
  const dados = aba.getDataRange().getValues();

  const COLUNA_DATA_REFERENCIA = 12; // Ajuste aqui para o índice correto (ex.: coluna L = 12 -> índice = 11)
  
  const cabecalho = dados[0];
  let colPrevisao = cabecalho.indexOf("Previsão de faturamento (Mês)");

  if (colPrevisao === -1) {
    colPrevisao = cabecalho.length;
    aba.getRange(1, colPrevisao + 1).setValue("Previsão de faturamento (Mês)");
  }

  for (let i = 1; i < dados.length; i++) {
    const linha = dados[i];
    const dataValorDaCelula = linha[COLUNA_DATA_REFERENCIA]; 

    let data;
    if (dataValorDaCelula instanceof Date) {
        data = dataValorDaCelula;
    } else {
        try {
            data = Utilities.parseDate(dataValorDaCelula, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "dd/MM/yyyy");
        } catch (e) {
            data = new Date(NaN); 
        }
    }

    if (isNaN(data.getTime())) {
      aba.getRange(i + 1, colPrevisao + 1).setValue(""); 
      continue;
    }

    const previsao = new Date(data);
    previsao.setMonth(previsao.getMonth() + 1); 

    const mes = previsao.getMonth() + 1; 

    aba.getRange(i + 1, colPrevisao + 1).setValue(mes);
  }

  const ultimaLinha = aba.getLastRow();
  if (ultimaLinha > 1) {
    aba.getRange(2, colPrevisao + 1, ultimaLinha - 1).setNumberFormat("00");
  }
  showAlert("Coluna 'Previsão de faturamento (Mês)' preenchida.");
}

function formatarColunaMComoData(){
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const aba = planilha.getSheetByName("Relatórios");

  const ultimaLinha = aba.getLastRow();

  if (ultimaLinha > 1){
    showAlert("A aba origem não está disponível ou está vazia!")
  }
}

function preencherMesesComValor() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Relatórios");
  
  // Configuração inicial com tratamento de erros
  if (!sheet) {
    showAlert("A aba 'Relatórios' não foi encontrada!");
    return;
  }

  // 1. Preparação dos meses
  const nomesMeses = [
    "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
    "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
  ];

  const meses = nomesMeses.map(mes => `${mes} (2025)`);
  meses.push("Janeiro (2026)");

  // 2. Obtenção dos dados da planilha (uma única leitura)
  const ultimaLinha = sheet.getLastRow();
  const ultimaColuna = sheet.getLastColumn();
  
  if (ultimaLinha < 2) {
    showAlert("Nenhum dado encontrado para processar.");
    return;
  }

  const cabecalho = sheet.getRange(1, 1, 1, ultimaColuna).getValues()[0];
  
  // 3. Localização das colunas com tratamento de erro
  const colunasNecessarias = {
    "MRR CALCULADO": cabecalho.indexOf("MRR CALCULADO"),
    "Valor": cabecalho.indexOf("Valor"),
    "Previsão de faturamento": cabecalho.indexOf("Previsão de faturamento"),
    "Duração contrato": cabecalho.indexOf("Duração contrato")
  };

  for (const [nome, indice] of Object.entries(colunasNecessarias)) {
    if (indice === -1) {
      showAlert(`Coluna '${nome}' não encontrada.`);
      return;
    }
  }

  // 4. Criação de novas colunas se necessário
  let novaColuna = ultimaColuna;
  const colunasExistentes = new Set();
  
  meses.forEach(mes => {
    if (cabecalho.indexOf(mes) === -1 && !colunasExistentes.has(mes)) {
      novaColuna++;
      sheet.getRange(1, novaColuna).setValue(mes);
      colunasExistentes.add(mes);
    }
  });

  // 5. Processamento dos dados
  const [novoCabecalho] = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues();
  const dados = sheet.getRange(2, 1, ultimaLinha - 1, novoCabecalho.length).getValues();

  // Limpeza e preenchimento
  for (let i = 0; i < dados.length; i++) {
    const linha = dados[i];
    
    // Limpa valores anteriores
    meses.forEach(mes => {
      const colMes = novoCabecalho.indexOf(mes);
      if (colMes !== -1) linha[colMes] = "";
    });

    // Processa apenas linhas válidas
    const valorG = parseNumeroBR(linha[colunasNecessarias["MRR CALCULADO"]]);
    const valorH = parseNumeroBR(linha[colunasNecessarias["Valor"]]);
    const dataPrevisao = linha[colunasNecessarias["Previsão de faturamento"]];
    const qtdMesesContrato = parseInt(linha[colunasNecessarias["Duração contrato"]], 10);

    if (!(dataPrevisao instanceof Date) || isNaN(qtdMesesContrato) || qtdMesesContrato <= 0) {
      continue;
    }

    const valorMensal = (!isNaN(valorG) && valorG !== 0) ? valorG : (!isNaN(valorH) ? valorH : null);
    if (valorMensal === null) continue;

    // Preenche os meses
    const dataClone = new Date(dataPrevisao);
    for (let m = 0; m < qtdMesesContrato; m++) {
      const mesIndex = dataClone.getMonth();
      const ano = dataClone.getFullYear();

      if (ano > 2026 || (ano === 2026 && mesIndex > 0)) break;

      const nomeMes = `${nomesMeses[mesIndex]} (${ano})`;
      const colMes = novoCabecalho.indexOf(nomeMes);

      if (colMes !== -1) {
        linha[colMes] = valorMensal;
      }

      dataClone.setMonth(dataClone.getMonth() + 1);
    }
  }

  // 6. Escrita dos dados de volta à planilha
  sheet.getRange(2, 1, dados.length, novoCabecalho.length).setValues(dados);
  
  // 7. Formatação das colunas
  meses.forEach(mes => {
    const colMes = novoCabecalho.indexOf(mes);
    if (colMes !== -1) {
      sheet.getRange(2, colMes + 1, dados.length).setNumberFormat("#,##0.00");
    }
  });

  showAlert("Valores preenchidos com sucesso!");
}