function gerarSnapshot() {
  const ss = SpreadsheetApp.getActive();
  const consulta = ss.getSheetByName("Consulta");
  const base = ss.getSheetByName("Relatórios");
  
  const dataRef = consulta.getRange("A1").getValue();
  if (!(dataRef instanceof Date)) {
    SpreadsheetApp.getUi().alert("Por favor, selecione uma data válida em Consulta!A1.");
    return;
  }
  
  const dados = base.getDataRange().getValues();
  const header = dados.shift();
  
  const idxId = header.indexOf("ID da oportunidade");
  const idxData = header.indexOf("Data Adição");
  const idxFase = header.indexOf("Fase");
  const idxMRR = header.indexOf("MRR");
  const idxValor = header.indexOf("Valor");
  
  const porId = {};
  
  dados.forEach(linha => {
    const id = linha[idxId];
    const dataAdic = linha[idxData];
    if (!(dataAdic instanceof Date)) return;
    if (dataAdic <= dataRef) {
      if (!porId[id] || dataAdic > porId[id].data) {
        porId[id] = { data: dataAdic, row: linha };
      }
    }
  });
  
  const resumo = {};
  
  Object.values(porId).forEach(({ row }) => {
    const fase = row[idxFase] || "Sem fase";
    const mrr = parseFloat(row[idxMRR]) || 0;
    const valor = parseFloat(row[idxValor]) || 0;
    
    if (!resumo[fase]) {
      resumo[fase] = { mrr: 0, valor: 0, qtd: 0 };
    }
    
    resumo[fase].mrr += mrr;
    resumo[fase].valor += valor;
    resumo[fase].qtd += 1;
  });
  
  const saida = [["Fase", "Quantidade", "Soma MRR", "Soma Valor"]];
  Object.keys(resumo).sort().forEach(fase => {
    const r = resumo[fase];
    saida.push([fase, r.qtd, r.mrr, r.valor]);
  });
  
  const dataFormatada = Utilities.formatDate(dataRef, ss.getSpreadsheetTimeZone(), "dd/MM/yyyy");
  consulta.getRange("A3").setValue(`📌 Pipeline em ${dataFormatada}`);
  
  consulta.getRange("A4:Z1000").clearContent();
  consulta.getRange(4, 1, saida.length, saida[0].length).setValues(saida);
}

function gerarSnapshotHoje() {
  const ss = SpreadsheetApp.getActive();
  const consulta = ss.getSheetByName("Consulta");
  const base = ss.getSheetByName("Relatórios");

  const hoje = new Date();
  hoje.setHours(0, 0, 0, 0);

  const dataReferencia = consulta.getRange("A1").getValue();
  const usarComparacaoCriacao = dataReferencia instanceof Date;

  const dados = base.getDataRange().getValues();
  const header = dados.shift();

  const idxId = header.indexOf("ID da oportunidade");
  const idxData = header.indexOf("Data Adição");
  const idxFase = header.indexOf("Fase");
  const idxMRR = header.indexOf("MRR");
  const idxValor = header.indexOf("Valor");

  const porId = {};

  dados.forEach(linha => {
    const id = linha[idxId];
    const dataAdic = linha[idxData];
    if (!(dataAdic instanceof Date)) return;
    if (!porId[id] || dataAdic > porId[id].data) {
      porId[id] = { data: dataAdic, row: linha };
    }
  });

  const resumo = {};
  let ganhos = 0;
  let perdidos = 0;
  let criados = 0;

  Object.values(porId).forEach(({ row, data }) => {
    const fase = row[idxFase] || "Sem fase";
    const mrr = parseFloat(row[idxMRR]) || 0;
    const valor = parseFloat(row[idxValor]) || 0;

    if (!resumo[fase]) {
      resumo[fase] = { mrr: 0, valor: 0, qtd: 0 };
    }

    resumo[fase].mrr += mrr;
    resumo[fase].valor += valor;
    resumo[fase].qtd += 1;

    if (fase === "Fechado e ganho") ganhos++;
    if (fase === "Fechado e perdido") perdidos++;
    if (usarComparacaoCriacao && data > dataReferencia) criados++;
  });

  const saida = [["Fase", "Quantidade", "Soma MRR", "Soma Valor"]];
  Object.keys(resumo).sort().forEach(fase => {
    const r = resumo[fase];
    saida.push([fase, r.qtd, r.mrr, r.valor]);
  });

  const dataFormatada = Utilities.formatDate(hoje, ss.getSpreadsheetTimeZone(), "dd/MM/yyyy");
  consulta.getRange("F3").setValue(`📌 Resumo atual gerado em ${dataFormatada}`);

  consulta.getRange("F4:H1000").clearContent();
  consulta.getRange(4, 6, saida.length, saida[0].length).setValues(saida);

  const destaques = [
    ["🏁 Oportunidades Ganhas", ganhos],
    ["❌ Oportunidades Perdidas", perdidos],
    ["✳️ Oportunidades Criadas", criados]
  ];
  consulta.getRange("J4:K6").clearContent();
  consulta.getRange("J4:K6").setValues(destaques);
}

function gerarAmbosSnapshots() {
  gerarSnapshot();
  gerarSnapshotHoje();
}
