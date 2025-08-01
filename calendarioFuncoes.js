let tipoFiltroGlobal = "";

function abrirCalendario() {
  const html = HtmlService
    .createTemplateFromFile('calendario')
    .evaluate()
    .setWidth(400)
    .setHeight(200)
    .setTitle('Selecionar Data');

  SpreadsheetApp.getUi().showModalDialog(html, 'Selecionar Data');
}

function obterTipoFiltro() {
  return tipoFiltroGlobal;
}

function aplicarFiltroPorData(data, tipoFiltro) {
  const dataSelecionada = new Date(data);

  if (tipoFiltro === "semana") {
    filtrarPorSemanaComData(dataSelecionada);
  } else if (tipoFiltro === "criadas") {
    filtrarOportunidadesCriadasComData(dataSelecionada);
  }
}
