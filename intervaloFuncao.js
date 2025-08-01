function abrirCalendarioGanhas() {
  const template = HtmlService.createTemplateFromFile("intervalo")
  const html = template.evaluate()
    .setWidth(350)
    .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, "Selecionar intervalo");
}