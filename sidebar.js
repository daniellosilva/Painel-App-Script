function abrirSidebarRetratil() {
  const template = HtmlService.createTemplateFromFile('sidebarRetratil');
  const html = template.evaluate()
    .setTitle("Painel Sauter")
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}
