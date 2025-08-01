function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Menu")
    .addItem("Abrir Painel Sauter", "abrirSidebarRetratil")
    .addToUi();
}
