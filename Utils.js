function parseData(dataStr) {
  const partes = dataStr.split('/');
  if (partes.length !== 3) return null;
  const dia = parseInt(partes[0], 10);
  const mes = parseInt(partes[1], 10) - 1;
  const ano = parseInt(partes[2], 10);
  return new Date(ano, mes, dia);
}

function parseNumeroBR(valor) {
  if (typeof valor === "string") {
    return Number(valor.replace(/\./g, "").replace(",", "."));
  }
  return Number(valor);
}

function formatarData(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "dd/MM/yyyy");
}

function formatarMoeda(valor) {
  return "R$ " + valor.toFixed(2).replace(".", ",").replace(/(\d)(?=(\d{3})+\,)/g, "$1.");
}

function showAlert(message) {
  try {
    Logger.log("ALERTA: " + message);
  } catch (e) {
    Logger.log("ALERTA SUPRIMIDO: " + message);
  }
}

function createRowKey(row) {
  const key = row.map(cell => {

    if (cell === null || typeof cell === 'undefined') {
      return '';
    }
    
    if (cell instanceof Date) {

      return Utilities.formatDate(cell, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "yyyy-MM-dd");
    } else if (typeof cell === 'string') {

      return cell.trim().replace(/\s+/g, ' ').replace(/[\n\r]+/g, ' '); 
    } else if (typeof cell === 'number' || typeof cell === 'boolean') {
      return String(cell);
    }
    return String(cell); 
  }).join("§"); 
  
  
  return key;
}