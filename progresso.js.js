// Versão melhorada com sincronização de estado
const CACHE_KEY = "progressState_v2";

function atualizarProgresso(percent, message) {
  const state = {
    percent: Math.min(100, Math.max(0, percent)), // Garante 0-100%
    message: message || "Processando...",
    lastUpdated: new Date().toISOString()
  };
  
  // Atualiza ambos os lugares sincronizadamente
  const stateJSON = JSON.stringify(state);
  CacheService.getScriptCache().put(CACHE_KEY, stateJSON, 21600); // 6 horas
}

function getProgressState() {
  try {
    const stateJSON = CacheService.getScriptCache().get(CACHE_KEY);
    return stateJSON ? JSON.parse(stateJSON) : {
      percent: 0,
      message: "Nenhuma operação ativa",
      lastUpdated: null
    };
  } catch (e) {
    console.error("Erro ao ler progresso:", e);
    return {
      percent: 0,
      message: "Erro ao carregar estado",
      lastUpdated: null
    };
  }
}

function limparProgresso() {
  CacheService.getScriptCache().remove(CACHE_KEY);
}

// Exemplo de uso:
function processamentoLongo() {
  atualizarProgresso(0, "Iniciando importação...");
  
  for (let i = 1; i <= 10; i++) {
    Utilities.sleep(1000); // Simula trabalho
    atualizarProgresso(i * 10, `Processando item ${i} de 10`);
  }
  
  atualizarProgresso(100, "Concluído!");
  Utilities.sleep(2000);
  limparProgresso();
}