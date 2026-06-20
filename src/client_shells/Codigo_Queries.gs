/**
 * ARQUIVO LOCAL (CASCA)
 */

function onOpen() {
  RHC_Core.instalarMenu();
}

function abrirSidebarTratamento() {
  RHC_Core.abrirSidebarTratamento();
}

// =========================================================================
// FUNÇÕES DE PONTE (BRIDGE) 
// Essenciais para que o HTML da biblioteca consiga falar com o GS da biblioteca
// =========================================================================

function obterContextoAtual() { return RHC_Core.obterContextoAtual(); }
function carregarEstadosBloqueio() { return RHC_Core.carregarEstadosBloqueio(); }
function executarMigracaoQParaPasso01() { return RHC_Core.executarMigracaoQParaPasso01(); }
function gerarProximoPasso() { return RHC_Core.gerarProximoPasso(); }
function pintarLinhaSelecionada(hex) { return RHC_Core.pintarLinhaSelecionada(hex); }
function filtrarLinhasPorCoresMulti(cores) { return RHC_Core.filtrarLinhasPorCoresMulti(cores); }
function verificarHeaders() { return RHC_Core.verificarHeaders(); }
function copiarBDExemplo() { return RHC_Core.copiarBDExemplo(); }
function aplicarValidacoesIniciais() { return RHC_Core.aplicarValidacoesIniciais(); }
function inserirColunasClinicas() { return RHC_Core.inserirColunasClinicas(); }
function aplicarValidacoesClinicas() { return RHC_Core.aplicarValidacoesClinicas(); }
function executarVerificacaoDuplicados(data) { return RHC_Core.executarVerificacaoDuplicados(data); }
function sugerirMesReferencia() { return RHC_Core.sugerirMesReferencia(); }
function salvarEstadoBloqueio(id, status) { return RHC_Core.salvarEstadoBloqueio(id, status); }

// =========================================================================
// PONTES DO PASSO 4 (GESTOR CLÍNICO)
// =========================================================================
function abrirGestorClinico() { return RHC_Core.abrirGestorClinico(); }
function obterEstatisticasPasso4() { return RHC_Core.obterEstatisticasPasso4(); }
function obterPacientesPasso4() { return RHC_Core.obterPacientesPasso4(); }
function salvarPacienteGestor(linha, dados) { return RHC_Core.salvarPacienteGestor(linha, dados); }

// ---> AS DUAS FUNÇÕES QUE FALTAVAM <---
function obterListasDropdown() { return RHC_Core.obterListasDropdown(); }
function obterDadosPaciente(linha) { return RHC_Core.obterDadosPaciente(linha); }
