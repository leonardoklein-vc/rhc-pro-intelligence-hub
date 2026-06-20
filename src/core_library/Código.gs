// =========================================================================
// RHC PRO V2 - AUTOMAÇÃO E TRATAMENTO DE DADOS
// =========================================================================

function instalarMenu() {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('🤖 RHC Pro v2')
      .addItem('🛠️ Abrir Automações', 'abrirSidebarTratamento');

  var nomeArquivo = SpreadsheetApp.getActiveSpreadsheet().getName().toLowerCase().replace(/\s/g, '');
  if (nomeArquivo.includes('passo4')) {
     menu.addSeparator();
     menu.addItem('💻 Abrir Registro RHC (Passo 4)', 'abrirRegistroRHC');
  }
  
  menu.addToUi();
}

function abrirRegistroRHC() {
  var html = HtmlService.createTemplateFromFile('Gestor')
      .evaluate()
      .setWidth(1150)
      .setHeight(700)
      .setTitle('RHC Pro v2 - Registro de Câncer');
      
  SpreadsheetApp.getUi().showModelessDialog(html, 'Registro RHC');
}

// 3. Função para gerar os dados do Sidebar (O que falta fazer?)
// =========================================================================
// ESTATÍSTICAS DA BARRA LATERAL (INDEX)
// =========================================================================
function obterEstatisticasPasso4() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getGuiaIndicadores(ss); // Busca a guia correta
  
  var dados = sheet.getDataRange().getValues();
  if (dados.length < 2) return { total: 0, concluidos: 0, pendentes: 0, pct: 0 };
  
  var headers = dados[0];
  var colEstadiamento = headers.indexOf("Estadiamento cT"); 
  
  if (colEstadiamento === -1) return { erro: "Colunas clínicas ausentes. Rode o Passo 3 primeiro." };

  var totalValidos = 0;
  var concluidos = 0;

  for (var i = 1; i < dados.length; i++) {
    if (!dados[i][0]) continue; 
    totalValidos++;
    
    if (dados[i][colEstadiamento] && dados[i][colEstadiamento].toString().trim() !== "") {
      concluidos++;
    }
  }

  var pct = totalValidos === 0 ? 0 : Math.round((concluidos / totalValidos) * 100);
  
  return {
    total: totalValidos,
    concluidos: concluidos,
    pendentes: totalValidos - concluidos,
    pct: pct
  };
}

function abrirSidebarTratamento() {
  // A biblioteca gera o HTML a partir dos arquivos dela mesma
  var html = HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Ferramentas Auxiliares');
  SpreadsheetApp.getUi().showSidebar(html);
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🤖 RHC Pro v2')
      .addItem('🛠️ Abrir Automações', 'abrirSidebarTratamento')
      .addToUi();
}

function abrirSidebarTratamento() {
  // ATENÇÃO: Como usamos a tag <?!= include('style'); ?> no HTML, 
  // precisamos usar 'createTemplateFromFile' e '.evaluate()' em vez de HtmlOutput direto.
  var html = HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Ferramentas Auxiliares');
      
  SpreadsheetApp.getUi().showSidebar(html);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// =========================================================================
// INTELIGÊNCIA DE CONTEXTO E FLUXO DE ARQUIVOS (NOVO)
// =========================================================================

/**
 * Identifica o contexto do arquivo e verifica a memória de geração de próximos passos.
 */
function obterContextoAtual() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getDocumentProperties();
  const nomeArquivo = ss.getName();
  
  // Leitura da Memória da Planilha
  const urlProximo = props.getProperty('URL_PROXIMO_PASSO');
  const proximoGerado = props.getProperty('PROXIMO_GERADO') === 'true';

  const headers = ss.getActiveSheet().getRange(1, 1, 1, 28).getValues()[0];
  const isQueryHeaders = (headers[0] === "MES_ANO" && headers[1] === "NOME" && headers[2] === "PRONTUARIO" && headers[27] === "CONSENTIMENTO");
  
  let passoIdentificado = "BLOQUEADO";
  let nomePastaMes = "DESCONHECIDO";
  let nomePastaPai = "DESCONHECIDO";

  try {
    const file = DriveApp.getFileById(ss.getId());
    const pais = file.getParents();
    
    if (pais.hasNext()) {
      const pastaMes = pais.next();
      nomePastaMes = pastaMes.getName();
      const avos = pastaMes.getParents();
      if (avos.hasNext()) {
        const pastaPai = avos.next();
        nomePastaPai = pastaPai.getName();

        if (nomePastaPai.includes("Queries") && isQueryHeaders) {
          passoIdentificado = "Q";
        } else if (nomePastaPai.includes("Passo 1")) {
          passoIdentificado = "1";
        } else if (nomePastaPai.includes("Passo 2")) {
          passoIdentificado = "2";
        } else if (nomePastaPai.includes("Passo 3")) {
          passoIdentificado = "3";
        } else if (nomePastaPai.includes("Passo 4")) {
          passoIdentificado = "4";
        }
      }
    }
  } catch (e) {
    // Fallback por nome se o DriveApp falhar
    const nomeLower = nomeArquivo.toLowerCase().replace(/\s/g, '');
    if (nomeLower.includes("query") && isQueryHeaders) passoIdentificado = "Q";
    else if (nomeLower.includes("passo1")) passoIdentificado = "1";
    else if (nomeLower.includes("passo2")) passoIdentificado = "2";
    else if (nomeLower.includes("passo3")) passoIdentificado = "3";
    else if (nomeLower.includes("passo4")) passoIdentificado = "4";
  }

  // Define qual seria o próximo número de passo para o botão
  let proximoNum = "1";
  if (passoIdentificado !== "Q" && passoIdentificado !== "BLOQUEADO") {
    proximoNum = (parseInt(passoIdentificado) + 1).toString();
  }

  return {
    passo: passoIdentificado,
    nomeArquivo: nomeArquivo,
    pastaMes: nomePastaMes,
    pastaPai: nomePastaPai,
    proximoGerado: proximoGerado,
    urlProximo: urlProximo,
    passoSeguinte: proximoNum
  };
}

/**
 * PASSO 0: Migra Query para Passo 1 e grava o link na memória.
 */
function executarMigracaoQParaPasso01() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const file = DriveApp.getFileById(ss.getId());
    const props = PropertiesService.getDocumentProperties();
    
    let folderMesAtual, folderQueries, folderAno;
    try {
      folderMesAtual = file.getParents().next(); 
      folderQueries = folderMesAtual.getParents().next(); 
      folderAno = folderQueries.getParents().next(); 
    } catch (e) {
      return { erro: "Erro de hierarquia: O arquivo deve estar em Ano > Queries > Mês." };
    }
    
    let folderPasso1 = getFolder(folderAno, "Passo 1");
    const nomeMesExtenso = folderMesAtual.getName();
    const anoStr = folderQueries.getName().replace(/\D/g,''); 
    const mesCod = converterMesParaMM(nomeMesExtenso);
    const nomePastaDestino = mesCod + (anoStr || new Date().getFullYear());
    
    let folderDestino = getFolder(folderPasso1, nomePastaDestino);
    const novoNome = file.getName().replace(/Query/i, "Passo1");
    
    // Verificação física no Drive
    const arquivosExistentes = folderDestino.getFilesByName(novoNome);
    if (arquivosExistentes.hasNext()) {
      const arqExistente = arquivosExistentes.next();
      props.setProperty('PROXIMO_GERADO', 'true');
      props.setProperty('URL_PROXIMO_PASSO', arqExistente.getUrl());
      return { erro: "O arquivo já existe no Passo 1. Botão atualizado.", url: arqExistente.getUrl(), passoGerado: "1" };
    }
    
    const novaPlanilha = file.makeCopy(novoNome, folderDestino);
    
    // GRAVAÇÃO DA MEMÓRIA
    props.setProperty('PROXIMO_GERADO', 'true');
    props.setProperty('URL_PROXIMO_PASSO', novaPlanilha.getUrl());

    return {
      msg: "Sucesso! Migrado para Passo 1 > " + nomePastaDestino,
      url: novaPlanilha.getUrl(),
      passoGerado: "1"
    };
  } catch (e) {
    return { erro: "Erro crítico: " + e.message };
  }
}

/**
 * GERA PRÓXIMO PASSO: Cria cópia para a pasta seguinte e grava na memória.
 */
function gerarProximoPasso() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const file = DriveApp.getFileById(ss.getId());
    const props = PropertiesService.getDocumentProperties();
    const currentContext = obterContextoAtual(); 
    
    if (currentContext.passo === "BLOQUEADO" || currentContext.passo === "Q") {
      return { erro: "Ação não permitida para este arquivo." };
    }
    
    const passoAtualNum = parseInt(currentContext.passo);
    const proximoNum = passoAtualNum + 1;
    const nomeProximoPasso = "Passo " + proximoNum;
    
    let folderPassoAtual, folderAno;
    try {
      folderPassoAtual = file.getParents().next().getParents().next();
      folderAno = folderPassoAtual.getParents().next();
    } catch (e) {
      return { erro: "Erro ao acessar pastas do Drive." };
    }
    
    const folderPassoDestino = getFolder(folderAno, nomeProximoPasso);
    const folderMesDestino = getFolder(folderPassoDestino, currentContext.pastaMes);
    
    const novoNome = file.getName().replace(/Passo\s*0?\d/i, "Passo" + proximoNum);
    
    // Verificação física no Drive
    const arquivosExistentes = folderMesDestino.getFilesByName(novoNome);
    if (arquivosExistentes.hasNext()) {
      const arqExistente = arquivosExistentes.next();
      props.setProperty('PROXIMO_GERADO', 'true');
      props.setProperty('URL_PROXIMO_PASSO', arqExistente.getUrl());
      return { erro: "Arquivo já existe no próximo passo. Memória atualizada.", url: arqExistente.getUrl(), passoGerado: proximoNum.toString() };
    }
    
    const novaPlanilha = file.makeCopy(novoNome, folderMesDestino);
    
    // GRAVAÇÃO DA MEMÓRIA
    props.setProperty('PROXIMO_GERADO', 'true');
    props.setProperty('URL_PROXIMO_PASSO', novaPlanilha.getUrl());

    return {
      msg: "Cópia para o Passo " + proximoNum + " gerada!",
      url: novaPlanilha.getUrl(),
      passoGerado: proximoNum.toString()
    };
  } catch (e) {
    return { erro: "Erro ao gerar: " + e.message };
  }
}

// Helpers que você já tem (MANTIDOS)
function getFolder(parent, name) {
  const iter = parent.getFoldersByName(name);
  return iter.hasNext() ? iter.next() : parent.createFolder(name);
}

function converterMesParaMM(extenso) {
  const meses = {
    'JANEIRO': '01', 'FEVEREIRO': '02', 'MARÇO': '03', 'ABRIL': '04', 
    'MAIO': '05', 'JUNHO': '06', 'JULHO': '07', 'AGOSTO': '08', 
    'SETEMBRO': '09', 'OUTUBRO': '10', 'NOVEMBRO': '11', 'DEZEMBRO': '12'
  };
  return meses[extenso.toString().toUpperCase().trim()] || '00';
}

// =========================================================================
// LÓGICA DO PASSO 1: PREPARAÇÃO
// =========================================================================

function verificarHeaders() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.setName("Indicadores");

  var headersDesejados = [
    "OBSERVAÇÕES", 
    "DIAGNÓSTICO E TRATAMENTOS ANTERIORES", 
    "CID DIAGNÓSTICO", 
    "DATA DA PRIMEIRA CONSULTA", 
    "DATA DIAGNÓSTICO", 
    "LEI 60 DIAS", 
    "DATA DO PRIMEIRO TRATAMENTO"
  ];
  
  var headersAtuais = sheet.getRange(1, 4, 1, 7).getValues()[0];
  var tudoOk = true;
  
  for (var i = 0; i < headersDesejados.length; i++) {
    if (headersAtuais[i] !== headersDesejados[i]) {
      tudoOk = false;
      break;
    }
  }
  
  if (!tudoOk) {
    sheet.insertColumnsAfter(3, 7);
    sheet.getRange(1, 4, 1, 7).setValues([headersDesejados]);
    sheet.getRange(1, 4, 1, 7).setFontWeight("bold").setBackground("#f3f3f3");
  }

  SpreadsheetApp.flush();
  Utilities.sleep(200);

  var maxCol = Math.min(35, sheet.getLastColumn());
  for (var c = 1; c <= maxCol; c++) {
    sheet.autoResizeColumn(c);
  }
  
  return tudoOk ? "Tudo ok! Estrutura já estava correta." : "Colunas inseridas e guia redimensionada.";
}

function copiarBDExemplo() {
  var ssAtivo = SpreadsheetApp.getActiveSpreadsheet();
  var idFonte = "1kzn7fHWQcCimDEWRLOf5ARVK5iUJOZ89aXbyQGuuQQ8";
  
  try {
    var ssFonte = SpreadsheetApp.openById(idFonte);
    var abaFonte = ssFonte.getSheetByName("BD Exemplo");
    
    if (!abaFonte) return "Erro: Aba 'BD Exemplo' não encontrada na planilha fonte.";
    
    var abaExistente = ssAtivo.getSheetByName("BD");
    if (abaExistente) {
      ssAtivo.deleteSheet(abaExistente);
    }
    
    var novaAba = abaFonte.copyTo(ssAtivo);
    novaAba.setName("BD");
    novaAba.hideSheet(); 
    
    return "Comando 2 concluído: Banco de Dados 'BD' importado e atualizado.";
  } catch (e) {
    return "Erro ao acessar planilha externa: Verifique as permissões.";
  }
}

function aplicarValidacoesIniciais() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetAtiva = ss.getActiveSheet();
  
  if (sheetAtiva.getName() !== "Indicadores") {
    return "Erro: Esta ação só pode ser executada na aba 'Indicadores'.";
  }

  var sheetBD = ss.getSheetByName("BD");
  if (!sheetBD) return "Erro: Importe o BD primeiro (Comando 2).";
  
  var lastRow = sheetAtiva.getLastRow();
  if (lastRow < 2) return "Aba vazia. Nenhuma validação foi aplicada.";
  
  var rangeE = sheetAtiva.getRange(2, 5, lastRow - 1, 1);
  var ruleE = SpreadsheetApp.newDataValidation().requireValueInRange(sheetBD.getRange("Q2:Q5"), true).build();
  rangeE.setDataValidation(ruleE);
  
  var rangeF = sheetAtiva.getRange(2, 6, lastRow - 1, 1);
  var maxRowR = sheetBD.getRange("R:R").getValues().filter(String).length; 
  var ruleF = SpreadsheetApp.newDataValidation().requireValueInRange(sheetBD.getRange("R2:R" + maxRowR), true).build();
  rangeF.setDataValidation(ruleF);
  
  var colsParaData = [7, 8, 10, 12, 15, 31, 32, 34];
  if (sheetAtiva.getMaxColumns() < 35) {
    sheetAtiva.insertColumnsAfter(sheetAtiva.getMaxColumns(), 35 - sheetAtiva.getMaxColumns());
  }
  colsParaData.forEach(function(col) {
    sheetAtiva.getRange(2, col, lastRow - 1, 1).setNumberFormat("dd/mm/yyyy");
  });

  sheetAtiva.getRange("I1").setFormula('={"LEI 60 DIAS"; ARRAYFORMULA(IF((J2:J="") + (H2:H=""); ""; J2:J - H2:H))}');
  sheetAtiva.getRange("I2:I").setNumberFormat("0"); 
  
  try {
    var protecao = sheetAtiva.getRange("I1").protect().setDescription('Cálculo Automático - Não Editar');
    protecao.removeEditors(protecao.getEditors());
    if (protecao.canDomainEdit()) protecao.setDomainEdit(false);
  } catch (e) { }

  var filtroExistente = sheetAtiva.getFilter();
  if (filtroExistente) filtroExistente.remove();
  sheetAtiva.getRange(1, 1, lastRow, 35).createFilter();
  
  return "Comando 3 concluído: Validações iniciais e Filtro Mestre aplicados.";
}

// =========================================================================
// LÓGICA DO PASSO 2: INDICADORES (Pintura e Filtro de Cores)
// =========================================================================

function pintarLinhaSelecionada(corHex) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sheet.getActiveRange();
  var startRow = range.getRow();
  var numRows = range.getNumRows();
  sheet.getRange(startRow, 1, numRows, 35).setBackground(corHex);
}

function filtrarLinhasPorCoresMulti(coresHexArray) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return "Erro: A planilha está vazia.";
  
  sheet.showRows(2, lastRow - 1);
  if (coresHexArray.indexOf("TODAS") !== -1) return "Todas as linhas exibidas.";
  
  var bgColors = sheet.getRange(2, 9, lastRow - 1, 1).getBackgrounds();
  var coresValidas = coresHexArray.map(function(cor) { return cor.toLowerCase(); });
  var hideStart = -1;
  var hideCount = 0;
  
  for (var i = 0; i < bgColors.length; i++) {
    var row = i + 2;
    var corCelula = bgColors[i][0].toLowerCase();
    
    if (coresValidas.indexOf(corCelula) === -1) {
       if (hideStart === -1) { hideStart = row; hideCount = 1; } 
       else { hideCount++; }
    } else {
       if (hideStart !== -1) {
         sheet.hideRows(hideStart, hideCount);
         hideStart = -1; hideCount = 0;
       }
    }
  }
  if (hideStart !== -1) sheet.hideRows(hideStart, hideCount);
  
  return "Tabela filtrada (Base: Coluna I).";
}

// =========================================================================
// LÓGICA DO PASSO 3: CAMPOS CLÍNICOS E VALIDAÇÕES ESPECÍFICAS
// =========================================================================

function inserirColunasClinicas() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (sheet.getName() !== "Indicadores") return "Erro: Esta ação deve ser feita na aba 'Indicadores'.";
  
  var colsClinicas = [
    "Convênio", 
    "Localização do tumor primário", 
    "Tipo histológico (CID-O)", 
    "Base mais importante do diagnóstico", 
    "Estadiamento cT", 
    "Estadiamento cN", 
    "Estadiamento cM", 
    "Estadiamento TNM patológico", 
    "Outro estadiamento", 
    "Estadiamento patológico - T (0 - in situ; X-desconhecido)", 
    "Estadiamento patológico - N (0- Ausência em linfonodos regionais; X-desconhecido)", 
    "Estadiamento patológico - M (0 - Ausência de metástase a distância; X-desconhecido)", 
    "Localização de metástase à distância I", 
    "Localização de metástase à distância II", 
    "Localização de metástase à distância III", 
    "Localização de metástase à distância IV", 
    "Clínica do início do tratamento", 
    "Principal razão para não realização do primeiro tratamento no hospital", 
    "1º tratamento realizado no hospital", 
    "Estado da doença ao final do primeiro tratamento", 
    "Data do óbito", 
    "Óbito devido ao câncer"
  ];
  
  sheet.insertColumnsAfter(10, colsClinicas.length);
  var range = sheet.getRange(1, 11, 1, colsClinicas.length);
  
  range.setValues([colsClinicas]);
  range.setFontWeight("bold").setBackground("#cfe2ff").setHorizontalAlignment("center");
  
  SpreadsheetApp.flush();
  
  for (var i = 11; i < 11 + colsClinicas.length; i++) {
    sheet.autoResizeColumn(i);
  }
  
  return "Comando 1 concluído: 22 novas colunas clínicas inseridas entre J e K.";
}

function aplicarValidacoesClinicas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var sheetBD = ss.getSheetByName("BD");
  
  if (!sheetBD) return "Erro: Aba 'BD' não encontrada para validação.";
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return "Sem dados para validar.";
  
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  // MAPA DE REGRAS ATUALIZADO (Tumor em AB, Histológico em X)
  var mapaRegras = [
    { label: "Convênio", bd: "A2:A5" },
    { label: "Base mais importante do diagnóstico", bd: "D2:D10" },
    { label: "Localização do tumor primário", bd: "AB2:AB350" }, // ATUALIZADO
    { label: "Tipo histológico (CID-O)", bd: "X2:X800" },        // ATUALIZADO
    { label: "Estadiamento cT", bd: "AD2:AD40" },
    { label: "Estadiamento cN", bd: "AE2:AE40" },
    { label: "Estadiamento cM", bd: "AF2:AF40" },
    { label: "Estadiamento TNM patológico", bd: "AG2:AG40" },
    { label: "Estadiamento patológico - T (0 - in situ; X-desconhecido)", bd: "AH2:AH40" },
    { label: "Estadiamento patológico - N (0- Ausência em linfonodos regionais; X-desconhecido)", bd: "AI2:AI40" },
    { label: "Estadiamento patológico - M (0 - Ausência de metástase a distância; X-desconhecido)", bd: "AJ2:AJ40" },
    { label: "Localização de metástase à distância I", bd: "R2:R100" },
    { label: "Localização de metástase à distância II", bd: "R2:R100" },
    { label: "Localização de metástase à distância III", bd: "R2:R100" },
    { label: "Localização de metástase à distância IV", bd: "R2:R100" },
    { label: "Clínica do início do tratamento", bd: "AK2:AK30" },
    { label: "Principal razão para não realização do primeiro tratamento no hospital", bd: "M2:M10" },
    { label: "1º tratamento realizado no hospital", bd: "N2:N11" },
    { label: "Estado da doença ao final do primeiro tratamento", bd: "O2:O10" },
    { label: "Óbito devido ao câncer", bd: "P2:P5" }
  ];
  
  mapaRegras.forEach(function(regra) {
    var colIdx = headers.indexOf(regra.label);
    if (colIdx !== -1) {
      var cellRange = sheet.getRange(2, colIdx + 1, lastRow - 1, 1);
      var validationRule = SpreadsheetApp.newDataValidation()
          .requireValueInRange(sheetBD.getRange(regra.bd), true)
          .setAllowInvalid(false)
          .build();
      cellRange.setDataValidation(validationRule);
    }
  });
  
  var colObitoIdx = headers.indexOf("Data do óbito");
  if (colObitoIdx !== -1) {
    var rangeObito = sheet.getRange(2, colObitoIdx + 1, lastRow - 1, 1);
    var dateRule = SpreadsheetApp.newDataValidation()
        .requireDate()
        .setAllowInvalid(false)
        .setHelpText("Insira uma data válida (DD/MM/AAAA)")
        .build();
    rangeObito.setDataValidation(dateRule);
    rangeObito.setNumberFormat("dd/mm/yyyy");
  }
  
  return "Comando 2 concluído: Todas as validações clínicas foram aplicadas com sucesso.";
}

// =========================================================================
// VARREDURA DE DUPLICADOS E ACESSO A DADOS EXTERNOS
// =========================================================================

function sugerirMesReferencia() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var valoresA = sheet.getRange(2, 1, Math.min(20, sheet.getLastRow()), 1).getValues();
  var sugestao = "";
  var regexData = /(jan|fev|mar|abr|mai|jun|jul|ago|set|out|nov|dez)[a-z]*[-/\s_]*(\d{2,4})/i;

  for (var i = 0; i < valoresA.length; i++) {
    var val = valoresA[i][0] ? String(valoresA[i][0]).trim() : "";
    if (!val || val === "DUP" || val === "C44" || val.match(/^\d{2}\/\d{4}$/) || val.match(/^C44 - \d{2}\/\d{4}$/)) continue;
    var match = val.match(regexData);
    if (match) {
      var mes = match[1].toLowerCase();
      var ano = match[2];
      if (ano.length === 2) ano = "20" + ano;
      var mesesNum = {jan:'01',fev:'02',mar:'03',abr:'04',mai:'05',jun:'06',jul:'07',ago:'08',set:'09',out:'10',nov:'11',dez:'12'};
      sugestao = mesesNum[mes] + "/" + ano;
      break;
    }
  }
  return sugestao;
}

function executarVerificacaoDuplicados(mesAnoRefStr) {
  var debugLogs = [];
  debugLogs.push("--- INÍCIO DA DEPURAÇÃO: REF " + mesAnoRefStr + " ---");

  if (!mesAnoRefStr) {
    debugLogs.push("ERRO: Mês de referência inválido.");
    return { msg: "Erro: Mês de referência inválido.", logs: debugLogs };
  }

  var partes = mesAnoRefStr.split('/');
  var mesRef = parseInt(partes[0], 10);
  var anoRef = parseInt(partes[1], 10);
  
  var dataRefNum = anoRef * 100 + mesRef; 
  var dataLimiteInfNum = (anoRef - 1) * 100 + mesRef; 

  var sheetLocal = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var idExterna = "1kzn7fHWQcCimDEWRLOf5ARVK5iUJOZ89aXbyQGuuQQ8";
  var extTimezone;

  try {
    var ssExt = SpreadsheetApp.openById(idExterna);
    var sheetExt = ssExt.getSheetByName("BD_REGISTROS");
    if (!sheetExt) throw new Error("Aba BD_REGISTROS não encontrada");
    extTimezone = ssExt.getSpreadsheetTimeZone();
  } catch(e) {
    debugLogs.push("ERRO: Não foi possível acessar a planilha externa.");
    return { msg: "Erro ao acessar a planilha de consulta externa.", logs: debugLogs };
  }

  var headersLocal = sheetLocal.getRange(1, 1, 1, sheetLocal.getLastColumn()).getDisplayValues()[0];
  var hLocal = mapearColunasDuplicados(headersLocal);

  var headersExt = sheetExt.getRange(1, 1, 1, sheetExt.getLastColumn()).getDisplayValues()[0];
  var hExt = mapearColunasDuplicados(headersExt);

  if ((hLocal.cid_tratado === -1 && hLocal.cid_orig === -1) || (hExt.cid_tratado === -1 && hExt.cid_orig === -1)) {
    debugLogs.push("ERRO: Colunas de CID ausentes.");
    return { msg: "Erro: Coluna de 'CID' não encontrada.", logs: debugLogs };
  }

  if (hExt.data === -1) {
    debugLogs.push("ERRO: Coluna de DATA ausente no Histórico.");
    return { msg: "Erro Crítico: Coluna C não mapeada na BD_REGISTROS.", logs: debugLogs };
  }

  var dadosLocal = sheetLocal.getDataRange().getDisplayValues();
  var dadosExt = sheetExt.getDataRange().getDisplayValues();
  var ocorrenciasExternas = [];

  function getMesAnoIntSeguroTextual(valorCelula, timezone) {
    if (!valorCelula) return null;
    if (valorCelula instanceof Date) {
      var yyyyMM = Utilities.formatDate(valorCelula, timezone, "yyyyMM");
      return parseInt(yyyyMM, 10);
    }
    var str = String(valorCelula).trim().toLowerCase();
    var regexMes = /(jan|fev|mar|abr|mai|jun|jul|ago|set|out|nov|dez)[a-z]*[\.\-\/\s_]*(\d{2,4})/i;
    var matchMes = str.match(regexMes);
    if (matchMes) {
      var mesesNum = {jan:'01',fev:'02',mar:'03',abr:'04',mai:'05',jun:'06',jul:'07',ago:'08',set:'09',out:'10',nov:'11',dez:'12'};
      var ano = matchMes[2];
      if (ano.length === 2) ano = "20" + ano;
      return parseInt(ano + mesesNum[matchMes[1]], 10);
    }
    var matchBr = str.match(/(?:^|\D)(0[1-9]|[12]\d|3[01])[\/\-](0[1-9]|1[0-2])[\/\-](\d{4})/);
    if (matchBr) return parseInt(matchBr[3] + matchBr[2], 10);
    var matchMY = str.match(/(?:^|\D)(0[1-9]|1[0-2])[\/\-](\d{4})/);
    if (matchMY) return parseInt(matchMY[2] + matchMY[1], 10);
    return null;
  }

  function numToMesAnoStr(num) {
    var s = String(num);
    return s.substring(4,6) + "/" + s.substring(0,4); 
  }

  function extrairRaizCid(cidStr) {
    if (!cidStr) return "";
    var match = String(cidStr).trim().match(/^[A-Z][0-9]{2}/i);
    return match ? match[0].toUpperCase() : String(cidStr).trim().toUpperCase();
  }

  function isStandardCID(cidStr) {
    var str = String(cidStr).trim();
    return /^[A-Z]\d{2}(\.\d)?(\s|-|$)/i.test(str);
  }

  function checkIdentity(p1, c1, n1, p2, c2, n2) {
    return ((p1 && p1 === p2) || (c1 && c1 === c2) || (n1 && n1 === n2));
  }

  for (var i = 1; i < dadosExt.length; i++) {
    var linha = dadosExt[i];
    var valorTextoBruto = linha[hExt.data]; 
    var valRegistroNum = getMesAnoIntSeguroTextual(valorTextoBruto, extTimezone);
    if (valRegistroNum !== null) {
      if (valRegistroNum >= dataLimiteInfNum && valRegistroNum < dataRefNum) {
        var extCidTratado = hExt.cid_tratado !== -1 ? linha[hExt.cid_tratado] : "";
        var extCidOrig = hExt.cid_orig !== -1 ? linha[hExt.cid_orig] : "";
        var extCid = extCidTratado ? extCidTratado : extCidOrig;
        ocorrenciasExternas.push({
          prontuario: hExt.pront !== -1 ? String(linha[hExt.pront]).trim() : "",
          pac_codigo: hExt.pac_cod !== -1 ? String(linha[hExt.pac_cod]).trim() : "",
          nome: hExt.nome !== -1 ? String(linha[hExt.nome]).trim().toUpperCase() : "",
          cidRaiz: extrairRaizCid(extCid),
          mesOrigem: numToMesAnoStr(valRegistroNum),
          dadoBrutoLog: valorTextoBruto 
        });
      }
    }
  }

  var backgrounds = sheetLocal.getDataRange().getBackgrounds();
  var valuesToUpdateA = [];
  var bgToUpdate = [];
  var countDups = 0;
  var countC44 = 0;

  for (var i = 1; i < dadosLocal.length; i++) {
    var linha = dadosLocal[i];
    var lPront = hLocal.pront !== -1 ? String(linha[hLocal.pront]).trim() : "";
    var lPacCod = hLocal.pac_cod !== -1 ? String(linha[hLocal.pac_cod]).trim() : "";
    var lNome = hLocal.nome !== -1 ? String(linha[hLocal.nome]).trim().toUpperCase() : "";
    if (!lPront && !lPacCod && !lNome) {
      valuesToUpdateA.push([linha[0]]);
      bgToUpdate.push(backgrounds[i]);
      continue;
    }
    var cidTratado = hLocal.cid_tratado !== -1 ? linha[hLocal.cid_tratado] : "";
    var cidOrig = hLocal.cid_orig !== -1 ? linha[hLocal.cid_orig] : "";
    var lCidStr = cidTratado ? String(cidTratado) : String(cidOrig);
    var lCidRaiz = extrairRaizCid(lCidStr);
    var lCidIsStandard = isStandardCID(lCidStr);
    var dupLocal = false, dupExtMes = "", dupDadoBruto = "", ehC44 = false;
    var nomeLog = lNome ? lNome : ("Pront. " + lPront);
    var identicalPrior = false, hasBetterElsewhere = false;

    for (var j = 1; j < dadosLocal.length; j++) {
      if (i === j) continue;
      var rOutro = dadosLocal[j];
      var oPront = hLocal.pront !== -1 ? String(rOutro[hLocal.pront]).trim() : "";
      var oPacCod = hLocal.pac_cod !== -1 ? String(rOutro[hLocal.pac_cod]).trim() : "";
      var oNome = hLocal.nome !== -1 ? String(rOutro[hLocal.nome]).trim().toUpperCase() : "";
      if (checkIdentity(lPront, lPacCod, lNome, oPront, oPacCod, oNome)) {
        if (j < i) identicalPrior = true;
        var oCidTratado = hLocal.cid_tratado !== -1 ? rOutro[hLocal.cid_tratado] : "";
        var oCidOrig = hLocal.cid_orig !== -1 ? rOutro[hLocal.cid_orig] : "";
        var oCidStr = oCidTratado ? String(oCidTratado) : String(oCidOrig);
        if (isStandardCID(oCidStr)) { hasBetterElsewhere = true; }
      }
    }
    if (!lCidIsStandard) {
      if (hasBetterElsewhere) { dupLocal = true; } 
      else if (identicalPrior) { dupLocal = true; }
    }
    if (!dupLocal && lCidRaiz !== "") {
      for (var k = 0; k < ocorrenciasExternas.length; k++) {
        var ext = ocorrenciasExternas[k];
        if (ext.cidRaiz === lCidRaiz) {
          if (checkIdentity(lPront, lPacCod, lNome, ext.prontuario, ext.pac_codigo, ext.nome)) {
            dupExtMes = ext.mesOrigem; 
            dupDadoBruto = ext.dadoBrutoLog; 
            if (lCidRaiz === "C44") { ehC44 = true; }
            break;
          }
        }
      }
    }

    var newColA = linha[0]; 
    var bgRow = backgrounds[i]; 
    if (dupLocal) {
      newColA = mesAnoRefStr; countDups++;
      for (var c = 0; c < bgRow.length; c++) bgRow[c] = "#808080"; 
      debugLogs.push("Paciente " + nomeLog + ", localizada duplicação LOCAL (" + mesAnoRefStr + ")");
    } else if (dupExtMes !== "") {
      if (ehC44) {
        newColA = "C44 - " + dupExtMes; countC44++;
        for (var c = 0; c < bgRow.length; c++) bgRow[c] = "#d3d3d3"; 
        debugLogs.push("Paciente " + nomeLog + ", localizada duplicação C44, conforme " + dupExtMes + " [Lido: '" + dupDadoBruto + "']");
      } else {
        newColA = dupExtMes; countDups++;
        for (var c = 0; c < bgRow.length; c++) bgRow[c] = "#808080"; 
        debugLogs.push("Paciente " + nomeLog + ", localizada duplicação, conforme " + dupExtMes + " [Lido: '" + dupDadoBruto + "']");
      }
    } else { debugLogs.push("Paciente " + nomeLog + ", não localizada duplicação."); }
    valuesToUpdateA.push([newColA]);
    bgToUpdate.push(bgRow);
  }

  if (valuesToUpdateA.length > 0) {
    sheetLocal.getRange(2, 1, valuesToUpdateA.length, 1).setValues(valuesToUpdateA);
    sheetLocal.getRange(2, 1, bgToUpdate.length, bgToUpdate[0].length).setBackgrounds(bgToUpdate);
  }

  var msgResumo = "Verificação concluída: " + countDups + " Duplicados | " + countC44 + " C44.";
  debugLogs.push(msgResumo);
  return { msg: msgResumo, logs: debugLogs };
}

function mapearColunasDuplicados(headers) {
  var map = { pront: -1, pac_cod: -1, nome: -1, cid_tratado: -1, cid_orig: -1, data: 2 };
  for (var i = 0; i < headers.length; i++) {
    var h = String(headers[i]).trim().toUpperCase();
    if (h === "PRONTUARIO" || h === "PRONTUÁRIO") map.pront = i;
    else if (h === "PAC_CODIGO" || h === "CÓDIGO PACIENTE") map.pac_cod = i;
    else if (h === "NOME" || h === "PACIENTE" || h === "NOME DO PACIENTE") map.nome = i;
    else if (h === "CID DIAGNÓSTICO" || h === "CID DIAGNOSTICO") map.cid_tratado = i;
    else if (h === "CID") map.cid_orig = i;
  }
  if (map.cid_tratado === -1 && map.cid_orig === -1) {
    for (var j = 0; j < headers.length; j++) {
      if (String(headers[j]).trim().toUpperCase().indexOf("CID") !== -1) { map.cid_orig = j; break; }
    }
  }
  return map;
}

// =========================================================================
// MEMÓRIA DE ESTADO (CADEADOS E BLOQUEIOS)
// =========================================================================

function salvarEstadoBloqueio(passoId, isLocked) {
  var props = PropertiesService.getDocumentProperties();
  props.setProperty('passo_' + passoId + '_bloqueado', isLocked ? 'true' : 'false');
  return true;
}

function carregarEstadosBloqueio() {
  var props = PropertiesService.getDocumentProperties();
  return {
    passo1: props.getProperty('passo_1_bloqueado') === 'true',
    passo2: props.getProperty('passo_2_bloqueado') === 'true',
    passo3: props.getProperty('passo_3_bloqueado') === 'true'
  };
}

// =========================================================================
// BACKEND DO GESTOR CLÍNICO (PASSO 4)
// =========================================================================

function obterListasDropdown() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetBD = ss.getSheetByName("BD");
  if (!sheetBD) return { erro: "Aba BD não encontrada." };

  // Usando getDisplayValues para NUNCA converter tumor em data
  function getUniques(rangeA1) {
    var vals = sheetBD.getRange(rangeA1).getDisplayValues();
    var list = [];
    for (var i = 0; i < vals.length; i++) {
      var val = vals[i][0];
      if (val !== null && val !== undefined && val.toString().trim() !== "") {
        list.push(val.toString().trim());
      }
    }
    return list;
  }

  return {
    convenios: getUniques("A2:A5"),
    baseDiag: getUniques("D2:D10"),
    tumor: getUniques("AB2:AB350"),       // ATUALIZADO
    histologico: getUniques("X2:X800"),   // ATUALIZADO
    cT: getUniques("AD2:AD40"),
    cN: getUniques("AE2:AE40"),
    cM: getUniques("AF2:AF40"),
    tnmPat: getUniques("AG2:AG40"),
    pT: getUniques("AH2:AH40"),
    pN: getUniques("AI2:AI40"),
    pM: getUniques("AJ2:AJ40"),
    metastase: getUniques("R2:R100"),
    clinica: getUniques("AK2:AK30"),
    razaoNaoTrat: getUniques("M2:M10"),
    primTrat: getUniques("N2:N11"),
    estadoDoenca: getUniques("O2:O10"),
    obitoCancer: getUniques("P2:P5")
  };
}

// =========================================================================
// O NOVO SISTEMA DE CACHE DO REGISTRO RHC (COM FILTRO CORRIGIDO)
// =========================================================================
function obterPacientesPasso4() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getGuiaIndicadores(ss); // Busca a guia correta
  
  var dados = sheet.getDataRange().getDisplayValues(); 
  if (dados.length < 2) return {};

  var headers = dados[0];
  
  var idxNome = headers.indexOf("Nome do paciente");
  if(idxNome === -1) idxNome = headers.indexOf("NOME");
  if(idxNome === -1) idxNome = headers.indexOf("PACIENTE");
  
  var idxPront = headers.indexOf("PRONTUARIO");
  if(idxPront === -1) idxPront = headers.indexOf("PRONTUÁRIO");

  var idxStatus = headers.indexOf("Estadiamento cT");

  var pacientesMap = {}; 
  
  for (var i = 1; i < dados.length; i++) {
    
    // --- O NOVO FILTRO INTELIGENTE ---
    // Pega o nome e prontuário da linha atual
    var valNome = idxNome !== -1 ? dados[i][idxNome].toString().trim() : "";
    var valPront = idxPront !== -1 ? dados[i][idxPront].toString().trim() : "";
    
    // Se a coluna A, o Nome E o Prontuário estiverem TODOS vazios, aí sim é uma linha inútil, então pula.
    if (!dados[i][0] && valNome === "" && valPront === "") continue; 
    // ---------------------------------
    
    var objLinha = {};
    for (var j = 0; j < headers.length; j++) {
       if (headers[j]) objLinha[headers[j].toString().trim()] = dados[i][j];
    }
    
    var nomeFinal = valNome !== "" ? valNome : "Sem Nome";
    
    objLinha._linha = i + 1;
    objLinha._nome = nomeFinal;
    objLinha._prontuario = valPront;
    objLinha._concluido = (idxStatus !== -1 && dados[i][idxStatus].toString().trim() !== "");
    
    pacientesMap[i + 1] = objLinha; 
  }
  return pacientesMap;
}

// =========================================================================
// SALVAMENTO SEGURO
// =========================================================================
function salvarPacienteGestor(linha, dadosForm) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = getGuiaIndicadores(ss); // Busca a guia correta
    
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    for (var key in dadosForm) {
      var colIdx = headers.indexOf(key);
      if (colIdx !== -1) {
        var val = dadosForm[key];
        var cell = sheet.getRange(linha, colIdx + 1);
        
        if (key === "Data do óbito" && val) {
           var partes = val.split("-");
           if(partes.length === 3) val = partes[2] + "/" + partes[1] + "/" + partes[0]; 
        }
        cell.setValue(val);
      }
    }
    return { sucesso: true, msg: "Ficha clínica salva com sucesso!" };
  } catch (e) {
    return { sucesso: false, erro: e.message };
  }
}

// =========================================================================
// FUNÇÃO AUXILIAR DE BUSCA DA ABA "INDICADORES"
// =========================================================================
function getGuiaIndicadores(ss) {
  var abas = ss.getSheets();
  for (var i = 0; i < abas.length; i++) {
    // Procura por qualquer aba que contenha a palavra "indicadores" no nome (ignorando maiúsculas/minúsculas)
    if (abas[i].getName().toLowerCase().indexOf('indicadores') !== -1) {
      return abas[i];
    }
  }
  // Se por acaso alguém deletou ou renomeou tudo errado, tenta a ativa como última esperança
  return ss.getActiveSheet(); 
}
