function forcarAutorizacaoTotal() {
  // Isso obriga o motor do Google a gerar um token global para tudo
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  DriveApp.getFiles();
  console.log("Token mestre gerado e salvo. O painel está liberado.");
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🏢 RHC Pro')
      .addItem('🚀 Abrir Painel Central', 'abrirPainelCentral')
      .addSeparator()
      .addItem('⚙️ Preparar Abas', 'prepararAbas')
      .addSeparator()
      .addItem('🔐 Autorizar Acessos (Destravar)', 'forcarAutorizacaoTotal')
      .addSeparator() // <-- Novo separador
      .addItem('🔍 Rastrear Casos Novos', 'destacarNovosCasos') // <-- Nova ferramenta adicionada aqui
      .addToUi();

  // Tenta abrir automaticamente
  try {
    abrirPainelCentral();
  } catch (e) {
    console.warn("Aguardando autorização do usuário.");
  }
}

function abrirPainelCentral() {
  var html = HtmlService.createTemplateFromFile('PainelCentral')
      .evaluate()
      .setTitle('Hub Central de Registros - HCPA')
      .setWidth(1150) 
      .setHeight(750);
  
  // Usar ModalDialog como você já usa hoje
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// =========================================================================
// CONFIGURAÇÃO DAS ABAS COM VALIDAÇÃO SEGURA (INCLUINDO POPS, CÓDIGOS E SNIPPETS)
// =========================================================================

function prepararAbas() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var sheetLinks = ss.getSheetByName("DB_LINKS");
  var sheetRegs = ss.getSheetByName("BD_REGISTROS");
  var sheetPops = ss.getSheetByName("DB_POPS");
  var sheetCodigos = ss.getSheetByName("DB_CODIGOS");
  var sheetSnippets = ss.getSheetByName("DB_SNIPPETS"); // <-- Aba do Cofre
  
  var faltantes = [];
  if (!sheetLinks) faltantes.push("DB_LINKS (Hub de Controle)");
  if (!sheetRegs) faltantes.push("BD_REGISTROS (Mesa de Dados)");
  if (!sheetPops) faltantes.push("DB_POPS (Manuais e Procedimentos)");
  if (!sheetCodigos) faltantes.push("DB_CODIGOS (Repositório de Scripts)");
  if (!sheetSnippets) faltantes.push("DB_SNIPPETS (Cofre de Referências Rápidas)"); // <-- Aviso do Cofre

  if (faltantes.length === 0) {
    ui.alert("Tudo Pronto!", "As abas do sistema já existem na sua planilha. Nenhuma ação é necessária e nenhum dado foi sobrescrito.", ui.ButtonSet.OK);
    return;
  }

  var msg = "Faltam as seguintes abas para o sistema funcionar:\n\n- " + faltantes.join("\n- ") + "\n\nDeseja criá-las agora? Isso NÃO apagará suas outras abas atuais.";
  var resposta = ui.alert("Instalação do Sistema", msg, ui.ButtonSet.YES_NO);

  if (resposta == ui.Button.YES) {
    
    if (!sheetLinks) {
      sheetLinks = ss.insertSheet("DB_LINKS");
      var headersLinks = ["ID", "DATA_CADASTRO", "ANO", "MES", "PASSO", "LINK_PLANILHA", "QTD_LINHAS", "STATUS"];
      sheetLinks.getRange(1, 1, 1, headersLinks.length).setValues([headersLinks]).setFontWeight("bold").setBackground("#1e293b").setFontColor("white");
      sheetLinks.setFrozenRows(1);
      sheetLinks.autoResizeColumns(1, headersLinks.length);
    }
    
    if (!sheetRegs) {
      sheetRegs = ss.insertSheet("BD_REGISTROS");
      var headersRegs = obterColunasAlvo();
      sheetRegs.getRange(1, 1, 1, headersRegs.length).setValues([headersRegs]).setFontWeight("bold").setBackground("#8b5cf6").setFontColor("white");
      sheetRegs.setFrozenRows(1);
    }
    
    if (!sheetPops) {
      sheetPops = ss.insertSheet("DB_POPS");
      var headersPops = ["PASSO", "TITULO", "CONTEUDO", "RESPONSAVEL"];
      sheetPops.getRange(1, 1, 1, headersPops.length).setValues([headersPops]).setFontWeight("bold").setBackground("#0f172a").setFontColor("white");
      sheetPops.setFrozenRows(1);
      popularPopsIniciais(sheetPops);
    }
    
    if (!sheetCodigos) {
      sheetCodigos = ss.insertSheet("DB_CODIGOS");
      var headersCodigos = ["ID", "PROJETO", "TITULO", "CONTEUDO", "DATA_CRIACAO"];
      sheetCodigos.getRange(1, 1, 1, headersCodigos.length).setValues([headersCodigos]).setFontWeight("bold").setBackground("#059669").setFontColor("white");
      sheetCodigos.setFrozenRows(1);
    }

    if (!sheetSnippets) {
      sheetSnippets = ss.insertSheet("DB_SNIPPETS");
      var headersSnippets = ["ID", "NOME", "CONTEUDO", "OBSERVACAO", "DATA"];
      sheetSnippets.getRange(1, 1, 1, headersSnippets.length).setValues([headersSnippets]).setFontWeight("bold").setBackground("#ea580c").setFontColor("white");
      sheetSnippets.setFrozenRows(1);
    }

    ui.alert("Sucesso!", "Estrutura montada com sucesso. Você pode ocultar essas abas se desejar.", ui.ButtonSet.OK);
  }
}

// FUNÇÕES DO COFRE
function buscarSnippets() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DB_SNIPPETS");
  if (!sheet) return [];
  const data = sheet.getDataRange().getDisplayValues();
  if (data.length <= 1) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    let obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
}

function salvarSnippet(id, nome, conteudo, obs) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DB_SNIPPETS");
  const dataStr = Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy HH:mm");
  
  if (id) { // Edição
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === id) {
        sheet.getRange(i + 1, 2, 1, 3).setValues([[nome, conteudo, obs]]);
        return { sucesso: "Snippet atualizado!" };
      }
    }
  } else { // Novo
    const novoId = "SNP" + Utilities.getUuid().split('-')[0].toUpperCase();
    sheet.appendRow([novoId, nome, conteudo, obs, dataStr]);
    return { sucesso: "Snippet guardado!" };
  }
}

function excluirSnippet(id) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DB_SNIPPETS");
  const data = sheet.getRange("A:A").getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      sheet.deleteRow(i + 1);
      return true;
    }
  }
}

function popularPopsIniciais(sheet) {
  var initialData = [
    ["Preparação", "Configuração Inicial", "Aguardando conteúdo...", ""],
    ["Passo 1", "Tratamento Inicial e Automatização", "Novo Passo 1\nResponsável: Léu\nTratamento Inicial e Automatização\n\n- Conversão de XLS para Google Planilhas;\n- Inclusão dos Scripts de Automação;\n- Inclusão das Colunas Úteis para o Passo 2 e da Guia de BD;\n- Revisão das Formatações de Número de todas as colunas;\n- Eliminação de Duplicados;", "Léu"],
    ["Passo 2", "Revisão e Qualificação", "Aguardando conteúdo...", ""],
    ["Passo 3", "Consolidação e Registro", "Aguardando conteúdo...", ""],
    ["Passo 4", "Análise de Consistência", "Aguardando conteúdo...", ""],
    ["Passo 5", "Finalização", "Aguardando conteúdo...", ""]
  ];
  sheet.getRange(2, 1, initialData.length, 4).setValues(initialData);
}

const COLUNAS_SISTEMA = {
  "Passo 1": ["MES_ANO", "NOME", "PRONTUARIO", "PAC_CODIGO", "DT_NASCIMENTO", "CIDDIAGNOSTICO", "DATA", "CPF", "CNS", "NOME_MAE", "IDADE", "ETNIA", "SEXO", "NATURALIDADE", "UF_NASCIMENTO", "ENDERECO", "BAIRRO", "CEP", "CIDADE_MORADIA", "TELEFONE_RES", "TELEFONE_REC", "TELEFONE_CEL", "DT_PRIM_INTERNACAO_PREVIA", "DT_OBITO", "CID_OBITO", "DATA_INICIO_QUIMIO", "CONSENTIMENTO"],
  "Passo 2": ["Revisor", "NOME", "PRONTUARIO", "OBSERVAÇÃO", "DIAGNÓSTICO E TRATAMENTOS ANTERIORES", "CID DIAGNÓSTICO", "DATA DA PRIMEIRA CONSULTA", "DATA DIAGNÓSTICO", "LEI 60 DIAS", "DATA DO PRIMEIRO TRATAMENTO", "PAC_CODIGO", "DT_NASCIMENTO", "CIDDIAGNOSTICO", "DATA", "CPF", "CNS", "NOME_MAE", "IDADE", "ETNIA", "SEXO", "NATURALIDADE", "UF_NASCIMENTO", "ENDERECO", "BAIRRO", "CEP", "CIDADE_MORADIA", "TELEFONE_RES", "TELEFONE_REC", "TELEFONE_CEL", "DT_PRIM_INTERNACAO_PREVIA", "DT_OBITO", "CID_OBITO", "DATA_INICIO_QUIMIO", "MES_ANO"],
  "Passo 3": ["MES_ANO", "Registrador", "NOME", "PRONTUARIO", "OBSERVAÇÕES", "DIAGNÓSTICO E TRATAMENTOS ANTERIORES", "CID DIAGNÓSTICO", "DATA DA PRIMEIRA CONSULTA", "DATA DIAGNÓSTICO", "LEI 60 DIAS", "DATA DO PRIMEIRO TRATAMENTO", "TIREÓIDE", "Convênio", "Base mais importante do diagnóstico", "Localização do tumor primário", "CID - O", "Estadiamento cT", "Estadiamento cN", "Estadiamento cM", "Estadiamento TNM", "Outro estadiamento", "Estadiamento patológico - T (0 - in situ; X-desconhecido)", "Estadiamento patológico - N (0- Ausência em linfonodos regionais; X-desconhecido)", "Estadiamento patológico - M (0 - Ausência de metástase a distância; X-desconhecido)", "Localização de Mtx à Distância", "Localização de Mtx à Distância", "Localização de Mtx à Distância", "Localização de Mtx à Distância", "Clínica do início do tratamento", "Razão para não realizar", "1º tratamento no hospital", "1º tratamento realizado no hospital", "Estado da doença ao final do primeiro tratamento", "Data do óbito", "Óbito devido ao câncer", "DATA DO PRIMEIRO TRATAMENTO", "DT_NASCIMENTO", "CID", "DIAGNOSTICO", "DATA", "CPF", "CNS", "NOME_MAE", "IDADE", "ETNIA", "SEXO", "NATURALIDADE", "UF_NASCIMENTO", "ENDERECO", "BAIRRO", "CEP", "CIDADE_MORADIA", "TELEFONE_RES", "TELEFONE_REC", "TELEFONE_CEL", "DT_PRIM_INTERNACAO_PREVIA", "DT_OBITO", "CID_OBITO", "DATA_INICIO_QUIMIOCONSENTIMENTO"]
};

// =========================================================================
// API FRONT-END E CONTROLE DE PROGRESSO
// =========================================================================

function atualizarStatusProgresso(msg, pct) {
  var cache = CacheService.getUserCache();
  cache.put("hub_msg", msg, 300);
  cache.put("hub_pct", pct.toString(), 300);
}

function lerProgresso() {
  var cache = CacheService.getUserCache();
  return { msg: cache.get("hub_msg") || "", pct: cache.get("hub_pct") || "0" };
}

function limparProgresso() {
  var cache = CacheService.getUserCache();
  cache.remove("hub_msg");
  cache.remove("hub_pct");
}

function buscarLinksRegistrados() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DB_LINKS");
  if (!sheet) return [];
  var data = sheet.getDataRange().getDisplayValues();
  if (data.length <= 1) return [];
  var headers = data[0], registros = [];
  for (var i = 1; i < data.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) obj[headers[j]] = data[i][j];
    registros.push(obj);
  }
  return registros.reverse();
}

function removerLink(id) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DB_LINKS");
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      sheet.deleteRow(i + 1);
      return { sucesso: "Registro removido com sucesso." };
    }
  }
  return { erro: "ID não encontrado." };
}

// =========================================================================
// INSERÇÃO MANUAL E VARREDURA
// =========================================================================

function salvarNovoLink(dados) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("DB_LINKS");
    if (!sheet) return { erro: "Aba 'DB_LINKS' não encontrada. Instale no menu 'Preparar Abas'." };

    var url = dados.link.trim();
    var urlsExistentes = sheet.getRange(2, 6, Math.max(1, sheet.getLastRow())).getValues().flat();
    if (urlsExistentes.includes(url)) return { erro: "Este link já está registrado no Hub." };

    var ws;
    try { ws = SpreadsheetApp.openByUrl(url); } catch(e) { return { erro: "Acesso negado ou link inválido." }; }

    var abaTrabalho = ws.getSheets()[0];
    var qtdLinhas = abaTrabalho.getLastRow() - 1;
    if (qtdLinhas < 0) qtdLinhas = 0;

    var status = "Ativo (Manual)";
    var headersArquivo = abaTrabalho.getRange(1, 1, 1, Math.max(1, abaTrabalho.getLastColumn())).getValues()[0];
    var headersEsperados = COLUNAS_SISTEMA[dados.passo];

    if (headersEsperados && !validarColunasHub(headersArquivo, headersEsperados)) status += " | Aviso: Colunas Divergentes";

    var novoId = Utilities.getUuid().split('-')[0].toUpperCase();
    var dataCadastro = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");

    sheet.appendRow([novoId, dataCadastro, dados.ano, dados.mes, dados.passo, url, qtdLinhas, status]);
    return { sucesso: "Planilha registrada", linhas: qtdLinhas };
  } catch (erro) {
    return { erro: "Erro interno: " + erro.message };
  }
}

function converterMesTexto(nomePasta) {
  var texto = nomePasta.toString().trim().toUpperCase();
  var matchLetras = texto.match(/^([A-Z]{3})/);
  if (matchLetras) return matchLetras[1];
  var matchNum = texto.match(/^(\d{2})/);
  if (matchNum) {
    var mapa = {"01":"JAN", "02":"FEV", "03":"MAR", "04":"ABR", "05":"MAI", "06":"JUN", "07":"JUL", "08":"AGO", "09":"SET", "10":"OUT", "11":"NOV", "12":"DEZ"};
    return mapa[matchNum[1]] || texto;
  }
  var matchNumSingle = texto.match(/^(\d{1})(?:\D|$)/);
  if (matchNumSingle) {
    var mapaSingle = {"1":"JAN", "2":"FEV", "3":"MAR", "4":"ABR", "5":"MAI", "6":"JUN", "7":"JUL", "8":"AGO", "9":"SET"};
    return mapaSingle[matchNumSingle[1]] || texto;
  }
  return texto; 
}

function validarColunasHub(headersEncontrados, headersEsperados) {
  if (!headersEncontrados || headersEncontrados.length === 0) return false;
  var acertos = 0;
  var encontradosStr = headersEncontrados.map(function(h) { return h.toString().trim().toUpperCase(); });
  headersEsperados.forEach(function(esp) {
    if (encontradosStr.indexOf(esp.trim().toUpperCase()) !== -1) acertos++;
  });
  return (acertos / headersEsperados.length) >= 0.7; 
}

// =========================================================================
// FUNÇÕES DE EXTRAÇÃO SEGURA ANTI-503
// =========================================================================

function extrairPastasSeguro(pasta) {
  let lista = [];
  try {
    let iterador = pasta.getFolders();
    while (true) {
      try {
        if (!iterador.hasNext()) break;
        lista.push(iterador.next());
      } catch (eIterador) {
        console.warn("Item corrompido saltado em: " + pasta.getName());
        break; // Aborta a iteração desta pasta mas devolve o que já recolheu
      }
    }
  } catch (eGeral) {
    console.error("Não foi possível aceder aos subdiretórios de " + (pasta ? pasta.getName() : "Desconhecido"));
  }
  return lista;
}

function extrairArquivosSeguro(pasta) {
  let lista = [];
  try {
    let iterador = pasta.getFiles();
    while (true) {
      try {
        if (!iterador.hasNext()) break;
        lista.push(iterador.next());
      } catch (eIterador) {
        console.warn("Ficheiro corrompido saltado em: " + pasta.getName());
        break;
      }
    }
  } catch (eGeral) {
    console.error("Não foi possível aceder aos ficheiros de " + (pasta ? pasta.getName() : "Desconhecido"));
  }
  return lista;
}

// =========================================================================
// MOTOR DE VARREDURA AUTOMÁTICA
// =========================================================================

function varreduraAutomaticaDrive() {
  limparProgresso();
  atualizarStatusProgresso("Iniciando conexão segura com Drive...", 5);
  
  const rootId = "1CeOQcEMQiLStgw-yx9vDg13lm13ufnaB";
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DB_LINKS");
  if (!sheet) return { erro: "Aba DB_LINKS não encontrada." };

  // =======================================================
  // 🛡️ BLINDAGEM DO CABEÇALHO (GARANTE QUE A LINHA 1 É INTACTA)
  // =======================================================
  if (sheet.getLastRow() === 0 || sheet.getRange("A1").getValue() === "") {
    // Se a aba estiver totalmente vazia ou faltar o ID na A1, desenha o cabeçalho!
    const cabecalho = [["ID", "DATA_CADASTRO", "ANO", "MES", "PASSO", "LINK_PLANILHA", "QTD_LINHAS", "STATUS"]];
    sheet.getRange("A1:H1").setValues(cabecalho)
         .setFontWeight("bold")
         .setBackground("#2c3e50")
         .setFontColor("white");
  }

  // Previne erros ao ler URLs se a aba só tiver o cabeçalho (getLastRow == 1)
  let urlsExistentes = [];
  const lastRowAtual = sheet.getLastRow();
  if (lastRowAtual > 1) {
    urlsExistentes = sheet.getRange(2, 6, lastRowAtual - 1).getValues().flat();
  }
  // =======================================================

  let rootFolder;
  try { 
    rootFolder = DriveApp.getFolderById(rootId);
  } catch (e) { 
    return { erro: "Acesso negado à pasta raiz no Drive Compartilhado." }; 
  }

  // Utiliza a extração blindada para evitar o erro 503 inicial
  let pastasAnos = extrairPastasSeguro(rootFolder);
  let listaAnos = pastasAnos.filter(f => f.getName().match(/^20\d{2}$/));
  listaAnos.sort((a, b) => a.getName().localeCompare(b.getName()));
  
  let countAdicionados = 0;
  
  for (let i = 0; i < listaAnos.length; i++) {
    let pastaAno = listaAnos[i];
    let anoStr = pastaAno.getName(); 
    
    console.log(`[AUDITORIA] Entrando no Ano: ${anoStr}`);
    let pctBase = 10 + Math.round((i / listaAnos.length) * 80);
    atualizarStatusProgresso("Mapeando Ano: " + anoStr, pctBase);

    // 1. VARREDURA DOS PASSOS PADRÃO
    let pastasPassos = extrairPastasSeguro(pastaAno);
    
    for (let j = 0; j < pastasPassos.length; j++) {
      let pastaPasso = pastasPassos[j];
      let matchPasso = pastaPasso.getName().match(/Passo\s*(\d)/i);
      if (!matchPasso) continue; 
      let passoOficial = "Passo " + matchPasso[1]; 

      let pastasMeses = extrairPastasSeguro(pastaPasso);
      
      for (let k = 0; k < pastasMeses.length; k++) {
        let pastaMes = pastasMeses[k];
        let mesStr = converterMesTexto(pastaMes.getName()); 
        
        let arquivos = extrairArquivosSeguro(pastaMes);
        
        for (let a = 0; a < arquivos.length; a++) {
          let arquivo = arquivos[a];
          let mime = arquivo.getMimeType();
          let url = arquivo.getUrl();

          if (urlsExistentes.includes(url)) continue;

          if (mime === MimeType.GOOGLE_SHEETS || mime.includes('spreadsheetml') || mime.includes('excel')) {
            console.log(`Processando: ${arquivo.getName()} (${anoStr})`);
            let status = "Ativo", qtdLinhas = 0;
            let isExcel = mime !== MimeType.GOOGLE_SHEETS;

            try {
              let ws = null;
              if (isExcel) {
                try {
                  let resource = { title: "TMP_" + arquivo.getName(), mimeType: MimeType.GOOGLE_SHEETS };
                  let tempFile = Drive.Files.insert(resource, arquivo.getBlob());
                  ws = SpreadsheetApp.openById(tempFile.id);
                  status = "Ativo (XLS)";
                } catch (errConv) {
                  status = "Erro: Conversão Drive Falhou";
                }
              } else {
                ws = SpreadsheetApp.openById(arquivo.getId());
              }

              if (ws) {
                let aba = ws.getSheets()[0];
                qtdLinhas = aba.getLastRow() - 1;
                let headersArquivo = aba.getRange(1, 1, 1, Math.max(1, aba.getLastColumn())).getValues()[0];
                let headersEsperados = typeof COLUNAS_SISTEMA !== 'undefined' ? COLUNAS_SISTEMA[passoOficial] : null;
                if (headersEsperados && typeof validarColunasHub !== 'undefined' && !validarColunasHub(headersArquivo, headersEsperados)) {
                  status += " | Aviso: Cabeçalho Divergente";
                }
                if (isExcel) Drive.Files.remove(ws.getId());
              }

              let novoId = Utilities.getUuid().split('-')[0].toUpperCase();
              let dataCadastro = Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy HH:mm:ss");
              
              sheet.appendRow([novoId, dataCadastro, anoStr, mesStr, passoOficial, url, (qtdLinhas < 0 ? 0 : qtdLinhas), status]);
              urlsExistentes.push(url);
              countAdicionados++;
              Utilities.sleep(300); 

            } catch (e) {
              let novoIdFallback = Utilities.getUuid().split('-')[0].toUpperCase();
              sheet.appendRow([novoIdFallback, Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy HH:mm:ss"), anoStr, mesStr, passoOficial, url, 0, "Erro de Leitura"]);
              urlsExistentes.push(url);
            }
          }
        }
      }
    }

    // 2. VARREDURA DE QUERIES
    console.log(`[AUDITORIA] Buscando pasta 'Queries' em ${anoStr}`);
    let novasQueries = rastrearQueriesHierarquico(pastaAno, anoStr, urlsExistentes);
    
    if (novasQueries.length > 0) {
      novasQueries.forEach(q => {
        sheet.appendRow([q.ID, q.DATA_CADASTRO, q.ANO, q.MES, q.PASSO, q.LINK_PLANILHA, q.QTD_LINHAS, q.STATUS]);
        urlsExistentes.push(q.LINK_PLANILHA);
        countAdicionados++;
      });
    }
  }

  atualizarStatusProgresso("Concluído!", 100);
  return { sucesso: `Varredura finalizada. ${countAdicionados} novos arquivos encontrados.` };
}

// =========================================================================
// FUNÇÕES AUXILIARES DAS QUERIES (COM EXTRAÇÃO SEGURA)
// =========================================================================

function rastrearQueriesHierarquico(pastaAno, anoNome, urlsExistentes) {
  let queriesEncontradas = [];
  let pastasFilhas = extrairPastasSeguro(pastaAno);
  let pastaQueries = null;

  for (let i = 0; i < pastasFilhas.length; i++) {
    if (pastasFilhas[i].getName().toLowerCase().includes("queri") || pastasFilhas[i].getName().toLowerCase().includes("brut")) {
      pastaQueries = pastasFilhas[i];
      break;
    }
  }

  if (!pastaQueries) return [];

  let itensQueries = extrairPastasSeguro(pastaQueries);
  
  // CASO A: Subpastas de Mês
  for (let i = 0; i < itensQueries.length; i++) {
    let pastaMes = itensQueries[i];
    let mesDetectado = extrairMesDePastaQuery(pastaMes.getName());
    
    if (mesDetectado) {
      let arquivos = extrairArquivosSeguro(pastaMes);
      for (let j = 0; j < arquivos.length; j++) {
        let arq = arquivos[j];
        let mime = arq.getMimeType();
        
        if (mime === MimeType.GOOGLE_SHEETS || mime.includes('spreadsheetml') || mime.includes('excel')) {
          if (urlsExistentes.includes(arq.getUrl())) continue;
          queriesEncontradas.push(montarObjetoQuery(arq, mesDetectado, anoNome, mime));
          Utilities.sleep(200);
        }
      }
    }
  }

  // CASO B: Arquivos Soltos
  let arquivosSoltos = extrairArquivosSeguro(pastaQueries);
  for (let i = 0; i < arquivosSoltos.length; i++) {
    let arq = arquivosSoltos[i];
    let mime = arq.getMimeType();
    
    if (mime === MimeType.GOOGLE_SHEETS || mime.includes('spreadsheetml') || mime.includes('excel')) {
      if (urlsExistentes.includes(arq.getUrl())) continue;
      
      let dataInterna = tentarExtrairDataDaCelulaA2(arq.getId());
      
      if (dataInterna && dataInterna.mes && dataInterna.ano) {
        queriesEncontradas.push(montarObjetoQuery(arq, dataInterna.mes, dataInterna.ano, mime));
      } else {
        let dataCriacao = arq.getDateCreated();
        let mesesExt = ["JAN", "FEV", "MAR", "ABR", "MAI", "JUN", "JUL", "AGO", "SET", "OUT", "NOV", "DEZ"];
        queriesEncontradas.push(montarObjetoQuery(arq, mesesExt[dataCriacao.getMonth()], dataCriacao.getFullYear().toString(), mime));
      }
      Utilities.sleep(200);
    }
  }

  return queriesEncontradas;
}

/**
 * Extrator flexível de mês a partir de nomes de pastas de Queries.
 */
function extrairMesDePastaQuery(nomePasta) {
  if (!nomePasta) return null;
  const texto = String(nomePasta).toUpperCase().trim();
  const mesesTextuais = ["JAN", "FEV", "MAR", "ABR", "MAI", "JUN", "JUL", "AGO", "SET", "OUT", "NOV", "DEZ"];
  
  // Ex: "01.janeiro", "01/JAN", "012026"
  const matchNum = texto.match(/^(0[1-9]|1[0-2])/);
  if (matchNum) {
    const indice = parseInt(matchNum[1], 10) - 1;
    return mesesTextuais[indice];
  }
  
  // Ex: "JANEIRO", "JAN"
  for (let i = 0; i < mesesTextuais.length; i++) {
    if (texto.includes(mesesTextuais[i])) return mesesTextuais[i];
  }
  
  return null;
}

/**
 * Lê célula A2 da Query para deduzir a data.
 */
function tentarExtrairDataDaCelulaA2(idArquivo) {
  try {
    const file = DriveApp.getFileById(idArquivo);
    const mime = file.getMimeType();
    let ws, isTemp = false;

    if (mime === MimeType.MICROSOFT_EXCEL || mime.includes('spreadsheetml') || mime.includes('excel')) {
      const temp = Drive.Files.insert({title: "Temp_DateCheck", mimeType: MimeType.GOOGLE_SHEETS}, file.getBlob());
      ws = SpreadsheetApp.openById(temp.id);
      isTemp = true;
    } else if (mime === MimeType.GOOGLE_SHEETS) {
      ws = SpreadsheetApp.openById(idArquivo);
    } else {
      return null;
    }

    const aba = ws.getSheets()[0];
    const valorA2 = String(aba.getRange("A2").getDisplayValue()).trim().toUpperCase();
    
    if (isTemp) Drive.Files.remove(ws.getId());

    if (!valorA2) return null;

    const matchData = valorA2.match(/(?:0[1-9]|[12]\d|3[01])?[\/\-]*(0[1-9]|1[0-2]|[A-Z]{3})[\/\-]*(20\d{2})/);
    if (matchData) {
      const mesStr = matchData[1];
      const ano = matchData[2];
      
      let mesExt = null;
      if (isNaN(parseInt(mesStr))) {
        mesExt = mesStr.substring(0,3); 
      } else {
        const mesesT = ["JAN", "FEV", "MAR", "ABR", "MAI", "JUN", "JUL", "AGO", "SET", "OUT", "NOV", "DEZ"];
        mesExt = mesesT[parseInt(mesStr) - 1]; 
      }
      return { mes: mesExt, ano: ano };
    }
    
    return null;
  } catch (e) {
    return null;
  }
}

function montarObjetoQuery(arquivo, mes, ano, mime) {
  return {
    ID: arquivo.getId().substring(0, 8).toUpperCase(),
    DATA_CADASTRO: Utilities.formatDate(arquivo.getDateCreated(), "GMT-3", "dd/MM/yyyy HH:mm"),
    ANO: ano,
    MES: mes,
    PASSO: "Query",
    QTD_LINHAS: "Dados Brutos",
    STATUS: mime === MimeType.GOOGLE_SHEETS ? "Pronto" : "Pronto (XLS)",
    LINK_PLANILHA: arquivo.getUrl(),
    TIPO: "Q" // <--- Adicionado para compatibilidade com o front-end
  };
}

// =========================================================================
// LEITURA DE REGISTROS E GRAVAÇÃO INTELIGENTE (COM FILTRO DE BRANCOS)
// =========================================================================

const COLUNAS_ALVO_EXTRACAO = ["NOME", "PRONTUARIO", "OBSERVAÇÃO", "DIAGNÓSTICO E TRATAMENTOS ANTERIORES", "CID DIAGNÓSTICO", "DATA DA PRIMEIRA CONSULTA", "DATA DIAGNÓSTICO", "LEI 60 DIAS", "DATA DO PRIMEIRO TRATAMENTO", "PAC_CODIGO", "DT_NASCIMENTO", "CID", "DIAGNOSTICO", "DATA", "CPF", "CNS", "NOME_MAE", "IDADE", "ETNIA", "SEXO", "NATURALIDADE", "UF_NASCIMENTO", "ENDERECO", "BAIRRO", "CEP", "CIDADE_MORADIA", "TELEFONE_RES", "TELEFONE_REC", "TELEFONE_CEL", "DT_PRIM_INTERNACAO_PREVIA", "DT_OBITO", "CID_OBITO", "DATA_INICIO_QUIMIO", "MES_ANO"];

function obterColunasAlvo() {
  return ["_ID_BASE", "_PASSO", "_MES_ANO"].concat(COLUNAS_ALVO_EXTRACAO);
}

function obterIdsExtraidos() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("BD_REGISTROS");
  if (!sheet) return [];
  if (sheet.getLastRow() <= 1) return [];
  
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  var ids = new Set();
  data.forEach(function(r) { if(r[0]) ids.add(r[0].toString()); });
  return Array.from(ids);
}

function buscarRegistrosLidos() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("BD_REGISTROS");
  if (!sheet) return [];
  var data = sheet.getDataRange().getDisplayValues();
  if (data.length <= 1) return [];
  
  var headers = data[0], registros = [];
  for (var i = 1; i < data.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      obj[headers[j]] = data[i][j];
    }
    registros.push(obj);
  }
  return registros;
}

function extrairRegistros(idsSelecionados) {
  limparProgresso();
  if(!idsSelecionados || idsSelecionados.length === 0) return {erro: "Nenhuma base selecionada."};

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetLinks = ss.getSheetByName("DB_LINKS");
  var sheetRegs = ss.getSheetByName("BD_REGISTROS");
  
  if (!sheetRegs) return {erro: "Aba 'BD_REGISTROS' não existe. Vá no menu e clique em Preparar Abas."};

  var dataLinks = sheetLinks.getDataRange().getValues();
  var basesParaLer = [];
  for (var i = 1; i < dataLinks.length; i++) {
    if (idsSelecionados.includes(dataLinks[i][0].toString())) {
      basesParaLer.push({
        id: dataLinks[i][0].toString(), 
        url: dataLinks[i][5], 
        passo: dataLinks[i][4],
        ano: dataLinks[i][2],
        mes: dataLinks[i][3]
      });
    }
  }

  var registrosExtraidos = [];
  var cabecalhoFinal = obterColunasAlvo();
  
  for (var k = 0; k < basesParaLer.length; k++) {
    var base = basesParaLer[k];
    var pct = Math.round((k / basesParaLer.length) * 80); 
    atualizarStatusProgresso("Lendo Base #" + base.id + " (" + base.mes + "/" + base.ano + ")", pct);

    try {
      var ws = SpreadsheetApp.openByUrl(base.url);
      var aba = ws.getSheets()[0];
      var dados = aba.getDataRange().getDisplayValues();
      if (dados.length <= 1) continue;

      var cabecalhoBase = dados[0].map(function(h) { return h.toString().trim().toUpperCase(); });
      var indicesExtraidos = {};

      COLUNAS_ALVO_EXTRACAO.forEach(function(col) {
        var idx = cabecalhoBase.indexOf(col.toUpperCase());
        if (idx !== -1) indicesExtraidos[col] = idx;
      });

      for (var r = 1; r < dados.length; r++) {
        var rowData = dados[r];
        
        // VALIDAÇÃO: Ignorar registros com NOME ou PRONTUARIO em branco/nulos
        var valNome = indicesExtraidos["NOME"] !== undefined && rowData[indicesExtraidos["NOME"]] ? rowData[indicesExtraidos["NOME"]].toString().trim() : "";
        var valProntuario = indicesExtraidos["PRONTUARIO"] !== undefined && rowData[indicesExtraidos["PRONTUARIO"]] ? rowData[indicesExtraidos["PRONTUARIO"]].toString().trim() : "";

        if (valNome === "" || valProntuario === "") {
          continue; // Pula essa linha, não será importada para o BD
        }

        var obj = { "_ID_BASE": base.id, "_PASSO": base.passo, "_MES_ANO": base.mes + "/" + base.ano };
        COLUNAS_ALVO_EXTRACAO.forEach(function(col) {
          obj[col] = indicesExtraidos[col] !== undefined ? rowData[indicesExtraidos[col]] : "";
        });
        registrosExtraidos.push(obj);
      }
    } catch (e) {
      // Continua para a proxima
    }
  }

  atualizarStatusProgresso("Conciliando com o histórico existente...", 85);
  
  // ESTRATÉGIA DE PRESERVAÇÃO DE HISTÓRICO
  var allDataRegs = sheetRegs.getDataRange().getValues();
  var linhasPreservadas = [];
  
  for (var j = 1; j < allDataRegs.length; j++) {
     if (!idsSelecionados.includes(allDataRegs[j][0].toString())) {
         linhasPreservadas.push(allDataRegs[j]);
     }
  }

  var matrizNovos = registrosExtraidos.map(function(reg) {
    return cabecalhoFinal.map(function(colName) {
      return reg[colName] !== undefined ? reg[colName] : "";
    });
  });

  var matrizFinal = linhasPreservadas.concat(matrizNovos);

  atualizarStatusProgresso("Gravando na planilha...", 95);
  
  if (sheetRegs.getLastRow() > 1) {
    sheetRegs.getRange(2, 1, sheetRegs.getLastRow() - 1, sheetRegs.getLastColumn()).clearContent();
  }

  if (matrizFinal.length > 0) {
    sheetRegs.getRange(2, 1, matrizFinal.length, cabecalhoFinal.length).setValues(matrizFinal);
  }

  atualizarStatusProgresso("Consolidação concluída!", 100);
  return { sucesso: true, total: registrosExtraidos.length };
}

function verificarEstadoVarredura() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DB_LINKS");
  if (!sheet || sheet.getLastRow() <= 1) return { realizado: false };
  return { realizado: true, ultimaData: sheet.getRange(sheet.getLastRow(), 2).getDisplayValue() };
}

// =========================================================================
// FUNÇÕES DE COMUNICAÇÃO: POPS E CÓDIGOS
// =========================================================================

function buscarPops() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DB_POPS");
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  var headers = data[0];
  return data.slice(1).map(function(row) {
    var obj = {};
    headers.forEach(function(h, i) { obj[h] = row[i]; });
    return obj;
  });
}

/**
 * Busca os códigos na planilha blindado contra falhas de serialização (Erro null do Google)
 */
function buscarCodigos() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("DB_CODIGOS");
    
    if (!sheet) {
      return { erro: "Aba DB_CODIGOS não encontrada.", dados: [] };
    }

    // O SEGREDO ESTÁ AQUI: getDisplayValues() transforma datas e números em Texto puro
    // Isso impede que a interface do Google (google.script.run) trave e retorne 'null'
    var data = sheet.getDataRange().getDisplayValues();
    
    if (data.length <= 1) {
      return { sucesso: true, dados: [] };
    }

    var headers = data[0].map(function(h) { return String(h).toUpperCase().trim(); });
    
    // Mapeamento dinâmico (imune à mudança de ordem das colunas)
    var idxId = headers.indexOf("ID");
    var idxProj = headers.indexOf("PROJETO");
    var idxTit = headers.indexOf("TITULO");
    var idxCont = headers.indexOf("CONTEUDO");
    var idxData = headers.indexOf("DATA_CRIACAO");

    var lista = [];
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      
      // Só adiciona se tiver um Título preenchido
      if (idxTit !== -1 && row[idxTit] && row[idxTit].trim() !== "") {
        lista.push({
          // Se for um código velho que não tem ID, cria um ID falso na tela para não quebrar os botões
          ID: (idxId !== -1 && row[idxId]) ? String(row[idxId]).trim() : "TEMP_" + i,
          PROJETO: (idxProj !== -1 && row[idxProj]) ? String(row[idxProj]).trim() : "Sem Categoria",
          TITULO: String(row[idxTit]).trim(),
          CONTEUDO: (idxCont !== -1 && row[idxCont]) ? String(row[idxCont]) : "",
          DATA_CRIACAO: (idxData !== -1 && row[idxData]) ? String(row[idxData]) : "Sem registro"
        });
      }
    }
    
    return { 
      sucesso: true, 
      dados: lista, 
      debug: `Sucesso. ${lista.length} scripts formatados como string.` 
    };

  } catch (e) {
    return { erro: "Falha interna no servidor: " + e.toString(), dados: [] };
  }
}

/**
 * Salva um código associado a um Projeto. Cria a aba automaticamente se não existir.
 */
function salvarCodigo(projeto, titulo, conteudo) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("DB_CODIGOS");
  
  // PRÉ-VERIFICAÇÃO: Se a aba não existir, cria ela no padrão correto
  if (!sheet) {
    sheet = ss.insertSheet("DB_CODIGOS");
    var headersCodigos = ["ID", "PROJETO", "TITULO", "CONTEUDO", "DATA_CRIACAO"];
    sheet.getRange(1, 1, 1, headersCodigos.length).setValues([headersCodigos])
         .setFontWeight("bold").setBackground("#059669").setFontColor("white");
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, 5);
  }

  var data = sheet.getDataRange().getValues();
  var id = Utilities.getUuid().split('-')[0].toUpperCase();
  var dataStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");

  // Verifica se o código já existe (mesmo projeto e mesmo título) para atualizar
  var index = -1;
  for (var i = 1; i < data.length; i++) {
    // Como você usava um modelo antigo antes, evitamos erro verificando o tamanho da linha
    if (data[i].length >= 3 && data[i][1] === projeto && data[i][2] === titulo) {
      index = i + 1;
      break;
    }
  }

  if (index !== -1) {
    sheet.getRange(index, 4).setValue(conteudo);
    sheet.getRange(index, 5).setValue(dataStr);
    return { sucesso: "Script atualizado com sucesso!" };
  } else {
    // Insere novo
    sheet.appendRow([id, projeto, titulo, conteudo, dataStr]);
    return { sucesso: "Novo script guardado no repositório!" };
  }
}

// =========================================================================
// LEITURA DE INDICADORES DIRETAMENTE DAS FONTES (CONSOLIDADA)
// =========================================================================
function buscarIndicadoresDasBases(idsSelecionados, anoRHC, tipoAno) {
  var idsAlvo = idsSelecionados || [];
  var anoNum = parseInt(anoRHC || new Date().getFullYear());
  var isGlobal = (anoRHC === "GLOBAL");
  var modelo = tipoAno || "RHC"; 
  
  var dadosProcessados = [];
  var erros = [];

  // 1. LEITURA DAS BASES DO HUB (A partir de Out/2025)
  if (idsAlvo.length > 0) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetLinks = ss.getSheetByName("DB_LINKS");
    var dataLinks = sheetLinks.getDataRange().getValues();

    var basesParaLer = [];
    for (var i = 1; i < dataLinks.length; i++) {
      if (idsAlvo.includes(dataLinks[i][0].toString())) {
        basesParaLer.push({
          id: dataLinks[i][0].toString(),
          url: dataLinks[i][5],
          mesAno: dataLinks[i][3] + "/" + dataLinks[i][2]
        });
      }
    }

    for (var k = 0; k < basesParaLer.length; k++) {
      var base = basesParaLer[k];
      var isExcel = false;
      var wsId = null;

      try {
        var ws;
        var matchId = base.url.match(/[-\w]{25,}/);
        if (!matchId) throw new Error("URL do arquivo é inválida.");
        
        var fileId = matchId[0];
        var arquivo = DriveApp.getFileById(fileId);
        var mime = arquivo.getMimeType();
        
        isExcel = (mime === MimeType.MICROSOFT_EXCEL || mime.includes("excel") || mime.includes("spreadsheetml"));

        if (isExcel) {
          var blob = arquivo.getBlob();
          var tempFile = Drive.Files.insert({title: "Temp_Ind_" + base.id, mimeType: MimeType.GOOGLE_SHEETS}, blob);
          wsId = tempFile.id;
          ws = SpreadsheetApp.openById(wsId);
        } else {
          ws = SpreadsheetApp.openByUrl(base.url);
        }

        var aba = ws.getSheets()[0];
        var dataRange = aba.getDataRange();
        var values = dataRange.getValues();
        var displayValues = dataRange.getDisplayValues();
        var backgrounds = dataRange.getBackgrounds();

        if (values.length < 2) {
           if (isExcel && wsId) Drive.Files.remove(wsId);
           continue;
        }

        var headers = values[0].map(function(h) { return String(h).toUpperCase().trim(); });
        
        var idxLei60 = headers.indexOf('LEI 60 DIAS');
        var idxCid = headers.indexOf('CID DIAGNÓSTICO') !== -1 ? headers.indexOf('CID DIAGNÓSTICO') : (headers.indexOf('CID DIAGNOSTICO') !== -1 ? headers.indexOf('CID DIAGNOSTICO') : headers.indexOf('CID'));
        var idxProntuario = headers.indexOf('PRONTUARIO') !== -1 ? headers.indexOf('PRONTUARIO') : (headers.indexOf('PRONTUÁRIO') !== -1 ? headers.indexOf('PRONTUÁRIO') : headers.indexOf('PRONTUAR'));
        var idxDiagTrat = headers.indexOf('DIAGNÓSTICO E TRATAMENTOS ANTERIORES');
        var idxNome = headers.indexOf('NOME') !== -1 ? headers.indexOf('NOME') : headers.indexOf('PACIENTE');
        var idxDtDiag = headers.indexOf('DTDIAGNO') !== -1 ? headers.indexOf('DTDIAGNO') : (headers.indexOf('DATA DIAGNÓSTICO') !== -1 ? headers.indexOf('DATA DIAGNÓSTICO') : headers.indexOf('DATA DIAGNOSTICO'));
        var idxSexo = headers.indexOf('SEXO');
        var idxEtnia = headers.indexOf('ETNIA');
        var idxIdade = headers.indexOf('IDADE');
        var idxNasc = headers.indexOf('DT_NASCIMENTO') !== -1 ? headers.indexOf('DT_NASCIMENTO') : headers.indexOf('DATA DE NASCIMENTO');
        var idxCidade = headers.indexOf('CIDADE_MORADIA');
        var idxObito = headers.indexOf('DT_OBITO') !== -1 ? headers.indexOf('DT_OBITO') : headers.indexOf('DATA DO ÓBITO');
        var idxBaseDiag = headers.indexOf('BASE MAIS IMPORTANTE DO DIAGNÓSTICO');
        var idxClinica = headers.indexOf('CLÍNICA DO INÍCIO DO TRATAMENTO');
        var idxRazaoNao = headers.indexOf('PRINCIPAL RAZÃO PARA NÃO REALIZAÇÃO DO PRIMEIRO TRATAMENTO NO HOSPITAL');
        var idxPrimTrat = headers.indexOf('1º TRATAMENTO REALIZADO NO HOSPITAL');
        var idxEstado = headers.indexOf('ESTADO DA DOENÇA AO FINAL DO PRIMEIRO TRATAMENTO');

        // Fallbacks
        if (idxLei60 === -1) idxLei60 = 8;
        if (idxCid === -1) idxCid = 5;
        if (idxProntuario === -1) idxProntuario = 2;
        if (idxDiagTrat === -1) idxDiagTrat = 4;
        if (idxNome === -1) idxNome = 1;

        for (var row = 1; row < values.length; row++) {
          var linha = values[row];
          if (linha.join('').trim() === '') continue;

          var cor = '#ffffff';
          if (idxLei60 !== -1 && backgrounds[row][idxLei60] && backgrounds[row][idxLei60] !== '#ffffff') cor = backgrounds[row][idxLei60];
          else if (backgrounds[row][0] && backgrounds[row][0] !== '#ffffff') cor = backgrounds[row][0];
          else if (backgrounds[row][2] && backgrounds[row][2] !== '#ffffff') cor = backgrounds[row][2];

          var valNome = linha[idxNome] ? String(linha[idxNome]).trim() : "";
          var valPront = linha[idxProntuario] ? String(linha[idxProntuario]).trim() : "";
          if (valNome === "" || valPront === "") continue;

          var displayLei = (idxLei60 !== -1 && displayValues[row][idxLei60]) ? String(displayValues[row][idxLei60]).trim() : '';
          var lei60 = null;
          var inconclusivo = false;

          if (displayLei !== '' && displayLei !== '-') {
            if (displayLei.startsWith('#') || displayLei.toUpperCase().includes('NAN') || displayLei.toUpperCase().includes('ERRO')) inconclusivo = true;
            else {
              var num = parseFloat(displayLei.replace(',', '.'));
              if (isNaN(num)) inconclusivo = true; else lei60 = num < 0 ? 0 : num;
            }
          }

          var obito = false;
          if (idxObito !== -1 && linha[idxObito] != null) {
              var valObito = String(linha[idxObito]).trim().toUpperCase();
              if (valObito !== "" && valObito !== "N/A" && valObito !== "-" && valObito !== "NÃO" && valObito !== "NAO") obito = true;
          }

          dadosProcessados.push({
            idBase: base.id,
            corSimulada: cor.toLowerCase(),
            lei60: lei60,
            cid: (idxCid !== -1 && linha[idxCid]) ? String(linha[idxCid]).trim() : 'Sem CID',
            dtDiag: (idxDtDiag !== -1 && displayValues[row][idxDtDiag]) ? String(displayValues[row][idxDtDiag]).trim() : "",
            inconclusivo: inconclusivo,
            nome: valNome,
            prontuario: valPront,
            diagTrat: (idxDiagTrat !== -1 && linha[idxDiagTrat]) ? String(linha[idxDiagTrat]).trim() : '',
            mesAno: base.mesAno,
            sexo: (idxSexo !== -1 && linha[idxSexo] != null) ? String(linha[idxSexo]).trim().toUpperCase() : "",
            etnia: (idxEtnia !== -1 && linha[idxEtnia] != null) ? String(linha[idxEtnia]).trim().toUpperCase() : "",
            idade: (idxIdade !== -1 && linha[idxIdade] != null && linha[idxIdade] !== "") ? parseInt(linha[idxIdade], 10) : null,
            nascimento: (idxNasc !== -1 && displayValues[row][idxNasc] != null) ? String(displayValues[row][idxNasc]).trim() : null,
            cidade: (idxCidade !== -1 && linha[idxCidade] != null) ? String(linha[idxCidade]).trim() : "",
            obito: obito,
            baseDiag: (idxBaseDiag !== -1 && linha[idxBaseDiag]) ? String(linha[idxBaseDiag]).trim() : "",
            clinicaTrat: (idxClinica !== -1 && linha[idxClinica]) ? String(linha[idxClinica]).trim() : "",
            razaoNaoTrat: (idxRazaoNao !== -1 && linha[idxRazaoNao]) ? String(linha[idxRazaoNao]).trim() : "",
            primTrat: (idxPrimTrat !== -1 && linha[idxPrimTrat]) ? String(linha[idxPrimTrat]).trim() : "",
            estadoDoenca: (idxEstado !== -1 && linha[idxEstado]) ? String(linha[idxEstado]).trim() : ""
          });
        }
        if (isExcel && wsId) Drive.Files.remove(wsId);
      } catch (e) {
        erros.push("Falha na base #" + base.id + ": " + e.message);
        if (isExcel && wsId) { try { Drive.Files.remove(wsId); } catch(ex) {} }
      }
    }
  }

  // 2. INTEGRAÇÃO DA BASE HISTÓRICA (Apenas se ano <= 2025)
  if (isGlobal || anoNum <= 2025) {
    try {
      var paramHist = isGlobal ? "GLOBAL" : anoNum;
      // TRUE = Aplica trava de segurança set/2025 para base antiga
      var histResult = buscarHistoricoDash(paramHist, true); 
      if (histResult && histResult.dados) {
          dadosProcessados = dadosProcessados.concat(histResult.dados);
          if (histResult.erros && histResult.erros.length > 0) erros = erros.concat(histResult.erros);
      }
    } catch(e) {
      erros.push("Falha ao integrar base histórica: " + e.message);
    }
  }

  return { dados: dadosProcessados, erros: erros };
}

function gerarPdfRelatorio(htmlConteudo) {
  // Cria o PDF a partir do HTML recebido
  const blob = HtmlService.createHtmlOutput(htmlConteudo)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .getAs('application/pdf')
    .setName('Relatorio_Proporcionalidade_RHC.pdf');
  
  // Retorna os dados como Base64 para o navegador iniciar o download
  return Utilities.base64Encode(blob.getBytes());
}

/**
 * Sinaliza para o servidor parar qualquer backup em andamento
 */
function cancelarRotinaBackup() {
  const cache = CacheService.getUserCache();
  cache.put("sinal_cancelar_backup", "true", 300); // 5 minutos de validade
  return { msg: "Sinal de cancelamento enviado. Aguardando o robô parar..." };
}

/**
 * Realiza o backup segmentado por Ano, salvando na pasta destino específica.
 * COM BOTÃO DE PARADA DE EMERGÊNCIA (ABORT)
 * @param {string} tipo - Define se o backup é "Manual" ou "Automático"
 */
function realizarBackupDiarioQueries(tipo) {
  // Se a função for chamada sem parâmetro (ex: pelo gatilho de tempo do Google),
  // assume que é um backup automático.
  if (!tipo) {
    tipo = "Automático";
  }

  const SOURCE_FOLDER_ID = "1CeOQcEMQiLStgw-yx9vDg13lm13ufnaB"; 
  const DEST_PARENT_FOLDER_ID = "14pyzsErHJBekAq3uFnj3b6EskvM_9_tA"; 
  const DIAS_RETENCAO = 30;
  const LIMITE_TAMANHO_BYTES = 45 * 1024 * 1024; 
  const cache = CacheService.getUserCache();

  // Limpa o sinal de cancelamento antes de começar
  cache.remove("sinal_cancelar_backup");

  try {
    const pastaRaiz = DriveApp.getFolderById(SOURCE_FOLDER_ID);
    const pastaDestinoMestre = DriveApp.getFolderById(DEST_PARENT_FOLDER_ID);
    
    atualizarStatusProgresso("Organizando ambiente de destino...", 10);

    let pastaBackupsGeral;
    const buscaBackups = pastaDestinoMestre.getFoldersByName("Backups");
    pastaBackupsGeral = buscaBackups.hasNext() ? buscaBackups.next() : pastaDestinoMestre.createFolder("Backups");

    const dataHoje = Utilities.formatDate(new Date(), "GMT-3", "yyyy-MM-dd_HHmm");
    
    // AGORA FUNCIONA! A variável 'tipo' já existe e tem o valor "Manual" ou "Automático".
    const pastaDeHoje = pastaBackupsGeral.createFolder(`Backup_${dataHoje}_${tipo}`);

    const subpastasAnos = pastaRaiz.getFolders();
    let anosProcessados = 0;

    while (subpastasAnos.hasNext()) {
      // 1. CHECAGEM DE CANCELAMENTO ENTRE ANOS
      if (cache.get("sinal_cancelar_backup") === "true") {
         pastaDeHoje.setTrashed(true); // Apaga o que já fez hoje pra não deixar lixo
         return { sucesso: false, erro: "Operação CANCELADA pelo usuário." };
      }

      let pastaAno = subpastasAnos.next();
      let nomeAno = pastaAno.getName();
      let idPastaAno = pastaAno.getId();

      if (idPastaAno === DEST_PARENT_FOLDER_ID) continue;

      atualizarStatusProgresso(`Coletando arquivos de ${nomeAno}...`, 20);

      let blobsDoAno = [];
      let nomesUtilizados = new Set();
      let isCancelado = false; // Flag local para a recursão

      function coletarRecurso(pastaAtual, caminhoPai) {
        // 2. CHECAGEM DE CANCELAMENTO DENTRO DAS PASTAS
        if (isCancelado || cache.get("sinal_cancelar_backup") === "true") {
          isCancelado = true;
          return; 
        }

        if (pastaAtual.getId() === DEST_PARENT_FOLDER_ID) return;

        let arquivos = pastaAtual.getFiles();
        while (arquivos.hasNext()) {
          // 3. CHECAGEM DE CANCELAMENTO A CADA ARQUIVO (Super Rápido)
          if (cache.get("sinal_cancelar_backup") === "true") {
             isCancelado = true;
             return;
          }

          let arq = arquivos.next();
          let nome = arq.getName();
          
          if (arq.getSize() > LIMITE_TAMANHO_BYTES || 
              nome.toLowerCase().endsWith('.mp4') || 
              nome.toLowerCase().endsWith('.zip')) continue;

          let caminhoCompleto = caminhoPai + nome;
          if (nomesUtilizados.has(caminhoCompleto)) {
             caminhoCompleto = caminhoPai + arq.getId().substring(0,4) + "_" + nome;
          }
          nomesUtilizados.add(caminhoCompleto);

          let b = arq.getBlob();
          b.setName(caminhoCompleto);
          blobsDoAno.push(b);
        }

        let pastasFilhas = pastaAtual.getFolders();
        while (pastasFilhas.hasNext()) {
          let pFilha = pastasFilhas.next();
          coletarRecurso(pFilha, caminhoPai + pFilha.getName() + "/");
        }
      }

      coletarRecurso(pastaAno, "");

      // Se foi cancelado durante a varredura recursiva, aborta tudo
      if (isCancelado) {
        pastaDeHoje.setTrashed(true);
        return { sucesso: false, erro: "Operação CANCELADA pelo usuário." };
      }

      if (blobsDoAno.length > 0) {
        atualizarStatusProgresso(`Gerando ZIP: ${nomeAno}...`, 50);
        let zipAno = Utilities.zip(blobsDoAno, `Arquivos_${nomeAno}.zip`);
        pastaDeHoje.createFile(zipAno);
      }
      anosProcessados++;
    }

    atualizarStatusProgresso("Limpando histórico antigo...", 90);
    const pastasAntigas = pastaBackupsGeral.getFolders();
    const limiteData = new Date(new Date().getTime() - (DIAS_RETENCAO * 24 * 60 * 60 * 1000));
    
    while (pastasAntigas.hasNext()) {
      let p = pastasAntigas.next();
      if (p.getDateCreated() < limiteData) p.setTrashed(true);
    }

    atualizarStatusProgresso("Backup por Ano Concluído!", 100);
    return { sucesso: true };

  } catch (e) {
    return { sucesso: false, erro: "Falha no Backup: " + e.toString() };
  }
}

/**
 * Atalho do Botão: Faz o backup, espera o Drive "respirar" e retorna a lista
 */
function dispararBackupManual() {
  limparProgresso(); 
  atualizarStatusProgresso("Iniciando rotina de clonagem...", 5);
  
  // AQUI É A MÁGICA: Passamos o texto "Manual" para a função principal
  var resultado = realizarBackupDiarioQueries("Manual"); 
  
  if (!resultado.sucesso) {
    throw new Error(resultado.erro);
  }
  
  Utilities.sleep(2500); 
  return listarBackupsQueries();
}
/**
 * Lista as SESSÕES de backup (as pastas que contêm os ZIPs por ano)
 */
function listarBackupsQueries() {
  const RAIZ_ID = "14pyzsErHJBekAq3uFnj3b6EskvM_9_tA";
  
  try {
    const pastaRaiz = DriveApp.getFolderById(RAIZ_ID);
    const pastasBackups = pastaRaiz.getFoldersByName("Backups");
    
    if (!pastasBackups.hasNext()) return [];
    
    const pastaMestre = pastasBackups.next();
    const sessoes = pastaMestre.getFolders(); // Pega as pastas: Backup_2026-04...
    const lista = [];
    
    while (sessoes.hasNext()) {
      let pastaSessao = sessoes.next();
      let nomeSessao = pastaSessao.getName();
      
      // Conta quantos arquivos ZIP tem dentro dessa sessão
      let arquivos = pastaSessao.getFiles();
      let qtdZips = 0;
      while (arquivos.hasNext()) { arquivos.next(); qtdZips++; }
      
      lista.push({
        id: pastaSessao.getId(),
        nome: nomeSessao, // Ex: Backup_2026-04-20_1344_Manual
        dataRaw: pastaSessao.getDateCreated().getTime(),
        // Formata a data de criação da pasta para exibição
        dataCriacao: Utilities.formatDate(pastaSessao.getDateCreated(), "GMT-3", "dd/MM/yyyy HH:mm"),
        // Usaremos o campo "tamanho" para mostrar quantos anos foram salvos
        tamanho: qtdZips + " arquivos (.zip)",
        url: pastaSessao.getUrl()
      });
    }
    
    // Ordena pela data, o mais recente no topo
    return lista.sort((a, b) => b.dataRaw - a.dataRaw);
    
  } catch (e) {
    throw new Error("Falha ao acessar estrutura de pastas: " + e.message);
  }
}

function excluirBackup(idArquivo) {
  try {
    DriveApp.getFileById(idArquivo).setTrashed(true);
    return { sucesso: true };
  } catch (e) {
    return { sucesso: false, erro: e.toString() };
  }
}

/**
 * Extrator de Mês Flexível
 * Lê nomes de pastas como "012026", "01/JAN", "01.janeiro", "11. novembro", etc.
 */
function extrairMesDePasta(nomePasta) {
  if (!nomePasta) return null;
  const texto = String(nomePasta).toUpperCase().trim();
  
  const mesesTextuais = ["JAN", "FEV", "MAR", "ABR", "MAI", "JUN", "JUL", "AGO", "SET", "OUT", "NOV", "DEZ"];
  
  // 1. Tenta achar o mês pelo número inicial (01 a 12)
  // Casos: "01.janeiro", "01/JAN", "012026", "01 janeiro"
  const matchNum = texto.match(/^(0[1-9]|1[0-2])/);
  if (matchNum) {
    const indice = parseInt(matchNum[1], 10) - 1;
    return mesesTextuais[indice];
  }
  
  // 2. Tenta achar o mês pelas três primeiras letras (se não tiver número na frente)
  // Casos: "JANEIRO", "JAN", "Outubro"
  for (let i = 0; i < mesesTextuais.length; i++) {
    if (texto.includes(mesesTextuais[i])) {
      return mesesTextuais[i];
    }
  }
  
  return null;
}

/**
 * Lê o arquivo e procura a data em A2
 * Espera encontrar algo como "01/2026", "JAN/2026"
 */
function tentarExtrairDataDaCelulaA2(idArquivo) {
  try {
    const file = DriveApp.getFileById(idArquivo);
    const mime = file.getMimeType();
    let ws, isTemp = false;

    // Se for Excel, converte temporariamente
    if (mime === MimeType.MICROSOFT_EXCEL || mime.includes('spreadsheetml') || mime.includes('excel')) {
      const temp = Drive.Files.insert({title: "Temp_DateCheck", mimeType: MimeType.GOOGLE_SHEETS}, file.getBlob());
      ws = SpreadsheetApp.openById(temp.id);
      isTemp = true;
    } else if (mime === MimeType.GOOGLE_SHEETS) {
      ws = SpreadsheetApp.openById(idArquivo);
    } else {
      return null;
    }

    const aba = ws.getSheets()[0];
    const valorA2 = String(aba.getRange("A2").getDisplayValue()).trim().toUpperCase(); // Tenta ler a A2
    
    if (isTemp) Drive.Files.remove(ws.getId());

    if (!valorA2) return null;

    // Procura por DD/MM/AAAA ou MM/AAAA
    const matchData = valorA2.match(/(?:0[1-9]|[12]\d|3[01])?[\/\-]*(0[1-9]|1[0-2]|[A-Z]{3})[\/\-]*(20\d{2})/);
    
    if (matchData) {
      const mesStr = matchData[1];
      const ano = matchData[2];
      
      let mesExt = null;
      if (isNaN(parseInt(mesStr))) {
        mesExt = mesStr.substring(0,3); // Se for "JAN", já pega as 3 letras
      } else {
        const mesesT = ["JAN", "FEV", "MAR", "ABR", "MAI", "JUN", "JUL", "AGO", "SET", "OUT", "NOV", "DEZ"];
        mesExt = mesesT[parseInt(mesStr) - 1]; // Se for "01", vira "JAN"
      }
      return { mes: mesExt, ano: ano };
    }
    
    return null;
  } catch (e) {
    return null;
  }
}

/**
 * Rastreador Universal de Queries (Baseado na Estrutura de Pastas: Ano -> Queries -> Mês)
 */
function rastrearQueriesAprimorado() {
  const PASTA_RAIZ_ID = "1CeOQcEMQiLStgw-yx9vDg13lm13ufnaB";
  const pastaRaiz = DriveApp.getFolderById(PASTA_RAIZ_ID);
  let queriesColetadas = [];

  // 1. Encontrar as Pastas de ANO
  const pastasAnos = pastaRaiz.getFolders();
  
  while (pastasAnos.hasNext()) {
    const pastaAno = pastasAnos.next();
    const anoNome = pastaAno.getName();
    
    // Verifica se a pasta tem um nome de Ano (ex: 2025, 2026)
    if (!anoNome.match(/^(20\d{2})$/)) continue;
    
    const anoAtual = anoNome;

    // 2. Encontrar a pasta "Queries" (ou "queri") DENTRO deste ano
    const pastasFilhasDoAno = pastaAno.getFolders();
    let pastaQueries = null;
    
    while (pastasFilhasDoAno.hasNext()) {
      const pf = pastasFilhasDoAno.next();
      if (pf.getName().toLowerCase().includes("queri")) {
        pastaQueries = pf;
        break;
      }
    }
    
    if (!pastaQueries) continue; // Se este ano não tem pasta de Queries, salta.

    // Função interna para varrer arquivos dentro da pasta Queries (e subpastas de meses)
    function processarArquivosQuery(pasta, mesHerdadoDaPasta) {
      // Tenta extrair o mês do nome da pasta atual (ex: se a pasta se chamar "01.janeiro")
      let mesDestaPasta = extrairMesDePasta(pasta.getName());
      // Se não achar, mantém o mês que já vinha da pasta mãe
      let mesMestre = mesDestaPasta || mesHerdadoDaPasta;

      // Lê os Arquivos
      const arquivos = pasta.getFiles();
      while (arquivos.hasNext()) {
        const arq = arquivos.next();
        const mime = arq.getMimeType();
        
        if (mime === MimeType.GOOGLE_SHEETS || mime.includes('spreadsheetml') || mime.includes('ms-excel')) {
          
          let mesDefinitivo = mesMestre;
          let anoDefinitivo = anoAtual;

          // Se a pasta não nos deu o mês, PLANO B: Abrir o arquivo e ler a célula A2
          if (!mesDefinitivo) {
            let dadosInternos = tentarExtrairDataDaCelulaA2(arq.getId());
            if (dadosInternos) {
              mesDefinitivo = dadosInternos.mes;
              anoDefinitivo = dadosInternos.ano; // Pode sobrescrever o ano da pasta mãe, se necessário
            }
          }

          // Se, após tudo isto, ainda tiver mês e ano válidos, regista a Query
          if (mesDefinitivo && anoDefinitivo) {
            queriesColetadas.push({
              ID: arq.getId().substring(0, 8).toUpperCase(),
              DATA_CADASTRO: Utilities.formatDate(arq.getDateCreated(), "GMT-3", "dd/MM/yyyy HH:mm"),
              ANO: anoDefinitivo,
              MES: mesDefinitivo,
              PASSO: "Query",
              QTD_LINHAS: "Dados Brutos", 
              STATUS: mime === MimeType.GOOGLE_SHEETS ? "Pronto" : "Pronto (XLS)",
              LINK_PLANILHA: arq.getUrl(),
              TIPO: "Q"
            });
          }
        }
      }
      
      // Lê as subpastas (as pastas dos meses) recursivamente
      const subpastas = pasta.getFolders();
      while (subpastas.hasNext()) {
        processarArquivosQuery(subpastas.next(), mesMestre);
      }
    }

    // Inicia a varredura dentro da pasta Queries deste ano
    processarArquivosQuery(pastaQueries, null);
  }

  return queriesColetadas;
}

// =========================================================================
// AUDITORIA CRUZADA (COM FILTRO HISTÓRICO SILENCIOSO DE 12 MESES)
// =========================================================================

function extrairIdDaUrl(url) {
  if (!url) return null;
  var match = url.match(/[-\w]{25,}/);
  return match ? match[0] : null;
}

/**
 * Converte "MAR", "2026" num número inteiro sequencial para cálculo de tempo.
 */
function calcularIndiceMes(mesTexto, anoTexto) {
  if (!mesTexto || !anoTexto) return 0;
  const meses = { "JAN": 1, "FEV": 2, "MAR": 3, "ABR": 4, "MAI": 5, "JUN": 6, "JUL": 7, "AGO": 8, "SET": 9, "OUT": 10, "NOV": 11, "DEZ": 12 };
  const numMes = meses[String(mesTexto).trim().toUpperCase()];
  const numAno = parseInt(String(anoTexto).trim(), 10);
  if (!numMes || isNaN(numAno)) return 0;
  return (numAno * 12) + numMes;
}

/**
 * Lê Prontuário, Nome e CID de forma blindada.
 */
function obterDadosPlanilhaAudit(url) {
  var id = extrairIdDaUrl(url);
  if (!id) throw new Error("URL inválida: " + url);

  var file = DriveApp.getFileById(id);
  var mime = file.getMimeType();
  var ws, isTemp = false;

  if (mime === MimeType.MICROSOFT_EXCEL || mime === "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" || mime === "application/vnd.ms-excel") {
    var tempFile = Drive.Files.insert({title: "Temp_Audit_" + id, mimeType: MimeType.GOOGLE_SHEETS}, file.getBlob());
    ws = SpreadsheetApp.openById(tempFile.id);
    isTemp = true;
  } else {
    ws = SpreadsheetApp.openById(id);
  }

  var abas = ws.getSheets();
  var dadosTotais = [];
  var auditoriaHeaders = []; 

  abas.forEach(function(aba) {
    var dados = aba.getDataRange().getDisplayValues();
    if (dados && dados.length > 1) {
       var headersNorm = dados[0].map(function(h) { return String(h || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase().trim(); });
       auditoriaHeaders.push({ aba: aba.getName(), headersEncontrados: headersNorm });

       var idxPront = headersNorm.findIndex(h => h.includes("PRONTUARIO") || h === "PRONT" || h.includes("REGISTRO") || h === "PAC_CODIGO");
       var idxNome = headersNorm.findIndex(h => h.includes("NOME") || h.includes("PACIENTE"));
       var idxCid = headersNorm.findIndex(h => h.includes("CID DIAGNÓSTICO") || h === "CID"); // Lê o CID para o filtro!

       if (idxPront !== -1) {
         for (var i = 1; i < dados.length; i++) {
           var pVal = dados[i][idxPront];
           if (pVal !== undefined && pVal !== null && String(pVal).trim() !== "") {
             var pront = String(pVal).trim().replace(/\.0$/, '').replace(/^0+/, '');
             var nome = (idxNome !== -1 && dados[i][idxNome]) ? String(dados[i][idxNome]).trim() : "Nome não identificado";
             var cid = (idxCid !== -1 && dados[i][idxCid]) ? String(dados[i][idxCid]).trim() : "";
             dadosTotais.push({ prontuario: pront, nome: nome, cid: cid, aba: aba.getName() });
           }
         }
       }
    }
  });

  if (isTemp) Drive.Files.remove(ws.getId());
  return { registros: dadosTotais, relatorioAbas: auditoriaHeaders };
}

/**
 * Normalizador Universal de Meses
 * Converte JAN, 01, Outubro, ago., etc, em um índice numérico de 1 a 12.
 * O SEGREDO AQUI: O replace(/[^A-Z0-9]/g, '') arranca o ponto de "ago." e vira "AGO".
 */
function converterMesParaNumeral(mes) {
  if (!mes) return 0;
  const m = String(mes).toUpperCase().replace(/[^A-Z0-9]/g, ''); 
  const mapa = {
    'JAN':1, 'FEV':2, 'MAR':3, 'ABR':4, 'MAI':5, 'JUN':6, 'JUL':7, 'AGO':8, 'SET':9, 'OUT':10, 'NOV':11, 'DEZ':12,
    '01':1, '02':2, '03':3, '04':4, '05':5, '06':6, '07':7, '08':8, '09':9, '10':10, '11':11, '12':12,
    '1':1, '2':2, '3':3, '4':4, '5':5, '6':6, '7':7, '8':8, '9':9
  };
  // Tenta o match exato, ou os 3 primeiros caracteres (ex: "AGOSTO" -> "AGO")
  return mapa[m] || mapa[m.substring(0,3)] || 0;
}

/**
 * Auditoria com Filtro Automático (-12 a +2) via Competência do Hub
 */
function auditarCruzamentoBases(idAuditada, referencia) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetLinks = ss.getSheetByName("DB_LINKS");
    const sheetRegs = ss.getSheetByName("BD_REGISTROS");
    const dataLinks = sheetLinks.getDataRange().getValues();
    
    let urlA = "", urlB = "", meta = null;

    // 1. Identifica a Competência (Mês/Ano) Automaticamente da Linha do Hub
    for(let i = 1; i < dataLinks.length; i++) {
      if (String(dataLinks[i][0]) === String(idAuditada)) {
        urlA = dataLinks[i][5];
        meta = { mes: dataLinks[i][3], ano: dataLinks[i][2] };
        break;
      }
    }

    let refStr = String(referencia).trim();
    if (refStr.toLowerCase().indexOf("http") !== -1) {
      urlB = refStr;
    } else {
      for(let j = 1; j < dataLinks.length; j++) {
        if (String(dataLinks[j][0]) === refStr) { urlB = dataLinks[j][5]; break; }
      }
    }
    
    if (!urlA || !urlB || !meta) throw new Error("Falha ao localizar dados da base ou competência (Verifique o DB_LINKS).");

    // 2. Extração Inicial (A vs B)
    const resA = obterDadosPlanilhaAudit(urlA);
    const prontsAuditados = new Set(resA.registros.map(d => d.prontuario));
    const resB = obterDadosPlanilhaAudit(urlB);

    // 3. A Matemática do Tempo
    const nMesRef = converterMesParaNumeral(meta.mes);
    const nAnoRef = parseInt(meta.ano);
    const indiceAlvo = (nAnoRef * 12) + nMesRef;

    // 4. PASSO 1 e 2: Gerar Lista Y (Diferença Bruta)
    let listaY = [];
    let trackerY = new Set();
    resB.registros.forEach(p => {
      if (!prontsAuditados.has(p.prontuario) && !trackerY.has(p.prontuario)) {
        listaY.push(p);
        trackerY.add(p.prontuario);
      }
    });

    // 5. PASSO 3 e 4: Filtragem Temporal (A Máquina do Tempo real)
    const dadosBD = sheetRegs.getDataRange().getDisplayValues();
    const headersBD = dadosBD[0].map(h => String(h).toUpperCase().trim());
    
    const idxPrt = headersBD.indexOf("PRONTUARIO") !== -1 ? headersBD.indexOf("PRONTUARIO") : headersBD.indexOf("PAC_CODIGO");
    const idxCid = headersBD.findIndex(h => h.includes("CID DIAGNÓSTICO") || h === "CID");
    const idxMes = headersBD.indexOf("_MES_ANO");

    let listaZ = [];
    let jPassado = 0, jFuturo = 0;
    let logDebug = ["Mês Alvo: " + meta.mes + "/" + meta.ano + " (Índice Matemático: " + indiceAlvo + ")"];

    // Percorre os Faltantes (Os 351 do seu exemplo)
    listaY.forEach(p => {
      let justificado = false;
      
      // Limpa o CID para comparar "C50.9" com "C509" sem frescuras
      const pCid = String(p.cid || "").replace(/[^A-Z0-9]/ig, '').toUpperCase(); 

      for (let r = 1; r < dadosBD.length; r++) {
        // Traz as variáveis do BD com segurança para não gerar "UNDEFINED"
        const hPrt = String(dadosBD[r][idxPrt] || "").trim().replace(/\.0$/, '').replace(/^0+/, '');
        const hCidBruto = idxCid !== -1 ? String(dadosBD[r][idxCid] || "") : "";
        const hCid = hCidBruto.replace(/[^A-Z0-9]/ig, '').toUpperCase();
        const hMesAno = String(dadosBD[r][idxMes] || "");

        // 1ª Regra: Prontuário Bate? 
        // 2ª Regra: CID está vazio num dos lados (ignoramos) OU contém um ao outro (ex: C50 e C509 batem)?
        if (hPrt === p.prontuario && (pCid === "" || hCid === "" || hCid.includes(pCid) || pCid.includes(hCid))) {
          const partes = hMesAno.split('/');
          
          if (partes.length === 2) {
            const indHist = (parseInt(partes[1]) * 12) + converterMesParaNumeral(partes[0]);
            const diff = indHist - indiceAlvo;

            // Diferença está entre -12 meses (Passado) e +2 meses (Futuro)? E não é 0 (Mês Atual)?
            if (diff >= -12 && diff <= 2 && diff !== 0) {
              justificado = true;
              if (diff < 0) jPassado++; else jFuturo++;
              logDebug.push(`Justificado: ${p.prontuario} achado em ${hMesAno} (Diferença: ${diff} meses)`);
              break; // Pára de procurar este paciente no BD, já foi salvo!
            }
          }
        }
      }
      
      // Se não justificou, vai para a lista crítica!
      if (!justificado) {
         listaZ.push(p);
      }
    });

    return {
      sucesso: true,
      perdidos: listaZ,
      log: logDebug,
      estatisticas: {
        ref: resB.registros.length,
        bruta: listaY.length,
        justPassado: jPassado,
        justFuturo: jFuturo,
        final: listaZ.length
      }
    };
  } catch (e) { return { erro: e.message }; }
}

/**
 * Cria a planilha Google com os Faltantes (Geração Blindada)
 */
function gerarPlanilhaPerdasAudit(idAuditada, idReferencia) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetLinks = ss.getSheetByName("DB_LINKS");
    var dataLinks = sheetLinks.getDataRange().getValues();
    
    var urlB = "";
    var refStr = String(idReferencia).trim();
    if (refStr.toLowerCase().indexOf("http") !== -1 || refStr.toLowerCase().indexOf("docs.google.com") !== -1) {
      urlB = refStr; 
    } else {
      for(var j = 1; j < dataLinks.length; j++) {
        if (String(dataLinks[j][0]) === refStr) {
          urlB = dataLinks[j][5];
          break;
        }
      }
    }

    if (!urlB) throw new Error("URL da base de referência não encontrada.");

    // Usa a auditoria inteligente para saber QUEM exatamente vamos exportar (já sem os filtrados)
    var resultadoAuditoria = auditarCruzamentoBases(idAuditada, idReferencia);
    if (resultadoAuditoria.erro || !resultadoAuditoria.perdidos || resultadoAuditoria.perdidos.length === 0) {
      return { erro: "Nenhum dado pendente válido encontrado para exportação." };
    }
    
    // Cria um Set apenas com os pacientes reais que sobraram para extrair
    var prontuariosReaisPerdidos = new Set(resultadoAuditoria.perdidos.map(function(d) { return d.prontuario; }));

    var idRef = extrairIdDaUrl(urlB);
    var fileRef = DriveApp.getFileById(idRef);
    var wsRef, isTemp = false;

    if (fileRef.getMimeType().includes('spreadsheetml') || fileRef.getMimeType().includes('excel')) {
      var temp = Drive.Files.insert({title: "Export_Temp_" + idRef, mimeType: MimeType.GOOGLE_SHEETS}, fileRef.getBlob());
      wsRef = SpreadsheetApp.openById(temp.id);
      isTemp = true;
    } else {
      wsRef = SpreadsheetApp.openById(idRef);
    }

    var abaOrigem = wsRef.getSheets()[0];
    var dadosBrutos = abaOrigem.getDataRange().getDisplayValues();
    
    // Blindagem de array vazio
    if (!dadosBrutos || dadosBrutos.length <= 1) {
      if (isTemp) Drive.Files.remove(wsRef.getId());
      return { erro: "Planilha de referência vazia." };
    }

    var cabecalhoFinal = dadosBrutos[0].map(h => String(h || ""));
    var headersNorm = cabecalhoFinal.map(h => h.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase().trim());
    var idxPrt = headersNorm.findIndex(h => h.includes("PRONTUARIO") || h === "PAC_CODIGO" || h === "PRONT" || h.includes("REGISTRO"));

    var linhasParaExportar = [];
    var tracker = new Set();

    for (var r = 1; r < dadosBrutos.length; r++) {
      if (!dadosBrutos[r]) continue;
      var pVal = dadosBrutos[r][idxPrt];
      if (pVal !== undefined && pVal !== null) {
         var pront = String(pVal).trim().replace(/\.0$/, '').replace(/^0+/, '');
         
         // Se for um dos pacientes que a IA confirmou que está realmente perdido
         if (pront && prontuariosReaisPerdidos.has(pront) && !tracker.has(pront)) {
           var linhaAjustada = dadosBrutos[r].slice(0, cabecalhoFinal.length);
           while(linhaAjustada.length < cabecalhoFinal.length) linhaAjustada.push(""); // Blindagem
           linhasParaExportar.push(linhaAjustada);
           tracker.add(pront);
         }
      }
    }

    if (isTemp) Drive.Files.remove(wsRef.getId());
    if (linhasParaExportar.length === 0) return { erro: "Os pacientes foram identificados, mas houve falha ao copiar as linhas da origem." };

    // Construção da Matriz Segura (Evita erro de length na criação do Excel)
    var totalColunas = cabecalhoFinal.length;
    var totalLinhas = linhasParaExportar.length;
    var matrizPerfeita = [];

    for (var l = 0; l < totalLinhas; l++) {
      var linhaPolida = [];
      for (var c = 0; c < totalColunas; c++) {
        linhaPolida.push(String(linhasParaExportar[l][c] || ""));
      }
      matrizPerfeita.push(linhaPolida);
    }

    var dataHora = Utilities.formatDate(new Date(), "GMT-3", "dd-MM-yyyy_HHmm");
    var novaPlanilha = SpreadsheetApp.create("FALTANTES_" + dataHora);
    var abaDest = novaPlanilha.getSheets()[0];
    
    abaDest.getRange(1, 1, 1, totalColunas).setValues([cabecalhoFinal]).setFontWeight("bold").setBackground("#6366f1").setFontColor("white");
    abaDest.getRange(2, 1, totalLinhas, totalColunas).setValues(matrizPerfeita);
    abaDest.setFrozenRows(1);
    abaDest.autoResizeColumns(1, totalColunas);

    return novaPlanilha.getUrl();

  } catch(e) {
    return { erro: "Falha ao gerar planilha: " + e.message };
  }
}

// =========================================================================
// MÓDULO DE ACOMPANHAMENTO E AUDITORIA (ANO RHC)
// =========================================================================

function pertenceAoAnoRHC(mesTexto, anoTexto, anoReferencia, tipoAno) {
  if (!mesTexto || !anoTexto || !anoReferencia) return false;
  const meses = { "JAN":1, "FEV":2, "MAR":3, "ABR":4, "MAI":5, "JUN":6, "JUL":7, "AGO":8, "SET":9, "OUT":10, "NOV":11, "DEZ":12 };
  
  const nMes = meses[String(mesTexto).trim().toUpperCase().substring(0,3)];
  const nAno = parseInt(anoTexto, 10);
  const ref = parseInt(anoReferencia, 10);
  
  if (!nMes || isNaN(nAno) || isNaN(ref)) return false;

  // 🛡️ NOVA REGRA: Ano Civil
  if (tipoAno === "CIVIL") {
    return nAno === ref;
  }

  // Ano RHC (Padrão): Vai de Outubro do ano anterior (ref - 1) até Setembro do ano alvo (ref)
  if (nMes >= 10 && nAno === (ref - 1)) return true;
  if (nMes <= 9 && nAno === ref) return true;
  
  return false;
}

function buscarDadosAcompanhamento(anoRHC) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetLinks = ss.getSheetByName("DB_LINKS");
    if (!sheetLinks) return { erro: "Aba DB_LINKS não encontrada." };
    
    const dataLinks = sheetLinks.getDataRange().getValues();
    const basesAlvo = [];
    
    // 1. Filtragem rápida de bases no Hub
    for (let i = 1; i < dataLinks.length; i++) {
      let passo = String(dataLinks[i][4]).trim().toUpperCase();
      let status = String(dataLinks[i][7]).toUpperCase();
      if (passo.includes("PASSO 2") && !status.includes("XLS") && !status.includes("ERRO")) {
        if (pertenceAoAnoRHC(dataLinks[i][3], dataLinks[i][2], anoRHC)) {
          basesAlvo.push({ id: dataLinks[i][0], url: dataLinks[i][5], mes: dataLinks[i][3], ano: dataLinks[i][2] });
        }
      }
    }

    let resumoGlobal = { ciano: 0, magenta: 0, brancoPreto: 0, brancoVermelho: 0, inconsistente: 0, total: 0 };
    let dadosPorMes = {}; 
    let listaCompleta = []; 

    const ordemRHC = ["OUT", "NOV", "DEZ", "JAN", "FEV", "MAR", "ABR", "MAI", "JUN", "JUL", "AGO", "SET"];
    ordemRHC.forEach(m => {
      dadosPorMes[m] = { ciano: 0, magenta: 0, brancoPreto: 0, brancoVermelho: 0, inconsistente: 0, total: 0, link: "", pacientes: [] };
    });

    for (let k = 0; k < basesAlvo.length; k++) {
      let base = basesAlvo[k];
      let ws = SpreadsheetApp.openByUrl(base.url);
      let aba = ws.getSheets()[0];
      
      // TRAVA DE SEGURANÇA: No máximo 1001 linhas (Header + 1000 pacientes)
      let totalLinhasPlanilha = aba.getLastRow();
      let limiteLinhas = Math.min(totalLinhasPlanilha, 1001);
      
      if (limiteLinhas < 2) continue;

      // 2. Extração Ultra-Rápida de Valores (Apenas as colunas necessárias: A até I)
      // Lemos até a coluna 9 (I) para pegar Nome, Prontuário e Dias.
      let values = aba.getRange(1, 1, limiteLinhas, 9).getValues();
      
      // Índices fixos baseados na estrutura padrão:
      // Coluna B (1) = Nome | Coluna C (2) = Prontuário | Coluna I (8) = Lei 60 Dias
      const idxNome = 1;
      const idxPront = 2;
      const idxLei = 8; 

      dadosPorMes[base.mes].link = base.url;

      // 3. EXTRAÇÃO SNIPER DE FORMATAÇÃO: Apenas Coluna I (9)
      let rangeI = aba.getRange(1, 9, limiteLinhas, 1);
      let bgs = rangeI.getBackgrounds();
      let fColors = rangeI.getFontColors();
      let fWeights = rangeI.getFontWeights();

      for (let i = 1; i < values.length; i++) {
        let nome = String(values[i][idxNome]);
        let pront = String(values[i][idxPront]);

        if (!nome || !pront || nome === "undefined") continue;

        // Pegamos a formatação da linha atual na Coluna I
        let bg = String(bgs[i][0]).toLowerCase();
        let fc = String(fColors[i][0]).toLowerCase();
        let fw = String(fWeights[i][0]).toLowerCase();
        
        let diasStr = String(values[i][idxLei]);
        let dias = parseFloat(diasStr.replace(',', '.'));

        let motivo = null;
        let corIdentificada = bg;

        // --- LÓGICA DE AUDITORIA ON-THE-FLY ---
        
        // 1. Ciano: Sem Conduta
        if (['#00ffff', '#0ff'].includes(bg)) {
          motivo = "Sem Conduta Fechada";
          resumoGlobal.ciano++;
          dadosPorMes[base.mes].ciano++;
        } 
        // 2. Magenta: Revisão (Futuro)
        else if (['#ff00ff', '#9900ff', '#8e7cc3'].includes(bg)) {
          motivo = "Revisão (Futuro)";
          resumoGlobal.magenta++;
          dadosPorMes[base.mes].magenta++;
        }
        // 3. Branco + Fonte Preta: Não Trabalhado
        else if ((bg === '#ffffff' || bg === 'white') && (fc === '#000000' || fc === 'black' || fc === '#000')) {
          motivo = "Não Trabalhado";
          resumoGlobal.brancoPreto++;
          dadosPorMes[base.mes].brancoPreto++;
        }
        // 4. Branco + Fonte Vermelha Negrito: Revisão de Excluídos
        else if ((bg === '#ffffff' || bg === 'white') && (fc.includes('ff0000') || fc.includes('red')) && fw === 'bold') {
          motivo = "Revisão de Excluídos";
          resumoGlobal.brancoVermelho++;
          dadosPorMes[base.mes].brancoVermelho++;
        }
        // 5. Inconsistências Matemáticas
        else if (!isNaN(dias)) {
          let isVerde = ['#00ff00', '#34a853', '#d9ead3'].some(c => bg.includes(c));
          let isVermelho = ['#ff0000', '#ea4335', '#f4cccc'].some(c => bg.includes(c));
          
          if (isVerde && dias > 60) {
            motivo = "Inconsistência (>60 em Verde)";
            corIdentificada = "#00ff00";
          } else if (isVermelho && dias <= 60) {
            motivo = "Inconsistência (<=60 em Vermelho)";
            corIdentificada = "#ff0000";
          }

          if (motivo) {
            resumoGlobal.inconsistente++;
            dadosPorMes[base.mes].inconsistente++;
          }
        }

        if (motivo) {
          let objPac = { 
            prontuario: pront, 
            nome: nome, 
            dias: diasStr, 
            cor: corIdentificada, 
            motivo: motivo, 
            mesAno: base.mes+"/"+base.ano, 
            url: base.url 
          };
          listaCompleta.push(objPac);
          dadosPorMes[base.mes].pacientes.push(objPac);
        }
      }
    }

    return { sucesso: true, resumoGlobal, dadosPorMes, listaCompleta, anoRHC: anoRHC };
  } catch(e) { return { erro: e.message }; }
}

/**
 * Renomeia/Exclui o agrupador (Projeto/Pasta) ou o Título do Script na planilha
 */
function renomearItemRHC(id, novoNome, tipo) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DB_CODIGOS");
  var data = sheet.getDataRange().getValues();
  
  if (tipo === 'pasta') {
    // Aqui o 'id' é o nome antigo do projeto
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] === id) sheet.getRange(i + 1, 2).setValue(novoNome);
    }
  } else {
    // Aqui o 'id' é o UUID único do script
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === id) {
        sheet.getRange(i + 1, 3).setValue(novoNome);
        break;
      }
    }
  }
  return true;
}

function excluirItemRHC(id, tipo) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DB_CODIGOS");
  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (tipo === 'pasta') {
      if (data[i][1] === id) sheet.deleteRow(i + 1);
    } else {
      if (data[i][0] === id) { sheet.deleteRow(i + 1); break; }
    }
  }
  return true;
}

/**
 * Ponto de entrada para o Web App (Tela Cheia)
 * Não afeta em nada a abertura via onOpen
 */
function doGet(e) {
  // Retorna o novo arquivo HTML que criaremos para o Dashboard
  return HtmlService.createTemplateFromFile('DashboardView')
      .evaluate()
      .setTitle('Painel RHC HCPA- Dashboard BI')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Função utilitária para obter a URL do Web App e passar para o botão do Painel
 */
function obterUrlWebApp() {
  return ScriptApp.getService().getUrl();
}

/**
 * Função Extratora Unificada (Suporte para Base Legada e Novas Bases)
 * @param {number} anoRHC - O ano solicitado
 * @param {boolean} isLegacy - Se true, aplica o filtro de corte (pré-outubro/2025). 
 * Para 2026 em diante, passe 'false' na chamada.
 */
function buscarHistoricoDash(anoRHC, isLegacy = true) {
  var dadosProcessados = [];
  var erros = [];
  try {
    var idExterna = "1kzn7fHWQcCimDEWRLOf5ARVK5iUJOZ89aXbyQGuuQQ8";
    var ws = SpreadsheetApp.openById(idExterna);
    var sheet = ws.getSheetByName("Dados");
    if (!sheet) throw new Error("Aba 'Dados' não encontrada.");

    var dataRange = sheet.getDataRange();
    var values = dataRange.getValues();
    var displayValues = dataRange.getDisplayValues();
    var backgrounds = dataRange.getBackgrounds();

    if (values.length < 2) return { dados: [], erros: [] };

    var headers = values[0].map(function(h) { return String(h).toUpperCase().trim(); });

    // Mapeamento dinâmico histórico
    var map = {
      lei60: headers.indexOf('LEI 60 DIAS'),
      cid: headers.indexOf('CID DIAGNÓSTICO') !== -1 ? headers.indexOf('CID DIAGNÓSTICO') : headers.indexOf('CID'),
      prontuario: headers.indexOf('PRONTUARIO') !== -1 ? headers.indexOf('PRONTUARIO') : headers.indexOf('PRONTUAR'),
      diagTrat: headers.indexOf('DIAGNÓSTICO E TRATAMENTOS ANTERIORES'),
      nome: headers.indexOf('NOME DO PACIENTE') !== -1 ? headers.indexOf('NOME DO PACIENTE') : headers.indexOf('NOME'),
      dataDiag: headers.indexOf('DATA DIAGNÓSTICO') !== -1 ? headers.indexOf('DATA DIAGNÓSTICO') : headers.indexOf('DTDIAGNO'),
      dataCons: headers.indexOf('DATA DA PRIMEIRA CONSULTA') !== -1 ? headers.indexOf('DATA DA PRIMEIRA CONSULTA') : headers.indexOf('DTPRICON'),
      sexo: headers.indexOf('SEXO'),
      etnia: headers.indexOf('ETNIA'),
      idade: headers.indexOf('IDADE'),
      nasc: headers.indexOf('DT_NASCIMENTO') !== -1 ? headers.indexOf('DT_NASCIMENTO') : headers.indexOf('DATANASC'),
      cidade: headers.indexOf('CIDADE_MORADIA') !== -1 ? headers.indexOf('CIDADE_MORADIA') : headers.indexOf('CIDADE'),
      obito1: headers.indexOf('DT_OBITO') !== -1 ? headers.indexOf('DT_OBITO') : headers.indexOf('DTOBITO'),
      obito2: headers.indexOf('DATA DO ÓBITO'),
      baseDiag: headers.indexOf('BASE MAIS IMPORTANTE DO DIAGNÓSTICO') !== -1 ? headers.indexOf('BASE MAIS IMPORTANTE DO DIAGNÓSTICO') : headers.indexOf('BASMAIMP'),
      clinicaTrat: headers.indexOf('CLÍNICA DO INÍCIO DO TRATAMENTO') !== -1 ? headers.indexOf('CLÍNICA DO INÍCIO DO TRATAMENTO') : headers.indexOf('CLINPRIM'),
      razaoNaoTrat: headers.indexOf('PRINCIPAL RAZÃO PARA NÃO REALIZAÇÃO DO PRIMEIRO TRATAMENTO NO HOSPITAL') !== -1 ? headers.indexOf('PRINCIPAL RAZÃO PARA NÃO REALIZAÇÃO DO PRIMEIRO TRATAMENTO NO HOSPITAL') : headers.indexOf('RAZAONT'),
      primTrat: headers.indexOf('1º TRATAMENTO REALIZADO NO HOSPITAL') !== -1 ? headers.indexOf('1º TRATAMENTO REALIZADO NO HOSPITAL') : headers.indexOf('PRITRATH'),
      estadoDoenca: headers.indexOf('ESTADO DA DOENÇA AO FINAL DO PRIMEIRO TRATAMENTO') !== -1 ? headers.indexOf('ESTADO DA DOENÇA AO FINAL DO PRIMEIRO TRATAMENTO') : headers.indexOf('ESTDFIMT'),
      locTumor: headers.indexOf('LOCALIZAÇÃO DO TUMOR PRIMÁRIO') !== -1 ? headers.indexOf('LOCALIZAÇÃO DO TUMOR PRIMÁRIO') : headers.indexOf('LOCTUPRI'),
      tipoHist: headers.indexOf('TIPO HISTOLÓGICO (CID-O)') !== -1 ? headers.indexOf('TIPO HISTOLÓGICO (CID-O)') : headers.indexOf('TIPOHIST')
    };

    function getMesAnoInt(val) {
      if (!val) return null;
      let mes = 0, ano = 0;
      
      // Correção adicionada: tratar formato YYYY-MM-DD
      if (typeof val === 'string' && val.includes('-') && val.length === 10) {
          let partes = val.split('-');
          mes = parseInt(partes[1], 10);
          ano = parseInt(partes[0], 10);
          return { mes: mes, ano: ano };
      }

      if (val instanceof Date) {
        mes = val.getMonth() + 1;
        ano = val.getFullYear();
      } else {
        let str = String(val).trim().toLowerCase();
        let matchBr = str.match(/(?:^|\D)(0[1-9]|[12]\d|3[01])[\/\-](0[1-9]|1[0-2])[\/\-](\d{4})/);
        if (matchBr) {
          mes = parseInt(matchBr[2], 10);
          ano = parseInt(matchBr[3], 10);
        } else {
          let matchText = str.match(/(jan|fev|mar|abr|mai|jun|jul|ago|set|out|nov|dez)[\w]*[\/\-\s_]*(\d{4})/i);
          const mesesNum = {jan:1, fev:2, mar:3, abr:4, mai:5, jun:6, jul:7, ago:8, set:9, out:10, nov:11, dez:12};
          if (matchText) {
            mes = mesesNum[matchText[1].toLowerCase()];
            ano = parseInt(matchText[2], 10);
          } else {
            let matchMy = str.match(/(0[1-9]|1[0-2])[\/\-](\d{4})/);
            if (matchMy) {
              mes = parseInt(matchMy[1], 10);
              ano = parseInt(matchMy[2], 10);
            }
          }
        }
      }
      if (mes > 0 && ano > 0) return { mes: mes, ano: ano };
      return null;
    }

    const mesesMap = { 1:"JAN", 2:"FEV", 3:"MAR", 4:"ABR", 5:"MAI", 6:"JUN", 7:"JUL", 8:"AGO", 9:"SET", 10:"OUT", 11:"NOV", 12:"DEZ" };
    var isGlobal = (String(anoRHC).toUpperCase() === "GLOBAL");
    var anoFiltro = parseInt(anoRHC, 10);

    for (var i = 1; i < values.length; i++) {
      var linha = values[i];
      var valDataAProcessar = map.dataDiag !== -1 ? linha[map.dataDiag] : (map.dataCons !== -1 ? linha[map.dataCons] : null);
      var parsedData = getMesAnoInt(valDataAProcessar);
      
      if (!parsedData) continue;

      // Filtro Opcional: Se isLegacy for true, descarta dados pós-setembro/2025
      if (isLegacy) {
          if (parsedData.ano > 2025 || (parsedData.ano === 2025 && parsedData.mes >= 10)) continue;
      }

      // Filtro de competência (sempre necessário para a visão do BI)
      if (!isGlobal && !isNaN(anoFiltro)) {
          if (!((parsedData.ano === anoFiltro - 1 && parsedData.mes >= 10) || (parsedData.ano === anoFiltro && parsedData.mes <= 9))) {
              continue;
          }
      }

      var mesAnoStr = mesesMap[parsedData.mes] + "/" + parsedData.ano;
      var corLei = map.lei60 !== -1 ? backgrounds[i][map.lei60] : '#ffffff';
      var cor = corLei && corLei !== '#ffffff' ? corLei : '#ffffff';

      var valNome = map.nome !== -1 && linha[map.nome] ? String(linha[map.nome]).trim() : "";
      var valPront = map.prontuario !== -1 && linha[map.prontuario] ? String(linha[map.prontuario]).trim() : "";
      if (valNome === "" || valPront === "") continue;

      var displayLei = map.lei60 !== -1 && displayValues[i][map.lei60] ? String(displayValues[i][map.lei60]).trim() : '';
      var lei60 = null;
      var inconclusivo = false;

      if (displayLei !== '' && displayLei !== '-') {
          if (displayLei.startsWith('#') || displayLei.toUpperCase().includes('NAN') || displayLei.toUpperCase().includes('ERRO')) {
              inconclusivo = true;
          } else {
              var numStr = displayLei.replace(',', '.');
              var num = parseFloat(numStr);
              if (isNaN(num)) inconclusivo = true;
              else lei60 = num < 0 ? 0 : num;
          }
      }

      var cid = map.cid !== -1 && linha[map.cid] ? String(linha[map.cid]).trim() : 'Sem CID';
      var diagTrat = map.diagTrat !== -1 && linha[map.diagTrat] ? String(linha[map.diagTrat]).trim() : '';
      var sexo = (map.sexo !== -1 && linha[map.sexo] != null) ? String(linha[map.sexo]).trim().toUpperCase() : "";
      var etnia = (map.etnia !== -1 && linha[map.etnia] != null) ? String(linha[map.etnia]).trim().toUpperCase() : "";
      var cidade = (map.cidade !== -1 && linha[map.cidade] != null) ? String(linha[map.cidade]).trim() : "";
      
      var obito = false;
      var valObito1 = map.obito1 !== -1 && linha[map.obito1] != null ? String(linha[map.obito1]).trim().toUpperCase() : "";
      var valObito2 = map.obito2 !== -1 && linha[map.obito2] != null ? String(linha[map.obito2]).trim().toUpperCase() : "";
      var valObitoFinal = valObito2 ? valObito2 : valObito1;
      
      if (valObitoFinal !== "" && valObitoFinal !== "N/A" && valObitoFinal !== "-" && valObitoFinal !== "NÃO" && valObitoFinal !== "NAO") {
          obito = true;
      }

      var idade = null;
      if (map.idade !== -1 && linha[map.idade] != null && linha[map.idade] !== "") {
          var idadeTentativa = parseInt(linha[map.idade], 10);
          if (!isNaN(idadeTentativa)) idade = idadeTentativa;
      }

      var nasc = (map.nasc !== -1 && displayValues[i][map.nasc] != null && displayValues[i][map.nasc] !== "") ? String(displayValues[i][map.nasc]).trim() : null;
      var dtDiagStr = (map.dataDiag !== -1 && displayValues[i][map.dataDiag]) ? String(displayValues[i][map.dataDiag]).trim() : "";

      dadosProcessados.push({
          idBase: "HISTORICO",
          corSimulada: cor.toLowerCase(),
          lei60: lei60,
          cid: cid,
          inconclusivo: inconclusivo,
          nome: valNome,
          prontuario: valPront,
          diagTrat: diagTrat,
          dtDiag: dtDiagStr, 
          mesAno: mesAnoStr,
          sexo: sexo,
          etnia: etnia,
          idade: idade,
          nascimento: nasc,
          cidade: cidade,
          obito: obito,
          baseDiag: map.baseDiag !== -1 && linha[map.baseDiag] ? String(linha[map.baseDiag]).trim() : "",
          clinicaTrat: map.clinicaTrat !== -1 && linha[map.clinicaTrat] ? String(linha[map.clinicaTrat]).trim() : "",
          razaoNaoTrat: map.razaoNaoTrat !== -1 && linha[map.razaoNaoTrat] ? String(linha[map.razaoNaoTrat]).trim() : "",
          primTrat: map.primTrat !== -1 && linha[map.primTrat] ? String(linha[map.primTrat]).trim() : "",
          estadoDoenca: map.estadoDoenca !== -1 && linha[map.estadoDoenca] ? String(linha[map.estadoDoenca]).trim() : "",
          locTumor: map.locTumor !== -1 && linha[map.locTumor] ? String(linha[map.locTumor]).trim() : "",
          tipoHist: map.tipoHist !== -1 && linha[map.tipoHist] ? String(linha[map.tipoHist]).trim() : ""
      });
    }
  } catch (e) {
    erros.push("Erro no processamento (" + anoRHC + "): " + e.message);
  }
  return { dados: dadosProcessados, erros: erros };
}

// =========================================================================
// FERRAMENTA DE TRIAGEM RÁPIDA (NOVOS VS BASE)
// =========================================================================

/**
 * Função principal que compara a aba "Novos" com a "Base"
 * Com filtro de alta precisão (ignora zeros à esquerda e decimais .0)
 * Avalia 5 Chaves: PRONTUAR, IDTUMOR, NOME, SEXO, DATANASC
 */
function destacarNovosCasos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaBase = ss.getSheetByName("Base");
  const abaNovos = ss.getSheetByName("Novos");
  
  if (!abaBase || !abaNovos) {
    SpreadsheetApp.getUi().alert("⚠️ Erro: As abas 'Base' e 'Novos' precisam existir com esses nomes exatos.");
    return;
  }
  
  const ui = SpreadsheetApp.getUi();
  ui.alert("⏳ Iniciando análise de alta precisão...", "Lendo as células como texto puro e removendo acentos. Isso pode levar alguns segundos.", ui.ButtonSet.OK);

  // O SEGREDO MUDOU AQUI: getDisplayValues pega o que os olhos veem, não o código da célula!
  const dadosBase = abaBase.getDataRange().getDisplayValues();
  const intervaloNovos = abaNovos.getDataRange();
  const dadosNovos = intervaloNovos.getDisplayValues();
  const coresAtuais = intervaloNovos.getBackgrounds();
  const corDestaque = "#fff2cc"; 
  
  // =========================================================
  // 🛡️ MOTOR DE LIMPEZA BLINDADO
  // =========================================================
  function normalizar(valor) {
    if (!valor) return "";
    return String(valor)
      .normalize("NFD").replace(/[\u0300-\u036f]/g, "") // Remove todos os acentos (JOÃO vira JOAO)
      .replace(/\s+/g, ' ')    // Transforma espaços duplos em um só
      .trim()                  // Remove espaços antes e depois
      .toUpperCase()           // Padroniza letras maiúsculas
      .replace(/\.0$/, '')     // Remove o ".0" fantasma que o Excel costuma exportar
      .replace(/^0+/, '');     // Remove os zeros à esquerda
  }

  // =========================================================
  // 🔍 LOCALIZADOR DINÂMICO DE COLUNAS
  // =========================================================
  const headersBase = dadosBase[0].map(h => String(h).toUpperCase().trim());
  const iPront = headersBase.indexOf("PRONTUAR") !== -1 ? headersBase.indexOf("PRONTUAR") : 2;
  const iTumor = headersBase.indexOf("IDTUMOR") !== -1 ? headersBase.indexOf("IDTUMOR") : 3;
  const iNome = headersBase.indexOf("NOME");
  const iSexo = headersBase.indexOf("SEXO");
  const iNasc = headersBase.indexOf("DATANASC");

  const headersNovos = dadosNovos[0].map(h => String(h).toUpperCase().trim());
  const nPront = headersNovos.indexOf("PRONTUAR") !== -1 ? headersNovos.indexOf("PRONTUAR") : 2;
  const nTumor = headersNovos.indexOf("IDTUMOR") !== -1 ? headersNovos.indexOf("IDTUMOR") : 3;
  const nNome = headersNovos.indexOf("NOME");
  const nSexo = headersNovos.indexOf("SEXO");
  const nNasc = headersNovos.indexOf("DATANASC");

  // 1. Criar o "Dicionário" (Set) com as chaves limpas da Base
  const chavesBase = new Set();
  
  for (let i = 1; i < dadosBase.length; i++) {
    let prontuario = normalizar(dadosBase[i][iPront]);
    let idTumor = normalizar(dadosBase[i][iTumor]);
    let nome = iNome !== -1 ? normalizar(dadosBase[i][iNome]) : "";
    let sexo = iSexo !== -1 ? normalizar(dadosBase[i][iSexo]) : "";
    let nasc = iNasc !== -1 ? normalizar(dadosBase[i][iNasc]) : "";
    
    if (prontuario !== "") {
      chavesBase.add(prontuario + "|" + idTumor + "|" + nome + "|" + sexo + "|" + nasc);
    }
  }
  
  // 2. Analisar a aba Novos
  let contadorNovos = 0;
  
  for (let i = 1; i < dadosNovos.length; i++) {
    let prontuarioNovos = normalizar(dadosNovos[i][nPront]);
    let idTumorNovos = normalizar(dadosNovos[i][nTumor]);
    let nomeNovos = nNome !== -1 ? normalizar(dadosNovos[i][nNome]) : "";
    let sexoNovos = nSexo !== -1 ? normalizar(dadosNovos[i][nSexo]) : "";
    let nascNovos = nNasc !== -1 ? normalizar(dadosNovos[i][nNasc]) : "";
    
    if (prontuarioNovos !== "") {
      let chaveBusca = prontuarioNovos + "|" + idTumorNovos + "|" + nomeNovos + "|" + sexoNovos + "|" + nascNovos;
      
      if (!chavesBase.has(chaveBusca)) {
        for (let j = 0; j < coresAtuais[i].length; j++) {
          coresAtuais[i][j] = corDestaque;
        }
        contadorNovos++;
      }
    }
  }
  
  // 3. Aplica todas as cores de uma vez só na planilha
  intervaloNovos.setBackgrounds(coresAtuais);
  
  // 4. Feedback final
  if (contadorNovos > 0) {
    ui.alert("✅ Análise Refinada Concluída!", `Foram encontrados e destacados ${contadorNovos} registros REALMENTE inéditos na guia 'Novos'.`, ui.ButtonSet.OK);
  } else {
    ui.alert("✅ Análise Refinada Concluída!", "Nenhum caso novo foi encontrado. Todos os registros já constam na 'Base'.", ui.ButtonSet.OK);
  }
}
