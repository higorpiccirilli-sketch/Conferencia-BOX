/**
 * ===================================================================================
 * MB2_Importador – Metabase → Google Sheets (cards 1328, 1317, 1330, 1318)
 * Arquivo: MB2_Importador.gs
 * -----------------------------------------------------------------------------------
 * Criadores: Higor Piccirilli & GPT-5 Thinking (ChatGPT)
 * Versão: v3.2.0  |  Data: 2025-11-28  |  TZ: America/Sao_Paulo
 * -----------------------------------------------------------------------------------
 * O que faz?
 * - Login no Metabase
 * - Busca JSON dos cards
 * - Escreve/atualiza as abas alvo
 * - Adiciona NOTA na célula A1 com data/hora da atualização
 *
 * Integração LogRT (Minimalist Edition):
 * - Modal limpo (apenas texto), sem botões, sem histórico persistente.
 * - Cache temporário apenas durante a execução.
 * ===================================================================================
 */

// -----------------------------------------------------------------
// CREDENCIAIS (Script Properties; SEM senha em código)
// -----------------------------------------------------------------
var MB2_URL_BASE      = (PropertiesService.getScriptProperties().getProperty('MB_URL')         || '').trim();
var MB2_USUARIO       = (PropertiesService.getScriptProperties().getProperty('MB_USER')        || '').trim();
var MB2_SENHA         = (PropertiesService.getScriptProperties().getProperty('MB_PASSWORD')    || '').trim();
var MB2_EMAIL_ALERTA  = (PropertiesService.getScriptProperties().getProperty('MB_ALERT_EMAIL') || '').trim();

// -----------------------------------------------------------------
// LISTA DE RELATÓRIOS / ABAS ALVO
// -----------------------------------------------------------------
var MB2_RELATORIOS = [
  { cardId: 1328, nomeAba: '_CadastroProdutosBox' }, // A:M
  { cardId: 1317, nomeAba: '_BoxQuantidade' },       // A:D
  { cardId: 1330, nomeAba: '_SaidasBox' },
  { cardId: 1318, nomeAba: '_ShopQuantidade' }       // <--- RENOMEADO
];

// -----------------------------------------------------------------
// MENU
// -----------------------------------------------------------------
function MB2_onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Metabase (2)')
      .addItem('Atualizar TUDO (com Log)', 'MB2_abrirLog_e_AtualizarTodos')
      .addSeparator()
      .addItem('_CadastroProdutosBox (sem Log)', 'MB2_atualizar1328')
      .addItem('_BoxQuantidade (sem Log)',       'MB2_atualizar1317')
      .addItem('_SaidasBox (sem Log)',           'MB2_atualizarSaidasBox')
      .addItem('_ShopQuantidade (sem Log)',      'MB2_atualizarShopQuantidade')
    .addToUi();
}

// -----------------------------------------------------------------
// ENTRADA COM LOG (Interface Minimalista)
// -----------------------------------------------------------------
function MB2_abrirLog_e_AtualizarTodos() {
  MB2_assertCredenciais();

  var execId = MB2_newExecId_();
  RT.clear(execId); // Garante limpeza inicial

  RT.iniciando(execId, 'Iniciando atualização completa...');

  // Prepara o HTML do modal minimalista
  var t = HtmlService.createTemplateFromFile('RTLog');
  t.execId         = execId;
  t.serverFunction = 'MB2_atualizarTodos_comLog'; 
  
  // Altura e largura ajustadas para foco no conteúdo
  var html = t.evaluate().setWidth(800).setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, 'Log de Execução');
}

// -----------------------------------------------------------------
// EXECUÇÃO COM LOG
// -----------------------------------------------------------------
function MB2_atualizarTodos_comLog(execId) {
  MB2_assertCredenciais();
  execId = execId || MB2_newExecId_();

  var ok = 0, fail = 0;

  MB2_emit(execId, 'andamento', 'Autenticando...');
  var token = MB2_loginMetabase_(2);
  if (!token) {
    MB2_emit(execId, 'erroFatal', 'Falha de autenticação.');
    throw new Error('Token vazio.');
  }
  
  MB2_RELATORIOS.forEach(function(r) {
    var etapa = r.nomeAba + ' (ID ' + r.cardId + ')';
    try {
      MB2_emit(execId, 'andamento', 'Consultando: ' + etapa);
      var result = MB2_buscarDadosDoMetabase(token, r.cardId, r.nomeAba);

      if (result && result.ok) {
        MB2_emit(execId, 'ok', 'Sucesso: ' + etapa + ' [' + result.rows + ' linhas]');
        ok++;
      } else {
        MB2_emit(execId, 'erro', 'Falha no retorno: ' + etapa);
        fail++;
      }
    } catch (e) {
      fail++;
      MB2_emit(execId, 'erro', 'Erro em ' + etapa + ': ' + (e.message || e));
    }
  });

  var resumo = 'Processo finalizado. Sucesso: ' + ok + ' | Falhas: ' + fail;
  MB2_emit(execId, 'final', resumo);
}

// -----------------------------------------------------------------
// AÇÕES INDIVIDUAIS (SEM LOG - TOAST ONLY)
// -----------------------------------------------------------------
function MB2_atualizar1328() {
  MB2_executarIndividual(1328);
}

function MB2_atualizar1317() {
  MB2_executarIndividual(1317);
}

function MB2_atualizarSaidasBox() {
  MB2_executarIndividual(1330);
}

function MB2_atualizarShopQuantidade() {
  MB2_executarIndividual(1318);
}

// Helper para execuções individuais
function MB2_executarIndividual(cardId) {
  MB2_assertCredenciais();
  var token = MB2_loginMetabase_(2);
  var r = MB2_RELATORIOS.find(function(x){ return x.cardId === cardId; });
  if (!r) {
    SpreadsheetApp.getActive().toast('Configuração não encontrada para ID ' + cardId);
    return;
  }
  var result = MB2_buscarDadosDoMetabase(token, r.cardId, r.nomeAba);
  var msg = (result && result.ok) ? r.nomeAba + ' atualizada!' : 'Erro ao atualizar ' + r.nomeAba;
  SpreadsheetApp.getActive().toast(msg);
}


// -----------------------------------------------------------------
// CORE – Consulta e Escrita
// -----------------------------------------------------------------
function MB2_buscarDadosDoMetabase(sessionToken, cardId, nomeDaAba) {
  var queryOptions = {
    method: 'post',
    headers: { 'X-Metabase-Session': sessionToken },
    contentType: 'application/json',
    muteHttpExceptions: true
  };
  var url = MB2_URL_BASE + '/api/card/' + cardId + '/query/json';
  var queryResp = UrlFetchApp.fetch(url, queryOptions);
  MB2_throwIfHttpError_(queryResp, 'Falha card ' + cardId);

  var text = queryResp.getContentText() || '[]';
  var dados;
  try { dados = JSON.parse(text); } catch (e) { throw new Error('JSON inválido'); }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(nomeDaAba) || ss.insertSheet(nomeDaAba);
  sheet.clearContents();

  var timestamp = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm:ss');
  var notaTexto = "Última importação:\n" + timestamp;

  if (!dados || !dados.length) {
    var cell = sheet.getRange(1, 1);
    cell.setValue('Vazio (Card ' + cardId + ')');
    cell.setNote(notaTexto);
    return { ok: true, rows: 0, cols: 1 };
  }

  var headers = Object.keys(dados[0]);
  var valores = [headers];
  dados.forEach(function(linha) {
    var arr = headers.map(function(h) {
      var v = linha[h];
      return (v === null || v === undefined) ? '' : v;
    });
    valores.push(arr);
  });

  sheet.getRange(1, 1, valores.length, headers.length).setValues(valores);
  sheet.getRange(1, 1).setNote(notaTexto);
  // sheet.autoResizeColumns(1, headers.length); // Opcional: desativado para performance se quiser
  SpreadsheetApp.flush();

  return { ok: true, rows: dados.length, cols: headers.length };
}

// -----------------------------------------------------------------
// LOGIN & UTILS
// -----------------------------------------------------------------
function MB2_loginMetabase_(tentativas) {
  tentativas = Math.max(1, parseInt(tentativas || 1, 10));
  for (var i = 1; i <= tentativas; i++) {
    try {
      var opts = {
        method: 'post', contentType: 'application/json',
        payload: JSON.stringify({ username: MB2_USUARIO, password: MB2_SENHA }),
        muteHttpExceptions: true
      };
      var resp = UrlFetchApp.fetch(MB2_URL_BASE + '/api/session', opts);
      MB2_throwIfHttpError_(resp, 'Login falhou');
      return JSON.parse(resp.getContentText()).id;
    } catch (e) { Utilities.sleep(500); }
  }
  return null;
}

function MB2_throwIfHttpError_(resp, msg) {
  var code = resp ? resp.getResponseCode() : 0;
  if (code >= 400) throw new Error(msg + ' (HTTP ' + code + ')');
}

function MB2_assertCredenciais() {
  if (!MB2_URL_BASE || !MB2_USUARIO || !MB2_SENHA) throw new Error('Credenciais MB incompletas nas Script Properties.');
}

function MB2_newExecId_() {
  return 'MB2-' + new Date().getTime();
}

function MB2_emit(execId, nivel, mensagem) {
  if (execId) {
    try {
      if (nivel === 'final') RT.final(execId, mensagem);
      else if (nivel === 'erro' || nivel === 'erroFatal') RT.erro(execId, mensagem);
      else if (nivel === 'ok') RT.ok(execId, mensagem);
      else RT.andamento(execId, mensagem);
    } catch (_) {}
  }
}
