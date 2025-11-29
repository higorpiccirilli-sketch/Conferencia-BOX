/**
 * ===================================================================================
 * Alerta_Falhas_Preco – Planilha -> E-mail (Alerta de Preço Zerado)
 * Arquivo: Alerta_Falhas.gs
 * -----------------------------------------------------------------------------------
 * Criadores: Higor & Gemini
 * Versão: v1.0.0  |  Data: 13/11/2025  |  TZ: America/Sao_Paulo
 * -----------------------------------------------------------------------------------
 * O que faz?
 * - Função principal: verificarEAlertarFalhas()
 * - É executado por um acionador de tempo (ex: a cada 1 hora).
 * - Verifica a aba "⚠️Falhas_Preço" para identificar se há dados de erro.
 * - Se houver dados (linhas > 3), busca e-mails em células específicas (G1:H2).
 * - Envia um e-mail de alerta para os destinatários encontrados.
 *
 * Configurações & Lógica:
 * - Aba Alvo: "⚠️Falhas_Preço" (variável NOME_DA_ABA)
 * - Linha de Cabeçalho: 3 (variável LINHA_DO_CABECALHO)
 * - Condição de Alerta: Se `aba.getLastRow()` for maior que a LINHA_DO_CABECALHO.
 * - Destinatários Dinâmicos: E-mails são lidos do intervalo G1:H2 da aba alvo.
 * - Wrappers (Logger): [Falhas encontradas], [Nenhuma falha encontrada], [E-mail enviado], [Erro: Aba não encontrada]
 * ===================================================================================
 */
function verificarEAlertarFalhas() {
  
  // 1. Definições Iniciais
  var NOME_DA_ABA = "⚠️Falhas_Preço";
  var LINHA_DO_CABECALHO = 3;
  
  // 2. Acessar a planilha e a aba
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var aba = planilha.getSheetByName(NOME_DA_ABA);

  // Se a aba não existir, pare o script
  if (!aba) {
    Logger.log("Erro: A aba '" + NOME_DA_ABA + "' não foi encontrada.");
    return;
  }

  // 3. Verificar se há dados
  var ultimaLinhaComDados = aba.getLastRow();
  
  if (ultimaLinhaComDados > LINHA_DO_CABECALHO) {
    // CONDIÇÃO VERDADEIRA: Existem dados de falha, precisamos alertar.
    Logger.log("Falhas encontradas. Verificando destinatários...");

    // 4. Buscar os e-mails de destino (G1:H2)
    var intervaloEmails = aba.getRange("G1:H2");
    var emailsArray2D = intervaloEmails.getValues(); // Retorna [['g1', 'h1'], ['g2', 'h2']]
    
    // Converte o array 2D em uma lista 1D e remove células vazias
    var listaEmails = emailsArray2D.flat().filter(function(email) {
      return email !== "";
    });

    // Se não houver e-mails preenchidos, pare o script
    if (listaEmails.length === 0) {
      Logger.log("Falhas encontradas, mas nenhum e-mail de destinatário foi configurado em G1:H2.");
      return;
    }
    
    // 5. Preparar e enviar o e-mail
    var emailsParaEnvio = listaEmails.join(","); // "email1@g.com,email2@g.com"
    var assunto = "ALERTA: Falha de Preço de Produto Encontrada";
    var urlPlanilha = planilha.getUrl();
    var corpoEmail = "Foram encontrados produtos com saída de estoque e preço de custo ou venda zerado.\n\n" +
                     "Por favor, verifique a aba '" + NOME_DA_ABA + "' para tomar uma ação.\n\n" +
                     "Link para a planilha:\n" +
                     urlPlanilha;
    
    // Envia o e-mail
    MailApp.sendEmail(emailsParaEnvio, assunto, corpoEmail);
    Logger.log("E-mail de alerta enviado para: " + emailsParaEnvio);

  } else {
    // CONDIÇÃO FALSA: Nenhuma falha encontrada (a última linha é 3 ou menos)
    Logger.log("Nenhuma falha de preço encontrada. Nenhum e-mail enviado.");
  }
}
