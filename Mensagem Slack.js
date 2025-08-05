// =================================================================
// CONFIGURAÇÕES GERAIS
// =================================================================

// URL do Webhook do Slack.
const SLACK_WEBHOOK_URL = "https://hooks.slack.com/services/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";

// ID da planilha de DADOS (a que contém "Meta 2025" e "Realizado 2025").
const SPREADSHEET_ID = "1jDgyRlmQXx5lYesUZKEbfJlhvxLDa2ax";

// Nomes das abas na planilha de DADOS.
const ABA_METAS = "Meta 2025";
const ABA_REALIZADO = "Realizado 2025";

// Coluna para análise. 'I' representa Julho.
const COLUNA_ANALISE = 'I';

// Mapeamento de todas as métricas que vamos analisar.
// Formato: [ "Nome da Métrica", "Linha", "Tipo de dado ('numero' ou 'moeda')" ]
const METRICAS = [
  // Seção de Vendas em PDVs
  ["Novas Vendas (PDVs)", "4", "numero"],
  ["Adições (PDVs)", "15", "numero"],
  ["TOTAL (PDVs)", "26", "numero"],
  // Seção de Vendas em MRR
  ["MRR de Novas Vendas", "44", "moeda"],
  ["MRR de Adições (CS)", "45", "moeda"],
  ["MRR TOTAL", "43", "moeda"],
  // Seção de Setup
  ["Setup (PDVs)", "48", "numero"],
  ["Setup (MRR)", "53", "moeda"],
];


// =================================================================
// FUNÇÃO PRINCIPAL
// =================================================================

/**
 * Busca os dados de metas e realizados, calcula o progresso
 * e envia um relatório formatado para o Slack.
 */
function enviarRelatorioDeMetas() {
  try {
    // Abre a planilha de DADOS específica pelo seu ID.
    // Isto garante que ele sempre leia a planilha correta, não importa onde o script esteja.
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);

    const sheetMetas = spreadsheet.getSheetByName(ABA_METAS);
    const sheetRealizado = spreadsheet.getSheetByName(ABA_REALIZADO);

    if (!sheetMetas) {
      throw new Error(`A aba "${ABA_METAS}" não foi encontrada na planilha de DADOS (ID: ${SPREADSHEET_ID}). Verifique o nome da aba.`);
    }
    if (!sheetRealizado) {
      throw new Error(`A aba "${ABA_REALIZADO}" não foi encontrada na planilha de DADOS (ID: ${SPREADSHEET_ID}). Verifique o nome da aba.`);
    }

    const nomeDoMes = "Julho";
    let mensagemSlack = `*📊 Relatório de Metas - ${nomeDoMes} 2025 📊*\n\nOpa time, segue a atualização dos nossos resultados até agora em ${nomeDoMes}!\n\n--- \n\n`;

    METRICAS.forEach(metricaInfo => {
      const nome = metricaInfo[0];
      const linha = metricaInfo[1];
      const tipo = metricaInfo[2];
      const celula = COLUNA_ANALISE + linha;
      const meta = sheetMetas.getRange(celula).getValue();
      const realizado = sheetRealizado.getRange(celula).getValue();
      mensagemSlack += gerarBlocoDeMetrica(nome, meta, realizado, tipo);
    });
    
    enviarMensagemSlack(mensagemSlack);
    Logger.log("Sucesso! O relatório de metas foi enviado para o Slack.");

  } catch (e) {
    Logger.log("Ocorreu um erro ao gerar o relatório. Detalhes: " + e.toString());
  }
}


// =================================================================
// FUNÇÕES AUXILIARES
// =================================================================

function gerarBlocoDeMetrica(nome, meta, realizado, tipo) {
  const metaNum = Number(meta) || 0;
  const realizadoNum = Number(realizado) || 0;
  let percentual = 0;
  if (metaNum > 0) {
    percentual = (realizadoNum / metaNum) * 100;
  } else if (realizadoNum > 0) {
    percentual = 100;
  }
  let emoji = "⏳";
  let statusTexto = "";
  const diferenca = metaNum - realizadoNum;
  if (percentual >= 100) {
    emoji = "✅ *META BATIDA!*";
    statusTexto = `Parabéns, superamos a meta em ${formatarValor(realizadoNum - metaNum, tipo)}!`;
  } else {
    emoji = "🎯";
    statusTexto = `Faltam *${formatarValor(diferenca, tipo)}* para atingir a meta.`;
  }
  let bloco = `*${nome}*\n`;
  bloco += `• *Meta:* ${formatarValor(metaNum, tipo)} | *Realizado:* ${formatarValor(realizadoNum, tipo)}\n`;
  bloco += `• *Progresso:* ${percentual.toFixed(1)}% ${emoji}\n`;
  bloco += `• _${statusTexto}_\n\n`;
  return bloco;
}

function formatarValor(valor, tipo) {
  if (tipo === 'moeda') {
    return valor.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
  }
  return Math.round(valor).toString();
}

function enviarMensagemSlack(texto) {
  const payload = {
    "text": texto,
    "username": "Robô de Metas",
    "icon_emoji": ":chart_with_upwards_trend:",
  };
  const options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload)
  };
  UrlFetchApp.fetch(SLACK_WEBHOOK_URL, options);
}
