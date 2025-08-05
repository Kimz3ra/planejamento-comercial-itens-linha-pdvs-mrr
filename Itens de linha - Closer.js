
const PIPELINE_NOMES_ATUAL = { 'default': 'Closer' };
const PIPELINE_IDS_ATUAL = ['default'];
const SHEET_NAME_ATUAL = 'Line Items - Closer';
const ETAPA_GANHO_ID_ATUAL = '151188407';
let DEALSTAGE_MAP_ATUAL = null;

function importarLineItems_MesAtual() {
  const hoje = new Date();
  const ano = hoje.getFullYear();
  const mes = hoje.getMonth();
  const inicio = new Date(ano, mes, 1).toISOString();
  const fim = new Date(ano, mes + 1, 0, 23, 59, 59, 999).toISOString();
  importarLineItemsPorPeriodoAtual(inicio, fim, PIPELINE_IDS_ATUAL);
}

function importarLineItemsPorPeriodoAtual(startStr, endStr, pipelineIDs) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME_ATUAL) || ss.insertSheet(SHEET_NAME_ATUAL);

  const header = [
    'ID do Negócio', 'Nome do Negócio', 'Origem', 'Pipeline', 'Etapa do Negócio',
    'Data de Criação', 'Data de Fechamento', 'ID do Item',
    'Produto', 'Classificação do Produto', 'Quantidade', 'Preço Líquido'
  ];

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(header);
  }

  const cabecalho = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const colDataFechamento = cabecalho.indexOf('Data de Fechamento') + 1;
  const mesAtualFormatado = Utilities.formatDate(new Date(startStr), 'GMT-3', 'MM/yyyy');

  const dados = sheet.getDataRange().getValues();
  for (let i = dados.length - 1; i > 0; i--) {
    const data = dados[i][colDataFechamento - 1];
    if (data && Utilities.formatDate(new Date(data), 'GMT-3', 'MM/yyyy') === mesAtualFormatado) {
      sheet.deleteRow(i + 1);
    }
  }

  if (!DEALSTAGE_MAP_ATUAL) {
    DEALSTAGE_MAP_ATUAL = obterMapeamentoEtapasPorPipelineAtual();
  }

  const dealsUrl = 'https://api.hubapi.com/crm/v3/objects/deals/search';
  const headers = { Authorization: 'Bearer ' + HUBSPOT_TOKEN_ATUAL };
  const buffer = [];

  for (const pipelineId of pipelineIDs) {
    let after = null;

    do {
      const payload = {
        filterGroups: [{
          filters: [
            { propertyName: "pipeline", operator: "EQ", value: pipelineId },
            { propertyName: "closedate", operator: "GTE", value: startStr },
            { propertyName: "closedate", operator: "LTE", value: endStr },
            { propertyName: "dealstage", operator: "EQ", value: ETAPA_GANHO_ID_ATUAL }
          ]
        }],
        properties: ["dealname", "createdate", "closedate", "pipeline", "dealstage", "origem"],
        limit: 100,
        sorts: [{ propertyName: "closedate", direction: "DESCENDING" }]
      };

      if (after) payload.after = after;

      const response = UrlFetchApp.fetch(dealsUrl, {
        method: 'post',
        contentType: 'application/json',
        headers: headers,
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      });

      if (response.getResponseCode() !== 200) {
        Logger.log('Erro: ' + response.getContentText());
        break;
      }

      const data = JSON.parse(response.getContentText());
      after = data.paging?.next?.after;

      for (const deal of data.results) {
        const p = deal.properties || {};
        const etapa = formatarNomeEtapaAtual(p.dealstage);
        const lineItems = buscarLineItemsDoNegocioAtual(deal.id);
        if (!lineItems.length) continue;

        for (const itemId of lineItems) {
          const props = buscarDetalhesDoLineItemAtual(itemId);
          if (!props) continue;

          const precoLiquido = parseFloat(props.amount || 0);

          buffer.push([
            deal.id,
            p.dealname || '',
            p.origem || '',
            PIPELINE_NOMES_ATUAL[p.pipeline] || p.pipeline || '',
            etapa,
            formatarDataBrasileiraAtual(p.createdate),
            formatarDataBrasileiraAtual(p.closedate),
            itemId,
            props.name || '',
            props['f360__tipo_de_produto'] || '',
            props.quantity || '',
            formatarPrecoAtual(precoLiquido)
          ]);
        }
      }
    } while (after);
  }

  if (buffer.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, buffer.length, 12).setValues(buffer);
  }
}

function obterMapeamentoEtapasPorPipelineAtual() {
  const url = 'https://api.hubapi.com/crm/v3/pipelines/deals?objectType=deal';
  const response = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { Authorization: 'Bearer ' + HUBSPOT_TOKEN_ATUAL },
    muteHttpExceptions: true
  });

  const content = response.getContentText();
  let data;

  try {
    data = JSON.parse(content);
  } catch (e) {
    Logger.log("Erro ao fazer parsing da resposta:");
    Logger.log(content);
    return {};
  }

  const map = {};
  if (!data.results || !Array.isArray(data.results)) return map;

  data.results.forEach(pipeline => {
    pipeline.stages.forEach(stage => {
      map[stage.id] = stage.label;
    });
  });

  return map;
}

function formatarNomeEtapaAtual(dealstageId) {
  return DEALSTAGE_MAP_ATUAL[dealstageId] || dealstageId;
}

function buscarLineItemsDoNegocioAtual(dealId) {
  const url = `https://api.hubapi.com/crm/v3/objects/deals/${dealId}/associations/line_items`;
  const response = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { Authorization: 'Bearer ' + HUBSPOT_TOKEN_ATUAL }
  });
  const data = JSON.parse(response.getContentText());
  return data.results.map(r => r.id);
}

function buscarDetalhesDoLineItemAtual(id) {
  const url = `https://api.hubapi.com/crm/v3/objects/line_items/${id}?properties=name,quantity,f360__tipo_de_produto,amount`;
  const response = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { Authorization: 'Bearer ' + HUBSPOT_TOKEN_ATUAL }
  });
  const data = JSON.parse(response.getContentText());
  return data.properties;
}

function formatarDataBrasileiraAtual(dataIso) {
  if (!dataIso) return '';
  const d = new Date(dataIso);
  return `${String(d.getDate()).padStart(2, '0')}/${String(d.getMonth() + 1).padStart(2, '0')}/${d.getFullYear()}`;
}

function formatarPrecoAtual(valor) {
  if (!valor || isNaN(valor)) return '';
  return parseFloat(valor).toFixed(2).replace('.', ',');
}








