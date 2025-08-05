const PIPELINE_NOME_FUNIL = 'Closer';
const PIPELINE_ID_FUNIL = 'default';
const SHEET_NAME_FUNIL = 'Classificacao funil Closer';
let DEALSTAGE_MAP_FUNIL = null;
const HUBSPOT_TEAM_MAP = {
  '42999253': 'BizDev',
  '42999254': 'Pré vendas',
  '42999259': 'Tecnologia/Produto',
  '42999263': 'Educa',
  '42999264': 'Vendas',
  '42999270': 'Financeiro',
  '42999280': 'Diretoria',
  '42999284': 'Marketing',
  '42999301': 'Onboarding',
  '44537670': 'Tropical Hub',
  '46397730': 'Fornecedores',
  '46853477': 'Customer Success',
  '47413098': 'Operações'
};

function importarClassificacaoFunilCloser() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME_FUNIL) || ss.insertSheet(SHEET_NAME_FUNIL);
  sheet.clearContents();
  sheet.appendRow([
    'ID do Negócio', 'Nome do Negócio', 'Origem', 'Pipeline',
    'Etapa do Negócio', 'Data de Criação', 'Data de Fechamento',
    'Classificação Funil', 'Número de CNPJs', 'HubSpot Team'
  ]);

  if (!DEALSTAGE_MAP_FUNIL) DEALSTAGE_MAP_FUNIL = obterMapeamentoEtapasPorPipeline_FUNIL();

  const dealsUrl = 'https://api.hubapi.com/crm/v3/objects/deals/search';
  const headers = { Authorization: 'Bearer ' + HUBSPOT_TOKEN_ATUAL };
  let after = null, buffer = [];

  do {
    const payload = {
      filterGroups: [{
        filters: [
          { propertyName: "pipeline", operator: "EQ", value: PIPELINE_ID_FUNIL },
          { propertyName: "createdate", operator: "GTE", value: "2024-01-01T00:00:00.000Z" }
        ]
      }],
      properties: [
        "dealname", "createdate", "closedate", "pipeline",
        "dealstage", "origem", "classificacao_funil", "numero_de_cnpjs", "hubspot_team_id"
      ],
      sorts: [{ propertyName: "createdate", direction: "DESCENDING" }],
      limit: 100
    };
    if (after) payload.after = after;

    const resp = UrlFetchApp.fetch(dealsUrl, {
      method: 'post',
      contentType: 'application/json',
      headers,
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    const data = JSON.parse(resp.getContentText());
    after = data.paging?.next?.after;

    for (const deal of data.results) {
      const p = deal.properties;
      const teamId = p.hubspot_team_id;
      const teamName = HUBSPOT_TEAM_MAP[teamId] || '';

      buffer.push([
        deal.id,
        p.dealname || '',
        p.origem || '',
        PIPELINE_NOME_FUNIL,
        formatarNomeEtapa_FUNIL(p.dealstage),
        formatarDataBrasileira_FUNIL(p.createdate),
        formatarDataBrasileira_FUNIL(p.closedate),
        p.classificacao_funil || '',
        p.numero_de_cnpjs || '',
        teamName
      ]);
    }
  } while (after);

  if (buffer.length > 0) {
    sheet.getRange(2, 1, buffer.length, buffer[0].length).setValues(buffer);
  }
}

function obterMapeamentoEtapasPorPipeline_FUNIL() {
  const url = 'https://api.hubapi.com/crm/v3/pipelines/deals?objectType=deal';
  const resp = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { Authorization: 'Bearer ' + HUBSPOT_TOKEN_ATUAL }
  });
  const data = JSON.parse(resp.getContentText());
  return Object.fromEntries(
    data.results.flatMap(p => p.stages.map(s => [s.id, s.label]))
  );
}

function formatarNomeEtapa_FUNIL(id) {
  return DEALSTAGE_MAP_FUNIL[id] || id;
}

function formatarDataBrasileira_FUNIL(dataIso) {
  if (!dataIso) return '';
  const d = new Date(dataIso);
  const dd = String(d.getDate()).padStart(2, '0');
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  const yyyy = d.getFullYear();
  return `${dd}/${mm}/${yyyy}`;
}




