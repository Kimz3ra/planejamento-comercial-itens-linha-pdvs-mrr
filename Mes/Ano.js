function preencherInicioMesFormatoTexto() {
  const abas = [
    'Line Items - Closer',
    'Line Items - Adição CS',
    'Line Items - Contábil CS',
    'Line Items - Cielo Conciliador',
    'Line Items - Educa'
    // 'Line Items - BizDev Enterprise' removido
    // 'Line Items - Consolidado' removido
  ];

  const meses = ['jan.', 'fev.', 'mar.', 'abr.', 'mai.', 'jun.', 'jul.', 'ago.', 'set.', 'out.', 'nov.', 'dez.'];
  const colMesAnoFechamento = 12; // Coluna M (13ª)

  abas.forEach(nomeAba => {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nomeAba);
    if (!sheet) return;

    const dados = sheet.getDataRange().getValues();
    const cabecalho = dados[0].map(v => v.toString().trim());
    const colDataFechamento = cabecalho.indexOf('Data de Fechamento');
    if (colDataFechamento === -1) return;

    // Define o título fixo da coluna M
    sheet.getRange(1, colMesAnoFechamento + 1).setValue('Mês/Ano (Fechamento)');

    const novaColuna = [];

    for (let i = 1; i < dados.length; i++) {
      const dataStr = dados[i][colDataFechamento];
      const data = new Date(dataStr);
      if (isNaN(data)) {
        novaColuna.push(['']);
      } else {
        const mes = meses[data.getMonth()];
        const ano = data.getFullYear();
        novaColuna.push([`${mes}-${ano}`]);
      }
    }

    sheet.getRange(2, colMesAnoFechamento + 1, novaColuna.length, 1).setValues(novaColuna);
  });
}




