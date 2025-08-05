function preencherMesAnoFechamento_BizDevEnterprise() {
  const abas = ['Line Items - BizDev Enterprise']; // Consolidado removido
  const colMesAnoFechamento = 13; // Coluna M

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  abas.forEach(nomeAba => {
    const sheet = ss.getSheetByName(nomeAba);
    if (!sheet) {
      Logger.log(`Aba '${nomeAba}' não encontrada.`);
      return;
    }

    const dados = sheet.getDataRange().getValues();
    if (dados.length < 2) return;

    const cabecalho = dados[0].map(c => c.toString().trim());
    const colDataFechamento = cabecalho.findIndex(t => t.toLowerCase() === 'data de fechamento');

    if (colDataFechamento === -1) {
      Logger.log(`Coluna 'Data de Fechamento' não encontrada na aba ${nomeAba}`);
      return;
    }

    // Define título fixo na Coluna M
    sheet.getRange(1, colMesAnoFechamento).setValue('Mês/Ano (Fechamento)');

    const novaColuna = [];

    for (let i = 1; i < dados.length; i++) {
      const dataOriginal = new Date(dados[i][colDataFechamento]);

      if (isNaN(dataOriginal)) {
        novaColuna.push(['']);
      } else {
        const ano = dataOriginal.getFullYear();
        const mes = dataOriginal.getMonth();
        const dataInicioMes = new Date(ano, mes, 1); // sempre dia 1º
        novaColuna.push([dataInicioMes]);
      }
    }

    const rangeDestino = sheet.getRange(2, colMesAnoFechamento, novaColuna.length, 1);
    rangeDestino.setValues(novaColuna);
    rangeDestino.setNumberFormat("dd/mm/yyyy"); // força exibição no padrão 01/01/2025
  });
}



