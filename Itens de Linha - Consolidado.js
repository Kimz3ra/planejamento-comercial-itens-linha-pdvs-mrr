function consolidarLineItemsComMesAno() {
  const abasOrigem = [
    'Line Items - Closer',
    'Line Items - Adição CS',
    'Line Items - Contábil CS',
    'Line Items - Cielo Conciliador',
    'Line Items - Educa',
    'Line Items - BizDev Enterprise'
  ];

  const abaDestinoNome = 'Line Items - Consolidado';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaDestino = ss.getSheetByName(abaDestinoNome) || ss.insertSheet(abaDestinoNome);
  abaDestino.clearContents();

  let linhaDestino = 1;
  let cabecalhoInserido = false;

  abasOrigem.forEach(nomeAba => {
    const aba = ss.getSheetByName(nomeAba);
    if (!aba) return;

    const dados = aba.getDataRange().getValues();
    if (dados.length <= 1) return;

    const colunasParaManter = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11]; // A:L

    if (!cabecalhoInserido) {
      const cabecalhoFiltrado = colunasParaManter.map(i => dados[0][i]);
      cabecalhoFiltrado.push("Mês/Ano");
      abaDestino.getRange(linhaDestino, 1, 1, cabecalhoFiltrado.length).setValues([cabecalhoFiltrado]);
      linhaDestino++;
      cabecalhoInserido = true;
    }

    const dadosFiltrados = dados.slice(1).map(row => {
      const linha = colunasParaManter.map(i => row[i]);

      const data = row[6]; // G = coluna 6 (ajuste se for outra)
      let mesAno = "";

      if (data instanceof Date) {
        const mes = (data.getMonth() + 1).toString().padStart(2, '0');
        const ano = data.getFullYear();
        mesAno = `${mes}/${ano}`;
      }

      linha.push(mesAno);
      return linha;
    });

    abaDestino.getRange(linhaDestino, 1, dadosFiltrados.length, dadosFiltrados[0].length).setValues(dadosFiltrados);
    linhaDestino += dadosFiltrados.length;
  });
}





