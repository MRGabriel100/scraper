const XLSX = require('xlsx');

async function getMetaData(indicador) {
  const url = `https://www.cidadessustentaveis.org.br/api/indicador/preenchidos/grafico/indicadores?indicador=${indicador}&cidades=3981&formulaidx=0`;
  
  try {
    const response = await fetch(url);
    if (!response.ok) throw new Error(`Erro HTTP ${response.status}`);
    const data = await response.json();
    return data.meta || '';
  } catch (error) {
    console.error(`Erro ao buscar meta para indicador ${indicador}:`, error.message);
    return '';
  }
}

async function getData(ods, startYear, endYear) {
  const url = `https://www.cidadessustentaveis.org.br/api/painel/indicadores?idOds=${ods}&idCidade=3981&anoInicial=${startYear}&anoFinal=${endYear}&indicadorPcs=true&indicadorComplementar=false&indicadorIndice=false`;
  
  try {
    const response = await fetch(url);
    if (!response.ok) throw new Error(`Erro HTTP ${response.status}`);
    return await response.json();
  } catch (error) {
    console.error(`Erro no ODS ${ods} (${startYear}-${endYear}):`, error.message);
    return null;
  }
}

async function processODS(ods) {
  // Busca dados para ambos os períodos
  const [data2017_2020, data2021_2024] = await Promise.all([
    getData(ods, 2017, 2020),
    getData(ods, 2021, 2024)
  ]);

  const combinedData = [];
  const processedIndicators = new Set(); 

  // Função para processar e combinar os dados
  const processData = async (data, startYear, endYear) => {
    if (!data) return;
    
    for (const item of data) {
      const indicatorId = item[0];
      if (processedIndicators.has(indicatorId)) continue;
      processedIndicators.add(indicatorId);

      const metaText = await getMetaData(indicatorId);
      
      let metaNum = '';
      let metaDesc = '';
      
      if (metaText) {
        const metaParts = metaText.split(' : ');
        if (metaParts.length > 1) {
          const metaMatch = metaParts[0].match(/\d+\.\d+/);
          metaNum = metaMatch ? metaMatch[0] : '';
          metaDesc = metaParts.slice(1).join(' : ').trim();
        } else {
          metaDesc = metaText;
        }
      } else {
        metaDesc = '';
      }

      const newItem = {
        'ODS nº': ods,
        'Meta Nº': metaNum,
        'Meta Descrição': metaDesc,
        'Discriminação': item[1],
        'Fonte': 'Cidades Sustentáveis'
      };
      
      for (let year = 2017; year <= 2024; year++) {
        newItem[year] = (year >= startYear && year <= endYear) 
          ? item[2 + (year - startYear)] 
          : '';
      }
      
      combinedData.push(newItem);
    }
  };

  await processData(data2017_2020, 2017, 2020);
  await processData(data2021_2024, 2021, 2024);

  return combinedData;
}

async function exportToExcel() {
  let allData = [];
  
  // Processa cada ODS sequencialmente (1 a 17)
  for (let ods = 1; ods < 18; ods++) {
    console.log(`Processando ODS ${ods}...`);
    const odsData = await processODS(ods);
    allData = [...allData, ...odsData];
  }

  // Ordena os dados por ODS e Meta
  allData.sort((a, b) => {
    if (a['ODS nº'] !== b['ODS nº']) return a['ODS nº'] - b['ODS nº'];
    return parseFloat(a['Meta Nº']) - parseFloat(b['Meta Nº']);
  });

  // Define a ordem exata das colunas
  const columnOrder = [
    'ODS nº',
    'Meta Nº',
    'Meta Descrição',
    'Discriminação',
    '2017', '2018', '2019', '2020', '2021', '2022', '2023', '2024',
    'Fonte'
  ];

  // Cria a planilha com a ordem de colunas especificada
  const ws = XLSX.utils.json_to_sheet(allData, {
    header: columnOrder
  });
  
  ws['!cols'] = [
    { width: 8 },   // ODS nº
    { width: 8 },   // Meta Nº
    { width: 60 },  // Meta Descrição
    { width: 40 },  // Discriminação
    { width: 8 },   // 2017
    { width: 8 },   // 2018
    { width: 8 },   // 2019
    { width: 8 },   // 2020
    { width: 8 },   // 2021
    { width: 8 },   // 2022
    { width: 8 },   // 2023
    { width: 8 },   // 2024
    { width: 20 }   // Fonte
  ];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Indicadores");
  XLSX.writeFile(wb, "indicadores_organizados.xlsx");
  console.log('Planilha gerada com sucesso!');
}

// Inicia o processo
exportToExcel().catch(error => {
  console.error('Erro no processo principal:', error);
  alert('Ocorreu um erro ao gerar a planilha. Verifique o console para detalhes.');
});