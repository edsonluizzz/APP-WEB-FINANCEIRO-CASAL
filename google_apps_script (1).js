// =====================================
// GOOGLE APPS SCRIPT
// Cole este código no Google Sheets
// =====================================

// INSTRUÇÕES:
// 1. Abra sua planilha: https://docs.google.com/spreadsheets/d/1hF93_DhTauLwfspzIKfj30uxadNmUbBIE94GLmK5rtw
// 2. Menu: Extensões > Apps Script
// 3. Cole todo este código
// 4. Clique em "Implantar" > "Nova implantação"
// 5. Tipo: "Aplicativo da Web"
// 6. Executar como: "Eu"
// 7. Quem tem acesso: "Qualquer pessoa"
// 8. Copie a URL gerada
// 9. Cole no dashboard HTML (tem um campo específico)

// =====================================
// FUNÇÃO PRINCIPAL - RETORNA DADOS
// =====================================
function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const params = e && e.parameter ? e.parameter : {};
  
  // Se não especificar aba, retorna resumo geral
  if (!params.sheet) {
    return retornarResumoGeral(ss);
  }
  
  // Se especificar aba, retorna dados daquela aba
  return retornarDadosAba(ss, params.sheet);
}

// =====================================
// RESUMO GERAL (DASHBOARD)
// =====================================
function retornarResumoGeral(ss) {
  try {
    const dashboardSheet = ss.getSheetByName('DASHBOARD');
    
    if (!dashboardSheet) {
      return ContentService.createTextOutput(JSON.stringify({
        success: false,
        error: 'Aba DASHBOARD não encontrada',
        abas_disponiveis: ss.getSheets().map(s => s.getName())
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    // Pegar TODOS os dados do dashboard
    const data = dashboardSheet.getDataRange().getValues();
    
    Logger.log('Total de linhas: ' + data.length);
    Logger.log('Primeira linha: ' + JSON.stringify(data[0]));
    Logger.log('Segunda linha: ' + JSON.stringify(data[1]));
    Logger.log('Terceira linha: ' + JSON.stringify(data[2]));
    
    // Processar meses
    const meses = [];
    const receitas = [];
    const despesas = [];
    const saldos = [];
    
    // Header está na linha 0
    const headers = data[0];
    
    // Encontrar onde começam as datas (pula a primeira coluna que geralmente é texto)
    for (let i = 1; i < headers.length; i++) {
      const celula = headers[i];
      
      // Verifica se é uma data
      if (celula && (celula instanceof Date || !isNaN(Date.parse(celula)))) {
        let dataObj = celula instanceof Date ? celula : new Date(celula);
        
        // Pegar mês e ano
        const mes = dataObj.getMonth(); // 0-11
        const ano = dataObj.getFullYear();
        
        // Nomes dos meses em português
        const mesesPt = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'];
        const mesNome = mesesPt[mes];
        
        meses.push(mesNome);
        
        // Pegar valores das linhas correspondentes
        // Linha 1 = RECEITA
        // Linha 2 = DESPESA
        // Linha 3 = Saldo projetado
        const receitaValor = data[1] && data[1][i] ? parseFloat(data[1][i]) : 0;
        const despesaValor = data[2] && data[2][i] ? parseFloat(data[2][i]) : 0;
        const saldoValor = data[3] && data[3][i] ? parseFloat(data[3][i]) : 0;
        
        receitas.push(receitaValor);
        despesas.push(despesaValor);
        saldos.push(saldoValor);
      }
    }
    
    const resultado = {
      success: true,
      timestamp: new Date().toISOString(),
      debug_info: {
        total_linhas: data.length,
        total_colunas: headers.length,
        primeira_linha: data[0].slice(0, 3),
        meses_encontrados: meses.length
      },
      data: {
        meses: meses,
        receitas: receitas,
        despesas: despesas,
        saldos: saldos
      }
    };
    
    Logger.log('Resultado final: ' + JSON.stringify(resultado));
    
    return ContentService.createTextOutput(JSON.stringify(resultado))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    Logger.log('ERRO: ' + error.toString());
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.toString(),
      stack: error.stack
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// =====================================
// DADOS DE UMA ABA ESPECÍFICA
// =====================================
function retornarDadosAba(ss, nomeAba) {
  try {
    const sheet = ss.getSheetByName(nomeAba);
    
    if (!sheet) {
      return ContentService.createTextOutput(JSON.stringify({
        success: false,
        error: `Aba ${nomeAba} não encontrada`,
        abas_disponiveis: ss.getSheets().map(s => s.getName())
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    const data = sheet.getDataRange().getValues();
    
    if (data.length === 0) {
      return ContentService.createTextOutput(JSON.stringify({
        success: false,
        error: 'Aba vazia'
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    const headers = data[0];
    
    Logger.log('Aba: ' + nomeAba);
    Logger.log('Headers: ' + JSON.stringify(headers));
    Logger.log('Total registros: ' + (data.length - 1));
    
    // Converter para array de objetos
    const registros = [];
    let totalReceita = 0;
    let totalDespesa = 0;
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      // Pular linhas vazias
      if (!row[0]) continue;
      
      const registro = {};
      
      for (let j = 0; j < headers.length; j++) {
        const header = headers[j];
        const valor = row[j];
        
        // Tratar valores especiais
        if (header && header.toString().trim() !== '') {
          registro[header] = valor;
        }
      }
      
      // Só adiciona se tiver descrição
      if (registro['Descrição']) {
        registros.push(registro);
        
        // Calcular totais
        const tipo = registro['Tipo'];
        const valor = parseFloat(registro['Valor']) || 0;
        
        if (tipo === 'Receita') {
          totalReceita += valor;
        } else if (tipo === 'Despesa') {
          totalDespesa += valor;
        }
      }
    }
    
    const resultado = {
      success: true,
      timestamp: new Date().toISOString(),
      aba: nomeAba,
      total_registros: registros.length,
      resumo: {
        receita: totalReceita,
        despesa: totalDespesa,
        saldo: totalReceita - totalDespesa
      },
      data: registros
    };
    
    Logger.log('Resumo: ' + JSON.stringify(resultado.resumo));
    
    return ContentService.createTextOutput(JSON.stringify(resultado))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    Logger.log('ERRO na aba ' + nomeAba + ': ' + error.toString());
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.toString(),
      stack: error.stack
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// =====================================
// FUNÇÃO POST - ATUALIZAR STATUS
// =====================================
function doPost(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const params = JSON.parse(e.postData.contents);
    
    const aba = params.aba;
    const descricao = params.descricao;
    const novoStatus = params.status;
    
    const sheet = ss.getSheetByName(aba);
    
    if (!sheet) {
      return ContentService.createTextOutput(JSON.stringify({
        success: false,
        error: 'Aba não encontrada'
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    // Procurar a linha com a descrição
    const data = sheet.getDataRange().getValues();
    const statusCol = data[0].indexOf('Status') + 1; // Coluna Status
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === descricao) { // Coluna Descrição
        sheet.getRange(i + 1, statusCol).setValue(novoStatus);
        
        return ContentService.createTextOutput(JSON.stringify({
          success: true,
          message: 'Status atualizado com sucesso'
        })).setMimeType(ContentService.MimeType.JSON);
      }
    }
    
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: 'Registro não encontrado'
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// =====================================
// TESTE - Executar para verificar
// =====================================
function testarAPI() {
  Logger.log('=== INICIANDO TESTE DA API ===');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log('Planilha ativa: ' + ss.getName());
  
  // Listar todas as abas
  const abas = ss.getSheets().map(s => s.getName());
  Logger.log('Abas disponíveis: ' + abas.join(', '));
  
  // Testar função principal
  const resultado = retornarResumoGeral(ss);
  Logger.log('=== RESULTADO ===');
  Logger.log(resultado.getContent());
  
  return resultado.getContent();
}

// =====================================
// TESTE RÁPIDO - Ver estrutura DASHBOARD
// =====================================
function verEstruturaDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboard = ss.getSheetByName('DASHBOARD');
  
  if (!dashboard) {
    Logger.log('ERRO: Aba DASHBOARD não encontrada!');
    return;
  }
  
  const data = dashboard.getDataRange().getValues();
  
  Logger.log('=== ESTRUTURA DA ABA DASHBOARD ===');
  Logger.log('Total de linhas: ' + data.length);
  Logger.log('Total de colunas: ' + data[0].length);
  Logger.log('');
  
  // Mostrar primeiras 4 linhas e 5 colunas
  for (let i = 0; i < Math.min(4, data.length); i++) {
    Logger.log('LINHA ' + i + ': ' + JSON.stringify(data[i].slice(0, 5)));
  }
}

// =====================================
// LISTAR TODAS AS ABAS
// =====================================
function listarAbas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  
  Logger.log('=== ABAS DISPONÍVEIS NA PLANILHA ===');
  
  sheets.forEach((sheet, index) => {
    Logger.log((index + 1) + '. ' + sheet.getName());
  });
  
  return sheets.map(s => s.getName());
}

// =====================================
// TESTAR ACESSO A UMA ABA ESPECÍFICA
// =====================================
function testarAbaMarco() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  Logger.log('=== TESTANDO ACESSO À ABA DE MARÇO ===');
  
  // Tentar diferentes variações
  const variacoes = ['0326', '03/26', '0326 ', 'Março', 'Marco'];
  
  variacoes.forEach(nome => {
    const sheet = ss.getSheetByName(nome);
    if (sheet) {
      Logger.log('✅ ENCONTRADA: "' + nome + '"');
      Logger.log('   Total de linhas: ' + sheet.getLastRow());
    } else {
      Logger.log('❌ NÃO ENCONTRADA: "' + nome + '"');
    }
  });
}
