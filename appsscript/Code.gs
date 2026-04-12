/**
 * EstoqueFarmacia – Google Apps Script Backend
 *
 * Setup:
 *  1. Open your Google Spreadsheet.
 *  2. Go to Extensions → Apps Script.
 *  3. Paste this entire file and save.
 *  4. Deploy as a Web App (Execute as: Me, Who has access: Anyone).
 *  5. Copy the deployment URL and paste it into docs/js/config.js.
 *
 * Spreadsheet structure expected:
 *   Sheet "Estoque"  → Column A: Medicine name | Column B: Current physical stock (number)
 *   Sheet "Vendas"   → Column A: Medicine name | Column B: Qty sold in the last N days (number)
 *                      Cell D1 of "Vendas": the number of days the sales data covers (e.g. 30)
 */

// ─── Sheet name constants ────────────────────────────────────────────────────
var SHEET_ESTOQUE = 'Estoque';
var SHEET_VENDAS  = 'Vendas';

// ─── Entry point ─────────────────────────────────────────────────────────────

/**
 * Handles all GET requests from the web app.
 * Supported actions: calcular, ping
 */
function doGet(e) {
  var params = e.parameter;
  var action = params.action || 'ping';

  try {
    if (action === 'calcular') {
      var dias = parseInt(params.dias, 10);
      if (isNaN(dias) || dias <= 0) {
        return jsonError('Parâmetro "dias" inválido.');
      }
      var resultado = calcularNecessidade(dias);
      return jsonOk(resultado);
    }

    if (action === 'relatorio') {
      var top = parseInt(params.top, 10);
      if (isNaN(top) || top < 0) top = 0; // 0 = todos
      var relatorio = gerarRelatorioMaisVendidos(top);
      return jsonOk(relatorio);
    }

    if (action === 'ping') {
      return jsonOk({ status: 'ok', timestamp: new Date().toISOString() });
    }

    return jsonError('Ação desconhecida: ' + action);
  } catch (err) {
    return jsonError(err.message);
  }
}

/**
 * Handles all POST requests from the web app.
 * Supported actions: importarEstoque, importarVendas
 */
function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    var action  = payload.action;

    if (action === 'importarEstoque') {
      importarDados(SHEET_ESTOQUE, payload.dados);
      return jsonOk({ importados: payload.dados.length });
    }

    if (action === 'importarVendas') {
      importarDados(SHEET_VENDAS, payload.dados);
      // Optionally store the period in D1
      if (payload.periodo) {
        var ss    = SpreadsheetApp.getActiveSpreadsheet();
        var sheet = ss.getSheetByName(SHEET_VENDAS);
        sheet.getRange('D1').setValue(payload.periodo);
      }
      return jsonOk({ importados: payload.dados.length });
    }

    return jsonError('Ação POST desconhecida: ' + action);
  } catch (err) {
    return jsonError(err.message);
  }
}

// ─── Core business logic ──────────────────────────────────────────────────────

/**
 * Reads both sheets, calculates moving-average purchase needs and returns
 * only items where the current stock does not cover the requested period.
 *
 * Formula per item:
 *   mediaDiaria  = vendasTotais / diasVendas
 *   necessidade  = (mediaDiaria × diasCobertura) - estoqueAtual
 *   → Only included if necessidade > 0
 *
 * @param {number} diasCobertura  Number of days to cover (7 / 14 / 28 …)
 * @returns {Object}
 */
function calcularNecessidade(diasCobertura) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── Load Estoque sheet ──────────────────────────────────────────────────
  var sheetEstoque = ss.getSheetByName(SHEET_ESTOQUE);
  if (!sheetEstoque) throw new Error('Aba "' + SHEET_ESTOQUE + '" não encontrada.');
  var dadosEstoque = sheetEstoque.getDataRange().getValues();

  // Build a Map: medicamento → estoqueAtual
  var mapaEstoque = {};
  for (var i = 0; i < dadosEstoque.length; i++) {
    var nome  = String(dadosEstoque[i][0]).trim();
    var qtd   = parseFloat(dadosEstoque[i][1]) || 0;
    if (nome && nome.toLowerCase() !== 'nome' && nome.toLowerCase() !== 'medicamento') {
      mapaEstoque[nome.toUpperCase()] = qtd;
    }
  }

  // ── Load Vendas sheet ───────────────────────────────────────────────────
  var sheetVendas = ss.getSheetByName(SHEET_VENDAS);
  if (!sheetVendas) throw new Error('Aba "' + SHEET_VENDAS + '" não encontrada.');
  var dadosVendas = sheetVendas.getDataRange().getValues();

  // D1 stores how many days the sales data covers; default 30
  var diasVendas = parseFloat(sheetVendas.getRange('D1').getValue()) || 30;

  // Build a Map: medicamento → vendasTotais
  var mapaVendas = {};
  for (var j = 0; j < dadosVendas.length; j++) {
    var nomeV = String(dadosVendas[j][0]).trim();
    var qtdV  = parseFloat(dadosVendas[j][1]) || 0;
    if (nomeV && nomeV.toLowerCase() !== 'nome' && nomeV.toLowerCase() !== 'medicamento') {
      mapaVendas[nomeV.toUpperCase()] = qtdV;
    }
  }

  // ── Cross-reference and calculate ──────────────────────────────────────
  var lista = [];
  var todosNomes = Object.keys(mapaVendas);

  for (var k = 0; k < todosNomes.length; k++) {
    var chave        = todosNomes[k];
    var vendasTotais = mapaVendas[chave]  || 0;
    var estoqueAtual = mapaEstoque[chave] || 0;

    var mediaDiaria  = vendasTotais / diasVendas;
    var projecao     = mediaDiaria * diasCobertura;
    var necessidade  = projecao - estoqueAtual;

    if (necessidade > 0) {
      lista.push({
        medicamento:  chave,
        estoqueAtual: estoqueAtual,
        mediaDiaria:  Math.round(mediaDiaria * 100) / 100,
        projecao:     Math.round(projecao * 10) / 10,
        comprar:      Math.ceil(necessidade)
      });
    }
  }

  // Sort by highest need first
  lista.sort(function(a, b) { return b.comprar - a.comprar; });

  return {
    diasCobertura: diasCobertura,
    diasVendas:    diasVendas,
    totalItens:    lista.length,
    geradoEm:      new Date().toISOString(),
    itens:         lista
  };
}

// ─── Import helpers ───────────────────────────────────────────────────────────

/**
 * Replaces all data in the given sheet with the provided rows.
 * @param {string}   sheetName  Target sheet name
 * @param {Array}    dados      Array of [name, qty] pairs (already parsed)
 */
function importarDados(sheetName, dados) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  } else {
    sheet.clearContents();
  }

  if (!dados || dados.length === 0) return;

  // Write header
  sheet.getRange(1, 1, 1, 2).setValues([['Medicamento', 'Quantidade']]);

  // Write data in a single batch call for performance (handles 900 rows easily)
  var rows = dados.map(function(row) {
    return [String(row[0]).trim(), parseFloat(row[1]) || 0];
  });

  sheet.getRange(2, 1, rows.length, 2).setValues(rows);
}

// ─── Report: most-sold items ──────────────────────────────────────────────────

/**
 * Reads the Vendas sheet and returns items sorted by mediaDiaria descending.
 * @param {number} top  Number of items to return; 0 = all items.
 * @returns {Object}
 */
function gerarRelatorioMaisVendidos(top) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var sheetVendas = ss.getSheetByName(SHEET_VENDAS);
  if (!sheetVendas) throw new Error('Aba "' + SHEET_VENDAS + '" não encontrada.');
  var dadosVendas = sheetVendas.getDataRange().getValues();

  var diasVendas = parseFloat(sheetVendas.getRange('D1').getValue()) || 30;
  if (diasVendas <= 0) diasVendas = 30;

  var lista = [];
  for (var j = 0; j < dadosVendas.length; j++) {
    var nome = String(dadosVendas[j][0]).trim();
    var qtd  = parseFloat(dadosVendas[j][1]) || 0;
    if (!nome || nome.toLowerCase() === 'nome' || nome.toLowerCase() === 'medicamento') continue;
    var media = Math.round((qtd / diasVendas) * 100) / 100;
    lista.push({ medicamento: nome.toUpperCase(), totalVendido: qtd, mediaDiaria: media });
  }

  lista.sort(function(a, b) { return b.mediaDiaria - a.mediaDiaria; });

  var itens = (top > 0) ? lista.slice(0, top) : lista;

  return {
    diasVendas: diasVendas,
    totalItens: itens.length,
    geradoEm:   new Date().toISOString(),
    itens:      itens
  };
}

// ─── Response helpers ─────────────────────────────────────────────────────────

function jsonOk(data) {
  var output = ContentService
    .createTextOutput(JSON.stringify({ ok: true, data: data }))
    .setMimeType(ContentService.MimeType.JSON);
  return output;
}

function jsonError(message) {
  var output = ContentService
    .createTextOutput(JSON.stringify({ ok: false, error: message }))
    .setMimeType(ContentService.MimeType.JSON);
  return output;
}
