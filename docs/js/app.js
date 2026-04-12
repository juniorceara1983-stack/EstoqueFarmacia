/**
 * docs/js/app.js
 * EstoqueFarmacia – Frontend logic
 *
 * Flow:
 *  1. User uploads CSV for "Estoque" → parsed and sent to Apps Script in batches
 *  2. User uploads CSV for "Vendas"  → same
 *  3. User picks coverage period (7 / 14 / 28 days)
 *  4. User clicks "Calcular" → GET request to Apps Script → renders table
 *  5. Export to PDF (jsPDF) or share via WhatsApp
 */

/* ── State ──────────────────────────────────────────────────────────────── */
const state = {
  periodoSelecionado: 14,
  topSelecionado: 10,
  estoqueCarregado: false,
  vendasCarregadas: false,
  resultado: null,
  relatorio: null,
  validadeDias: 60,
  validadeItens: null,
  validadeResultado: null,
  // Local data for in-browser calculation (avoids backend round-trip and
  // order-of-upload issues with Google Sheets).
  estoqueLocal: null,        // [[nome, qtd], …]
  vendasLocal: null,         // [[nome, qtd], …]
  diasVendasLocal: 7,
  estoqueVemSugestao: false, // true when stock came from a sugestão/rede file
};

/* ── DOM refs ────────────────────────────────────────────────────────────── */
const $ = (id) => document.getElementById(id);

/* ── Bootstrap ───────────────────────────────────────────────────────────── */
document.addEventListener('DOMContentLoaded', () => {
  setupPeriodButtons();
  setupTopButtons();
  setupValidadePeriodButtons();
  setupDropzone('estoqueDropzone', 'estoqueCsvInput', 'estoqueProgress',
                'estoqueStatus', 'Estoque', false);
  setupDropzone('vendasDropzone', 'vendasCsvInput', 'vendasProgress',
                'vendasStatus', 'Vendas', true);
  setupValidadeDropzone();
  $('btnCalcular').addEventListener('click', calcular);
  $('btnPdf').addEventListener('click', exportarPdf);
  $('btnWhatsapp').addEventListener('click', compartilharWhatsapp);
  $('btnRelatorio').addEventListener('click', gerarRelatorio);
  $('btnRelatorioPdf').addEventListener('click', exportarRelatorioPdf);
  $('btnValidadePdf').addEventListener('click', exportarValidadePdf);
  updateCalcularBtn();
});

/* ── Period selector ────────────────────────────────────────────────────── */
function setupPeriodButtons() {
  // Only select the main coverage-period buttons (not validade period buttons)
  document.querySelectorAll('.period-btn[data-dias]:not(.vld-btn)').forEach((btn) => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.period-btn[data-dias]:not(.vld-btn)').forEach((b) => b.classList.remove('active'));
      btn.classList.add('active');
      state.periodoSelecionado = parseInt(btn.dataset.dias, 10);
    });
  });
  // Default active
  document.querySelector('.period-btn[data-dias="14"]:not(.vld-btn)').classList.add('active');
}

/* ── Top-N selector for report ──────────────────────────────────────────── */
function setupTopButtons() {
  document.querySelectorAll('.top-btn').forEach((btn) => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.top-btn').forEach((b) => b.classList.remove('active'));
      btn.classList.add('active');
      state.topSelecionado = parseInt(btn.dataset.top, 10);
    });
  });
}

/* ── Dropzone helper ─────────────────────────────────────────────────────── */
/**
 * Wires up drag-and-drop and the file input change event for a dropzone.
 * The click-to-open behaviour is handled natively by the <label> element in
 * the HTML (no JS needed), which makes it reliably work on mobile browsers too.
 */
function setupDropzone(zoneId, inputId, progressId, statusId, tipo, isVendas) {
  const zone  = $(zoneId);
  const input = $(inputId);

  zone.addEventListener('dragover', (e) => {
    e.preventDefault();
    zone.classList.add('over');
  });
  zone.addEventListener('dragleave', () => zone.classList.remove('over'));
  zone.addEventListener('drop', (e) => {
    e.preventDefault();
    zone.classList.remove('over');
    const file = e.dataTransfer.files[0];
    if (file) handleCsvFile(file, progressId, statusId, tipo, isVendas);
  });

  input.addEventListener('change', () => {
    if (input.files[0]) handleCsvFile(input.files[0], progressId, statusId, tipo, isVendas);
    input.value = '';
  });
}

/* ── CSV parsing and upload ──────────────────────────────────────────────── */
function handleCsvFile(file, progressId, statusId, tipo, isVendas) {
  const isXls = file.name.match(/\.xlsx?$/i) ||
                /application\/vnd\.(ms-excel|openxmlformats-officedocument\.spreadsheetml\.sheet)/.test(file.type);
  const isCsv = file.name.match(/\.(csv|txt)$/i) ||
                /^text\/(csv|plain)$/.test(file.type);

  if (!isXls && !isCsv) {
    showToast('Arquivo inválido. Use CSV, XLS ou XLSX.', 'error');
    return;
  }

  const status = $(statusId);
  status.textContent = 'Lendo arquivo…';

  if (isXls) {
    if (typeof XLSX === 'undefined') {
      showToast('Erro: biblioteca XLS não carregou. Verifique sua conexão e recarregue a página.', 'error');
      status.textContent = '❌ Biblioteca XLS indisponível.';
      return;
    }
    const reader = new FileReader();
    reader.onload = async (e) => {
      try {
        const data   = new Uint8Array(e.target.result);
        const wb     = XLSX.read(data, { type: 'array' });
        const ws     = wb.Sheets[wb.SheetNames[0]];
        const parsed = parseXlsRows(ws);
        if (isParsedEmpty(parsed)) {
          status.textContent = '⚠️ Arquivo vazio ou sem dados.';
          showToast('O arquivo não contém dados válidos.', 'error');
          return;
        }
        await processRows(parsed, progressId, statusId, tipo, isVendas, status);
      } catch (err) {
        status.textContent = `❌ Erro: ${err.message}`;
        showToast(`Erro ao ler arquivo: ${err.message}`, 'error');
      }
    };
    reader.readAsArrayBuffer(file);
  } else {
    const reader = new FileReader();
    reader.onload = async (e) => {
      const parsed = parseCsv(e.target.result);
      if (isParsedEmpty(parsed)) {
        status.textContent = '⚠️ Arquivo vazio ou sem dados.';
        showToast('O arquivo não contém dados válidos.', 'error');
        return;
      }
      await processRows(parsed, progressId, statusId, tipo, isVendas, status);
    };
    reader.readAsText(file, 'UTF-8');
  }
}

async function processRows(parsed, progressId, statusId, tipo, isVendas, status) {
  // ── Rede format: one file populates both Estoque and Vendas ──────────────
  if (parsed.format === 'rede') {
    const total = parsed.estoqueRows.length;
    if (total === 0) {
      status.textContent = '⚠️ Nenhum item encontrado no arquivo.';
      showToast('O arquivo não contém dados válidos.', 'error');
      return;
    }

    // Store locally – rede file is authoritative for both stock and sales
    state.estoqueLocal       = parsed.estoqueRows.slice();
    state.vendasLocal        = parsed.vendasRows.slice();
    state.diasVendasLocal    = 30; // Vend. column represents monthly sales
    state.estoqueVemSugestao = true;
    state.estoqueCarregado   = true;
    state.vendasCarregadas   = true;
    updateCalcularBtn();

    const otherProgressId = progressId === 'estoqueProgress' ? 'vendasProgress' : 'estoqueProgress';
    status.textContent = `Formato Rede detectado – ${total} itens. Importando estoque e vendas…`;
    showProgress(progressId);
    showProgress(otherProgressId);

    try {
      await uploadEmLotes(parsed.estoqueRows, 'Estoque', progressId, null);
      $('estoqueStatus').textContent = `✅ ${total} itens de estoque importados!`;

      // Vendas period: 30 days (Vend. column represents monthly sales)
      await uploadEmLotes(parsed.vendasRows, 'Vendas', otherProgressId, 30);
      $('vendasStatus').textContent = `✅ ${total} itens de vendas importados (período: 30 dias)!`;

      showToast(`Arquivo de Rede importado: ${total} itens (estoque + vendas).`, 'success');
    } catch (err) {
      status.textContent = `❌ Erro ao enviar ao servidor: ${err.message}`;
      showToast(`Arquivo carregado localmente. Erro ao sincronizar com servidor: ${err.message}`, 'error');
    }
    return;
  }

  // ── Sugestão de Compras format: one file populates both Estoque and Vendas ─
  if (parsed.format === 'sugestao') {
    const total = parsed.estoqueRows.length;
    if (total === 0) {
      status.textContent = '⚠️ Nenhum item encontrado no arquivo.';
      showToast('O arquivo não contém dados válidos.', 'error');
      return;
    }

    // Store locally – sugestão is authoritative for both stock and sales.
    // The "Estoq" column in the sugestão is the definitive current-stock source.
    state.estoqueLocal       = parsed.estoqueRows.slice();
    state.vendasLocal        = parsed.vendasRows.slice();
    state.diasVendasLocal    = parsed.periodoDias;
    state.estoqueVemSugestao = true;
    state.estoqueCarregado   = true;
    state.vendasCarregadas   = true;
    updateCalcularBtn();

    const otherProgressId = progressId === 'estoqueProgress' ? 'vendasProgress' : 'estoqueProgress';
    status.textContent = `Sugestão de Compras detectada – ${total} itens (período: ${parsed.periodoDias} dias). Importando…`;
    showProgress(progressId);
    showProgress(otherProgressId);

    try {
      await uploadEmLotes(parsed.estoqueRows, 'Estoque', progressId, null);
      $('estoqueStatus').textContent = `✅ ${total} itens de estoque importados!`;

      await uploadEmLotes(parsed.vendasRows, 'Vendas', otherProgressId, parsed.periodoDias);
      $('vendasStatus').textContent = `✅ ${total} itens de vendas importados (período: ${parsed.periodoDias} dias)!`;

      showToast(`Sugestão de Compras importada: ${total} itens (${parsed.periodoDias} dias).`, 'success');
    } catch (err) {
      status.textContent = `❌ Erro ao enviar ao servidor: ${err.message}`;
      showToast(`Arquivo carregado localmente. Erro ao sincronizar com servidor: ${err.message}`, 'error');
    }
    return;
  }

  // ── NF or simple format: uses .rows ──────────────────────────────────────
  const rows = parsed.rows || [];
  if (rows.length === 0) {
    status.textContent = '⚠️ Arquivo vazio ou sem dados.';
    showToast('O arquivo não contém dados válidos.', 'error');
    return;
  }

  // NF files always go to Estoque regardless of which dropzone was used
  const tipoEfetivo      = parsed.format === 'nf' ? 'Estoque' : tipo;
  const isVendasEfetivo  = tipoEfetivo !== 'Estoque';

  if (parsed.format === 'nf' && isVendas) {
    showToast('Nota Fiscal detectada: importada como Estoque.', 'success');
  }

  // Store locally.
  // NF files carry delivery quantities, not the live stock balance.
  // Only use them for stock if no authoritative sugestão data is present yet.
  if (!isVendasEfetivo) {
    if (!state.estoqueVemSugestao) {
      state.estoqueLocal = rows.slice();
    }
    state.estoqueCarregado = true;
  } else {
    state.vendasLocal     = rows.slice();
    state.diasVendasLocal = state.periodoSelecionado;
    state.vendasCarregadas = true;
  }
  updateCalcularBtn();

  status.textContent = `${rows.length} linhas encontradas. Enviando…`;
  showProgress(progressId);

  try {
    await uploadEmLotes(rows, tipoEfetivo, progressId, isVendasEfetivo ? state.periodoSelecionado : null);
    status.textContent = `✅ ${rows.length} itens importados com sucesso!`;
    showToast(`${tipoEfetivo} importado: ${rows.length} itens.`, 'success');
  } catch (err) {
    status.textContent = `❌ Erro ao enviar ao servidor: ${err.message}`;
    showToast(`Arquivo carregado localmente. Erro ao sincronizar com servidor: ${err.message}`, 'error');
  }
}

/** Returns true when a parsed result contains no usable rows. */
function isParsedEmpty(parsed) {
  if (parsed.format === 'vazio')    return true;
  if (parsed.format === 'rede')     return parsed.estoqueRows.length === 0;
  if (parsed.format === 'sugestao') return parsed.estoqueRows.length === 0;
  return (parsed.rows || []).length === 0;
}

/**
 * Parses a SheetJS worksheet, auto-detecting the file format.
 * See detectAndParseRows() for supported formats.
 */
function parseXlsRows(ws) {
  const ref = ws['!ref'];
  if (!ref) return { format: 'vazio', rows: [] };

  const range   = XLSX.utils.decode_range(ref);
  const allRows = [];
  for (let r = range.s.r; r <= range.e.r; r++) {
    const row = [];
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      row.push(cell ? cell.v : '');
    }
    allRows.push(row);
  }
  return detectAndParseRows(allRows);
}

/**
 * Parses a CSV text, auto-detecting the file format (NF, Rede, or simple 2-column).
 * See detectAndParseRows() for supported formats.
 */
function parseCsv(text) {
  const lines   = text.split(/\r?\n/).map((l) => l.trim()).filter(Boolean);
  const sep     = lines[0]?.includes(';') ? ';' : ',';
  const allRows = lines.map((line) => splitCsvLine(line, sep));
  return detectAndParseRows(allRows);
}

/**
 * Splits a single CSV line, respecting double-quoted fields.
 */
function splitCsvLine(line, sep) {
  const result = [];
  let cur = '';
  let inQ = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (ch === '"') {
      // Doubled quotes inside a quoted field represent a literal quote character
      if (inQ && line[i + 1] === '"') { cur += '"'; i++; }
      else { inQ = !inQ; }
    } else if (ch === sep && !inQ) {
      result.push(cur); cur = '';
    } else {
      cur += ch;
    }
  }
  result.push(cur);
  return result;
}

/**
 * Inspects the first rows to detect the file format, then parses accordingly.
 *
 * Supported formats:
 *   'sugestao'– DROGAMAIS "Sugestão de Compras" report.
 *              → returns { format: 'sugestao', estoqueRows, vendasRows, periodoDias }
 *   'nf'     – Nota Fiscal / Invoice: columns Código, Descrição, QTD.
 *              → returns { format: 'nf', rows: [[nome, qtd], …] }
 *   'rede'   – Rede/Estoque report: columns Código, Nome, Estoq, Vend.
 *              → returns { format: 'rede', estoqueRows: […], vendasRows: […] }
 *   'simples'– Plain 2-column file (col A = name, col B = qty)
 *              → returns { format: 'simples', rows: [[nome, qtd], …] }
 *   'vazio'  – No data found
 *              → returns { format: 'vazio', rows: [] }
 */
function detectAndParseRows(allRows) {
  if (allRows.length === 0) return { format: 'vazio', rows: [] };

  // Check first 15 rows for DROGAMAIS "Sugestão de Compras" report
  for (let i = 0; i < Math.min(15, allRows.length); i++) {
    const cells  = allRows[i].map((c) => String(c == null ? '' : c).trim());
    const joined = cells.join('|').toUpperCase();
    if (joined.includes('SUGEST') && joined.includes('COMPRAS')) {
      return parseSugestaoRows(allRows);
    }
  }

  for (let i = 0; i < Math.min(10, allRows.length); i++) {
    const cells  = allRows[i].map((c) => String(c == null ? '' : c).trim());
    const joined = cells.join('|').toUpperCase();

    // Rede / stock-report format: header contains "ESTOQ" and "VEND"
    if (joined.includes('ESTOQ') && joined.includes('VEND')) {
      return parseRedeRows(allRows, i, cells);
    }

    // DROGAMAIS CSV estoque export: empty col 0+1, "Código" at col 2 (merged-cell artifact)
    if (cells[0] === '' && cells[1] === '' && cells[2] && /C[ÓO]DIGO/i.test(cells[2])) {
      return parseDrogamaisEstoqueRows(allRows, i);
    }

    // Nota Fiscal format: header contains ("CÓDIGO" or "CODIGO") and ("DESCRI" or "QTD")
    if ((joined.includes('C\u00d3DIGO') || joined.includes('CODIGO')) &&
        (joined.includes('DESCRI') || joined.includes('QTD'))) {
      return parseNFRows(allRows, i, cells);
    }
  }

  return { format: 'simples', rows: parseSimpleRows(allRows) };
}

/**
 * Parses a Nota Fiscal file.
 * Extracts the Código column (product code), Descrição column (medicine name)
 * and the QTD. column (quantity). The code is prepended to the name so that
 * medicines with the same name but different codes can be distinguished.
 */
function parseNFRows(allRows, headerIdx, header) {
  const codeIdx = header.findIndex((h) => /C[ÓO]DIGO/i.test(h));
  const descIdx = header.findIndex((h) => /DESCRI/i.test(h));
  const qtdIdx  = header.findIndex((h) => /^QTD\.?$/i.test(h.trim()));

  if (descIdx < 0 || qtdIdx < 0) {
    return { format: 'simples', rows: parseSimpleRows(allRows) };
  }

  const rows = [];
  for (let i = headerIdx + 1; i < allRows.length; i++) {
    const cells = allRows[i].map((c) => String(c == null ? '' : c).trim());
    const code  = codeIdx >= 0 ? (cells[codeIdx] || '') : '';
    const nome  = cells[descIdx] || '';
    const qtd   = parseFloat((cells[qtdIdx] || '').replace(',', '.'));
    if (!nome || isNaN(qtd) || qtd <= 0) continue;
    rows.push([code ? `${code} – ${nome}` : nome, qtd]);
  }
  return { format: 'nf', rows };
}

/**
 * Parses a Rede / Estoque report file.
 * In this format the code is at column index 2 and the medicine name at index 3;
 * these columns carry no header label in the source report (the header row only
 * labels columns from "Estoq" onward). The Estoq and Vend. columns are located
 * dynamically by their header names.
 * Returns separate arrays for estoque and vendas so both sheets can be populated
 * from a single file. The vendas period is assumed to be 30 days (monthly column).
 * Note: valid product codes in this format are always numeric; non-numeric values
 * in that column indicate branch/group header rows that must be skipped.
 */
function parseRedeRows(allRows, headerIdx, header) {
  // Columns 2 (code) and 3 (name) have no header labels in this report format
  const codeIdx  = 2;
  const nomeIdx  = 3;
  const estoqIdx = header.findIndex((h) => /ESTOQ/i.test(h));
  const vendIdx  = header.findIndex((h) => /^VEND\.?$/i.test(h.trim()));

  if (estoqIdx < 0 || vendIdx < 0) {
    return { format: 'simples', rows: parseSimpleRows(allRows) };
  }

  const estoqueRows = [];
  const vendasRows  = [];

  for (let i = headerIdx + 1; i < allRows.length; i++) {
    const cells = allRows[i].map((c) => String(c == null ? '' : c).trim());
    const code  = cells[codeIdx] || '';
    const nome  = cells[nomeIdx] || '';

    // Skip branch/group header rows and blank rows (code is not a number)
    if (!nome || !code || isNaN(parseInt(code, 10))) continue;

    const estoq = parseFloat((cells[estoqIdx] || '0').replace(',', '.')) || 0;
    const vend  = parseFloat((cells[vendIdx]  || '0').replace(',', '.')) || 0;

    const nomeComCodigo = `${code} – ${nome}`;
    estoqueRows.push([nomeComCodigo, estoq]);
    vendasRows.push([nomeComCodigo, vend]);
  }

  return { format: 'rede', estoqueRows, vendasRows };
}

/**
 * Parses the DROGAMAIS CSV estoque export.
 * When the original XLS (with merged cells) is saved as CSV, each merged column
 * gains an extra empty padding column, shifting data positions:
 *   col[1] = Código, col[3] = Descrição, col[12] = QTD.
 * The product code is prepended to the name to help distinguish medicines that
 * share the same description.
 */
function parseDrogamaisEstoqueRows(allRows, headerIdx) {
  const rows = [];
  for (let i = headerIdx + 1; i < allRows.length; i++) {
    const cells = allRows[i].map((c) => String(c == null ? '' : c).trim());
    const code  = cells[1] || '';
    const nome  = cells[3] || '';
    const qtd   = parseFloat((cells[12] || '').replace(',', '.'));
    if (!nome || !code || isNaN(qtd) || qtd <= 0 || isNaN(parseInt(code, 10))) continue;
    rows.push([`${code} – ${nome}`, qtd]);
  }
  return { format: 'nf', rows };
}

/**
 * Parses the DROGAMAIS "Sugestão de Compras – Pelas Vendas no Período" report.
 * Works for both CSV and XLS exports of this report.
 *
 * Column layout is detected dynamically by scanning header rows:
 *   Cód.  column → product code
 *   next column → product name
 *   sub-header "Estoq" column → Saldo Estoque (current branch stock)
 *   sub-header "Vend." column → Qtd. Vendida (qty sold during the period)
 *
 * Fallback defaults if detection fails:
 *   cells[2]=code, cells[3]=name, cells[9]=saldo, cells[13]=qtdVend
 *
 * The sales period (in days) is extracted from the "Período:" date range in the
 * report header so the daily average calculation is accurate.
 *
 * Returns separate arrays for estoque and vendas so both sheets can be populated
 * from a single file.
 */
function parseSugestaoRows(allRows) {
  // Extract the period length (days) from the date range in the report header
  let periodoDias = 7;
  for (let i = 0; i < Math.min(15, allRows.length); i++) {
    const cells = allRows[i].map((c) => String(c == null ? '' : c).trim());
    for (const cell of cells) {
      // Matches date ranges like "04/04/2026 00:00:00 à 11/04/2026 23:59:59" (DD/MM/YYYY)
      const m = cell.match(/(\d{2})\/(\d{2})\/(\d{4}).*?(\d{2})\/(\d{2})\/(\d{4})/);
      if (m) {
        const [, day1, month1, year1, day2, month2, year2] = m;
        const d1   = new Date(+year1, +month1 - 1, +day1);
        const d2   = new Date(+year2, +month2 - 1, +day2);
        const diff = Math.round((d2 - d1) / 86400000);
        if (diff > 0) { periodoDias = diff; break; }
      }
    }
  }

  // ── Detect column positions from header rows ────────────────────────────
  // Defaults match the standard Drogamais report layout (range.s.c = 0)
  let codeCol  = 2;   // "Cód." column
  let nomeCol  = 3;   // Product name column (next after code)
  let saldoCol = 9;   // "Estoq" (Saldo Estoq) in the sub-header row
  let vendCol  = 13;  // "Vend." (Qtd. Vend.) in the sub-header row

  for (let i = 0; i < Math.min(20, allRows.length); i++) {
    const cells = allRows[i].map((c) => String(c == null ? '' : c).trim());

    // Detect code column from "Cód." header (row with main column labels)
    const cidx = cells.findIndex((c) => /^C[ÓO]D\.?$/i.test(c));
    if (cidx >= 0) {
      codeCol = cidx;
      nomeCol = cidx + 1;
    }

    // Detect stock and sales columns from the sub-header row:
    //   "Estoq" = Saldo Estoq (local branch stock)
    //   "Vend." = Qtd. Vendida (qty sold in the period)
    const eidx = cells.findIndex((c) => /^ESTOQ$/i.test(c));
    const vidx = cells.findIndex((c) => /^VEND\.?$/i.test(c));
    if (eidx >= 0 && vidx >= 0) {
      saldoCol = eidx;
      vendCol  = vidx;
      break; // sub-header row found – stop scanning
    }
  }

  const estoqueRows = [];
  const vendasRows  = [];

  for (let i = 0; i < allRows.length; i++) {
    const cells = allRows[i].map((c) => String(c == null ? '' : c).trim());
    const code  = cells[codeCol] || '';
    // Only process rows whose code column is a positive number (product rows)
    if (!code || isNaN(parseFloat(code)) || parseFloat(code) <= 0) continue;

    const nome = cells[nomeCol] || '';
    if (!nome) continue;

    const saldo   = parseFloat((cells[saldoCol] || '0').replace(',', '.')) || 0;
    const qtdVend = parseFloat((cells[vendCol]  || '0').replace(',', '.')) || 0;

    const nomeComCodigo = `${code} – ${nome}`;
    estoqueRows.push([nomeComCodigo, Math.max(0, saldo)]); // clamp negative stock to 0
    vendasRows.push([nomeComCodigo, qtdVend]);
  }

  return { format: 'sugestao', estoqueRows, vendasRows, periodoDias };
}

/**
 * Plain 2-column fallback parser: col A = name, col B = quantity.
 * Quote characters are already stripped by splitCsvLine / SheetJS at this point.
 */
function parseSimpleRows(allRows) {
  const rows = [];
  for (const row of allRows) {
    const nome   = String(row[0] == null ? '' : row[0]).trim();
    const qtdRaw = String(row[1] == null ? '' : row[1]).trim().replace(',', '.');
    const qtd    = parseFloat(qtdRaw);
    if (!nome || isNaN(qtd)) continue;
    rows.push([nome, qtd]);
  }
  return rows;
}

/**
 * Sends rows to the Apps Script in batches of CONFIG.BATCH_SIZE.
 * Only the first batch clears the sheet (clearFirst=true); subsequent batches append.
 */
async function uploadEmLotes(rows, tipo, progressId, periodo) {
  const action = tipo === 'Estoque' ? 'importarEstoque' : 'importarVendas';
  const total  = rows.length;
  let   sent   = 0;

  for (let i = 0; i < total; i += CONFIG.BATCH_SIZE) {
    const lote = rows.slice(i, i + CONFIG.BATCH_SIZE);
    const body = { action, dados: lote, clearFirst: i === 0 };
    if (periodo) body.periodo = periodo; // always send so Apps Script can store it on clearFirst

    const resp = await fetchPost(CONFIG.APPS_SCRIPT_URL, body);
    if (!resp.ok) throw new Error(resp.error || 'Erro desconhecido no servidor.');

    sent += lote.length;
    updateProgress(progressId, sent, total);
  }
}

/* ── Local calculation ───────────────────────────────────────────────────── */
/**
 * Mirrors the backend calcularNecessidade() logic entirely in the browser.
 * Uses the data already parsed from the uploaded files (state.estoqueLocal /
 * state.vendasLocal) so the result is always consistent with what was uploaded,
 * regardless of the order the files were dropped or whether the backend is
 * reachable.
 *
 * Stock check:
 *   mediaDiaria  = vendasTotais / diasVendas
 *   projecao     = mediaDiaria × diasCobertura
 *   necessidade  = projecao − estoqueAtual
 *   → included only when necessidade > 0 (stock is insufficient for the period)
 *
 * @param {number} diasCobertura  Number of days to cover.
 * @returns {Object}  Same shape as the backend response.
 */
function calcularLocal(diasCobertura) {
  // Build estoque map: NOME_UPPER → qty
  const mapaEstoque = {};
  if (state.estoqueLocal) {
    state.estoqueLocal.forEach(([nome, qtd]) => {
      const chave = String(nome == null ? '' : nome).trim().toUpperCase();
      if (chave) mapaEstoque[chave] = Number(qtd) || 0;
    });
  }

  const diasVendas = state.diasVendasLocal || 7;
  const lista = [];

  (state.vendasLocal || []).forEach(([nome, vendasTotais]) => {
    const chave      = String(nome == null ? '' : nome).trim().toUpperCase();
    if (!chave) return;
    const estoqueAtual = mapaEstoque[chave] ?? 0;
    const vendas       = Number(vendasTotais) || 0;
    const mediaDiaria  = vendas / diasVendas;
    const projecao     = mediaDiaria * diasCobertura;
    const necessidade  = projecao - estoqueAtual;

    if (necessidade > 0) {
      lista.push({
        medicamento:  chave,
        estoqueAtual: estoqueAtual,
        mediaDiaria:  Math.round(mediaDiaria * 100) / 100,
        projecao:     Math.round(projecao * 10) / 10,
        comprar:      Math.ceil(necessidade),
      });
    }
  });

  lista.sort((a, b) => b.comprar - a.comprar);

  return {
    diasCobertura,
    diasVendas,
    totalItens: lista.length,
    geradoEm:   new Date().toISOString(),
    itens:      lista,
  };
}

/* ── Calculate purchase list ─────────────────────────────────────────────── */
async function calcular() {
  const btn = $('btnCalcular');
  btn.disabled = true;
  btn.innerHTML = '<span class="spinner"></span> Calculando…';

  $('resultSection').style.display = 'none';

  try {
    // Prefer local calculation: faster, always uses the correct stock values
    // from the uploaded file's "Estoq" column, regardless of upload order.
    if (state.vendasLocal && state.vendasLocal.length > 0) {
      const resultado = calcularLocal(state.periodoSelecionado);
      state.resultado = resultado;
      renderResultado(resultado);
      showToast(`Lista gerada: ${resultado.totalItens} item(s) para comprar.`, 'success');
      return;
    }

    // Fallback: call the Apps Script backend (requires prior upload to succeed)
    const url  = `${CONFIG.APPS_SCRIPT_URL}?action=calcular&dias=${state.periodoSelecionado}`;
    const resp = await fetchGet(url);

    if (!resp.ok) throw new Error(resp.error || 'Erro no cálculo.');

    state.resultado = resp.data;
    renderResultado(resp.data);
    showToast(`Lista gerada: ${resp.data.totalItens} item(s) para comprar.`, 'success');
  } catch (err) {
    showToast(`Erro: ${err.message}`, 'error');
  } finally {
    btn.disabled = false;
    btn.innerHTML = '🔍 Calcular Lista de Compras';
  }
}

/* ── Render results table ────────────────────────────────────────────────── */
function renderResultado(data) {
  const section = $('resultSection');
  section.style.display = 'block';

  $('metaDias').textContent    = data.diasCobertura;
  $('metaTotal').textContent   = data.totalItens;
  $('metaGerado').textContent  = new Date(data.geradoEm).toLocaleString('pt-BR');

  const tbody = $('tabelaBody');
  tbody.innerHTML = '';

  if (!data.itens || data.itens.length === 0) {
    $('emptyState').style.display = 'block';
    $('tabelaWrap').style.display = 'none';
    return;
  }

  $('emptyState').style.display = 'none';
  $('tabelaWrap').style.display = '';

  data.itens.forEach((item, idx) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${idx + 1}</td>
      <td>${escHtml(item.medicamento)}</td>
      <td>${item.estoqueAtual}</td>
      <td>${item.mediaDiaria.toFixed(2)}/dia</td>
      <td>${item.projecao.toFixed(1)}</td>
      <td><span class="badge-comprar">${item.comprar}</span></td>`;
    tbody.appendChild(tr);
  });
}

/* ── PDF download helper (cross-device) ──────────────────────────────────── */
/**
 * Saves a jsPDF document.
 * Uses a blob URL + hidden anchor trick so the download works on both desktop
 * and mobile browsers (including iOS Safari which ignores the `download` attr
 * but can open a blob in a new tab where the user can then share/save the PDF).
 */
function downloadPdf(doc, filename) {
  try {
    const blob = doc.output('blob');
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement('a');
    a.href     = url;
    a.download = filename;
    a.target   = '_blank';
    a.rel      = 'noopener';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    // Release the object URL after a short delay to allow the download to start
    setTimeout(() => URL.revokeObjectURL(url), 500);
  } catch (e) {
    // Fallback: use jsPDF built-in save (works on most desktop browsers)
    doc.save(filename);
  }
}

/* ── PDF export ──────────────────────────────────────────────────────────── */
function exportarPdf() {
  if (!state.resultado) return;
  if (typeof window.jspdf === 'undefined' || typeof window.jspdf.jsPDF === 'undefined') {
    showToast('Biblioteca de PDF não carregou. Verifique sua conexão e recarregue a página.', 'error');
    return;
  }
  try {
  const { jsPDF } = window.jspdf;
  const doc  = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });
  const data = state.resultado;
  const farmacia = CONFIG.FARMACIA_NOME;
  const hoje = new Date().toLocaleDateString('pt-BR');

  // Header
  doc.setFillColor(26, 111, 196);
  doc.rect(0, 0, 210, 28, 'F');
  doc.setTextColor(255, 255, 255);
  doc.setFontSize(16);
  doc.setFont('helvetica', 'bold');
  doc.text(farmacia, 14, 12);
  doc.setFontSize(10);
  doc.setFont('helvetica', 'normal');
  doc.text(`Sugestão de Compra – Cobertura: ${data.diasCobertura} dias`, 14, 20);
  doc.text(`Gerado em: ${hoje}`, 155, 20);

  // Table
  doc.setTextColor(0, 0, 0);
  doc.setFontSize(8);
  const cols   = ['#', 'Cód. / Medicamento', 'Estoque', 'Média/dia', 'Projeção', 'Comprar'];
  const widths = [8, 90, 18, 22, 22, 17];
  let y = 36;

  // Column headers
  doc.setFont('helvetica', 'bold');
  doc.setFillColor(230, 238, 255);
  doc.rect(10, y - 5, 190, 8, 'F');
  let x = 14;
  cols.forEach((c, i) => {
    doc.text(c, x, y);
    x += widths[i];
  });
  doc.setFont('helvetica', 'normal');
  y += 6;

  // Rows
  data.itens.forEach((item, idx) => {
    if (y > 275) {
      doc.addPage();
      y = 20;
    }
    if (idx % 2 === 0) {
      doc.setFillColor(247, 250, 255);
      doc.rect(10, y - 4, 190, 7, 'F');
    }
    x = 14;
    const row = [
      String(idx + 1),
      item.medicamento.length > 48 ? item.medicamento.substring(0, 46) + '…' : item.medicamento,
      String(item.estoqueAtual),
      item.mediaDiaria.toFixed(2),
      item.projecao.toFixed(1),
      String(item.comprar),
    ];
    row.forEach((cell, i) => {
      if (i === 5) doc.setFont('helvetica', 'bold');
      doc.text(cell, x, y);
      doc.setFont('helvetica', 'normal');
      x += widths[i];
    });
    y += 7;
  });

  // Footer
  doc.setFontSize(7);
  doc.setTextColor(120, 120, 120);
  doc.text(`Total: ${data.totalItens} item(s) | Período de vendas: ${data.diasVendas} dias`, 14, 290);

  downloadPdf(doc, `lista-compras-${data.diasCobertura}dias-${hoje.replace(/\//g, '-')}.pdf`);
  showToast('PDF gerado com sucesso!', 'success');
  } catch (err) {
    showToast(`Erro ao gerar PDF: ${err.message}`, 'error');
  }
}

/* ── WhatsApp share ──────────────────────────────────────────────────────── */
function compartilharWhatsapp() {
  if (!state.resultado) return;
  const data = state.resultado;

  let msg = `*Sugestão de Compra – ${CONFIG.FARMACIA_NOME}*\n`;
  msg    += `_Cobertura: ${data.diasCobertura} dias | ${new Date().toLocaleDateString('pt-BR')}_\n\n`;

  data.itens.forEach((item) => {
    msg += `• ${item.medicamento}: *${item.comprar} un*\n`;
  });

  msg += `\n_Total: ${data.totalItens} item(s)_`;

  const url = `https://wa.me/?text=${encodeURIComponent(msg)}`;
  window.open(url, '_blank');
}

/* ── Report: most-sold items ─────────────────────────────────────────────── */
async function gerarRelatorio() {
  const btn = $('btnRelatorio');
  btn.disabled = true;
  btn.innerHTML = '<span class="spinner"></span> Gerando…';

  try {
    // Prefer local data when available
    if (state.vendasLocal && state.vendasLocal.length > 0) {
      const diasVendas = state.diasVendasLocal || 7;
      let lista = state.vendasLocal.map(([nome, qtd]) => {
        const total = Number(qtd) || 0;
        const media = Math.round((total / diasVendas) * 100) / 100;
        return { medicamento: String(nome).trim().toUpperCase(), totalVendido: total, mediaDiaria: media };
      });
      lista = lista.filter((i) => i.totalVendido > 0);
      lista.sort((a, b) => b.mediaDiaria - a.mediaDiaria);
      const top  = state.topSelecionado;
      const itens = top > 0 ? lista.slice(0, top) : lista;
      const data = { diasVendas, totalItens: itens.length, geradoEm: new Date().toISOString(), itens };
      state.relatorio = data;
      renderRelatorio(data);
      showToast(`Relatório gerado: ${data.totalItens} item(s).`, 'success');
      return;
    }

    // Fallback: call the Apps Script backend
    const url  = `${CONFIG.APPS_SCRIPT_URL}?action=relatorio&top=${state.topSelecionado}`;
    const resp = await fetchGet(url);
    if (!resp.ok) throw new Error(resp.error || 'Erro ao gerar relatório.');

    state.relatorio = resp.data;
    renderRelatorio(resp.data);
    showToast(`Relatório gerado: ${resp.data.totalItens} item(s).`, 'success');
  } catch (err) {
    showToast(`Erro: ${err.message}`, 'error');
  } finally {
    btn.disabled = false;
    btn.innerHTML = '📊 Gerar Relatório';
  }
}

function renderRelatorio(data) {
  $('relatorioResultado').style.display = 'block';

  const label = state.topSelecionado > 0 ? `Top ${state.topSelecionado}` : 'Todos';
  $('relatorioMeta').innerHTML =
    `<span>Exibindo: <strong>${label}</strong></span>` +
    `<span>Itens: <strong>${data.totalItens}</strong></span>` +
    `<span>Período de vendas: <strong>${data.diasVendas} dias</strong></span>` +
    `<span>Gerado em: <strong>${new Date(data.geradoEm).toLocaleString('pt-BR')}</strong></span>`;

  const tbody = $('relatorioBody');
  tbody.innerHTML = '';

  data.itens.forEach((item, idx) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td><span class="badge-rank">${idx + 1}º</span></td>
      <td>${escHtml(item.medicamento)}</td>
      <td>${item.totalVendido}</td>
      <td><strong>${item.mediaDiaria.toFixed(2)}</strong>/dia</td>`;
    tbody.appendChild(tr);
  });
}

function exportarRelatorioPdf() {
  if (!state.relatorio) return;
  if (typeof window.jspdf === 'undefined' || typeof window.jspdf.jsPDF === 'undefined') {
    showToast('Biblioteca de PDF não carregou. Verifique sua conexão e recarregue a página.', 'error');
    return;
  }
  try {
  const { jsPDF } = window.jspdf;
  const doc  = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });
  const data = state.relatorio;
  const farmacia = CONFIG.FARMACIA_NOME;
  const hoje = new Date().toLocaleDateString('pt-BR');
  const label = state.topSelecionado > 0 ? `Top ${state.topSelecionado}` : 'Todos';

  // Header
  doc.setFillColor(26, 111, 196);
  doc.rect(0, 0, 210, 28, 'F');
  doc.setTextColor(255, 255, 255);
  doc.setFontSize(16);
  doc.setFont('helvetica', 'bold');
  doc.text(farmacia, 14, 12);
  doc.setFontSize(10);
  doc.setFont('helvetica', 'normal');
  doc.text(`Relatório de Mais Vendidos – ${label} | Período: ${data.diasVendas} dias`, 14, 20);
  doc.text(`Gerado em: ${hoje}`, 155, 20);

  // Table
  doc.setTextColor(0, 0, 0);
  doc.setFontSize(8);
  const cols   = ['#', 'Cód. / Medicamento', 'Total Vendido', 'Média Diária'];
  const widths = [10, 110, 35, 35];
  let y = 36;

  doc.setFont('helvetica', 'bold');
  doc.setFillColor(230, 238, 255);
  doc.rect(10, y - 5, 190, 8, 'F');
  let x = 14;
  cols.forEach((c, i) => { doc.text(c, x, y); x += widths[i]; });
  doc.setFont('helvetica', 'normal');
  y += 6;

  data.itens.forEach((item, idx) => {
    if (y > 275) { doc.addPage(); y = 20; }
    if (idx % 2 === 0) {
      doc.setFillColor(247, 250, 255);
      doc.rect(10, y - 4, 190, 7, 'F');
    }
    x = 14;
    const row = [
      String(idx + 1) + 'º',
      item.medicamento.length > 60 ? item.medicamento.substring(0, 58) + '…' : item.medicamento,
      String(item.totalVendido),
      item.mediaDiaria.toFixed(2) + '/dia',
    ];
    row.forEach((cell, i) => {
      if (i === 3) doc.setFont('helvetica', 'bold');
      doc.text(cell, x, y);
      doc.setFont('helvetica', 'normal');
      x += widths[i];
    });
    y += 7;
  });

  doc.setFontSize(7);
  doc.setTextColor(120, 120, 120);
  doc.text(`Total: ${data.totalItens} item(s) | Período de vendas: ${data.diasVendas} dias`, 14, 290);

  downloadPdf(doc, `relatorio-mais-vendidos-${label.toLowerCase().replace(/ /g, '')}-${hoje.replace(/\//g, '-')}.pdf`);
  showToast('PDF do relatório gerado!', 'success');
  } catch (err) {
    showToast(`Erro ao gerar PDF: ${err.message}`, 'error');
  }
}

/* ── HTTP helpers ────────────────────────────────────────────────────────── */
async function fetchGet(url) {
  const resp = await fetch(url);
  if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
  return resp.json();
}

async function fetchPost(url, body) {
  const resp = await fetch(url, {
    method:  'POST',
    headers: { 'Content-Type': 'text/plain' }, // Apps Script requires text/plain for doPost
    body:    JSON.stringify(body),
  });
  if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
  return resp.json();
}

/* ── UI helpers ──────────────────────────────────────────────────────────── */
function updateCalcularBtn() {
  $('btnCalcular').disabled = !(state.estoqueCarregado && state.vendasCarregadas);
  $('btnRelatorio').disabled = !state.vendasCarregadas;
}

function showProgress(progressId) {
  const wrap = $(`${progressId}Wrap`);
  if (wrap) wrap.style.display = 'block';
}

function updateProgress(progressId, done, total) {
  const fg    = $(`${progressId}Fg`);
  const label = $(`${progressId}Label`);
  const pct   = Math.round((done / total) * 100);
  if (fg)    fg.style.width   = `${pct}%`;
  if (label) label.textContent = `${done} / ${total}`;
}

let toastTimer;
function showToast(msg, type = '') {
  const t = $('toast');
  t.textContent = msg;
  t.className   = `show ${type}`;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => { t.className = ''; }, 3500);
}

function escHtml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

/* ─────────────────────────────────────────────────────────────────────────── */
/* ── Validades (Expiry Dates) Feature ───────────────────────────────────── */
/* ─────────────────────────────────────────────────────────────────────────── */

/* ── Validade period selector ──────────────────────────────────────────── */
function setupValidadePeriodButtons() {
  document.querySelectorAll('.vld-btn').forEach((btn) => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.vld-btn').forEach((b) => b.classList.remove('active'));
      btn.classList.add('active');
      state.validadeDias = parseInt(btn.dataset.dias, 10);
      // Re-render if data is already loaded
      if (state.validadeItens) renderValidade(state.validadeItens);
    });
  });
  state.validadeDias = 60; // default
}

/* ── Validade dropzone setup ────────────────────────────────────────────── */
function setupValidadeDropzone() {
  const zone  = $('validadeDropzone');
  const input = $('validadeCsvInput');

  zone.addEventListener('dragover', (e) => {
    e.preventDefault();
    zone.classList.add('over');
  });
  zone.addEventListener('dragleave', () => zone.classList.remove('over'));
  zone.addEventListener('drop', (e) => {
    e.preventDefault();
    zone.classList.remove('over');
    const file = e.dataTransfer.files[0];
    if (file) handleValidadeFile(file);
  });

  input.addEventListener('change', () => {
    if (input.files[0]) handleValidadeFile(input.files[0]);
    input.value = '';
  });
}

/* ── Handle uploaded validade file ─────────────────────────────────────── */
function handleValidadeFile(file) {
  const isXls = file.name.match(/\.xlsx?$/i) ||
                /application\/vnd\.(ms-excel|openxmlformats-officedocument\.spreadsheetml\.sheet)/.test(file.type);
  const isCsv = file.name.match(/\.(csv|txt)$/i) ||
                /^text\/(csv|plain)$/.test(file.type);

  if (!isXls && !isCsv) {
    showToast('Arquivo inválido. Use CSV, XLS ou XLSX.', 'error');
    return;
  }

  const status = $('validadeStatus');
  status.textContent = 'Lendo arquivo…';

  if (isXls) {
    if (typeof XLSX === 'undefined') {
      showToast('Erro: biblioteca XLS não carregou. Verifique sua conexão e recarregue a página.', 'error');
      status.textContent = '❌ Biblioteca XLS indisponível.';
      return;
    }
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data   = new Uint8Array(e.target.result);
        const wb     = XLSX.read(data, { type: 'array', cellDates: true });
        const ws     = wb.Sheets[wb.SheetNames[0]];
        const itens  = parseValidadeXls(ws);
        processValidadeItens(itens, status);
      } catch (err) {
        status.textContent = `❌ Erro: ${err.message}`;
        showToast(`Erro ao ler arquivo: ${err.message}`, 'error');
      }
    };
    reader.readAsArrayBuffer(file);
  } else {
    const reader = new FileReader();
    reader.onload = (e) => {
      const itens = parseValidadeCsv(e.target.result);
      processValidadeItens(itens, status);
    };
    reader.readAsText(file, 'UTF-8');
  }
}

function processValidadeItens(itens, status) {
  if (itens.length === 0) {
    status.textContent = '⚠️ Nenhum produto com data de validade encontrado. Verifique o formato do arquivo.';
    showToast('Nenhum item de validade encontrado.', 'error');
    return;
  }
  state.validadeItens = itens;
  status.textContent = `✅ ${itens.length} itens carregados.`;
  renderValidade(itens);
  showToast(`Validades carregadas: ${itens.length} itens.`, 'success');
}

/**
 * Parses a SheetJS worksheet for expiry-date data.
 * Detects columns dynamically: Produto/Descrição, Lote, Validade, Qtd.
 */
function parseValidadeXls(ws) {
  const ref = ws['!ref'];
  if (!ref) return [];

  const range   = XLSX.utils.decode_range(ref);
  const allRows = [];
  for (let r = range.s.r; r <= range.e.r; r++) {
    const row = [];
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      // Preserve raw value for dates; cell.w is the formatted string
      row.push(cell ? (cell.w !== undefined ? cell.w : cell.v) : '');
    }
    allRows.push(row);
  }
  return parseValidadeRows(allRows);
}

/**
 * Parses a CSV text for expiry-date data.
 */
function parseValidadeCsv(text) {
  const lines   = text.split(/\r?\n/).map((l) => l.trim()).filter(Boolean);
  const sep     = lines[0]?.includes(';') ? ';' : ',';
  const allRows = lines.map((line) => splitCsvLine(line, sep));
  return parseValidadeRows(allRows);
}

/**
 * Finds the header row and column indices for a validade file, then extracts items.
 * Supports flexible column layouts by searching for keywords.
 *
 * Required: a column containing a date (DD/MM/YYYY or MM/YYYY)
 * Optional: Produto/Descrição, Lote, Quantidade
 */
function parseValidadeRows(allRows) {
  // ── Detect header row and column indices ──────────────────────────────────
  let prodCol     = -1;
  let loteCol     = -1;
  let validadeCol = -1;
  let qtdCol      = -1;
  let codeCol     = -1;
  let headerIdx   = -1;

  for (let i = 0; i < Math.min(20, allRows.length); i++) {
    const cells = allRows[i].map((c) => String(c == null ? '' : c).trim().toUpperCase());
    // Look for any header keyword
    const hasHeader = cells.some(
      (c) => /VALIDADE|VENCIMENTO|PRODUTO|DESCRIÇÃO|LOTE|QUANT/i.test(c)
    );
    if (!hasHeader) continue;

    headerIdx   = i;
    validadeCol = cells.findIndex((c) => /VALIDADE|VENCIMENTO/i.test(c));
    prodCol     = cells.findIndex((c) => /PRODUTO|DESCRI|MEDICAMENTO|NOME/i.test(c));
    loteCol     = cells.findIndex((c) => /^LOTE$/i.test(c));
    qtdCol      = cells.findIndex((c) => /QUANT|QTD/i.test(c));
    codeCol     = cells.findIndex((c) => /^C[ÓO]D(IGO)?\.?$/i.test(c));
    break;
  }

  // If no header found try to detect columns by data patterns
  if (validadeCol < 0) {
    // Scan data rows for a column that has date-like values
    for (let i = 0; i < Math.min(30, allRows.length); i++) {
      const cells = allRows[i].map((c) => String(c == null ? '' : c).trim());
      for (let j = 0; j < cells.length; j++) {
        if (/^\d{2}\/\d{2}\/\d{4}$/.test(cells[j]) || /^\d{2}\/\d{4}$/.test(cells[j])) {
          validadeCol = j;
          // Assume product name is the first text column before the date
          if (j > 0) prodCol = j - 1;
          headerIdx = -1; // no header row
          break;
        }
      }
      if (validadeCol >= 0) break;
    }
  }

  if (validadeCol < 0) return []; // no date column found

  const items = [];
  const dataStart = headerIdx >= 0 ? headerIdx + 1 : 0;

  for (let i = dataStart; i < allRows.length; i++) {
    const cells   = allRows[i].map((c) => String(c == null ? '' : c).trim());
    const valRaw  = cells[validadeCol] || '';
    if (!valRaw) continue;

    const validade = parseValidadeDate(valRaw);
    if (!validade) continue;

    const code  = codeCol  >= 0 ? (cells[codeCol]  || '') : '';
    const nome  = prodCol  >= 0 ? (cells[prodCol]   || '') : '';
    const lote  = loteCol  >= 0 ? (cells[loteCol]   || '') : '';
    const qtd   = qtdCol   >= 0 ? (parseFloat((cells[qtdCol] || '0').replace(',', '.')) || 0) : 0;

    if (!nome && !code) continue; // skip rows without product identification

    items.push({
      code,
      nome: code ? `${code} – ${nome}` : nome,
      lote,
      validade,
      qtd,
    });
  }
  return items;
}

/**
 * Parses a date string in DD/MM/YYYY or MM/YYYY format.
 * MM/YYYY is treated as the last day of that month.
 * Returns a Date object or null if parsing fails.
 */
function parseValidadeDate(str) {
  // DD/MM/YYYY
  let m = str.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (m) {
    const d = new Date(+m[3], +m[2] - 1, +m[1]);
    return isNaN(d.getTime()) ? null : d;
  }
  // MM/YYYY → last day of that month
  m = str.match(/^(\d{2})\/(\d{4})$/);
  if (m) {
    const d = new Date(+m[2], +m[1], 0); // day 0 of next month = last day of this month
    return isNaN(d.getTime()) ? null : d;
  }
  // YYYY-MM-DD (ISO)
  m = str.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) {
    const d = new Date(+m[1], +m[2] - 1, +m[3]);
    return isNaN(d.getTime()) ? null : d;
  }
  // Try native Date parsing as last resort
  const d = new Date(str);
  return isNaN(d.getTime()) ? null : d;
}

/* ── Render validade results ────────────────────────────────────────────── */
function renderValidade(itens) {
  const hoje       = new Date();
  hoje.setHours(0, 0, 0, 0);
  const diasAlerta = state.validadeDias || 60;
  const limiteAlerta = new Date(hoje.getTime() + diasAlerta * 86400000);

  const vencidos     = [];
  const proxVencer   = [];

  itens.forEach((item) => {
    const val = new Date(item.validade);
    val.setHours(0, 0, 0, 0);
    const diffDias = Math.round((val - hoje) / 86400000);
    if (diffDias < 0) {
      vencidos.push({ ...item, diffDias });
    } else if (val <= limiteAlerta) {
      proxVencer.push({ ...item, diffDias });
    }
  });

  // Sort: most expired first, then soonest-to-expire first
  vencidos.sort((a, b) => a.diffDias - b.diffDias);
  proxVencer.sort((a, b) => a.diffDias - b.diffDias);

  state.validadeResultado = { vencidos, proxVencer, diasAlerta };

  const section = $('validadeResultSection');
  section.style.display = 'block';

  $('validadeMeta').innerHTML =
    `<span>Total carregado: <strong>${itens.length}</strong></span>` +
    `<span>🔴 Vencidos: <strong>${vencidos.length}</strong></span>` +
    `<span>🟡 Vencem em até ${diasAlerta} dias: <strong>${proxVencer.length}</strong></span>` +
    `<span>Gerado em: <strong>${hoje.toLocaleDateString('pt-BR')}</strong></span>`;

  const hasItems = vencidos.length > 0 || proxVencer.length > 0;
  $('validadeEmpty').style.display = hasItems ? 'none' : 'block';

  // Vencidos table
  const vWrap = $('vencidosWrap');
  if (vencidos.length > 0) {
    vWrap.style.display = 'block';
    const tbody = $('vencidosBody');
    tbody.innerHTML = '';
    vencidos.forEach((item, idx) => {
      const tr = document.createElement('tr');
      tr.className = 'row-vencido';
      tr.innerHTML = `
        <td>${idx + 1}</td>
        <td>${escHtml(item.nome)}</td>
        <td>${escHtml(item.lote)}</td>
        <td>${item.validade.toLocaleDateString('pt-BR')}</td>
        <td>${item.qtd || '–'}</td>
        <td><span class="badge-vencido">${Math.abs(item.diffDias)} dias atrás</span></td>`;
      tbody.appendChild(tr);
    });
  } else {
    vWrap.style.display = 'none';
  }

  // Próximos a vencer table
  const pWrap = $('proxVencerWrap');
  if (proxVencer.length > 0) {
    pWrap.style.display = 'block';
    const tbody = $('proxVencerBody');
    tbody.innerHTML = '';
    proxVencer.forEach((item, idx) => {
      const tr = document.createElement('tr');
      tr.className = 'row-proxvencer';
      tr.innerHTML = `
        <td>${idx + 1}</td>
        <td>${escHtml(item.nome)}</td>
        <td>${escHtml(item.lote)}</td>
        <td>${item.validade.toLocaleDateString('pt-BR')}</td>
        <td>${item.qtd || '–'}</td>
        <td><span class="badge-proxvencer">${item.diffDias} dias</span></td>`;
      tbody.appendChild(tr);
    });
  } else {
    pWrap.style.display = 'none';
  }

  section.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

/* ── PDF export for validade ────────────────────────────────────────────── */
function exportarValidadePdf() {
  if (!state.validadeResultado) return;
  if (typeof window.jspdf === 'undefined' || typeof window.jspdf.jsPDF === 'undefined') {
    showToast('Biblioteca de PDF não carregou. Verifique sua conexão e recarregue a página.', 'error');
    return;
  }
  try {
    const { jsPDF } = window.jspdf;
    const doc       = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });
    const farmacia  = CONFIG.FARMACIA_NOME;
    const hoje      = new Date().toLocaleDateString('pt-BR');
    const { vencidos, proxVencer, diasAlerta } = state.validadeResultado;

    // Header
    doc.setFillColor(26, 111, 196);
    doc.rect(0, 0, 210, 28, 'F');
    doc.setTextColor(255, 255, 255);
    doc.setFontSize(16);
    doc.setFont('helvetica', 'bold');
    doc.text(farmacia, 14, 12);
    doc.setFontSize(10);
    doc.setFont('helvetica', 'normal');
    doc.text(`Relatório de Validades – Alerta: ${diasAlerta} dias`, 14, 20);
    doc.text(`Gerado em: ${hoje}`, 155, 20);

    const cols   = ['#', 'Cód. / Produto', 'Lote', 'Validade', 'Qtd', 'Situação'];
    const widths = [8, 76, 28, 24, 14, 30];

    function drawSection(title, items, colorFn) {
      let y = doc.internal.getCurrentPageInfo().pageNumber === 1 ? 36 : 20;
      // Section title
      doc.setFontSize(11);
      doc.setFont('helvetica', 'bold');
      doc.setTextColor(...colorFn('title'));
      doc.text(title, 14, y);
      y += 8;

      // Column headers
      doc.setFontSize(8);
      doc.setTextColor(0, 0, 0);
      doc.setFillColor(230, 238, 255);
      doc.rect(10, y - 5, 190, 8, 'F');
      let x = 14;
      cols.forEach((c, i) => { doc.text(c, x, y); x += widths[i]; });
      doc.setFont('helvetica', 'normal');
      y += 6;

      items.forEach((item, idx) => {
        if (y > 275) { doc.addPage(); y = 20; }
        if (idx % 2 === 0) {
          doc.setFillColor(...colorFn('rowBg'));
          doc.rect(10, y - 4, 190, 7, 'F');
        }
        x = 14;
        const situacao = item.diffDias < 0
          ? `Vencido há ${Math.abs(item.diffDias)}d`
          : `Vence em ${item.diffDias}d`;
        const row = [
          String(idx + 1),
          item.nome.length > 42 ? item.nome.substring(0, 40) + '…' : item.nome,
          item.lote || '–',
          item.validade.toLocaleDateString('pt-BR'),
          item.qtd ? String(item.qtd) : '–',
          situacao,
        ];
        row.forEach((cell, i) => {
          if (i === 5) {
            doc.setFont('helvetica', 'bold');
            doc.setTextColor(...colorFn('situacao'));
          }
          doc.text(cell, x, y);
          doc.setFont('helvetica', 'normal');
          doc.setTextColor(0, 0, 0);
          x += widths[i];
        });
        y += 7;
      });
    }

    if (vencidos.length > 0) {
      drawSection(`🔴 Produtos Vencidos (${vencidos.length})`, vencidos,
        (part) => part === 'title' ? [192, 57, 43] : part === 'rowBg' ? [255, 235, 235] : [192, 57, 43]);
    }
    if (proxVencer.length > 0) {
      if (vencidos.length > 0) doc.addPage();
      drawSection(`🟡 Próximos a Vencer em até ${diasAlerta} dias (${proxVencer.length})`, proxVencer,
        (part) => part === 'title' ? [180, 90, 0] : part === 'rowBg' ? [255, 248, 225] : [180, 90, 0]);
    }

    doc.setFontSize(7);
    doc.setTextColor(120, 120, 120);
    doc.text(`Total vencidos: ${vencidos.length} | Próximos a vencer: ${proxVencer.length}`, 14, 290);

    downloadPdf(doc, `validades-${hoje.replace(/\//g, '-')}.pdf`);
    showToast('PDF de validades gerado!', 'success');
  } catch (err) {
    showToast(`Erro ao gerar PDF: ${err.message}`, 'error');
  }
}
