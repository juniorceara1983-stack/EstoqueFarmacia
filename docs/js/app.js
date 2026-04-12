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
  estoqueCarregado: false,
  vendasCarregadas: false,
  resultado: null,
};

/* ── DOM refs ────────────────────────────────────────────────────────────── */
const $ = (id) => document.getElementById(id);

/* ── Bootstrap ───────────────────────────────────────────────────────────── */
document.addEventListener('DOMContentLoaded', () => {
  setupPeriodButtons();
  setupDropzone('estoqueDropzone', 'estoqueCsvInput', 'estoqueProgress',
                'estoqueStatus', 'Estoque', false);
  setupDropzone('vendasDropzone', 'vendasCsvInput', 'vendasProgress',
                'vendasStatus', 'Vendas', true);
  $('btnCalcular').addEventListener('click', calcular);
  $('btnPdf').addEventListener('click', exportarPdf);
  $('btnWhatsapp').addEventListener('click', compartilharWhatsapp);
  updateCalcularBtn();
});

/* ── Period selector ────────────────────────────────────────────────────── */
function setupPeriodButtons() {
  document.querySelectorAll('.period-btn').forEach((btn) => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.period-btn').forEach((b) => b.classList.remove('active'));
      btn.classList.add('active');
      state.periodoSelecionado = parseInt(btn.dataset.dias, 10);
    });
  });
  // Default active
  document.querySelector('.period-btn[data-dias="14"]').classList.add('active');
}

/* ── Dropzone helper ─────────────────────────────────────────────────────── */
function setupDropzone(zoneId, inputId, progressId, statusId, tipo, isVendas) {
  const zone  = $(zoneId);
  const input = $(inputId);

  zone.addEventListener('click', () => input.click());

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
  if (!file.name.match(/\.(csv|txt)$/i)) {
    showToast('Arquivo inválido. Use um arquivo CSV.', 'error');
    return;
  }

  const status = $(statusId);
  status.textContent = 'Lendo arquivo…';

  const reader = new FileReader();
  reader.onload = async (e) => {
    const rows = parseCsv(e.target.result);
    if (rows.length === 0) {
      status.textContent = '⚠️ Arquivo vazio ou sem dados.';
      showToast('O arquivo não contém dados válidos.', 'error');
      return;
    }

    status.textContent = `${rows.length} linhas encontradas. Enviando…`;
    showProgress(progressId);

    try {
      await uploadEmLotes(rows, tipo, progressId, isVendas ? state.periodoSelecionado : null);
      status.textContent = `✅ ${rows.length} itens importados com sucesso!`;
      if (isVendas) state.vendasCarregadas = true;
      else          state.estoqueCarregado = true;
      updateCalcularBtn();
      showToast(`${tipo} importado: ${rows.length} itens.`, 'success');
    } catch (err) {
      status.textContent = `❌ Erro: ${err.message}`;
      showToast(`Erro ao importar ${tipo}: ${err.message}`, 'error');
    }
  };
  reader.readAsText(file, 'UTF-8');
}

/**
 * Parses a CSV string into an array of [name, qty] pairs.
 * Handles semicolons and commas as delimiters.
 * Skips header row if first column is non-numeric text (e.g. "Medicamento").
 */
function parseCsv(text) {
  const lines = text.split(/\r?\n/).map((l) => l.trim()).filter(Boolean);
  const sep   = lines[0].includes(';') ? ';' : ',';
  const rows  = [];

  for (const line of lines) {
    const parts = line.split(sep);
    const nome  = (parts[0] || '').replace(/^"|"$/g, '').trim();
    const qtdRaw = (parts[1] || '').replace(/^"|"$/g, '').replace(',', '.').trim();
    const qtd   = parseFloat(qtdRaw);
    if (!nome || isNaN(qtd)) continue; // skip header or blank
    rows.push([nome, qtd]);
  }
  return rows;
}

/**
 * Sends rows to the Apps Script in batches of CONFIG.BATCH_SIZE.
 */
async function uploadEmLotes(rows, tipo, progressId, periodo) {
  const action = tipo === 'Estoque' ? 'importarEstoque' : 'importarVendas';
  const total  = rows.length;
  let   sent   = 0;

  for (let i = 0; i < total; i += CONFIG.BATCH_SIZE) {
    const lote = rows.slice(i, i + CONFIG.BATCH_SIZE);
    const body = { action, dados: lote };
    if (periodo) body.periodo = periodo;

    const resp = await fetchPost(CONFIG.APPS_SCRIPT_URL, body);
    if (!resp.ok) throw new Error(resp.error || 'Erro desconhecido no servidor.');

    sent += lote.length;
    updateProgress(progressId, sent, total);
  }
}

/* ── Calculate purchase list ─────────────────────────────────────────────── */
async function calcular() {
  const btn = $('btnCalcular');
  btn.disabled = true;
  btn.innerHTML = '<span class="spinner"></span> Calculando…';

  $('resultSection').style.display = 'none';

  try {
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

/* ── PDF export ──────────────────────────────────────────────────────────── */
function exportarPdf() {
  if (!state.resultado) return;
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
  const cols   = ['#', 'Medicamento', 'Estoque', 'Média/dia', 'Projeção', 'Comprar'];
  const widths = [8, 85, 18, 22, 22, 22];
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
      item.medicamento.length > 45 ? item.medicamento.substring(0, 43) + '…' : item.medicamento,
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

  doc.save(`lista-compras-${data.diasCobertura}dias-${hoje.replace(/\//g, '-')}.pdf`);
  showToast('PDF gerado com sucesso!', 'success');
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
