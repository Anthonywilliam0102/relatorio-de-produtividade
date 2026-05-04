/* ================================================================
   ADVBOX — Dashboard de Produtividade  |  script.js
   Bugs corrigidos:
   1. suggestColumn9IfLikely movida para escopo global
   2. inc() para countsRed e countsProgrammed adicionados no loop
   3. % Vermelho corrigido: red / (red + near + remote)
   ================================================================ */

'use strict';

const $ = (id) => document.getElementById(id);
let state = { resolved: null, open: null, chartCounts: null, chartScore: null, lastRows: [] };

/* ─── PESOS ─────────────────────────────────────────────────── */
const WEIGHTS_KEY = 'advbox_weights_v1';
const DEFAULT_WEIGHTS = { wDone: 1.0, wRed: 2.0, wNear: 0.3, wRemote: 0.1 };

function clampNum(x, fallback) {
  const n = Number(x);
  return Number.isFinite(n) ? n : fallback;
}

function loadWeights() {
  try {
    const raw = localStorage.getItem(WEIGHTS_KEY);
    if (!raw) return { ...DEFAULT_WEIGHTS };
    const obj = JSON.parse(raw);
    return {
      wDone:   clampNum(obj.wDone,   DEFAULT_WEIGHTS.wDone),
      wRed:    clampNum(obj.wRed,    DEFAULT_WEIGHTS.wRed),
      wNear:   clampNum(obj.wNear,   DEFAULT_WEIGHTS.wNear),
      wRemote: clampNum(obj.wRemote, DEFAULT_WEIGHTS.wRemote),
    };
  } catch (e) {
    return { ...DEFAULT_WEIGHTS };
  }
}

function applyWeightsToUI(w) {
  ['wDone','wRed','wNear','wRemote'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.value = String(w[id]);
  });
}

function getWeights() {
  return {
    wDone:   clampNum(document.getElementById('wDone')?.value,   DEFAULT_WEIGHTS.wDone),
    wRed:    clampNum(document.getElementById('wRed')?.value,    DEFAULT_WEIGHTS.wRed),
    wNear:   clampNum(document.getElementById('wNear')?.value,   DEFAULT_WEIGHTS.wNear),
    wRemote: clampNum(document.getElementById('wRemote')?.value, DEFAULT_WEIGHTS.wRemote),
  };
}

function saveWeights() {
  localStorage.setItem(WEIGHTS_KEY, JSON.stringify(getWeights()));
  showToast('ok', 'Pesos salvos com sucesso.');
  if (state.resolved && state.open) computeDashboard();
}

function resetWeights() {
  localStorage.setItem(WEIGHTS_KEY, JSON.stringify(DEFAULT_WEIGHTS));
  applyWeightsToUI(DEFAULT_WEIGHTS);
  showToast('ok', 'Pesos resetados para o padrão.');
  if (state.resolved && state.open) computeDashboard();
}

/* ─── TEMA ───────────────────────────────────────────────────── */
function initTheme() {
  const saved = localStorage.getItem('advbox_theme') || 'light';
  document.documentElement.setAttribute('data-theme', saved);
}

function toggleTheme() {
  const cur  = document.documentElement.getAttribute('data-theme') || 'light';
  const next = cur === 'light' ? 'dark' : 'light';
  document.documentElement.setAttribute('data-theme', next);
  localStorage.setItem('advbox_theme', next);
  if (state.lastRows.length) renderCharts(state.lastRows);
}

/* ─── TOAST ──────────────────────────────────────────────────── */
function showToast(kind, msg) {
  const area = $('toastArea');
  const toast = document.createElement('div');
  toast.className = `toast toast-${kind}`;
  toast.innerHTML = `<span class="toast-dot"></span>${msg}`;
  area.appendChild(toast);
  setTimeout(() => {
    toast.style.animation = 'fadeOut .3s ease forwards';
    toast.addEventListener('animationend', () => toast.remove());
  }, 3500);
}

/* ─── MOBILE SIDEBAR ─────────────────────────────────────────── */
function openSidebar() {
  $('sidebar').classList.add('open');
  $('sidebarOverlay').classList.add('active');
  document.body.style.overflow = 'hidden';
}

function closeSidebar() {
  $('sidebar').classList.remove('open');
  $('sidebarOverlay').classList.remove('active');
  document.body.style.overflow = '';
}

/* ─── UTILITÁRIOS ────────────────────────────────────────────── */
function pad2(n) { return String(n).padStart(2, '0'); }
function toISODate(d) { return `${d.getFullYear()}-${pad2(d.getMonth()+1)}-${pad2(d.getDate())}`; }
function cssVar(name) { return getComputedStyle(document.documentElement).getPropertyValue(name).trim(); }
function normStr(s) { return (s ?? '').toString().trim().toLowerCase().normalize('NFD').replace(/\p{Diacritic}/gu, ''); }
function normCompromisso(v) { return String(v ?? '').trim().replace(/\s+/g, ' '); }
function uniq(arr) { return [...new Set(arr)]; }

/* ─── PARSE DE DATAS ─────────────────────────────────────────── */
function parseDate(v) {
  if (v == null || v === '') return null;
  if (v instanceof Date) return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  if (typeof v === 'number' && window.XLSX && XLSX.SSF) {
    const p = XLSX.SSF.parse_date_code(v);
    if (p) return new Date(p.y, p.m - 1, p.d);
  }
  const s = String(v).trim();
  const mISO = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (mISO) return new Date(+mISO[1], +mISO[2] - 1, +mISO[3]);
  const mBR = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
  if (mBR) {
    let yy = +mBR[3]; if (yy < 100) yy += 2000;
    return new Date(yy, +mBR[2] - 1, +mBR[1]);
  }
  const iso = new Date(s);
  if (!isNaN(iso)) return new Date(iso.getFullYear(), iso.getMonth(), iso.getDate());
  return null;
}

function fmtBR(d) { return `${pad2(d.getDate())}/${pad2(d.getMonth()+1)}/${d.getFullYear()}`; }
function addDays(d, n) { const x = new Date(d); x.setDate(x.getDate() + n); return x; }

/* ─── JANELAS DE PERÍODO ─────────────────────────────────────── */
function weekSunSatContaining(day) {
  const d = new Date(day.getFullYear(), day.getMonth(), day.getDate());
  const start = addDays(d, -d.getDay());
  return { start, end: addDays(start, 6) };
}

function analyzedPeriod(mode, cutoff) {
  const c = new Date(cutoff.getFullYear(), cutoff.getMonth(), cutoff.getDate());
  if (mode === 'daily') {
    const d = addDays(c, -1);
    return { start: d, end: d };
  }
  if (mode === 'weekly') {
    const w = weekSunSatContaining(c);
    return { start: w.start, end: w.end };
  }
  // monthly: mês anterior
  const prevEnd = new Date(c.getFullYear(), c.getMonth(), 0);
  return {
    start: new Date(prevEnd.getFullYear(), prevEnd.getMonth(), 1),
    end:   new Date(prevEnd.getFullYear(), prevEnd.getMonth() + 1, 0),
  };
}

function remoteWindow(mode, cutoff) {
  const c = new Date(cutoff.getFullYear(), cutoff.getMonth(), cutoff.getDate());
  if (mode === 'daily') return { start: addDays(c, 1), end: null };
  if (mode === 'weekly') {
    const w = weekSunSatContaining(c);
    const start = addDays(w.start, 7);
    return { start, end: addDays(start, 6) };
  }
  return {
    start: new Date(c.getFullYear(), c.getMonth() + 1, 1),
    end:   new Date(c.getFullYear(), c.getMonth() + 2, 0),
  };
}

function modeHelp() {
  const mode = $('mode').value;
  const msgs = {
    daily:   '📅 Diário — Próximas = Data = cutoff | Remotas = cutoff + 1 em diante (sem limite).',
    weekly:  '📅 Semanal — Próximas = Data = cutoff | Remotas = semana subsequente (Dom→Sáb).',
    monthly: '📅 Mensal — Próximas = Data = cutoff | Remotas = mês subsequente.',
  };
  $('modeHelp').textContent = msgs[mode] || '';
}

/* ─── PESQUISA DE COLUNAS ────────────────────────────────────── */
function findColumn(headers, candidates) {
  const hnorm = headers.map(normStr);
  for (const cand of candidates) {
    const c = normStr(cand);
    let idx = hnorm.findIndex(h => h === c);
    if (idx >= 0) return idx;
    idx = hnorm.findIndex(h => h.includes(c));
    if (idx >= 0) return idx;
  }
  return -1;
}

function findDestinatarioIndex(headers) {
  return findColumn(headers, ['destinatario', 'destinatário']);
}

function getCollabIndexSafe(sheetHeaders, fallbackSelectId) {
  const idxDest = findDestinatarioIndex(sheetHeaders);
  if (idxDest >= 0) return { idx: idxDest, used: 'Destinatário' };
  const sel = document.getElementById(fallbackSelectId);
  const idx = sel ? parseInt(sel.value || '-1', 10) : -1;
  return { idx, used: 'Seleção manual' };
}

function findCompromissoIndex(headers) {
  return findColumn(headers, ['compromisso','compromissos','tipo de atividade','atividade','tarefa','descricao','descrição','assunto','título','titulo']);
}

function fallbackTextIndex(headers) {
  const bad = ['data','prazo','prazo fatal','destinatario','destinatário','remetente','colaborador','responsavel','responsável','status','situação','situacao'];
  const hnorm = headers.map(normStr);
  for (let i = 0; i < headers.length; i++) {
    const h = hnorm[i];
    if (!h) continue;
    if (bad.some(b => h.includes(normStr(b)))) continue;
    return i;
  }
  return -1;
}

/* ─── FIX 1: suggestColumn9IfLikely em escopo global ─────────── */
function suggestColumn9IfLikely(headers, data, selectId) {
  const sel = document.getElementById(selectId);
  if (!sel || headers.length < 9) return;
  const idx9 = 8;
  const header = String(headers[idx9] ?? '').trim();
  if (header !== '') return;
  let hits = 0;
  const sample = Math.min(data.length, 80);
  for (let i = 0; i < sample; i++) {
    const v = data[i]?.[headers[idx9]];
    if (String(v ?? '').trim()) hits++;
  }
  if (hits >= Math.max(5, Math.floor(sample * 0.2))) {
    sel.value = String(idx9);
  }
}

/* ─── LEITURA DE PLANILHA ────────────────────────────────────── */
async function readSheet(file) {
  const buf = await file.arrayBuffer();
  const wb  = XLSX.read(buf, { type: 'array', cellDates: true });
  const ws  = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: true });
}

function rowsToObjects(rows) {
  let headerRowIndex = 0;
  for (let i = 0; i < Math.min(rows.length, 15); i++) {
    if ((rows[i] || []).filter(x => String(x).trim() !== '').length >= 3) {
      headerRowIndex = i; break;
    }
  }
  const headers = (rows[headerRowIndex] || []).map(h => String(h).trim());
  const data = [];
  for (let r = headerRowIndex + 1; r < rows.length; r++) {
    const row = rows[r] || [];
    if (row.every(x => String(x).trim() === '')) continue;
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    data.push(obj);
  }
  return { headers, data };
}

function getCell(obj, headers, idx) {
  if (idx == null || idx < 0) return '';
  return obj[headers[idx]];
}

/* ─── SELECTS ────────────────────────────────────────────────── */
function populateSelect(selectId, headers, preferred) {
  const sel = $(selectId);
  sel.innerHTML = '';
  if (!headers.length) {
    sel.innerHTML = '<option value="-1">Nenhuma coluna detectada</option>';
    return;
  }
  headers.forEach((h, i) => {
    const opt = document.createElement('option');
    opt.value = i;
    opt.textContent = h || `(Coluna ${i + 1})`;
    sel.appendChild(opt);
  });
  const idx = findColumn(headers, preferred);
  sel.value = String(idx >= 0 ? idx : 0);
}

/* ─── COLABORADORES ──────────────────────────────────────────── */
function collectNamesFromData() {
  const names = [];
  if (state.resolved) {
    const idx = parseInt($('mapResolvedCollab').value || -1, 10);
    state.resolved.data.forEach(r => {
      const n = String(getCell(r, state.resolved.headers, idx)).trim();
      if (n) names.push(n.toUpperCase());
    });
  }
  if (state.open) {
    const idx = parseInt($('mapOpenCollab').value || -1, 10);
    state.open.data.forEach(r => {
      const n = String(getCell(r, state.open.headers, idx)).trim();
      if (n) names.push(n.toUpperCase());
    });
  }
  return uniq(names).sort();
}

function mostFrequentName() {
  const counts = new Map();
  const add = (name) => {
    const n = String(name || '').trim().toUpperCase();
    if (!n) return;
    counts.set(n, (counts.get(n) || 0) + 1);
  };
  if (state.resolved) {
    const idx = parseInt($('mapResolvedCollab').value || -1, 10);
    state.resolved.data.forEach(r => add(getCell(r, state.resolved.headers, idx)));
  }
  if (state.open) {
    const idx = parseInt($('mapOpenCollab').value || -1, 10);
    state.open.data.forEach(r => add(getCell(r, state.open.headers, idx)));
  }
  let best = null, bestCount = -1;
  for (const [k, v] of counts.entries()) {
    if (v > bestCount) { best = k; bestCount = v; }
  }
  return best;
}

function rebuildSelectedUser() {
  const sel = $('selectedUser');
  const current = sel.value || '__AUTO__';
  const names = collectNamesFromData();
  sel.innerHTML = '<option value="__AUTO__">Auto (mais frequente)</option>';
  names.forEach(n => {
    const opt = document.createElement('option');
    opt.value = n; opt.textContent = n;
    sel.appendChild(opt);
  });
  sel.value = names.includes(current) || current === '__AUTO__' ? current : '__AUTO__';
}

/* ─── CARREGAMENTO DE ARQUIVOS ───────────────────────────────── */
function updateUploadZone(zoneId, iconId, nameId, filename) {
  const zone = $(zoneId);
  const icon = $(iconId);
  const name = $(nameId);
  zone.classList.add('loaded');
  icon.textContent = '✅';
  name.textContent = filename;
}

async function handleResolvedFile() {
  const f = $('fileResolved').files[0];
  if (!f) return;
  try {
    state.resolved = rowsToObjects(await readSheet(f));
    populateSelect('mapResolvedCollab', state.resolved.headers, ['destinatario','destinatário','colaborador','responsavel','responsável','remetente']);
    populateSelect('mapResolvedDone',   state.resolved.headers, ['data conclusao','data conclusão','conclusao','conclusão']);
    populateSelect('mapResolvedDue',    state.resolved.headers, ['data','prazo fatal','prazo','data prazo','término','termino']);
    populateSelect('mapResolvedComp',   state.resolved.headers, ['compromisso']);
    suggestColumn9IfLikely(state.resolved.headers, state.resolved.data, 'mapResolvedComp');
    rebuildSelectedUser();
    updateUploadZone('zoneResolved', 'iconResolved', 'nameResolved', f.name);
    showToast('ok', `Resolvidas: ${state.resolved.data.length} registros carregados.`);
  } catch (err) {
    showToast('err', 'Erro ao carregar planilha de resolvidas: ' + err.message);
  }
}

async function handleOpenFile() {
  const f = $('fileOpen').files[0];
  if (!f) return;
  try {
    state.open = rowsToObjects(await readSheet(f));
    populateSelect('mapOpenCollab', state.open.headers, ['destinatario','destinatário','colaborador','responsavel','responsável','remetente']);
    populateSelect('mapOpenDue',    state.open.headers, ['data','prazo fatal','prazo','data prazo','término','termino']);
    populateSelect('mapOpenComp',   state.open.headers, ['compromisso']);
    suggestColumn9IfLikely(state.open.headers, state.open.data, 'mapOpenComp');
    rebuildSelectedUser();
    updateUploadZone('zoneOpen', 'iconOpen', 'nameOpen', f.name);
    showToast('ok', `Abertas: ${state.open.data.length} registros carregados.`);
  } catch (err) {
    showToast('err', 'Erro ao carregar planilha de abertas: ' + err.message);
  }
}

/* ─── TOP 10 ─────────────────────────────────────────────────── */
function topNFromCounts(counts, n = 10) {
  return [...counts.entries()]
    .filter(([k]) => k && k.trim() !== '')
    .sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]))
    .slice(0, n);
}

function renderTopSection(containerId, entries) {
  const container = document.getElementById(containerId);
  if (!container) return;
  if (!entries.length) {
    container.innerHTML = '<div class="no-data">Sem dados para exibir</div>';
    return;
  }
  const max = entries[0][1] || 1;
  container.innerHTML = entries.map(([label, count], i) => `
    <div class="rank-bar">
      <span class="rank-num">${i + 1}</span>
      <span class="rank-label" title="${label}">${label}</span>
      <span class="rank-count">${count}</span>
    </div>
  `).join('');
}

/* ─── COMPUTE DASHBOARD ──────────────────────────────────────── */
function computeDashboard() {
  if (!state.resolved || !state.open) {
    showToast('err', 'Selecione as duas planilhas antes de processar.');
    return;
  }

  const mode   = $('mode').value;
  const cutoff = parseDate($('refDate').value);
  if (!cutoff) {
    showToast('err', 'Data de referência inválida.');
    return;
  }

  const analyzed = analyzedPeriod(mode, cutoff);
  const remote   = remoteWindow(mode, cutoff);

  const rCollabInfo = getCollabIndexSafe(state.resolved.headers, 'mapResolvedCollab');
  const rCollab = rCollabInfo.idx;
  const rDone   = +$('mapResolvedDone').value;
  const rDue    = +$('mapResolvedDue').value;
  const oCollabInfo = getCollabIndexSafe(state.open.headers, 'mapOpenCollab');
  const oCollab = oCollabInfo.idx;
  const oDue    = +$('mapOpenDue').value;

  let chosen = $('selectedUser').value;
  if (chosen === '__AUTO__') chosen = mostFrequentName();
  if (!chosen) {
    showToast('err', 'Colaborador não detectado. Selecione manualmente.');
    return;
  }

  $('viewChip').textContent = chosen;

  const remoteLabel = (mode === 'daily')
    ? `≥ ${fmtBR(remote.start)} (sem limite)`
    : `${fmtBR(remote.start)} a ${fmtBR(remote.end)}`;

  $('periodLabel').textContent =
    `Cutoff: ${fmtBR(cutoff)} · Concluídas: ${fmtBR(analyzed.start)} → ${fmtBR(analyzed.end)} · Remotas: ${remoteLabel}`;

  const weights = getWeights();
  const s = { done:0, red:0, near:0, remote:0, openTotal:0, redPct:0, score:0,
    diag:{ label:'Monitore', cls:'warn' } };

  // Índices de Compromisso
  let rCompIdx = parseInt(document.getElementById('mapResolvedComp')?.value || '-1', 10);
  let oCompIdx = parseInt(document.getElementById('mapOpenComp')?.value || '-1', 10);
  if (rCompIdx < 0) rCompIdx = findCompromissoIndex(state.resolved.headers);
  if (oCompIdx < 0) oCompIdx = findCompromissoIndex(state.open.headers);
  if (rCompIdx < 0) rCompIdx = fallbackTextIndex(state.resolved.headers);
  if (oCompIdx < 0) oCompIdx = fallbackTextIndex(state.open.headers);

  const countsResolved   = new Map();
  const countsRed        = new Map();
  const countsProgrammed = new Map();

  function inc(map, key) {
    const k = normCompromisso(key);
    if (!k) return;
    map.set(k, (map.get(k) || 0) + 1);
  }

  let resBad = 0, openBad = 0, resMatch = 0, openMatch = 0;

  // ─ Loop resolvidas ─
  for (const row of state.resolved.data) {
    const name = String(getCell(row, state.resolved.headers, rCollab)).trim().toUpperCase();
    if (name !== chosen) continue;
    resMatch++;
    let d = parseDate(getCell(row, state.resolved.headers, rDone));
    if (!d) d = parseDate(getCell(row, state.resolved.headers, rDue));
    if (!d) { resBad++; continue; }
    if (d >= analyzed.start && d <= analyzed.end) {
      s.done++;
      if (rCompIdx >= 0) inc(countsResolved, getCell(row, state.resolved.headers, rCompIdx));
    }
  }

  // ─ Loop abertas ─
  // FIX 2: inc() para countsRed e countsProgrammed adicionados em todas as categorias
  for (const row of state.open.data) {
    const name = String(getCell(row, state.open.headers, oCollab)).trim().toUpperCase();
    if (name !== chosen) continue;
    openMatch++;
    const due = parseDate(getCell(row, state.open.headers, oDue));
    if (!due) { openBad++; continue; }

    if (due < cutoff) {
      s.red++;
      if (oCompIdx >= 0) inc(countsRed, getCell(row, state.open.headers, oCompIdx)); // FIX 2a
    } else if (due.getTime() === cutoff.getTime()) {
      s.near++;
      if (oCompIdx >= 0) inc(countsProgrammed, getCell(row, state.open.headers, oCompIdx)); // FIX 2b
    } else {
      if (mode === 'daily') {
        if (due >= remote.start) {
          s.remote++;
          if (oCompIdx >= 0) inc(countsProgrammed, getCell(row, state.open.headers, oCompIdx)); // FIX 2c
        }
      } else {
        if (due >= remote.start && due <= remote.end) {
          s.remote++;
          if (oCompIdx >= 0) inc(countsProgrammed, getCell(row, state.open.headers, oCompIdx)); // FIX 2c
        }
      }
    }
  }

  // FIX 3: % Vermelho = red / (red + near + remote), não red / (near + remote)
  s.openTotal = s.red + s.near + s.remote;
  s.redPct    = s.openTotal > 0 ? s.red / s.openTotal : 0;
  s.score     = (s.done * weights.wDone) - (s.red * weights.wRed) + (s.near * weights.wNear) + (s.remote * weights.wRemote);

  // Diagnóstico
  if (s.red >= 10)                                       s.diag = { label: 'Gargalo / urgência',   cls: 'bad' };
  else if (s.done === 0 && s.openTotal > 0)              s.diag = { label: 'Sem entrega',           cls: 'bad' };
  else if (s.done > 0 && s.red > 0)                     s.diag = { label: 'Entrega com risco',     cls: 'warn' };
  else if (s.done > 0 && s.red === 0 && !s.remote && s.near > 0) s.diag = { label: 'Carga imediata', cls: 'warn' };
  else if (s.done > 0 && s.red === 0 && s.remote > 0)   s.diag = { label: 'Saudável / previsível', cls: 'good' };
  else if ((s.done + s.openTotal) === 0)                 s.diag = { label: 'Sem dados',             cls: 'warn' };

  // ─ Atualiza KPIs com animação ─
  animateKPI('kDone',   s.done);
  animateKPI('kRed',    s.red);
  animateKPI('kNear',   s.near);
  animateKPI('kRemote', s.remote);

  // ─ Tabela ─
  const diagChip = $('diagChip');
  diagChip.className = `pill ${s.diag.cls}`;
  diagChip.textContent = s.diag.label;

  const tb = $('tbody');
  tb.innerHTML = '';
  const tr = document.createElement('tr');
  tr.innerHTML = `
    <td class="td-name">${chosen}</td>
    <td class="td-num">${s.done}</td>
    <td class="td-num"><span class="pill bad">${s.red}</span></td>
    <td class="td-num">${s.near}</td>
    <td class="td-num">${s.remote}</td>
    <td class="td-num">${(s.redPct * 100).toFixed(1)}%</td>
    <td class="td-num"><strong>${s.score.toFixed(1)}</strong></td>
    <td><span class="pill ${s.diag.cls}">${s.diag.label}</span></td>
  `;
  tb.appendChild(tr);

  state.lastRows = [[chosen, s]];
  renderCharts(state.lastRows);

  // ─ Top 10 ─
  renderTopSection('topResolved',   topNFromCounts(countsResolved, 10));
  renderTopSection('topRed',        topNFromCounts(countsRed, 10));        // FIX 2: agora populado
  renderTopSection('topProgrammed', topNFromCounts(countsProgrammed, 10)); // FIX 2: agora populado

  // ─ Auditoria ─
  $('auditText').innerHTML =
    `Resolvidas: ${resMatch} linhas do colaborador (${resBad} datas inválidas)\n` +
    `Abertas: ${openMatch} linhas (${openBad} datas inválidas)\n` +
    `Janela remota: ${remoteLabel}\n` +
    `Collab via: ${rCollabInfo.used} (Res.) | ${oCollabInfo.used} (Ab.)\n` +
    `Compromisso: col ${rCompIdx+1} "${state.resolved.headers[rCompIdx]??''}" (Res.) | col ${oCompIdx+1} "${state.open.headers[oCompIdx]??''}" (Ab.)`;

  showToast('ok', `Dashboard processado — ${s.done} concluídas | ${s.red} em vermelho.`);
}

/* ─── KPI ANIMAÇÃO ───────────────────────────────────────────── */
function animateKPI(id, target) {
  const el = $(id);
  if (!el) return;
  const start = parseInt(el.textContent) || 0;
  const duration = 500;
  const startTime = performance.now();
  const ease = t => t < .5 ? 2*t*t : -1+(4-2*t)*t;
  function step(now) {
    const p = Math.min((now - startTime) / duration, 1);
    el.textContent = Math.round(start + (target - start) * ease(p));
    if (p < 1) requestAnimationFrame(step);
  }
  requestAnimationFrame(step);
}

/* ─── CHARTS ─────────────────────────────────────────────────── */
function renderCharts(rows) {
  const labels     = rows.map(([n]) => n);
  const doneData   = rows.map(([, s]) => s.done);
  const redData    = rows.map(([, s]) => s.red);
  const nearData   = rows.map(([, s]) => s.near);
  const remoteData = rows.map(([, s]) => s.remote);
  const scoreData  = rows.map(([, s]) => Number(s.score.toFixed(2)));

  if (state.chartCounts) state.chartCounts.destroy();
  if (state.chartScore)  state.chartScore.destroy();

  const cGood    = cssVar('--good');
  const cBad     = cssVar('--red');
  const cWarn    = cssVar('--warn');
  const cAccent  = cssVar('--accent');
  const cPrimary = cssVar('--primary');
  const cText    = cssVar('--text');
  const cGrid    = cssVar('--border');

  const baseChartOpts = {
    responsive: true,
    animation: { duration: 600, easing: 'easeInOutQuart' },
    plugins: {
      legend: { labels: { color: cText, font: { family: "'Plus Jakarta Sans'", size: 11, weight: '700' }, padding: 12 } }
    },
    scales: {
      x: { ticks: { color: cssVar('--text-3'), font: { size: 11 } }, grid: { color: cGrid } },
      y: { ticks: { color: cssVar('--text-3'), font: { size: 11 } }, grid: { color: cGrid }, beginAtZero: true }
    }
  };

  state.chartCounts = new Chart($('chartCounts'), {
    type: 'bar',
    data: {
      labels,
      datasets: [
        { label: 'Concluídas', data: doneData,   backgroundColor: `${cGood}99`,   borderColor: cGood,   borderWidth: 1.5, borderRadius: 6 },
        { label: 'Vermelho',   data: redData,    backgroundColor: `${cBad}99`,    borderColor: cBad,    borderWidth: 1.5, borderRadius: 6 },
        { label: 'Próximas',   data: nearData,   backgroundColor: `${cWarn}99`,   borderColor: cWarn,   borderWidth: 1.5, borderRadius: 6 },
        { label: 'Remotas',    data: remoteData, backgroundColor: `${cAccent}99`, borderColor: cAccent, borderWidth: 1.5, borderRadius: 6 },
      ]
    },
    options: baseChartOpts
  });

  const scoreColor = scoreData[0] >= 0 ? cGood : cBad;
  state.chartScore = new Chart($('chartScore'), {
    type: 'bar',
    data: {
      labels,
      datasets: [{
        label: 'Score',
        data: scoreData,
        backgroundColor: `${scoreColor}99`,
        borderColor: scoreColor,
        borderWidth: 1.5,
        borderRadius: 6
      }]
    },
    options: { ...baseChartOpts, plugins: { ...baseChartOpts.plugins } }
  });
}

/* ─── EXPORTAR CSV ───────────────────────────────────────────── */
function exportCSV() {
  const rows = state.lastRows || [];
  if (!rows.length) { showToast('err', 'Processe o dashboard antes de exportar.'); return; }
  const lines = [['Colaborador','Concluidas','Vermelho','Proximas','Remotas','PctVermelho','Score','Diagnostico'].join(';')];
  rows.forEach(([name, s]) => {
    lines.push([name, s.done, s.red, s.near, s.remote,
      `${(s.redPct * 100).toFixed(1)}%`, s.score.toFixed(1), s.diag.label].join(';'));
  });
  const blob = new Blob([lines.join('\n')], { type: 'text/csv;charset=utf-8' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = 'dashboard_advbox_resumo.csv';
  a.click();
  URL.revokeObjectURL(a.href);
  showToast('ok', 'CSV exportado com sucesso.');
}

/* ─── INIT ───────────────────────────────────────────────────── */
initTheme();
modeHelp();
applyWeightsToUI(loadWeights());
$('refDate').value = toISODate(new Date());

// Tema
$('themeToggle').addEventListener('click', toggleTheme);

// Sidebar mobile
$('menuToggle')?.addEventListener('click', openSidebar);
$('sidebarOverlay')?.addEventListener('click', closeSidebar);

// Arquivos
$('fileResolved').addEventListener('change', async () => { await handleResolvedFile(); });
$('fileOpen').addEventListener('change',     async () => { await handleOpenFile(); });

// Modo e data
$('mode').addEventListener('change', () => { modeHelp(); if (state.resolved && state.open) computeDashboard(); });
$('refDate').addEventListener('change', () => { if (state.resolved && state.open) computeDashboard(); });

// Colaborador
['mapResolvedCollab', 'mapOpenCollab'].forEach(id => {
  const el = $(id);
  if (el) el.addEventListener('change', rebuildSelectedUser);
});
$('selectedUser').addEventListener('change', () => { if (state.resolved && state.open) computeDashboard(); });

// Botões principais
$('runBtn').addEventListener('click', computeDashboard);
$('exportBtn').addEventListener('click', exportCSV);
$('pdfBtn').addEventListener('click', () => window.print());

// Pesos
$('saveWeightsBtn').addEventListener('click', saveWeights);
$('resetWeightsBtn').addEventListener('click', resetWeights);
['wDone','wRed','wNear','wRemote'].forEach(id => {
  const el = $(id);
  if (el) el.addEventListener('change', () => { if (state.resolved && state.open) computeDashboard(); });
});

// Mapeamento avançado
['mapResolvedComp','mapOpenComp'].forEach(id => {
  const el = $(id);
  if (el) el.addEventListener('change', () => { if (state.resolved && state.open) computeDashboard(); });
});

showToast('ok', 'Dashboard pronto. Importe as planilhas para começar.');