// ─────────────────────────────────────────────────────────────────────────────
// SUPERA Dashboard — script.js
// ─────────────────────────────────────────────────────────────────────────────

// ── XLSX library (loaded from CDN via index.html) ─────────────────────────────
// We'll load SheetJS dynamically
(function loadSheetJS() {
  const s = document.createElement('script');
  s.src = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
  document.head.appendChild(s);
})();

// ── STATE ─────────────────────────────────────────────────────────────────────
let SETORES = [], REGIONAIS = {}, DISTRITAIS = {};
let BRASIL_LUC = {}, GR_LUC = {}, GD_LUC = {};
const MESES = ['JAN','FEV','MAR','ABR','MAI','JUN','JUL','AGO','SET','OUT','NOV','DEZ'];
const MESES_PTBR = {JAN:'JANEIRO',FEV:'FEVEREIRO',MAR:'MARÇO',ABR:'ABRIL',MAI:'MAIO',JUN:'JUNHO',JUL:'JULHO',AGO:'AGOSTO',SET:'SETEMBRO',OUT:'OUTUBRO',NOV:'NOVEMBRO',DEZ:'DEZEMBRO'};
let activeMeses = []; // months actually in the data
let DESEMP_BY_MES = {}; // desempenho data keyed by month: { JAN: {code: {...}}, FEV: {...}, ... }

let filtered = [], sortCol = 'classificacao', sortDir = 1, page = 0;
const PS = 100;
let activeTab = '', kpiF = '', viewM = 'ABR', negFilter = false, virouFilter = false, virouPosFilter = false;
let linhaVal = '';

function toggleLinhaDd(e) {
  e.stopPropagation();
  document.getElementById('linha-dd').classList.toggle('open');
}
function setLinha(e, val) {
  e.stopPropagation();
  linhaVal = val;
  const label = document.getElementById('linha-dd-label');
  if (val === 'PRIME') label.innerHTML = '<span class="linha-badge lb-prime">PRIME</span>';
  else if (val === 'INFINITY') label.innerHTML = '<span class="linha-badge lb-infinity">INFINITY</span>';
  else label.textContent = 'Todas';
  document.getElementById('linha-dd').classList.remove('open');
  applyFilters();
}
// Close dropdown when clicking outside
document.addEventListener('click', () => document.getElementById('linha-dd')?.classList.remove('open'));

// Upload state
let pendingLuc = []; // File objects
let pendingDesemp = []; // File objects (one per month)

// ── INIT ──────────────────────────────────────────────────────────────────────
window.addEventListener('DOMContentLoaded', () => {
  // Try to load from localStorage
  try {
    const saved = localStorage.getItem('supera_data');
    if (saved) {
      const parsed = JSON.parse(saved);
      loadDataset(parsed);
      return;
    }
  } catch(e) {}
  showNoData();
});

function showNoData() {
  document.getElementById('no-data').classList.add('show');
  document.getElementById('twrap').style.display = 'none';
}

function loadDataset(parsed) {
  SETORES    = parsed.setores    || [];
  REGIONAIS  = parsed.regionais  || {};
  DISTRITAIS = parsed.distritais || {};
  BRASIL_LUC = parsed.brasil_luc || {};
  GR_LUC     = parsed.gr_luc     || {};
  GD_LUC     = parsed.gd_luc     || {};
  DESEMP_BY_MES = parsed.desemp_by_mes || {};
  activeMeses = parsed.meses     || ['JAN','FEV','MAR','ABR'];
  viewM = activeMeses[activeMeses.length - 1]; // default to latest month

  if (!SETORES.length) { showNoData(); return; }

  // Re-filter to only setores present in the last month (fixes old data with 605)
  const _lastM = activeMeses[activeMeses.length - 1];
  if (_lastM) {
    const _validCodes = new Set(SETORES.filter(s => s['luc_' + _lastM] != null).map(s => s.code));
    if (_validCodes.size > 0 && _validCodes.size < SETORES.length) {
      SETORES = SETORES.filter(s => _validCodes.has(s.code));
    }
  }

  // Enrich
  SETORES.forEach(s => {
    s.regional_nome  = REGIONAIS[s.regional] || s.regional;
    const di = DISTRITAIS[s.distrital];
    s.distrital_nome = typeof di === 'string' ? di : (di?.name || s.distrital);
  });

  // Show table, hide no-data
  document.getElementById('no-data').classList.remove('show');
  document.getElementById('twrap').style.display = '';

  // Populate regional filter
  const fReg = document.getElementById('f-reg');
  fReg.innerHTML = '<option value="">Todas</option>';
  Object.entries(REGIONAIS).forEach(([code, name]) => {
    const o = document.createElement('option');
    o.value = code; o.textContent = name; fReg.appendChild(o);
  });

  // Sync month buttons to activeMeses
  syncMonthButtons();

  applyFilters();
}

function syncMonthButtons() {
  const msel = document.querySelector('.msel');
  // Remove old month buttons
  msel.querySelectorAll('.mb').forEach(b => b.remove());
  // Rebuild
  activeMeses.forEach(m => {
    const b = document.createElement('button');
    b.className = 'mb' + (m === viewM ? ' ma' : '');
    b.textContent = m;
    b.onclick = () => setVM(m);
    msel.appendChild(b);
  });
  document.getElementById('th-mes').textContent = viewM;
  const tgMes = document.getElementById('tg-mes-label');
  if (tgMes) tgMes.textContent = viewM;
}

// ── FILTERS ───────────────────────────────────────────────────────────────────
function onRegChange() {
  const reg = document.getElementById('f-reg').value;
  const fd  = document.getElementById('f-dist');
  fd.innerHTML = '<option value="">Todas</option>';
  const seen = new Set();
  SETORES.forEach(s => {
    if (reg && s.regional !== reg) return;
    if (seen.has(s.distrital)) return;
    seen.add(s.distrital);
    const o = document.createElement('option');
    o.value = s.distrital; o.textContent = s.distrital_nome; fd.appendChild(o);
  });
  applyFilters();
}

function setTab(tab, btn) {
  activeTab = tab;
  negFilter = false; virouFilter = false; virouPosFilter = false;
  document.getElementById('btn-neg')?.classList.remove('active');
  document.getElementById('btn-virou')?.classList.remove('active');
  document.getElementById('btn-virou-pos')?.classList.remove('active');
  document.querySelectorAll('.ctab').forEach(b => {
    if (!['btn-neg','btn-virou','btn-virou-pos'].includes(b.id)) b.className = 'ctab';
  });
  btn.className = 'ctab ' + (!tab ? 't-all' : tab === 'Saudável' ? 't-s' : tab === 'Atenção' ? 't-a' : 't-c');
  applyFilters();
}

function toggleNegFilter(btn) {
  negFilter = !negFilter;
  if (negFilter) { virouFilter = false; virouPosFilter = false; activeTab = ''; }
  btn.classList.toggle('active', negFilter);
  document.getElementById('btn-virou')?.classList.remove('active');
  document.getElementById('btn-virou-pos')?.classList.remove('active');
  document.querySelectorAll('.ctab:not(#btn-neg):not(#btn-virou):not(#btn-virou-pos)').forEach(b => b.className = 'ctab');
  applyFilters();
}

function toggleVirouFilter(btn) {
  virouFilter = !virouFilter;
  if (virouFilter) { negFilter = false; virouPosFilter = false; activeTab = ''; }
  btn.classList.toggle('active', virouFilter);
  document.getElementById('btn-neg')?.classList.remove('active');
  document.getElementById('btn-virou-pos')?.classList.remove('active');
  document.querySelectorAll('.ctab:not(#btn-neg):not(#btn-virou):not(#btn-virou-pos)').forEach(b => b.className = 'ctab');
  applyFilters();
}

function toggleVirouPosFilter(btn) {
  virouPosFilter = !virouPosFilter;
  if (virouPosFilter) { negFilter = false; virouFilter = false; activeTab = ''; }
  btn.classList.toggle('active', virouPosFilter);
  document.getElementById('btn-neg')?.classList.remove('active');
  document.getElementById('btn-virou')?.classList.remove('active');
  document.querySelectorAll('.ctab:not(#btn-neg):not(#btn-virou):not(#btn-virou-pos)').forEach(b => b.className = 'ctab');
  applyFilters();
}

function setVM(m) {
  viewM = m;
  document.querySelectorAll('.mb').forEach(b => b.classList.toggle('ma', b.textContent === m));
  document.getElementById('th-mes').textContent = m;
  // Update group header label too
  const tgMes = document.getElementById('tg-mes-label');
  if (tgMes) tgMes.textContent = m;
  // Inject desempenho data for the selected month into SETORES
  if (DESEMP_BY_MES && DESEMP_BY_MES[m]) {
    const dp = DESEMP_BY_MES[m];
    SETORES.forEach(s => {
      const d = dp[s.code] || {};
      // Clear existing desemp fields first
      ['ytd','abrabr','abramar'].forEach(prefix => {
        ['merc_ant','merc_atual','merc_var','sup_ant','sup_atual','sup_var','share_ant','share_atual','share_var'].forEach(f => {
          s[`${prefix}_${f}`] = null;
        });
      });
      Object.assign(s, d);
      // venda_media will be recalculated below
    });
    // Also update venda_media for the selected month
    const nMesesMap = {JAN:1,FEV:2,MAR:3,ABR:4,MAI:5,JUN:6,JUL:7,AGO:8,SET:9,OUT:10,NOV:11,DEZ:12};
    const nMeses = nMesesMap[m] || 1;
    SETORES.forEach(s => {
      const d = dp[s.code] || {};
      s.venda_media = d.ytd_sup_atual != null ? d.ytd_sup_atual / nMeses : null;
    });
  }
  applyFilters();
}

// KPI click — no scroll, only data update
function kpiClick(f, id) {
  if (kpiF === f) { clearKpi(); return; }
  kpiF = f;
  document.querySelectorAll('.kpi').forEach(k => k.classList.remove('ka'));
  document.getElementById(id).classList.add('ka');
  const labels = {
    'Saudável'        : '● Saudável',
    'Atenção'         : '● Atenção',
    'ALERTA'          : '● ALERTA',
    'Crítico'         : '● Crítico',
  };
  document.getElementById('afb-lbl').textContent = labels[f] || f;
  document.getElementById('afb').classList.add('show');
  applyFilters();
  setTimeout(updateStickyTop, 50); // just update data, no scroll
}

function clearKpi() {
  kpiF = '';
  document.querySelectorAll('.kpi').forEach(k => k.classList.remove('ka'));
  document.getElementById('afb').classList.remove('show');
  applyFilters();
  setTimeout(updateStickyTop, 50);
}

function getBase() {
  const reg   = document.getElementById('f-reg').value;
  const dist  = document.getElementById('f-dist').value;
  const neg   = parseInt(document.getElementById('f-neg').value) || 0;
  const share = document.getElementById('f-share').value;
  const srch  = document.getElementById('f-search').value.toLowerCase().trim();
  const prevM = (() => {
    const idx = activeMeses.indexOf(viewM);
    return idx > 0 ? activeMeses[idx - 1] : null;
  })();
  return SETORES.filter(s => {
    if (activeTab && s.classificacao !== activeTab) return false;
    if (negFilter && (s['luc_' + viewM] == null || s['luc_' + viewM] >= 0)) return false;
    if (linhaVal && s.linha !== linhaVal) return false;
    if (virouFilter) {
      const cur = s['luc_' + viewM], prev = prevM ? s['luc_' + prevM] : null;
      if (cur == null || prev == null || !(prev >= 0 && cur < 0)) return false;
    }
    if (virouPosFilter) {
      const cur = s['luc_' + viewM], prev = prevM ? s['luc_' + prevM] : null;
      if (cur == null || prev == null || !(prev < 0 && cur >= 0)) return false;
    }
    if (reg && s.regional !== reg) return false;
    if (dist && s.distrital !== dist) return false;
    if (neg === 4 && s.n_neg < 4) return false;
    else if (neg && neg < 4 && s.n_neg < neg) return false;
    if (share === 'beat' && !(s.ytd_share_var > 0)) return false;
    if (share === 'miss' && !(s.ytd_share_var < 0)) return false;
    if (srch && !s.nome.toLowerCase().includes(srch) && !s.code.includes(srch) &&
        !(s.regional_nome||'').toLowerCase().includes(srch) &&
        !(s.distrital_nome||'').toLowerCase().includes(srch)) return false;
    return true;
  });
}

function applyFilters() {
  const base = getBase();
  filtered = base.filter(s => {
    if (kpiF && s.sub_class !== kpiF && !(kpiF === 'Saudável' && s.sub_class === 'Excelente')) return false;
    return true;
  });
  filtered.sort((a, b) => {
    let va = a[sortCol], vb = b[sortCol];
    if (sortCol === 'hier') { va = a.regional_nome||''; vb = b.regional_nome||''; }
    if (va == null) va = sortDir > 0 ? Infinity : -Infinity;
    if (vb == null) vb = sortDir > 0 ? Infinity : -Infinity;
    if (typeof va === 'string') return va.localeCompare(vb) * sortDir;
    return (va - vb) * sortDir;
  });
  page = 0;
  updateHdr(base);
  updateKPIs(base);
  renderTable();
}

function sortBy(col) {
  if (sortCol === col) sortDir *= -1;
  else { sortCol = col; sortDir = col === 'nome' ? 1 : -1; }
  applyFilters();
}

function updateHdr(base) {
  document.getElementById('h-tot').textContent = base.length;
  document.getElementById('h-s').textContent   = base.filter(x => x.sub_class === 'Saudável' || x.sub_class === 'Excelente').length;
  document.getElementById('h-est').textContent = 0;
  document.getElementById('h-a').textContent   = base.filter(x => x.sub_class === 'Atenção' || x.sub_class === 'Atenção ao Mercado').length;
  document.getElementById('h-rec').textContent = 0;
  document.getElementById('h-alr').textContent = base.filter(x => x.sub_class === 'ALERTA').length;
  document.getElementById('h-c').textContent   = base.filter(x => x.sub_class === 'Crítico').length;
}

function updateKPIs(base) {
  if (!base.length) return;
  const lastM = activeMeses[activeMeses.length - 1];
  document.getElementById('kv-s').textContent   = base.filter(x => x.sub_class === 'Saudável' || x.sub_class === 'Excelente').length;
  document.getElementById('kv-est').textContent = 0;
  document.getElementById('kv-a').textContent   = base.filter(x => x.sub_class === 'Atenção' || x.sub_class === 'Atenção ao Mercado').length;
  document.getElementById('kv-rec').textContent = 0;
  document.getElementById('kv-alr').textContent = base.filter(x => x.sub_class === 'ALERTA').length;
  document.getElementById('kv-c').textContent   = base.filter(x => x.sub_class === 'Crítico').length;

  // Luc consolidada: busca no nível do filtro ativo (GD > GR > Brasil)
  const reg  = document.getElementById('f-reg').value;
  const dist = document.getElementById('f-dist').value;
  let consolidadaLuc = null;
  if (dist && GD_LUC && GD_LUC[lastM]) {
    consolidadaLuc = GD_LUC[lastM][dist] ?? null;
  } else if (reg && GR_LUC && GR_LUC[lastM]) {
    consolidadaLuc = GR_LUC[lastM][reg] ?? null;
  } else if (BRASIL_LUC) {
    consolidadaLuc = BRASIL_LUC[lastM] ?? null;
  }
  const me = document.getElementById('kv-med');
  me.textContent = fp(consolidadaLuc);
  me.style.color = (consolidadaLuc || 0) >= 0 ? 'var(--green)' : 'var(--red)';
  // Update label to reflect current filter level
  const lbl = document.getElementById('lbl-consolidada');
  if (lbl) {
    if (dist) {
      const distName = (DISTRITAIS[dist] || dist).split(' ')[0];
      lbl.textContent = 'Luc. ' + distName + ' ' + lastM;
    } else if (reg) {
      const regName = (REGIONAIS[reg] || reg).split(' ')[0];
      lbl.textContent = 'Luc. ' + regName + ' ' + lastM;
    } else {
      lbl.textContent = 'Luc. Brasil ' + lastM;
    }
  }
}

// ── TABLE ─────────────────────────────────────────────────────────────────────
function renderTable() {
  const tbody = document.getElementById('tbody');
  const slice = filtered.slice(page * PS, (page + 1) * PS);
  document.getElementById('tcnt').textContent = `${filtered.length} setores`;
  if (!filtered.length) {
    tbody.innerHTML = `<tr><td colspan="11"><div class="empty">📊 Nenhum setor encontrado.</div></td></tr>`;
    renderPag(); return;
  }

  tbody.innerHTML = slice.map(s => {
    // Badge com os 5 status usando sub_class
    const sub = s.sub_class || s.classificacao;
    const badgeCfg = {
      'Saudável':          { cls: 's',    label: 'Saudável' },
      'Excelente':         { cls: 's',    label: 'Saudável' },
      'Atenção':           { cls: 'a',    label: 'Atenção' },
      'Atenção ao Mercado':{ cls: 'a',    label: 'Atenção' },
      'ALERTA':            { cls: 'alr',  label: 'ALERTA' },
      'Crítico':           { cls: 'c',    label: 'Crítico' },
    };
    const bc = badgeCfg[sub] || badgeCfg[s.classificacao] || { cls: 'c', label: s.classificacao };

    // Dados para o tooltip — guardados em data-attributes para evitar problema de overflow
    const folga = s.folga_pe != null ? s.folga_pe.toFixed(1) + '%' : '—';
    const folgaNum = s.folga_pe != null ? s.folga_pe : null;
    const supVar  = s.ytd_sup_var  != null ? (s.ytd_sup_var  * 100).toFixed(2) + '%' : '—';
    const mercVar = s.ytd_merc_var != null ? (s.ytd_merc_var * 100).toFixed(2) + '%' : '—';
    const acimaMerc = s.ytd_sup_var != null && s.ytd_merc_var != null && s.ytd_sup_var >= s.ytd_merc_var;

    const motivoFin = folgaNum == null ? 'Dados de Ponto de Equilíbrio indisponíveis.' :
      folgaNum >= 8  ? `Folga de ${folga} acima do Ponto de Equilíbrio — colchão financeiro confortável.` :
      folgaNum >= 4  ? `Folga de ${folga} acima do Ponto de Equilíbrio — zona intermediária, próximo do limite.` :
      folgaNum >= 0  ? `Folga de ${folga} acima do Ponto de Equilíbrio — muito próximo do limite.` :
                       `Folga de ${folga} — abaixo do Ponto de Equilíbrio, operando no prejuízo.`;

    const motivoMerc = acimaMerc
      ? `Supera (${supVar}) crescendo acima do mercado (${mercVar}) ✔`
      : `Supera (${supVar}) crescendo abaixo do mercado (${mercVar}) ✖`;

    const badge = `<span class="bc ${bc.cls} badge-wrap"
      data-tip-title="${bc.label}"
      data-tip-fin="${motivoFin.replace(/"/g, '&quot;')}"
      data-tip-merc="${motivoMerc.replace(/"/g, '&quot;')}"
      onmouseenter="showBadgeTip(event,this)"
      onmouseleave="hideBadgeTip()"
    ><span class="d6 ${bc.cls}"></span>${bc.label}</span>`;

    const lv    = s['luc_' + viewM];
    const lvS   = (lv != null && Math.abs(lv) < 2) ? lv : null;
    const lvC   = (lv || 0) >= 0 ? 'pos' : 'neg';

    const dots  = activeMeses.map(m => {
      const v = s['luc_' + m];
      const dc = v == null ? 'na' : v < 0 ? 'neg' : 'pos';
      return `<div class="mwrap${m === viewM ? ' sel' : ''}" onclick="setVM('${m}')" title="${m}: ${v != null ? fp(v) : 'sem dado'}">
        <div class="md ${dc}"></div><div class="ml">${m}</div>
      </div>`;
    }).join('');

    const vm    = s.venda_media;
    const vHtml = vm != null ? `<span class="vbrl">${fBRL(vm)}</span>` : '<span style="color:var(--muted)">—</span>';

    const vp = v => {
      if (v == null) return '<span style="color:var(--muted)">—</span>';
      const c = Math.abs(v) < 0.005 ? 'fl' : v > 0 ? 'up' : 'dn';
      return `<span class="vp ${c}">${c === 'fl' ? '→' : v > 0 ? '▲' : '▼'} ${fp(v)}</span>`;
    };

    // ⭐ star indicator for neg but gaining share
    const sv2 = s.abrabr_share_var, spv = s.abrabr_sup_var, mv2 = s.abrabr_merc_var;
    const hasStar = s['luc_' + viewM] < 0 && (sv2||0) > 0 && spv != null && mv2 != null && spv > mv2;

    return `<tr>
      <td class="c-setor"><div class="sn">${s.nome}</div><div class="sc">${s.code}${s.linha ? `<span class="linha-badge lb-${s.linha.toLowerCase()}">${s.linha}</span>` : ''}</div></td>
      <td class="c-hier"><div class="hier-r">${s.regional_nome}</div><div class="hier-d">${s.distrital_nome}</div></td>
      <td class="c-class">${badge}</td>
      <td class="c-luc"><span class="lp ${lvC}">${lvS != null ? fp(lvS) : '—'}</span><div class="mw">${dots}</div></td>
      <td class="c-vm sup-col">${vHtml}</td>
      <td class="c-sytd sup-col gs-ytd">${vp(s.ytd_sup_var)}</td>
      <td class="c-mytd">${vp(s.ytd_merc_var)}</td>
      <td class="c-shytd sh-col">${fShare(s.ytd_share_var)}</td>
      <td class="c-sabr mes-col gs-mes">${vp(s.abrabr_sup_var)}</td>
      <td class="c-mabr mes-col">${vp(s.abrabr_merc_var)}</td>
      <td class="c-shabr sh-col mes-col">${fShare(s.abrabr_share_var)}${hasStar ? '<span style="font-size:11px"> ⭐</span>' : ''}</td>
    </tr>`;
  }).join('');
  renderPag();
  // Rebuild clone when table re-renders
  theadCloneBuilt = false;
  setTimeout(updateStickyTop, 50);
}

function renderPag() {
  const tp  = Math.ceil(filtered.length / PS);
  const pag = document.getElementById('pag');
  if (!filtered.length) { pag.innerHTML = ''; return; }
  const s = page * PS + 1, e = Math.min((page + 1) * PS, filtered.length);
  let pages = '';
  for (let i = 0; i < Math.min(tp, 7); i++)
    pages += `<button class="pb${i === page ? ' pba' : ''}" onclick="goPage(${i})">${i + 1}</button>`;
  if (tp > 7) pages += `<span style="color:var(--muted);padding:0 3px;font-size:10px">…</span>`;
  pag.innerHTML = `<span class="pi">Mostrando ${s}–${e} de ${filtered.length}</span>
    <div class="pbs">
      <button class="pb" onclick="goPage(${page - 1})" ${page === 0 ? 'disabled' : ''}>‹</button>
      ${pages}
      <button class="pb" onclick="goPage(${page + 1})" ${page >= tp - 1 ? 'disabled' : ''}>›</button>
    </div>`;
}

function goPage(p) {
  const tp = Math.ceil(filtered.length / PS);
  if (p < 0 || p >= tp) return;
  page = p; renderTable();
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

// ── STICKY THEAD TOP — recalculate dynamically ────────────────────────────────
function updateStickyTop() {
  const hdr  = document.querySelector('.hdr');
  const fbar = document.querySelector('.fbar');
  const kpis = document.querySelector('.kpis-wrap');
  if (!hdr || !fbar) return;

  const hdrH  = hdr.offsetHeight;
  const fbarH = fbar.offsetHeight;
  const kpisH = kpis ? kpis.offsetHeight : 0;
  const theadTop = hdrH + fbarH + kpisH;

  // Update sticky tops
  fbar.style.top = hdrH + 'px';
  if (kpis) kpis.style.top = (hdrH + fbarH) + 'px';

  // Thead rows stick below header + filters + kpis
  document.querySelectorAll('thead tr.thead-groups th').forEach(th => {
    th.style.top = theadTop + 'px';
  });
  // Get height of groups row for second row offset
  const groupsRow = document.querySelector('thead tr.thead-groups');
  const groupsH = groupsRow ? groupsRow.offsetHeight : 26;
  document.querySelectorAll('thead tr:not(.thead-groups) th').forEach(th => {
    th.style.top = (theadTop + groupsH) + 'px';
  });
}
window.addEventListener('load', updateStickyTop);
window.addEventListener('resize', updateStickyTop);


// ── UPLOAD PANEL ──────────────────────────────────────────────────────────────
function openUpload() {
  pendingLuc = [];
  pendingDesemp = [];
  renderFileList('luc');
  renderFileList('desemp');
  document.getElementById('upload-overlay').classList.add('show');
  document.getElementById('upload-progress').classList.remove('show');
  document.getElementById('btn-process').disabled = true;
}

function closeUpload() {
  document.getElementById('upload-overlay').classList.remove('show');
}

function closeUploadIfBg(e) {
  if (e.target === document.getElementById('upload-overlay')) closeUpload();
}

function dragOver(e, id) {
  e.preventDefault();
  document.getElementById(id).classList.add('drag');
}
function dragLeave(id) {
  document.getElementById(id).classList.remove('drag');
}
function dropFiles(e, type) {
  e.preventDefault();
  dragLeave(type + '-drop');
  addFiles(e.dataTransfer.files, type);
}

function addFiles(files, type) {
  const arr = Array.from(files).filter(f => f.name.toLowerCase().endsWith('.xlsx'));
  if (type === 'luc') {
    arr.forEach(f => {
      if (!pendingLuc.find(x => x.name === f.name)) pendingLuc.push(f);
    });
  } else {
    arr.forEach(f => {
      if (!pendingDesemp.find(x => x.name === f.name)) pendingDesemp.push(f);
    });
  }
  renderFileList(type);
  updateProcessBtn();
}

function renderFileList(type) {
  const list  = document.getElementById(type + '-list');
  const items = type === 'luc' ? pendingLuc : pendingDesemp;

  if (!items.length) {
    // Show currently saved files info
    try {
      const saved = localStorage.getItem('supera_data');
      if (saved) {
        const p = JSON.parse(saved);
        const savedFiles = type === 'luc' ? (p.sourceFiles?.luc || []) : (p.sourceFiles?.desemp || []);
        if (savedFiles.length) {
          list.innerHTML = savedFiles.map(n => `
            <div class="file-item ok">
              <svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="2"><path d="M13 5L6 12l-3-3"/></svg>
              <span class="fname">${n}</span>
              <span class="ftype">Salvo</span>
            </div>`).join('');
          return;
        }
      }
    } catch(e) {}
    list.innerHTML = '<div style="color:var(--muted);font-size:11px;padding:4px 0">Nenhum arquivo</div>';
    return;
  }

  list.innerHTML = items.map((f, i) => `
    <div class="file-item">
      <svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.5"><rect x="2" y="1" width="10" height="14" rx="1.5"/><path d="M5 5h6M5 8h6M5 11h4"/></svg>
      <span class="fname">${f.name}</span>
      <span class="ftype">XLSX</span>
      <span onclick="${type==='luc'?`removeLuc(${i})`:`removeDesemp(${i})`}" style="cursor:pointer;color:var(--muted);font-size:11px;padding:0 4px">✕</span>
    </div>`).join('');
}

function removeLuc(i)        { pendingLuc.splice(i, 1); renderFileList('luc'); updateProcessBtn(); }
function removeDesemp(i)     { pendingDesemp.splice(i, 1); renderFileList('desemp'); updateProcessBtn(); }

function updateProcessBtn() {
  document.getElementById('btn-process').disabled = !(pendingLuc.length > 0 && pendingDesemp.length > 0);
}

// ── XLSX PROCESSING ───────────────────────────────────────────────────────────
async function processFiles() {
  if (!window.XLSX) { alert('Aguarde o carregamento da biblioteca XLSX...'); return; }

  const prog = document.getElementById('upload-progress');
  const msg  = document.getElementById('progress-msg');
  prog.classList.add('show');
  document.getElementById('btn-process').disabled = true;

  try {
    msg.textContent = 'Lendo arquivos de lucratividade…';
    await sleep(50);

    // ── 1. Read all lucratividade files ──────────────────────────────────────
    const lucData = {}; // { 'JAN': { code: {luc, lucro, sup_ytd, mkt_ytd, nome, reg, dist} } }
    const regMap = {}, distMap = {};
    const brasilLucMap = {}, grLucMap = {}, gdLucMap = {}; // consolidated luc by level
    const fileNames = { luc: pendingLuc.map(f => f.name), desemp: pendingDesemp.map(f => f.name) };

    for (const file of pendingLuc) {
      msg.textContent = `Processando ${file.name}…`;
      await sleep(30);
      const wb   = XLSX.read(await file.arrayBuffer(), { type: 'array' });
      const ws   = wb.Sheets['LUCRATIVIDADE'];
      if (!ws) { alert(`Aba "LUCRATIVIDADE" não encontrada em ${file.name}`); continue; }
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });

      // Detect month from header row
      const header = rows[0].map(c => String(c || ''));
      const supIdx  = header.findIndex(c => c.includes('SUPERA R$'));
      const mktIdx  = header.findIndex(c => c.includes('MKT SH'));
      const lucIdx  = header.findIndex(c => c.includes('LUCRATIVIDADE'));
      const lucRIdx = header.findIndex(c => c.includes('LUCRO'));
      const vendIdx = header.findIndex(c => c.includes('VENDA'));
      const dfIdx   = header.findIndex(c => c.includes('DESPESA FIXA'));
      const dvIdx   = header.findIndex(c => c.includes('DESPESA VARIÁVEL'));

      // Detect mes abbreviation — tenta pelo cabeçalho da coluna, fallback pelo nome do arquivo
      const supHeader = header[supIdx] || '';
      let mes = detectMes(supHeader) || detectMes(file.name);
      if (!mes) { console.warn('Mês não detectado em', file.name, supHeader); continue; }

      lucData[mes] = {};
      if (!grLucMap[mes]) grLucMap[mes] = {};
      if (!gdLucMap[mes]) gdLucMap[mes] = {};

      for (let r = 1; r < rows.length; r++) {
        const row  = rows[r];
        const cargo= String(row[2] || '').trim();

        // Extract consolidated luc for BRASIL / GR / GD
        if (cargo === 'BRASIL') {
          brasilLucMap[mes] = safeNum(row[lucIdx]);
        } else if (cargo === 'GR') {
          const rcode = String(row[0] || '').replace('.0','').trim().padStart(6,'0');
          grLucMap[mes][rcode] = safeNum(row[lucIdx]);
          regMap[rcode] = String(row[4] || '').replace(/^\d+\s*-\s*/, '').trim();
        } else if (cargo === 'GD') {
          const dcode = String(row[1] || '').replace('.0','').trim().padStart(6,'0');
          gdLucMap[mes][dcode] = safeNum(row[lucIdx]);
          distMap[dcode] = String(row[4] || '').replace(/^\d+\s*-\s*/, '').trim();
        }

        if (cargo !== 'PV') continue;
        const fdv  = String(row[4] || '').trim();
        const code = fdv.match(/^(\d+)/)?.[1];
        if (!code) continue;
        const nome = fdv.replace(/^\d+\s*-\s*/, '').trim();
        const reg  = String(row[0] || '').trim();
        const dist = String(row[1] || '').trim();

        // Build maps from GR/GD rows for names
        if (cargo === 'GR') {
          const rcode = String(row[0] || '').replace('.0','').trim().padStart(6,'0');
          regMap[rcode] = fdv.replace(/^\d+\s*-\s*/, '').trim();
        }
        if (cargo === 'GD') {
          const dcode = String(row[1] || '').replace('.0','').trim().padStart(6,'0');
          distMap[dcode] = fdv.replace(/^\d+\s*-\s*/, '').trim();
        }

        const linha = String(row[3] || '').trim().toUpperCase();
        lucData[mes][code] = {
          nome, reg: reg.replace('.0','').padStart(6,'0'),
          dist: dist.replace('.0','').padStart(6,'0'),
          linha: linha || null,
          luc:       safeNum(row[lucIdx]),
          lucro:     safeNum(row[lucRIdx]),
          sup_ytd:   safeNum(row[supIdx]),
          mkt_ytd:   safeNum(row[mktIdx]),
          venda:     safeNum(row[vendIdx]),
          desp_fixa: safeNum(row[dfIdx]),
          desp_var:  safeNum(row[dvIdx]),
        };
      }
    }

    // GR/GD names and consolidated luc already extracted inline above

    // ── 2. Read desempenho files (one per month) ──────────────────────────────
    msg.textContent = 'Processando planilhas de desempenho…';
    await sleep(50);

    const desempByMes = {}; // { 'JAN': { code: {...} }, 'FEV': {...}, ... }
    let lastDesempMes = null;
    let nMeses = 1;

    const MESES_ABR = {JAN:1,FEV:2,MAR:3,ABR:4,MAI:5,JUN:6,JUL:7,AGO:8,SET:9,OUT:10,NOV:11,DEZ:12};

    for (const desempFile of pendingDesemp) {
      msg.textContent = `Processando ${desempFile.name}…`;
      await sleep(30);

      const wb2    = XLSX.read(await desempFile.arrayBuffer(), { type: 'array' });
      const desemp = {};
      const allSheets = wb2.SheetNames;
      // Detect sheet names dynamically
      const ytdSheet   = allSheets.find(s => s === 'YTD') || null;
      const mesXmesSheet = allSheets.find(s => {
        const p = s.split(' X ');
        return p.length === 2 && p[0].trim() === p[1].trim() && p[0].trim().length === 3;
      }) || null;
      const mesXprevSheet = allSheets.find(s => {
        const p = s.split(' X ');
        return p.length === 2 && p[0].trim() !== p[1].trim() && p[0].trim().length === 3 && p[1].trim().length === 3;
      }) || null;

      // Detect month from mesXmesSheet or file name
      const currentMesAbbr = mesXmesSheet
        ? mesXmesSheet.split(' X ')[0].trim()
        : detectMes(desempFile.name);
      const fileMesN = currentMesAbbr ? (MESES_ABR[currentMesAbbr] || 1) : 1;

      const sheetMap = [
        [ytdSheet,      'ytd'],
        [mesXmesSheet,  'abrabr'],
        [mesXprevSheet, 'abramar'],
      ];
      for (const [sheet, prefix] of sheetMap) {
        const ws2 = sheet ? wb2.Sheets[sheet] : null;
        if (!ws2) continue;
        const rows2 = XLSX.utils.sheet_to_json(ws2, { header: 1 });
        let hRow = 0;
        for (let r = 0; r < rows2.length; r++) {
          if (rows2[r].includes('CARGO')) { hRow = r; break; }
        }
        const h = rows2[hRow].map(c => String(c || ''));
        const nomI  = h.indexOf('NOME');
        const carI  = h.indexOf('CARGO');
        const maNI  = h.indexOf('MERC. ANT');
        const maCI  = h.indexOf('MERC. ATUAL');
        const mvI   = h.findIndex((c,i)=>c==='VARIAÇÃO %'&&i>0);
        const saNI  = h.indexOf('SUP. ANT');
        const saCI  = h.indexOf('SUP. ATUAL');
        const svI   = h.findIndex((c,i)=>c.startsWith('VARIAÇÃO')&&i>saNI);
        const shaNI = h.indexOf('SHARE ANT.');
        const shaCI = h.indexOf('SHARE ATUAL');
        // Share var: always the column immediately after SHARE ATUAL
        const shvI  = shaCI >= 0 ? shaCI + 1 : -1;

        for (let r = hRow + 1; r < rows2.length; r++) {
          const row = rows2[r];
          if (String(row[carI]||'').trim() !== 'PV') continue;
          const nom = String(row[nomI]||'').trim();
          const code = nom.match(/^(\d+)/)?.[1];
          if (!code) continue;
          if (!desemp[code]) desemp[code] = {};
          desemp[code][`${prefix}_merc_ant`]   = safeNum(row[maNI]);
          desemp[code][`${prefix}_merc_atual`] = safeNum(row[maCI]);
          desemp[code][`${prefix}_merc_var`]   = safeNum(row[mvI]);
          desemp[code][`${prefix}_sup_ant`]    = safeNum(row[saNI]);
          desemp[code][`${prefix}_sup_atual`]  = safeNum(row[saCI]);
          desemp[code][`${prefix}_sup_var`]    = safeNum(row[svI]);
          desemp[code][`${prefix}_share_ant`]  = safeNum(row[shaNI]);
          desemp[code][`${prefix}_share_atual`]= safeNum(row[shaCI]);
          desemp[code][`${prefix}_share_var`]  = safeNum(row[shvI]);
        }
      }

      // Store by detected month
      if (currentMesAbbr) {
        desempByMes[currentMesAbbr] = desemp;
        // Track latest desempenho month
        if (!lastDesempMes || fileMesN > (MESES_ABR[lastDesempMes] || 0)) {
          lastDesempMes = currentMesAbbr;
          nMeses = fileMesN;
        }
      }
    }

    // Use latest desempenho file as the base for current setores data
    const desemp = desempByMes[lastDesempMes] || {};

    // ── 3. Merge into setores ────────────────────────────────────────────────
    msg.textContent = 'Calculando lucratividade e classificações…';
    await sleep(50);

    const sortedMeses = MESES.filter(m => lucData[m]);
    // Usar apenas códigos do mês mais recente (evita setores desativados inflarem o total)
    const lastMes = sortedMeses[sortedMeses.length - 1];
    const allCodes = new Set(Object.keys(lucData[lastMes] || {}));

    const setores = [];
    for (const code of allCodes) {
      // Get nome/reg/dist from any month
      let nome = '', reg = '', dist = '';
      for (const m of sortedMeses) {
        if (lucData[m][code]) { nome = lucData[m][code].nome; reg = lucData[m][code].reg; dist = lucData[m][code].dist; break; }
      }
      let linha = '';
      for (const m of sortedMeses) {
        if (lucData[m][code]?.linha) { linha = lucData[m][code].linha; break; }
      }
      const s = { code, nome, regional: reg, distrital: dist, linha };

      // Per-month luc values
      let prevYtd = 0;
      for (let i = 0; i < sortedMeses.length; i++) {
        const m   = sortedMeses[i];
        const d   = lucData[m][code];
        s['luc_' + m]   = d?.luc   ?? null;
        s['lucro_' + m] = d?.lucro ?? null;

        // Venda mensal: JAN = ytd, resto = diferença
        const ytd = d?.sup_ytd ?? null;
        if (i === 0) {
          s['venda_mes_' + m] = ytd;
          prevYtd = ytd || 0;
        } else {
          s['venda_mes_' + m] = ytd != null ? ytd - prevYtd : null;
          prevYtd = ytd || prevYtd;
        }
      }

      // Venda média = YTD Supera (fonte oficial desempenho) ÷ número de meses
      const dp = desemp[code] || {};
      const lastM = sortedMeses[sortedMeses.length - 1];
      s.venda_media = dp.ytd_sup_atual != null ? dp.ytd_sup_atual / nMeses : null;

      // Desempenho fields
      Object.assign(s, dp);

      // n_neg
      s.n_neg = sortedMeses.filter(m => (s['luc_'+m] || 0) < 0).length;

      // ── PONTO DE EQUILÍBRIO (YTD do mês mais recente) ───────────────────────
      // BE = Despesa Fixa / (1 - Despesa Variável / Venda Líquida)
      const d_lastM = lucData[lastM][code];
      const _df  = d_lastM?.desp_fixa ?? null;
      const _dv  = d_lastM?.desp_var  ?? null;
      const _vl  = d_lastM?.venda     ?? null;
      let folga_pe = null; // % acima/abaixo do BE
      if (_df != null && _dv != null && _vl != null && _vl > 0) {
        const mc_pct = 1 - (_dv / _vl);          // margem de contribuição %
        const be_ytd = mc_pct > 0 ? _df / mc_pct : null;
        if (be_ytd != null && be_ytd > 0) {
          folga_pe = ((_vl - be_ytd) / be_ytd) * 100; // % folga acima do BE
        }
      }
      s.folga_pe = folga_pe; // expõe para uso futuro

      // ── EIXO 1: Saúde Financeira (Ponto de Equilíbrio) ──────────────────────
      // Três faixas de folga acima do BE:
      // saudavel_fin:     folga >= 8%
      // intermediario_fin: folga 4% a 7,9%
      // fragil_fin:       folga < 4% (inclui negativos)
      const PE_SAUDAVEL      = 8;
      const PE_INTERMEDIARIO = 4;
      let saudavel_fin, intermediario_fin, fragil_fin;
      if (folga_pe != null) {
        saudavel_fin      = folga_pe >= PE_SAUDAVEL;
        intermediario_fin = folga_pe >= PE_INTERMEDIARIO && folga_pe < PE_SAUDAVEL;
        fragil_fin        = folga_pe < PE_INTERMEDIARIO;
      } else {
        // fallback para % de lucratividade
        const luc_fb = s['luc_' + lastM] || 0;
        saudavel_fin      = luc_fb >= 0.05;
        intermediario_fin = luc_fb >= 0.02 && luc_fb < 0.05;
        fragil_fin        = luc_fb < 0.02;
      }

      // ── EIXO 2: Desempenho vs Mercado ────────────────────────────────────────
      const sup_var   = s.ytd_sup_var  || 0;
      const merc_var  = s.ytd_merc_var || 0;
      const share_var = s.ytd_share_var || 0;
      const acima_merc   = sup_var >= merc_var; // crescendo igual ou acima
      const ganhou_share = share_var > 0;

      // ── CLASSIFICAÇÃO FINAL (6 status) ───────────────────────────────────────
      // Saudável       → PE ≥ 8%
      // Estável        → PE 4–7,9% + crescendo igual/acima do mercado
      // Atenção        → PE 4–7,9% + abaixo do mercado
      // Em Recuperação → PE 0–3,9% + crescendo igual/acima do mercado
      // ALERTA         → PE 0–3,9% + abaixo do mercado  OU  PE negativo + acima do mercado
      // Crítico        → PE negativo + abaixo do mercado
      if (saudavel_fin) {
        s.classificacao = 'Saudável';
        s.sub_class     = 'Saudável';
      } else if (intermediario_fin && acima_merc) {
        s.classificacao = 'Atenção';
        s.sub_class     = 'Atenção';
      } else if (intermediario_fin && !acima_merc) {
        s.classificacao = 'Atenção';
        s.sub_class     = 'Atenção';
      } else if (fragil_fin && folga_pe >= 0 && acima_merc) {
        s.classificacao = 'Atenção';
        s.sub_class     = 'Atenção';
      } else if (fragil_fin && folga_pe >= 0 && !acima_merc) {
        s.classificacao = 'Atenção';
        s.sub_class     = 'ALERTA';
      } else if (folga_pe < 0 && acima_merc) {
        s.classificacao = 'Atenção';
        s.sub_class     = 'ALERTA';
      } else {
        s.classificacao = 'Crítico';
        s.sub_class     = 'Crítico';
      }

      setores.push(s);
    }

    // ── 4. Build dataset and save ────────────────────────────────────────────
    msg.textContent = 'Salvando dados…';
    await sleep(30);

    const dataset = {
      setores,
      regionais: regMap,
      distritais: distMap,
      brasil_luc: brasilLucMap,
      gr_luc: grLucMap,
      gd_luc: gdLucMap,
      meses: sortedMeses,
      desemp_by_mes: desempByMes,
      sourceFiles: fileNames,
      savedAt: new Date().toISOString(),
    };

    try {
      localStorage.setItem('supera_data', JSON.stringify(dataset));
    } catch(e) {
      // If too large for localStorage, store without venda_mes details
      const slim = { ...dataset };
      slim.setores = setores.map(s => {
        const r = {};
        for (const k of Object.keys(s)) {
          if (!k.startsWith('venda_mes_')) r[k] = s[k];
        }
        return r;
      });
      // Slim desemp_by_mes: only keep ytd and abrabr (drop abramar) to save space
      slim.desemp_by_mes = {};
      for (const [mes, dp] of Object.entries(desempByMes)) {
        slim.desemp_by_mes[mes] = {};
        for (const [code, d] of Object.entries(dp)) {
          slim.desemp_by_mes[mes][code] = {};
          for (const k of Object.keys(d)) {
            if (!k.startsWith('abramar_')) slim.desemp_by_mes[mes][code][k] = d[k];
          }
        }
      }
      localStorage.setItem('supera_data', JSON.stringify(slim));
    }

    prog.classList.remove('show');
    closeUpload();
    loadDataset(dataset);

  } catch(err) {
    prog.classList.remove('show');
    document.getElementById('btn-process').disabled = false;
    alert('Erro ao processar: ' + err.message);
    console.error(err);
  }
}

function clearAllData() {
  // Some environments block confirm() — use direct action with visual feedback
  const btn = document.getElementById('btn-clear');
  if (btn && btn.dataset.confirmed !== 'yes') {
    btn.textContent = '⚠ Clique novamente para confirmar';
    btn.dataset.confirmed = 'yes';
    setTimeout(() => { if(btn) { btn.textContent = '🗑 Limpar dados salvos'; btn.dataset.confirmed = ''; } }, 3000);
    return;
  }
  if (btn) { btn.textContent = '🗑 Limpar dados salvos'; btn.dataset.confirmed = ''; }
  localStorage.removeItem('supera_data');
  SETORES = []; REGIONAIS = {}; DISTRITAIS = {};
  BRASIL_LUC = {}; GR_LUC = {}; GD_LUC = {};
  DESEMP_BY_MES = {}; activeMeses = []; filtered = [];
  pendingLuc = []; pendingDesemp = []; negFilter = false; virouFilter = false; virouPosFilter = false;
  linhaVal = '';
  const lbl = document.getElementById('linha-dd-label'); if (lbl) lbl.textContent = 'Todas';
  // Reset header counts
  ['h-tot','h-s','h-a','h-c'].forEach(id => document.getElementById(id).textContent = '—');
  ['kv-s','kv-a','kv-c','kv-ns'].forEach(id => { const el=document.getElementById(id); if(el) el.textContent='—'; });
  // Reset KPI cards active state
  document.querySelectorAll('.kpi').forEach(k => k.classList.remove('ka'));
  document.getElementById('afb').classList.remove('show');
  kpiF = ''; activeTab = '';
  // Close upload panel and show no-data
  closeUpload();
  showNoData();
}

// ── XLSX HELPERS ──────────────────────────────────────────────────────────────
// Detecta mês em qualquer string (cabeçalho de coluna ou nome de arquivo)
// Aceita nome completo (JANEIRO) ou abreviação (JAN)
function detectMes(str) {
  if (!str) return null;
  str = str.toUpperCase()
           .normalize('NFD').replace(/[\u0300-\u036f]/g, ''); // remove acentos
  // Tenta nome completo primeiro
  for (const [abr, full] of Object.entries(MESES_PTBR)) {
    const fullNorm = full.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
    if (str.includes(fullNorm)) return abr;
  }
  // Tenta abreviação (3 letras) como palavra isolada
  for (const abr of Object.keys(MESES_PTBR)) {
    const re = new RegExp('\\b' + abr + '\\b');
    if (re.test(str)) return abr;
  }
  return null;
}

// Verifica se arquivo é de lucratividade (nome contém "lucratividade" ou "lucrativ")
function isLucFile(name) {
  const n = name.toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  return n.includes('LUCRATIVIDADE') || n.includes('LUCRATIV');
}

// Verifica se arquivo é de desempenho (nome contém "desempenho" ou "desemp")
function isDesempFile(name) {
  const n = name.toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  return n.includes('DESEMPENHO') || n.includes('DESEMP');
}

function safeNum(v) {
  if (v == null || v === '') return null;
  const n = parseFloat(v);
  return isNaN(n) ? null : n;
}

function polyfit1(vals) {
  const n = vals.length;
  const x = vals.map((_,i) => i);
  const sx = x.reduce((a,b)=>a+b,0);
  const sy = vals.reduce((a,b)=>a+b,0);
  const sxy = x.reduce((a,b,i)=>a+b*vals[i],0);
  const sx2 = x.reduce((a,b)=>a+b*b,0);
  return (n*sxy - sx*sy) / (n*sx2 - sx*sx);
}

function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

// ── BADGE TOOLTIP GLOBAL ──────────────────────────────────────────────────────
// Cria o elemento tooltip uma única vez no body
(function() {
  const el = document.createElement('div');
  el.id = 'badge-tooltip';
  el.innerHTML = '<div class="badge-tip-title" id="btt"></div><div class="badge-tip-row" id="btf"></div><div class="badge-tip-row" id="btm"></div>';
  document.body.appendChild(el);
})();

function showBadgeTip(e, el) {
  const tip = document.getElementById('badge-tooltip');
  document.getElementById('btt').textContent = el.dataset.tipTitle || '';
  document.getElementById('btf').textContent = el.dataset.tipFin  || '';
  document.getElementById('btm').textContent = el.dataset.tipMerc || '';

  tip.classList.add('show');

  const rect = el.getBoundingClientRect();
  const tipW = 260;
  const tipH = tip.offsetHeight || 90;

  // Posição horizontal — centrado no badge, sem sair da tela
  let left = rect.left + rect.width / 2 - tipW / 2;
  left = Math.max(8, Math.min(left, window.innerWidth - tipW - 8));

  // Posição vertical — aparece abaixo, mas se não couber aparece acima
  const spaceAbaixo = window.innerHeight - rect.bottom;
  const spaceAcima  = rect.top;
  let top, acima;

  if (spaceAbaixo >= tipH + 12) {
    top = rect.bottom + 8;
    acima = false;
  } else {
    top = rect.top - tipH - 8;
    acima = true;
  }

  tip.style.left = left + 'px';
  tip.style.top  = top + 'px';
  tip.style.width = tipW + 'px';

  // Ajusta a setinha para apontar na direção certa
  tip.style.setProperty('--arrow-top', acima ? 'auto' : '0');
  tip.style.setProperty('--arrow-bottom', acima ? '0' : 'auto');
  tip.style.setProperty('--arrow-border-top', acima ? `6px solid #1e3a5f` : '6px solid transparent');
  tip.style.setProperty('--arrow-border-bottom', acima ? '6px solid transparent' : `6px solid #1e3a5f`);
}

function hideBadgeTip() {
  const tip = document.getElementById('badge-tooltip');
  if (tip) tip.classList.remove('show');
}

// ── FORMAT HELPERS ────────────────────────────────────────────────────────────
function fp(v)   { if (v == null) return '—'; return (v * 100).toFixed(2) + '%'; }
function fBRL(v) { if (v == null) return '—'; return 'R$ ' + Math.round(v).toLocaleString('pt-BR'); }

// Var. Share = Var. Supera% - Var. Mercado% (positivo=verde, negativo=vermelho)
// fShare now receives the share_var value directly (read from spreadsheet)
function fShare(shareVar) {
  if (shareVar == null) return '<span style="color:var(--muted)">—</span>';
  const neutral = Math.abs(shareVar) < 0.0005;
  const pos     = shareVar > 0;
  const cls     = neutral ? 'fl' : pos ? 'pos' : 'neg';
  const arr     = neutral ? '→' : pos ? '▲' : '▼';
  return '<span class="vp ' + cls + '">' + arr + ' ' + (shareVar * 100).toFixed(2) + '%</span>';
}

// ── LEGEND MODAL ──────────────────────────────────────────────────────────────
function openLegend() {
  document.getElementById('legend-overlay').classList.add('show');
}
function closeLegend() {
  document.getElementById('legend-overlay').classList.remove('show');
}
function closeLegendIfBg(e) {
  if (e.target === document.getElementById('legend-overlay')) closeLegend();
}
