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
// Close dropdowns when clicking outside
document.addEventListener('click', () => {
  document.getElementById('linha-dd')?.classList.remove('open');
  document.getElementById('luc-linha-dd')?.classList.remove('open');
});

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
  const btnExp = document.getElementById('btn-export');
  if (btnExp) btnExp.style.display = '';
  lucViewM = viewM; // sync luc month to current month
  if (currentView === 'luc') { syncLucFilters(); renderLuc(); }

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
  document.getElementById('btn-process').disabled = !(pendingLuc.length > 0 || pendingDesemp.length > 0);
}

// ── XLSX PROCESSING ───────────────────────────────────────────────────────────
async function processFiles() {
  if (!window.XLSX) { alert('Aguarde o carregamento da biblioteca XLSX...'); return; }

  const prog = document.getElementById('upload-progress');
  const msg  = document.getElementById('progress-msg');
  prog.classList.add('show');
  document.getElementById('btn-process').disabled = true;

  try {
    if (pendingLuc.length > 0) {
      msg.textContent = 'Lendo arquivos de lucratividade…';
      await sleep(50);
    }

    // ── 1. Read all lucratividade files ──────────────────────────────────────
    const lucData = {};
    const regMap = {}, distMap = {};
    const brasilLucMap = {}, grLucMap = {}, gdLucMap = {};
    const fileNames = { luc: pendingLuc.map(f => f.name), desemp: pendingDesemp.map(f => f.name) };

    // If no luc files uploaded, reuse saved luc data from localStorage
    if (pendingLuc.length === 0) {
      try {
        const saved = JSON.parse(localStorage.getItem('supera_data') || '{}');
        if (saved.setores?.length) {
          // Rebuild lucData from saved setores
          const savedMeses = saved.meses || [];
          savedMeses.forEach(m => {
            lucData[m] = {};
            (saved.setores || []).forEach(s => {
              if (s['luc_' + m] != null || s['lucro_' + m] != null) {
                lucData[m][s.code] = {
                  nome: s.nome, reg: s.regional, dist: s.distrital, linha: s.linha,
                  luc:       s['luc_'   + m] ?? null,
                  lucro:     s['lucro_' + m] ?? null,
                  sup_ytd:   null, mkt_ytd: null,
                  venda:     null, desp_fixa: null, desp_var: null,
                };
              }
            });
          });
          // Restore maps
          Object.assign(regMap,  saved.regionais  || {});
          Object.assign(distMap, saved.distritais || {});
          Object.entries(saved.brasil_luc || {}).forEach(([m,v]) => brasilLucMap[m] = v);
          Object.entries(saved.gr_luc || {}).forEach(([m,v]) => { grLucMap[m] = v; });
          Object.entries(saved.gd_luc || {}).forEach(([m,v]) => { gdLucMap[m] = v; });
          // Preserve old file names for luc
          fileNames.luc = saved.sourceFiles?.luc || [];
        }
      } catch(e) { console.warn('Could not restore luc from saved data', e); }
    }

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

    // ── 2. Read desempenho files (one per month) — optional ──────────────────
    if (pendingDesemp.length > 0) {
      msg.textContent = 'Processando planilhas de desempenho…';
      await sleep(50);
    }

    const desempByMes = {};
    let lastDesempMes = null;
    let nMeses = 1;

    const MESES_ABR = {JAN:1,FEV:2,MAR:3,ABR:4,MAI:5,JUN:6,JUL:7,AGO:8,SET:9,OUT:10,NOV:11,DEZ:12};

    for (const desempFile of pendingDesemp) {
      msg.textContent = `Processando ${desempFile.name}…`;
      await sleep(30);

      const wb2    = XLSX.read(await desempFile.arrayBuffer(), { type: 'array' });
      const desemp = {};
      const allSheets = wb2.SheetNames;
      const ytdSheet   = allSheets.find(s => s === 'YTD') || null;
      const mesXmesSheet = allSheets.find(s => {
        const p = s.split(' X ');
        return p.length === 2 && p[0].trim() === p[1].trim() && p[0].trim().length === 3;
      }) || null;
      const mesXprevSheet = allSheets.find(s => {
        const p = s.split(' X ');
        return p.length === 2 && p[0].trim() !== p[1].trim() && p[0].trim().length === 3 && p[1].trim().length === 3;
      }) || null;

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

      if (currentMesAbbr) {
        desempByMes[currentMesAbbr] = desemp;
        if (!lastDesempMes || fileMesN > (MESES_ABR[lastDesempMes] || 0)) {
          lastDesempMes = currentMesAbbr;
          nMeses = fileMesN;
        }
      }
    }

    // Use latest desempenho file as base (empty object if none loaded)
    const desemp = desempByMes[lastDesempMes] || {};

    // ── 3. Merge into setores ────────────────────────────────────────────────
    msg.textContent = 'Calculando lucratividade e classificações…';
    await sleep(50);

    const sortedMeses = MESES.filter(m => lucData[m]);

    // If no luc files, build setores from desemp codes
    let allCodes;
    if (sortedMeses.length > 0) {
      const lastMes = sortedMeses[sortedMeses.length - 1];
      allCodes = new Set(Object.keys(lucData[lastMes] || {}));
    } else {
      // Only desemp loaded — build code list from desemp
      const allDesempCodes = new Set();
      Object.values(desempByMes).forEach(dm => Object.keys(dm).forEach(c => allDesempCodes.add(c)));
      allCodes = allDesempCodes;
    }

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
      const lastM = sortedMeses.length > 0 ? sortedMeses[sortedMeses.length - 1] : null;
      s.venda_media = dp.ytd_sup_atual != null ? dp.ytd_sup_atual / nMeses : null;

      // Desempenho fields
      Object.assign(s, dp);

      // n_neg
      s.n_neg = sortedMeses.filter(m => (s['luc_'+m] || 0) < 0).length;

      // ── PONTO DE EQUILÍBRIO (YTD do mês mais recente) ───────────────────────
      const d_lastM = lastM ? (lucData[lastM]?.[code]) : null;
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

// ── EXPORT EXCEL ─────────────────────────────────────────────────────────────
function exportExcel() {
  if (!window.XLSX) { alert('Aguarde o carregamento da biblioteca XLSX...'); return; }
  if (!filtered.length) { alert('Nenhum setor para exportar com os filtros atuais.'); return; }

  const wb = XLSX.utils.book_new();

  // ── Build data rows ──────────────────────────────────────────────────────
  const mesLabel = viewM;
  const prevM = (() => { const i = activeMeses.indexOf(viewM); return i > 0 ? activeMeses[i-1] : null; })();

  const headers = [
    'Setor', 'Código', 'Linha', 'Regional', 'Distrital', 'Classificação',
    `Luc. ${mesLabel}`,
    'Venda Média Mensal',
    'YTD Var. Supera%', 'YTD Var. Mercado%', 'YTD Var. Share',
    `${mesLabel} Var. Supera%`, `${mesLabel} Var. Mercado%`, `${mesLabel} Var. Share`,
    'Meses Negativos',
  ];
  activeMeses.forEach(m => headers.push(`Luc. ${m}`));

  const rows = filtered.map(s => {
    const row = [
      s.nome,
      s.code,
      s.linha || '',
      s.regional_nome || '',
      s.distrital_nome || '',
      s.sub_class || s.classificacao,
      s['luc_' + viewM] != null ? s['luc_' + viewM] : '',
      s.venda_media != null ? s.venda_media : '',
      s.ytd_sup_var  != null ? s.ytd_sup_var  : '',
      s.ytd_merc_var != null ? s.ytd_merc_var : '',
      s.ytd_share_var != null ? s.ytd_share_var : '',
      s.abrabr_sup_var  != null ? s.abrabr_sup_var  : '',
      s.abrabr_merc_var != null ? s.abrabr_merc_var : '',
      s.abrabr_share_var != null ? s.abrabr_share_var : '',
      s.n_neg || 0,
    ];
    activeMeses.forEach(m => row.push(s['luc_' + m] != null ? s['luc_' + m] : ''));
    return row;
  });

  const wsData = [headers, ...rows];
  const ws = XLSX.utils.aoa_to_sheet(wsData);

  // ── Column widths ────────────────────────────────────────────────────────
  ws['!cols'] = [
    {wch:32},{wch:9},{wch:10},{wch:22},{wch:26},{wch:16},
    {wch:10},{wch:18},
    {wch:14},{wch:16},{wch:14},
    {wch:14},{wch:16},{wch:14},
    {wch:8},
    ...activeMeses.map(()=>({wch:10})),
  ];

  // ── Styles via cell-level properties (xlsx supports limited styling) ─────
  const range = XLSX.utils.decode_range(ws['!ref']);
  const nCols = headers.length;

  // Header row styling
  for (let c = 0; c < nCols; c++) {
    const cell = ws[XLSX.utils.encode_cell({r:0, c})];
    if (!cell) continue;
    cell.s = {
      font: { bold: true, color: { rgb: 'FFFFFF' }, sz: 11, name: 'Arial' },
      fill: { fgColor: { rgb: '1E3A5F' } },
      alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
      border: { bottom: { style: 'thin', color: { rgb: '93C5FD' } } },
    };
  }

  // Classification color map
  const classColors = {
    'Saudável':       { bg: 'D1FAE5', fg: '065F46' },
    'Excelente':      { bg: 'D1FAE5', fg: '065F46' },
    'Atenção':        { bg: 'FEF9C3', fg: '92400E' },
    'ALERTA':         { bg: 'FFEDD5', fg: '9A3412' },
    'Crítico':        { bg: 'FEE2E2', fg: '991B1B' },
  };

  // Data rows styling
  for (let r = 1; r <= filtered.length; r++) {
    const s = filtered[r-1];
    const isEven = r % 2 === 0;
    const rowBg = isEven ? 'F8FAFC' : 'FFFFFF';

    for (let c = 0; c < nCols; c++) {
      const addr = XLSX.utils.encode_cell({r, c});
      if (!ws[addr]) ws[addr] = { t: 'z', v: '' };
      const cell = ws[addr];

      // Base style
      cell.s = {
        font: { sz: 10, name: 'Arial' },
        fill: { fgColor: { rgb: rowBg } },
        alignment: { vertical: 'center' },
        border: { bottom: { style: 'thin', color: { rgb: 'E2E8F0' } } },
      };

      // Classification column — colored
      if (c === 5) {
        const cc = classColors[cell.v] || {};
        if (cc.bg) {
          cell.s.fill = { fgColor: { rgb: cc.bg } };
          cell.s.font = { sz: 10, name: 'Arial', bold: true, color: { rgb: cc.fg } };
          cell.s.alignment = { horizontal: 'center', vertical: 'center' };
        }
      }

      // Percentage columns — format as % and color pos/neg
      const pctCols = [6, 8, 9, 10, 11, 12, 13];
      const lucStartCol = 15;
      const isLucCol = c >= lucStartCol;
      if (pctCols.includes(c) || isLucCol) {
        if (cell.v !== '' && cell.v != null) {
          cell.t = 'n';
          cell.z = '0.00%';
          const val = parseFloat(cell.v);
          const color = val >= 0 ? '166534' : '991B1B';
          cell.s.font = { ...cell.s.font, color: { rgb: color }, bold: true };
          cell.s.alignment = { ...cell.s.alignment, horizontal: 'center' };
        }
      }

      // Currency column
      if (c === 7) {
        if (cell.v !== '' && cell.v != null) {
          cell.t = 'n';
          cell.z = 'R$ #,##0';
          cell.s.alignment = { ...cell.s.alignment, horizontal: 'right' };
        }
      }

      // Meses neg column — center
      if (c === 14) {
        cell.s.alignment = { ...cell.s.alignment, horizontal: 'center' };
      }
    }
  }

  // ── Freeze top row ───────────────────────────────────────────────────────
  ws['!freeze'] = { xSplit: 0, ySplit: 1, topLeftCell: 'A2', activeCell: 'A2', sqref: 'A2' };

  // ── Sheet name with active filters ──────────────────────────────────────
  const filterParts = [];
  filterParts.push(mesLabel);
  if (linhaVal) filterParts.push(linhaVal);
  if (activeTab) filterParts.push(activeTab.substring(0,4));
  if (negFilter) filterParts.push('NEG');
  if (virouFilter) filterParts.push('VIROU-');
  if (virouPosFilter) filterParts.push('VIROU+');
  const sheetName = ('Setores_' + filterParts.join('_')).substring(0, 31);

  XLSX.utils.book_append_sheet(wb, ws, sheetName);

  // ── Download ─────────────────────────────────────────────────────────────
  const dateStr = new Date().toLocaleDateString('pt-BR').replace(/\//g,'-');
  XLSX.writeFile(wb, `SUPERA_${sheetName}_${dateStr}.xlsx`);
}

// ── LUCRATIVIDADE VIEW ────────────────────────────────────────────────────────
let currentView = 'geral';
let lucViewM = 'ABR';
let lucLinhaVal = '';
let lucNegFilter = false, lucVirouFilter = false, lucVirouPosFilter = false;
let lucFocusedCol = null; // which column is expanded
let lucSortDir = 'desc';  // 'desc' = maior para menor, 'asc' = menor para maior

function switchView(v) {
  currentView = v;
  document.getElementById('view-luc').style.display = v === 'luc' ? '' : 'none';

  // Hide/show geral-specific sections (fbar and main), but never the upload overlay
  const geralSections = document.querySelectorAll('.fbar, .main');
  geralSections.forEach(el => {
    if (!el.closest('#view-luc') && !el.closest('#upload-overlay')) {
      el.style.display = v === 'geral' ? '' : 'none';
    }
  });

  document.getElementById('vt-geral').classList.toggle('active', v === 'geral');
  document.getElementById('vt-luc').classList.toggle('active', v === 'luc');

  // Always keep export button accessible
  const btnExp = document.getElementById('btn-export');
  if (btnExp && SETORES.length) btnExp.style.display = '';

  if (v === 'luc') {
    syncLucFilters();
    renderLuc();
  }
}

function syncLucFilters() {
  // Sync month buttons
  const msel = document.getElementById('luc-msel');
  msel.innerHTML = '';
  activeMeses.forEach(m => {
    const b = document.createElement('button');
    b.className = 'mb' + (m === lucViewM ? ' ma' : '');
    b.textContent = m;
    b.onclick = () => { lucViewM = m; msel.querySelectorAll('.mb').forEach(x=>x.classList.toggle('ma',x.textContent===m)); renderLuc(); };
    msel.appendChild(b);
  });
  // Sync regional filter
  const lReg = document.getElementById('lf-reg');
  lReg.innerHTML = '<option value="">Todas</option>';
  Object.entries(REGIONAIS).forEach(([code,name]) => {
    const o = document.createElement('option'); o.value=code; o.textContent=name; lReg.appendChild(o);
  });
}

function onLucRegChange() {
  const reg = document.getElementById('lf-reg').value;
  const fd = document.getElementById('lf-dist');
  fd.innerHTML = '<option value="">Todas</option>';
  const seen = new Set();
  SETORES.forEach(s => {
    if (reg && s.regional !== reg) return;
    if (seen.has(s.distrital)) return;
    seen.add(s.distrital);
    const o = document.createElement('option');
    o.value = s.distrital; o.textContent = s.distrital_nome; fd.appendChild(o);
  });
  renderLuc();
}

function focusLucCol(cls) {
  lucFocusedCol = cls;
  lucSortDir = 'desc';
  applyLucFocus();
}
function closeLucFocus() {
  lucFocusedCol = null;
  const grid = document.getElementById('luc-grid');
  if (!grid) return;
  grid.classList.remove('has-focus');
  grid.querySelectorAll('.luc-col').forEach(c => c.classList.remove('focused'));
  grid.querySelectorAll('.luc-col-sort').forEach(el => el.remove());
}
function setLucSort(dir) {
  lucSortDir = dir;
  applyLucFocus();
}
function applyLucFocus() {
  const grid = document.getElementById('luc-grid');
  if (!lucFocusedCol) { closeLucFocus(); return; }
  grid.classList.add('has-focus');
  grid.querySelectorAll('.luc-col').forEach(col => {
    const isFocused = col.dataset.cls === lucFocusedCol;
    col.classList.toggle('focused', isFocused);
    if (isFocused) {
      // Add/update sort bar
      let sortBar = col.querySelector('.luc-col-sort');
      if (!sortBar) {
        sortBar = document.createElement('div');
        sortBar.className = 'luc-col-sort';
        col.querySelector('.luc-col-hdr').insertAdjacentElement('afterend', sortBar);
      }
      sortBar.innerHTML = `
        <span>Ordenar:</span>
        <button class="luc-sort-btn${lucSortDir==='desc'?' active':''}" onclick="setLucSort('desc')">↓ Maior → Menor</button>
        <button class="luc-sort-btn${lucSortDir==='asc'?' active':''}" onclick="setLucSort('asc')">↑ Menor → Maior</button>
        <button class="luc-col-close" onclick="closeLucFocus()">✕ Fechar</button>`;
      // Re-sort cards in body
      const body = col.querySelector('.luc-col-body');
      const cards = Array.from(body.querySelectorAll('.luc-card'));
      cards.sort((a, b) => {
        const va = parseFloat(a.dataset.lucro) || 0;
        const vb = parseFloat(b.dataset.lucro) || 0;
        return lucSortDir === 'desc' ? vb - va : va - vb;
      });
      body.innerHTML = '';
      cards.forEach(c => body.appendChild(c));
    }
  });
}

function setLucFilters(neg, virou, virouPos) {
  lucNegFilter      = neg;
  lucVirouFilter    = virou;
  lucVirouPosFilter = virouPos;
  // Sync button states by ID (elements always exist in DOM)
  const nb  = document.getElementById('lbtn-neg');
  const vb  = document.getElementById('lbtn-virou');
  const vpb = document.getElementById('lbtn-virou-pos');
  if (nb)  nb.classList.toggle('active', neg);
  if (vb)  vb.classList.toggle('active', virou);
  if (vpb) vpb.classList.toggle('active', virouPos);
  renderLuc();
}
function toggleLucNeg()      { setLucFilters(!lucNegFilter, false, false); }
function toggleLucVirou()    { setLucFilters(false, !lucVirouFilter, false); }
function toggleLucVirouPos() { setLucFilters(false, false, !lucVirouPosFilter); }

function toggleLucLinhaDd(e) {
  e.stopPropagation();
  document.getElementById('luc-linha-dd').classList.toggle('open');
}
function setLucLinha(e, val) {
  e.stopPropagation();
  lucLinhaVal = val;
  const label = document.getElementById('luc-linha-dd-label');
  if (val === 'PRIME') label.innerHTML = '<span class="linha-badge lb-prime">PRIME</span>';
  else if (val === 'INFINITY') label.innerHTML = '<span class="linha-badge lb-infinity">INFINITY</span>';
  else label.textContent = 'Todas';
  document.getElementById('luc-linha-dd').classList.remove('open');
  renderLuc();
}

function lucClassify(lucro) {
  if (lucro == null) return null;
  if (lucro < 0)      return 'critico';
  if (lucro < 10000)  return 'alerta';
  if (lucro <= 20000) return 'atencao';
  if (lucro < 25000)  return 'atencao-leve';
  return 'saudavel';
}

const lucClassLabels = {
  'critico':      'Crítico',
  'alerta':       'Alerta',
  'atencao':      'Atenção',
  'atencao-leve': 'Atenção Leve',
  'saudavel':     'Saudável',
};
const lucClassOrder = ['critico','alerta','atencao','atencao-leve','saudavel'];

function renderLuc() {
  if (!SETORES.length) return;
  const reg   = document.getElementById('lf-reg').value;
  const dist  = document.getElementById('lf-dist').value;
  const srch  = (document.getElementById('lf-search').value||'').toLowerCase().trim();

  // Filter setores
  const prevML = (() => { const i = activeMeses.indexOf(lucViewM); return i > 0 ? activeMeses[i-1] : null; })();

  let data = SETORES.filter(s => {
    if (reg  && s.regional  !== reg)  return false;
    if (dist && s.distrital !== dist) return false;
    if (lucLinhaVal && s.linha !== lucLinhaVal) return false;
    if (srch && !s.nome.toLowerCase().includes(srch) && !s.code.includes(srch)) return false;
    // Use lucro_ (R$) if available, fallback to sign of luc_ (%)
    const cur  = s['lucro_' + lucViewM] ?? null;
    const curL = s['luc_'   + lucViewM] ?? null;
    const curSign  = cur  != null ? cur  : (curL  != null ? (curL  >= 0 ? 1 : -1) : null);
    const prev = prevML ? (s['lucro_' + prevML] ?? null) : null;
    const prevL= prevML ? (s['luc_'   + prevML] ?? null) : null;
    const prevSign = prev != null ? prev : (prevL != null ? (prevL >= 0 ? 1 : -1) : null);
    if (lucNegFilter     && (curSign  == null || curSign  >= 0)) return false;
    if (lucVirouFilter   && (curSign  == null || prevSign == null || !(prevSign >= 0 && curSign  < 0))) return false;
    if (lucVirouPosFilter && (curSign == null || prevSign == null || !(prevSign <  0 && curSign  >= 0))) return false;
    return true;
  });

  // Attach lucro value and class for current month
  data = data.map(s => ({
    ...s,
    _lucro: s['lucro_' + lucViewM] ?? null,
    _luc:   s['luc_'   + lucViewM] ?? null,
    _class: lucClassify(s['lucro_' + lucViewM] ?? null),
  })).filter(s => s._class !== null);

  // Sort: critico first, then alerta, atencao, atencao-leve, saudavel; within each by lucro asc
  data.sort((a,b) => {
    const oi = lucClassOrder.indexOf(a._class) - lucClassOrder.indexOf(b._class);
    if (oi !== 0) return oi;
    return (a._lucro??0) - (b._lucro??0);
  });

  document.getElementById('luc-count').textContent = data.length + ' setores';

  // Summary strip
  const counts = {};
  lucClassOrder.forEach(c => counts[c] = 0);
  data.forEach(s => counts[s._class]++);

  const summaryLabels = {
    'critico':      ['Crítico','Lucro negativo'],
    'alerta':       ['Alerta','R$0 – R$9.999'],
    'atencao':      ['Atenção','R$10K – R$20K'],
    'atencao-leve': ['Atenção Leve','R$20K – R$24.999'],
    'saudavel':     ['Saudável','≥ R$25.000'],
  };
  document.getElementById('luc-summary').innerHTML = lucClassOrder.map(c => `
    <div class="luc-scard lsc-${c}">
      <div class="luc-scard-n">${counts[c]}</div>
      <div class="luc-scard-l">${summaryLabels[c][0]}</div>
      <div class="luc-scard-sub">${summaryLabels[c][1]}</div>
    </div>`).join('');

  // 5-column executive layout
  const fBR = v => v == null ? '—' : (v < 0 ? '−R$ ' : 'R$ ') + Math.abs(Math.round(v)).toLocaleString('pt-BR');

  const colConfig = {
    'critico':      { title: 'Crítico',       range: 'Lucro negativo' },
    'alerta':       { title: 'Alerta',        range: 'R$0 – R$9.999' },
    'atencao':      { title: 'Atenção',       range: 'R$10K – R$20K' },
    'atencao-leve': { title: 'Atenção Leve',  range: 'R$20K – R$24.999' },
    'saudavel':     { title: 'Saudável',      range: '≥ R$25.000' },
  };

  const grid = document.getElementById('luc-grid');
  grid.innerHTML = '';
  // Don't reset lucFocusedCol here — preserve across filter changes

  lucClassOrder.forEach(cls => {
    const group = data.filter(s => s._class === cls);
    const cfg = colConfig[cls];

    const col = document.createElement('div');
    col.className = `luc-col luc-col-${cls}`;
    col.dataset.cls = cls;

    const hdr = document.createElement('div');
    hdr.className = 'luc-col-hdr';
    hdr.style.cursor = 'pointer';
    hdr.title = 'Clique para expandir';
    hdr.onclick = () => lucFocusedCol === cls ? closeLucFocus() : focusLucCol(cls);
    hdr.innerHTML = `
      <div>
        <div class="luc-col-hdr-title">${cfg.title}</div>
        <div class="luc-col-hdr-range">${cfg.range}</div>
      </div>
      <div class="luc-col-hdr-count">${group.length}</div>`;
    col.appendChild(hdr);

    const body = document.createElement('div');
    body.className = 'luc-col-body';

    if (!group.length) {
      body.innerHTML = `<div class="luc-col-empty">Nenhum setor</div>`;
    } else {
      group.forEach(s => {
        const lucPct = s._luc != null ? `<span class="luc-card-pct">${(s._luc*100).toFixed(2)}%</span>` : '';
        const card = document.createElement('div');
        card.className = 'luc-card';
        card.dataset.lucro = s._lucro ?? '';
        card.innerHTML = `
          <div class="luc-card-nome" title="${s.nome}">${s.nome}</div>
          <div class="luc-card-code">${s.code}${s.linha ? `<span class="linha-badge lb-${s.linha.toLowerCase()}">${s.linha}</span>` : ''}</div>
          <div class="luc-card-val">${fBR(s._lucro)}${lucPct}</div>
          <div class="luc-card-reg">${s.regional} ${(s.regional_nome||'').split(' ')[0]} · ${s.distrital} ${(s.distrital_nome||'').split(' ')[0]}</div>`;
        body.appendChild(card);
      });
    }

    col.appendChild(body);
    grid.appendChild(col);
  });

  // Re-apply focus if one was active before re-render
  if (lucFocusedCol) applyLucFocus();
}

// ── EXPORT LUC EXCEL ─────────────────────────────────────────────────────────
async function exportLucExcel() {
  if (!window.ExcelJS) { alert('Aguarde o carregamento da biblioteca ExcelJS...'); return; }

  const reg  = document.getElementById('lf-reg').value;
  const dist = document.getElementById('lf-dist').value;
  const srch = (document.getElementById('lf-search').value||'').toLowerCase().trim();
  const prevML = (() => { const i = activeMeses.indexOf(lucViewM); return i > 0 ? activeMeses[i-1] : null; })();

  let data = SETORES.filter(s => {
    if (reg  && s.regional  !== reg)  return false;
    if (dist && s.distrital !== dist) return false;
    if (lucLinhaVal && s.linha !== lucLinhaVal) return false;
    if (srch && !s.nome.toLowerCase().includes(srch) && !s.code.includes(srch)) return false;
    const cur   = s['lucro_' + lucViewM] ?? null;
    const curL  = s['luc_'   + lucViewM] ?? null;
    const curSign  = cur  != null ? cur  : (curL  != null ? (curL  >= 0 ? 1 : -1) : null);
    const prev  = prevML ? (s['lucro_' + prevML] ?? null) : null;
    const prevL = prevML ? (s['luc_'   + prevML] ?? null) : null;
    const prevSign = prev != null ? prev : (prevL != null ? (prevL >= 0 ? 1 : -1) : null);
    if (lucNegFilter      && (curSign == null || curSign >= 0)) return false;
    if (lucVirouFilter    && (curSign == null || prevSign == null || !(prevSign >= 0 && curSign < 0))) return false;
    if (lucVirouPosFilter && (curSign == null || prevSign == null || !(prevSign < 0 && curSign >= 0))) return false;
    return true;
  }).map(s => ({
    ...s,
    _lucro: s['lucro_' + lucViewM] ?? null,
    _luc:   s['luc_'   + lucViewM] ?? null,
    _class: lucClassify(s['lucro_' + lucViewM] ?? null),
  })).filter(s => s._class !== null)
    .sort((a,b) => {
      const oi = lucClassOrder.indexOf(a._class) - lucClassOrder.indexOf(b._class);
      return oi !== 0 ? oi : (a._lucro??0) - (b._lucro??0);
    });

  if (!data.length) { alert('Nenhum setor para exportar.'); return; }

  const classLabel = {
    'critico':'Crítico','alerta':'Alerta',
    'atencao':'Atenção','atencao-leve':'Atenção Leve','saudavel':'Saudável'
  };
  const themeColor = {
    'critico':'BE123C','alerta':'C2410C','atencao':'A16207',
    'atencao-leve':'CA8A04','saudavel':'15803D'
  };
  const themeBg = {
    'critico':'FFF1F2','alerta':'FFF7ED','atencao':'FFFBEB',
    'atencao-leve':'FEFCE8','saudavel':'F0FDF4'
  };

  const wb = new ExcelJS.Workbook();
  const sheetName = `Luc_${lucViewM}`.substring(0,31);
  const ws = wb.addWorksheet(sheetName);

  // Column definitions
  ws.columns = [
    { key:'cls',    width: 14 },
    { key:'setor',  width: 56 },
    { key:'linha',  width:  9 },
    { key:'reg',    width: 28 },
    { key:'dist',   width: 40 },
    { key:'lucro',  width: 14 },
    { key:'pct',    width: 11 },
  ];

  // Header row
  const hdrRow = ws.addRow([
    'Classificação','Setor','Linha','Regional','Distrital',
    `Lucro R$ ${lucViewM}`, `Luc% ${lucViewM}`
  ]);
  hdrRow.height = 18;
  hdrRow.eachCell(cell => {
    cell.font      = { bold: true, color: { argb: 'FFFFFFFF' }, size: 12, name: 'Calibri' };
    cell.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF002060' } };
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
    cell.border    = { bottom: { style: 'thin', color: { argb: 'FFFFFFFF' } } };
  });

  // Freeze header
  ws.views = [{ state: 'frozen', xSplit: 0, ySplit: 1, topLeftCell: 'A2', activeCell: 'A2' }];

  // Data rows
  data.forEach(s => {
    const bg   = 'FF' + (themeBg[s._class]   || 'FFFFFF');
    const fg   = 'FF' + (themeColor[s._class] || '000000');
    const row  = ws.addRow([
      classLabel[s._class] || '',
      s.code + ' - ' + s.nome,
      s.linha || '',
      s.regional_nome || '',
      s.distrital_nome || '',
      s._lucro ?? null,
      s._luc   ?? null,
    ]);
    row.height = 16;

    row.eachCell({ includeEmpty: true }, (cell, colNum) => {
      cell.fill   = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
      cell.border = { bottom: { style: 'thin', color: { argb: 'FFD9D9D9' } } };

      const isLucro = colNum === 6;
      const isPct   = colNum === 7;
      const val     = cell.value;
      const isNeg   = typeof val === 'number' && val < 0;
      const numArgb = 'FF' + (isNeg ? 'B91C1C' : '15803D');

      if (isLucro) {
        cell.numFmt    = '\R\$\ #,##0;\-\R\$\ #,##0';
        cell.font      = { bold: true, size: 12, name: 'Calibri', color: { argb: numArgb } };
        cell.alignment = { horizontal: 'right', vertical: 'middle' };
      } else if (isPct) {
        cell.numFmt    = '0.00%';
        cell.font      = { bold: true, size: 12, name: 'Calibri', color: { argb: numArgb } };
        cell.alignment = { horizontal: 'right', vertical: 'middle' };
      } else if (colNum === 1) {
        cell.font      = { bold: true, size: 12, name: 'Calibri', color: { argb: fg } };
        cell.alignment = { horizontal: 'left', vertical: 'middle' };
      } else {
        cell.font      = { size: 12, name: 'Calibri', color: { argb: 'FF000000' } };
        cell.alignment = { horizontal: 'left', vertical: 'middle' };
      }
    });
  });

  // Download
  const buf  = await wb.xlsx.writeBuffer();
  const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  const date = new Date().toLocaleDateString('pt-BR').replace(/\//g,'-');
  a.href     = url;
  a.download = `SUPERA_Lucratividade_${lucViewM}_${date}.xlsx`;
  a.click();
  URL.revokeObjectURL(url);
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
