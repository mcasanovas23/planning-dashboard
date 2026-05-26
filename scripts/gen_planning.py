import openpyxl, json, sys, os
from datetime import datetime

EXCEL_PATH = r'C:\Users\mcasanovas\OneDrive - IVASCULAR, S.L.U\Planning General.xlsm'
TMP_PATH   = r'C:\Users\mcasanovas\AppData\Local\Temp\Planning_General_tmp.xlsm'
OUT_HTML   = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'index.html')

import shutil
shutil.copy2(EXCEL_PATH, TMP_PATH)

wb = openpyxl.load_workbook(TMP_PATH, read_only=True, keep_vba=False, data_only=True)

# ─── Family merge rules ────────────────────────────────────────────────────
# Families merged into ANGIOLIT (tab), hidden from dashboard as own rows
MERGE_INTO_ANGIOLIT = {'ANGIOBTK', 'PUTNIY', 'ENVIROS'}
# Families merged into ICOVER (tab), hidden from dashboard as own rows
MERGE_INTO_ICOVER   = {'TUVA'}

def canonical_tab(familia):
    if familia in MERGE_INTO_ANGIOLIT: return 'ANGIOLIT'
    if familia in MERGE_INTO_ICOVER:   return 'ICOVER'
    return familia

# ─── Datos ─────────────────────────────────────────────────────────────────
ws = wb['Datos']
all_rows = list(ws.iter_rows(values_only=True))

# Columns: a,b,c,d,f,h,k,m,o,s,t,u,v,w,ab
COL_INDICES = [0,1,2,3,5,7,10,12,14,18,19,20,21,22,27]
COL_KEYS    = ['article','medidas','familia','lot','setmana',
               'mat1','disp1','mat2','disp2','prep_linia',
               'pzas_lot','pzas_cmd','sota_cmd','fase','txt']
COL_HEADERS = ['Artículo','Medidas','Familia','Lote','Sem.',
               'Material 1','Disp. Mat1','Material 2','Disp. Mat2','Prep. Línea',
               'Pzas/Lote','Pzas/Cmd','Sota Cmd','Fase Actual','Txt']

rows_data = []
for row in all_rows[1:]:
    if not any(v is not None for v in row[:5]):
        continue
    rd = {}
    for i, ci in enumerate(COL_INDICES):
        v = row[ci] if ci < len(row) else None
        if isinstance(v, datetime): v = v.strftime('%d/%m/%Y')
        rd[COL_KEYS[i]] = v if v is not None else ''
    rd['_retras']  = bool(row[24] == '!!!') if len(row) > 24 else False
    rd['_tab']     = canonical_tab(rd['familia'])  # which tab this row belongs to
    rows_data.append(rd)

# Families that have any mat2 data → show mat2 columns in those tabs
fam_has_mat2 = set(
    rd['_tab'] for rd in rows_data
    if rd['mat2'] not in ('', None)
)

# Build rows grouped by _tab, sorted by setmana then article
rows_by_tab = {}
for rd in rows_data:
    tab = rd['_tab']
    if tab not in rows_by_tab:
        rows_by_tab[tab] = []
    rows_by_tab[tab].append(rd)

for tab in rows_by_tab:
    rows_by_tab[tab].sort(key=lambda r: (
        r['setmana'] if isinstance(r['setmana'], int) else 999,
        str(r['article'])
    ))

# ─── Resum ─────────────────────────────────────────────────────────────────
ws_resum = wb['Resum']
resum_rows = list(ws_resum.iter_rows(values_only=True))

week_row = resum_rows[6]
WEEKS = []
for ci in [4, 8, 12, 16, 20, 24]:
    wnum = week_row[ci]
    if wnum:
        WEEKS.append({'num': int(wnum), 'ci': ci})

# Families to SKIP in dashboard (merged/absorbed elsewhere)
SKIP_IN_DASH = MERGE_INTO_ANGIOLIT | MERGE_INTO_ICOVER

resum_data = []
family_order_resum = []
for row in resum_rows[8:28]:
    familia = row[2]
    if not familia:
        continue
    if familia in SKIP_IN_DASH:
        continue
    family_order_resum.append(familia)
    weeks_info = []
    for w in WEEKS:
        ci = w['ci']
        sem      = row[ci]   if len(row) > ci   else None
        planning = row[ci+1] if len(row) > ci+1 else None
        dif      = row[ci+2] if len(row) > ci+2 else None
        weeks_info.append({'sem': sem, 'planning': planning, 'dif': dif})
    resum_data.append({'familia': familia, 'weeks': weeks_info})

# Tabs ordered: Resum order, only tabs with data
all_tabs_with_data = set(rows_by_tab.keys())
familias_ordered = [f for f in family_order_resum if f in all_tabs_with_data]

# ─── FASES ─────────────────────────────────────────────────────────────────
ws_fases = wb['FASES']
fases_rows = list(ws_fases.iter_rows(values_only=True))
fases_by_family = {}
seen = {}
for row in fases_rows[1:]:
    if not row[0]: continue
    fam = row[0]; name = row[2]; order = row[3]
    if fam not in fases_by_family:
        fases_by_family[fam] = []; seen[fam] = set()
    if name and name not in seen[fam]:
        fases_by_family[fam].append({'name': str(name), 'order': order or 0})
        seen[fam].add(name)
for f in fases_by_family:
    fases_by_family[f].sort(key=lambda x: x['order'])

# Also make merged families share fases
for src in MERGE_INTO_ANGIOLIT:
    if src in fases_by_family and 'ANGIOLIT' not in fases_by_family:
        fases_by_family['ANGIOLIT'] = fases_by_family[src]
for src in MERGE_INTO_ICOVER:
    if src in fases_by_family and 'ICOVER' not in fases_by_family:
        fases_by_family['ICOVER'] = fases_by_family[src]

# ─── JSON blobs ─────────────────────────────────────────────────────────────
j_resum    = json.dumps(resum_data,      ensure_ascii=False)
j_weeks    = json.dumps([w['num'] for w in WEEKS], ensure_ascii=False)
j_fam_ord  = json.dumps(familias_ordered, ensure_ascii=False)
j_col_keys = json.dumps(COL_KEYS,        ensure_ascii=False)
j_col_hdr  = json.dumps(COL_HEADERS,     ensure_ascii=False)
j_fases    = json.dumps(fases_by_family, ensure_ascii=False)
j_all_rows = json.dumps(rows_by_tab,     ensure_ascii=False)
j_has_mat2 = json.dumps(list(fam_has_mat2), ensure_ascii=False)

ts = datetime.now().strftime('%d/%m/%Y %H:%M')

# ─── HTML ───────────────────────────────────────────────────────────────────
html = f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>Planning General – iVascular</title>
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:'Segoe UI',Arial,sans-serif;background:#f0f4f8;color:#1e293b;font-size:13px}}

/* ── Top bar ── */
.top-bar{{background:#0f2044;color:#fff;padding:10px 20px;display:flex;align-items:center;gap:16px;
  position:sticky;top:0;z-index:200;box-shadow:0 2px 8px rgba(0,0,0,.5)}}
.top-bar h1{{font-size:15px;font-weight:700;letter-spacing:.5px;white-space:nowrap}}
.top-bar .ts{{font-size:10px;color:#64748b;margin-left:auto;white-space:nowrap}}

/* ── Tab nav ── */
.tab-nav{{background:#162d5a;display:flex;flex-wrap:wrap;gap:2px;padding:0 8px;
  position:sticky;top:41px;z-index:190;box-shadow:0 2px 4px rgba(0,0,0,.3)}}
.tab-btn{{padding:7px 13px;border:none;background:transparent;color:#94a3b8;cursor:pointer;
  font-size:11px;font-weight:600;border-bottom:3px solid transparent;transition:all .15s;white-space:nowrap}}
.tab-btn:hover{{color:#e2e8f0;background:rgba(255,255,255,.06)}}
.tab-btn.active{{color:#fff;border-bottom-color:#3b82f6;background:rgba(255,255,255,.08)}}

/* ── Content ── */
.tab-content{{display:none;padding:20px}}
.tab-content.active{{display:block}}

/* ── Dashboard ── */
.dash-header{{display:flex;align-items:baseline;gap:12px;margin-bottom:14px}}
.dash-header h2{{font-size:17px;font-weight:700;color:#0f2044}}
.dash-meta{{font-size:11px;color:#64748b}}

.table-wrap{{overflow-x:auto;border-radius:10px;box-shadow:0 1px 6px rgba(0,0,0,.1)}}

/* ── Dashboard table ── */
.dash-table{{width:100%;border-collapse:collapse;background:#fff;font-size:12px}}
.dash-table thead tr.wk-header th{{
  background:#0f2044;color:#fff;padding:8px 10px;text-align:center;
  font-size:11px;font-weight:700;letter-spacing:.4px;white-space:nowrap;
  border-right:1px solid rgba(255,255,255,.1)}}
.dash-table thead tr.wk-header th:first-child{{text-align:left;border-right:2px solid rgba(255,255,255,.2);min-width:110px}}
.dash-table thead tr.wk-header th:last-child{{border-right:none}}
.dash-table tbody tr{{border-bottom:1px solid #f1f5f9;transition:background .1s}}
.dash-table tbody tr:hover{{background:#f0f7ff}}
.dash-table td{{padding:8px 10px;vertical-align:middle;border-right:1px solid #f1f5f9}}
.dash-table td:first-child{{font-weight:700;color:#0f2044;font-size:13px;border-right:2px solid #e2e8f0;white-space:nowrap}}
.dash-table td:last-child{{border-right:none;text-align:center;white-space:nowrap}}
.dash-table tbody tr:last-child td{{border-bottom:none}}

.week-cell{{display:flex;flex-direction:column;align-items:center;gap:3px;min-width:90px}}
.wc-top{{display:flex;align-items:baseline;gap:5px}}
.wc-planning{{font-weight:700;font-size:13px;color:#1e293b}}
.wc-sem{{font-size:10px;color:#94a3b8}}
.dif-chip{{display:inline-block;padding:1px 8px;border-radius:10px;font-size:10px;font-weight:700}}
.dif-chip.pos{{background:#dcfce7;color:#15803d}}
.dif-chip.neg{{background:#fee2e2;color:#b91c1c}}
.dif-chip.zero{{background:#f1f5f9;color:#94a3b8}}
.no-data-cell{{color:#d1d5db;font-size:11px;text-align:center}}

.detail-btn{{padding:4px 12px;background:#0f2044;color:#fff;border:none;border-radius:5px;
  font-size:11px;font-weight:600;cursor:pointer;transition:background .15s}}
.detail-btn:hover{{background:#1e3a6e}}

/* ── Family tabs ── */
.fam-controls{{display:flex;align-items:center;gap:10px;margin-bottom:14px;flex-wrap:wrap}}
.fam-controls h2{{font-size:17px;font-weight:700;color:#0f2044}}
.search-box{{padding:5px 11px;border:1px solid #cbd5e1;border-radius:6px;font-size:12px;
  width:190px;outline:none}}
.search-box:focus{{border-color:#3b82f6;box-shadow:0 0 0 2px rgba(59,130,246,.2)}}
.week-filter{{display:flex;gap:3px;flex-wrap:wrap}}
.week-btn{{padding:3px 9px;border:1px solid #cbd5e1;border-radius:5px;background:#fff;
  font-size:11px;cursor:pointer;font-weight:600;color:#475569}}
.week-btn:hover{{background:#e2e8f0}}
.week-btn.active{{background:#0f2044;color:#fff;border-color:#0f2044}}
.row-count{{font-size:11px;color:#64748b;margin-left:auto}}

/* ── Planning table ── */
.plan-table{{width:100%;border-collapse:collapse;background:#fff;font-size:12px}}
.plan-table thead tr:first-child th{{background:#0f2044;color:#fff;padding:8px 10px;text-align:left;
  white-space:nowrap;position:sticky;top:0;z-index:11;font-size:11px;font-weight:700;
  letter-spacing:.3px}}
.plan-table tbody tr{{border-bottom:1px solid #f1f5f9;transition:background .1s}}
.plan-table tbody tr:hover{{background:#f0f7ff}}
.plan-table tbody tr.retras{{background:#fff5f5}}
.plan-table tbody tr.retras:hover{{background:#fee2e2}}
.plan-table td{{padding:6px 10px;vertical-align:middle;white-space:nowrap}}

/* ── Cell styles ── */
.badge{{display:inline-block;padding:1px 7px;border-radius:10px;font-size:10px;font-weight:700;white-space:nowrap}}
.badge-asig  {{background:#dcfce7;color:#166534}}
.badge-stock {{background:#dbeafe;color:#1e40af}}
.badge-proc  {{background:#fef9c3;color:#854d0e}}
.badge-pdt   {{background:#fee2e2;color:#991b1b}}
.badge-pdtmq {{background:#fce7f3;color:#9d174d}}
.badge-txt   {{background:#f3f4f6;color:#374151;border:1px solid #e5e7eb}}
.badge-prio  {{background:#f97316;color:#fff}}

/* Sota comanda: always granate + white check */
.sota-check{{
  display:inline-flex;align-items:center;justify-content:center;
  width:20px;height:20px;border-radius:4px;
  background:#7B1D3A;color:#fff;font-size:13px;font-weight:700;
  line-height:1}}

/* Fase progress */
.fase-wrap{{display:flex;flex-direction:column;gap:2px}}
.fase-text{{font-size:11px;color:#374151;max-width:190px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}}
.fase-bar-bg{{height:3px;background:#e5e7eb;border-radius:2px;width:100px}}
.fase-bar{{height:3px;background:#3b82f6;border-radius:2px}}

/* ── Column filters ── */
.filter-row th{{background:#162d5a;padding:4px 6px;position:sticky;top:34px;z-index:9}}
.col-filter{{
  width:100%;padding:3px 6px;border:1px solid rgba(255,255,255,.18);
  border-radius:4px;background:rgba(255,255,255,.08);color:#e2e8f0;
  font-size:10px;outline:none;font-family:inherit}}
.col-filter::placeholder{{color:rgba(255,255,255,.35)}}
.col-filter:focus{{background:rgba(255,255,255,.18);border-color:#60a5fa}}

/* Scrollbar */
::-webkit-scrollbar{{height:6px;width:6px}}
::-webkit-scrollbar-track{{background:#f1f5f9}}
::-webkit-scrollbar-thumb{{background:#cbd5e1;border-radius:3px}}
</style>
</head>
<body>

<div class="top-bar">
  <h1>&#9783; Planning General – iVascular</h1>
  <span class="ts">Actualizado: {ts}</span>
</div>

<nav class="tab-nav" id="tab-nav">
  <button class="tab-btn active" data-id="dashboard" onclick="showTab('dashboard')">Dashboard</button>
</nav>

<div id="dashboard" class="tab-content active">
  <div class="dash-header">
    <h2>Resumen semanal</h2>
    <span class="dash-meta" id="dash-meta"></span>
  </div>
  <div class="table-wrap">
    <table class="dash-table" id="dash-table"></table>
  </div>
</div>

<script>
const RESUM_DATA   = {j_resum};
const WEEKS        = {j_weeks};
const FAM_ORDERED  = {j_fam_ord};
const COL_KEYS     = {j_col_keys};
const COL_HEADERS  = {j_col_hdr};
const FASES_MAP    = {j_fases};
const ALL_ROWS     = {j_all_rows};
const FAM_HAS_MAT2 = new Set({j_has_mat2});

// ── Tabs ──────────────────────────────────────────────────────────────────
const activeFilters  = {{}};
const searchState    = {{}};
const colFilterState = {{}};

function showTab(id) {{
  document.querySelectorAll('.tab-content').forEach(el => el.classList.remove('active'));
  document.querySelectorAll('.tab-btn').forEach(el => el.classList.remove('active'));
  if (!document.getElementById(id)) buildFamiliaTab(id);
  document.getElementById(id).classList.add('active');
  document.querySelectorAll('.tab-btn').forEach(btn => {{
    if (btn.dataset.id === id) btn.classList.add('active');
  }});
}}

function buildNav() {{
  const nav = document.getElementById('tab-nav');
  FAM_ORDERED.forEach(fam => {{
    const btn = document.createElement('button');
    btn.className = 'tab-btn';
    btn.dataset.id = fam;
    btn.textContent = fam;
    btn.onclick = () => showTab(fam);
    nav.appendChild(btn);
  }});
}}

// ── Dashboard ─────────────────────────────────────────────────────────────
function difChip(val) {{
  if (val === null || val === undefined) return '<span class="no-data-cell">—</span>';
  const n = Number(val);
  if (n > 0)  return `<span class="dif-chip pos">+${{n}}</span>`;
  if (n < 0)  return `<span class="dif-chip neg">${{n}}</span>`;
  return `<span class="dif-chip zero">0</span>`;
}}

function buildDashboard() {{
  const tbl = document.getElementById('dash-table');

  // Header row
  let thCols = `<th>Familia</th>`;
  WEEKS.forEach(w => {{ thCols += `<th>Semana ${{w}}</th>`; }});
  thCols += `<th></th>`;
  tbl.innerHTML = `<thead><tr class="wk-header">${{thCols}}</tr></thead><tbody id="dash-tbody"></tbody>`;

  const tbody = document.getElementById('dash-tbody');
  RESUM_DATA.forEach(fam => {{
    const hasTab = FAM_ORDERED.includes(fam.familia);
    const detailBtn = hasTab
      ? `<button class="detail-btn" onclick="showTab('${{fam.familia}}')">Ver detalle →</button>`
      : '';

    let cells = fam.weeks.map(w => {{
      if (w.sem === null && w.planning === null)
        return `<td><span class="no-data-cell">—</span></td>`;
      return `<td><div class="week-cell">
        <div class="wc-top">
          <span class="wc-planning">${{w.planning ?? '—'}}</span>
          <span class="wc-sem">/ ${{w.sem ?? '—'}}</span>
        </div>
        ${{difChip(w.dif)}}
      </div></td>`;
    }}).join('');

    tbody.insertAdjacentHTML('beforeend',
      `<tr><td>${{fam.familia}}</td>${{cells}}<td>${{detailBtn}}</td></tr>`);
  }});

  document.getElementById('dash-meta').textContent =
    `Semanas ${{WEEKS[0]}} – ${{WEEKS[WEEKS.length-1]}} · ${{RESUM_DATA.length}} familias`;
}}

// ── Family tab ────────────────────────────────────────────────────────────
function dispBadge(val) {{
  if (!val) return '';
  const v = (val+'').toLowerCase();
  if (v === 'asignado')           return `<span class="badge badge-asig">${{val}}</span>`;
  if (v === 'en stock')           return `<span class="badge badge-stock">${{val}}</span>`;
  if (v.includes('proceso'))      return `<span class="badge badge-proc">${{val}}</span>`;
  if (v.includes('pdt. mq'))      return `<span class="badge badge-pdtmq">${{val}}</span>`;
  if (v.startsWith('pdt'))        return `<span class="badge badge-pdt">${{val}}</span>`;
  return `<span class="badge">${{val}}</span>`;
}}

function txtBadge(val) {{
  if (!val) return '';
  if ((val+'').toLowerCase().includes('prio')) return `<span class="badge badge-prio">${{val}}</span>`;
  return `<span class="badge badge-txt">${{val}}</span>`;
}}

function faseCell(fam, fase) {{
  if (!fase) return '';
  const phases = FASES_MAP[fam] || [];
  const total  = phases.length;
  const idx    = phases.findIndex(p => fase.startsWith(p.name.substring(0, 12)));
  const pct    = (total > 0 && idx >= 0) ? Math.round(((idx + 1) / total) * 100) : 0;
  return `<div class="fase-wrap">
    <span class="fase-text" title="${{fase}}">${{fase}}</span>
    ${{total > 0 && idx >= 0 ? `<div class="fase-bar-bg"><div class="fase-bar" style="width:${{pct}}%"></div></div>` : ''}}
  </div>`;
}}

function renderTable(fam) {{
  const rows    = ALL_ROWS[fam] || [];
  const search  = (searchState[fam] || '').toLowerCase();
  const activeW = activeFilters[fam];
  const showMat2 = FAM_HAS_MAT2.has(fam);

  const colFilters = colFilterState[fam] || {{}};
  const filtered = rows.filter(r => {{
    if (activeW && activeW.size > 0 && !activeW.has(r.setmana)) return false;
    if (search && !Object.values(r).join(' ').toLowerCase().includes(search)) return false;
    for (const [k, v] of Object.entries(colFilters)) {{
      if (!v) continue;
      if (!String(r[k] ?? '').toLowerCase().includes(v.toLowerCase())) return false;
    }}
    return true;
  }});

  document.getElementById('rowcount-' + fam).textContent =
    `${{filtered.length}} / ${{rows.length}} filas`;

  const tbody = document.getElementById('tbody-' + fam);
  tbody.innerHTML = filtered.map(r => {{
    const cls = r._retras ? 'retras' : '';

    const cells = COL_KEYS.map((k, i) => {{
      // Skip mat2/disp2 columns if family has no mat2 data
      if ((k === 'mat2' || k === 'disp2') && !showMat2) return '';

      const val = r[k];
      if (k === 'disp1' || k === 'disp2')  return `<td>${{dispBadge(val)}}</td>`;
      if (k === 'txt')                      return `<td>${{txtBadge(val)}}</td>`;
      if (k === 'fase')                     return `<td>${{faseCell(r.familia || fam, val)}}</td>`;
      if (k === 'article')                  return `<td style="font-family:monospace;font-size:11px">${{val}}</td>`;
      if (k === 'sota_cmd')                 return `<td style="text-align:center">${{val === 'X' ? '<span class="sota-check">✓</span>' : ''}}</td>`;
      if (k === 'prep_linia')               return `<td style="text-align:center">${{val === 'X' ? '<span class="sota-check" style="background:#3b82f6">✓</span>' : ''}}</td>`;
      if (k === 'pzas_lot' || k === 'pzas_cmd') return `<td style="text-align:right;font-weight:600">${{val}}</td>`;
      if (k === 'setmana')                  return `<td style="text-align:center;font-weight:700;color:#0f2044">${{val}}</td>`;
      if (k === 'lot')                      return `<td style="font-family:monospace;font-size:11px">${{val}}</td>`;
      return `<td>${{val ?? ''}}</td>`;
    }}).join('');

    return `<tr class="${{cls}}">${{cells}}</tr>`;
  }}).join('');
}}

function buildFamiliaTab(fam) {{
  const rows    = ALL_ROWS[fam] || [];
  const showMat2 = FAM_HAS_MAT2.has(fam);
  const weeks   = [...new Set(rows.map(r => r.setmana).filter(Boolean))].sort((a,b)=>a-b);

  const weekBtns = ['Todas', ...weeks].map((w, idx) =>
    `<button class="week-btn ${{idx===0?'active':''}}" onclick="toggleWeek('${{fam}}',${{idx===0?'null':w}},this)">${{idx===0?'Todas':'Sem. '+w}}</button>`
  ).join('');

  const thCols = COL_KEYS.map((k, i) => {{
    if ((k === 'mat2' || k === 'disp2') && !showMat2) return '';
    return `<th>${{COL_HEADERS[i]}}</th>`;
  }}).join('');

  const filterCells = COL_KEYS.map((k) => {{
    if ((k === 'mat2' || k === 'disp2') && !showMat2) return '';
    return `<th><input class="col-filter" placeholder="▼" oninput="onColFilter('${{fam}}','${{k}}',this.value)"></th>`;
  }}).join('');

  const div = document.createElement('div');
  div.id = fam; div.className = 'tab-content';
  div.innerHTML = `
    <div class="fam-controls">
      <h2>${{fam}}</h2>
      <input class="search-box" placeholder="Buscar en todo..." oninput="onSearch('${{fam}}',this.value)">
      <div class="week-filter">${{weekBtns}}</div>
      <span class="row-count" id="rowcount-${{fam}}"></span>
    </div>
    <div class="table-wrap">
      <table class="plan-table">
        <thead>
          <tr>${{thCols}}</tr>
          <tr class="filter-row">${{filterCells}}</tr>
        </thead>
        <tbody id="tbody-${{fam}}"></tbody>
      </table>
    </div>`;
  document.body.appendChild(div);
  renderTable(fam);
}}

function toggleWeek(fam, week, btn) {{
  if (!activeFilters[fam]) activeFilters[fam] = new Set();
  const filter = activeFilters[fam];
  if (week === null) {{
    filter.clear();
    document.querySelectorAll(`#${{CSS.escape(fam)}} .week-btn`).forEach(b => b.classList.remove('active'));
    btn.classList.add('active');
  }} else {{
    document.querySelector(`#${{CSS.escape(fam)}} .week-btn:first-child`).classList.remove('active');
    filter.has(week) ? filter.delete(week) : filter.add(week);
    btn.classList.toggle('active', filter.has(week));
    if (filter.size === 0)
      document.querySelector(`#${{CSS.escape(fam)}} .week-btn:first-child`).classList.add('active');
  }}
  renderTable(fam);
}}

function onSearch(fam, val) {{
  searchState[fam] = val;
  renderTable(fam);
}}

function onColFilter(fam, key, val) {{
  if (!colFilterState[fam]) colFilterState[fam] = {{}};
  colFilterState[fam][key] = val;
  renderTable(fam);
}}

// ── Init ──────────────────────────────────────────────────────────────────
buildNav();
buildDashboard();
</script>
</body>
</html>"""

with open(OUT_HTML, 'w', encoding='utf-8') as f:
    f.write(html)

print(f"OK: {OUT_HTML}")
print(f"Familias dashboard: {len(resum_data)}")
print(f"Tabs con datos: {familias_ordered}")
print(f"Tabs con mat2: {sorted(fam_has_mat2)}")
