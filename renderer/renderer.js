'use strict';

// ─── State ────────────────────────────────────────────────────────────────────
let config    = { workspaces: [], savedQueries: [], settings: { githubToken: '', model: 'gpt-4o-mini', azPath: '', theme: 'night' } };
let lastRows  = [];
let lastCols  = [];
let sortState = { col: null, dir: 'asc' };
let hiddenCols = new Set();

// ─── Pagination state ─────────────────────────────────────────────────────────
let currentPage = 0;
let totalChunks = 0;
let totalRows   = 0;

// Columns to show by default (lowercased for matching)
const DEFAULT_VISIBLE = new Set([
  '_workspace', 'message', 'innermessage', 'outermessage',
  'timegenerated', 'timestamp', 'time',
  'severitylevel', 'level', 'itemtype', 'type',
  'problemid', 'operationname', 'operation_name',
]);

const WS_COLORS = [
  '#4c9be8', '#4ec994', '#f5a623', '#e05c6e',
  '#a78bfa', '#f472b6', '#34d399', '#fb923c',
  '#60a5fa', '#e879f9', '#facc15', '#2dd4bf',
];

function nextColor() {
  const used = new Set(config.workspaces.map(w => w.color));
  return WS_COLORS.find(c => !used.has(c)) || WS_COLORS[config.workspaces.length % WS_COLORS.length];
}

const DEFAULT_KQL = `union AppTraces, AppExceptions
| order by TimeGenerated desc`;

// ─── KQL limit detection ──────────────────────────────────────────────────────
const KQL_LIMIT_RE = /\|\s*(take|top)\s+\d+/i;

function checkKqlLimit() {
  const warn = document.getElementById('kqlLimitWarn');
  warn.style.display = KQL_LIMIT_RE.test(document.getElementById('kqlInput').value) ? '' : 'none';
}

// Strip take/top limits before sending to backend
function stripKqlLimits(kql) {
  return kql.replace(/\|\s*(take|top)\s+\d+[^\n]*/gi, '').trim();
}

// ─── Time range ───────────────────────────────────────────────────────────────
function toDatetimeLocal(date) {
  // Returns "YYYY-MM-DDTHH:MM" in local time for datetime-local inputs
  const pad = n => String(n).padStart(2, '0');
  return `${date.getFullYear()}-${pad(date.getMonth()+1)}-${pad(date.getDate())}T${pad(date.getHours())}:${pad(date.getMinutes())}`;
}

function setTimeRangeInputs(startDate, endDate) {
  document.getElementById('timeStart').value = toDatetimeLocal(startDate);
  document.getElementById('timeEnd').value   = toDatetimeLocal(endDate);
}

function getTimespan() {
  const s = document.getElementById('timeStart').value;
  const e = document.getElementById('timeEnd').value;
  if (s && e) {
    return `${new Date(s).toISOString()}/${new Date(e).toISOString()}`;
  }
  return 'P1D'; // fallback
}

function initTimeRange() {
  // Set default: last 24 h
  applyPreset(24);

  // Preset buttons
  document.querySelectorAll('.btn-time-preset').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.btn-time-preset').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      applyPreset(parseInt(btn.dataset.h, 10));
    });
  });

  // Manual edits deactivate preset highlight
  ['timeStart', 'timeEnd'].forEach(id => {
    document.getElementById(id).addEventListener('change', () => {
      document.querySelectorAll('.btn-time-preset').forEach(b => b.classList.remove('active'));
    });
  });
}

function applyPreset(hours) {
  const end   = new Date();
  const start = new Date(end.getTime() - hours * 3600 * 1000);
  setTimeRangeInputs(start, end);
}


// ─── Init ─────────────────────────────────────────────────────────────────────
(async () => {
  config = await api.config.load();
  renderWorkspaceList();
  checkLogin();

  if (!config.savedQueries) config.savedQueries = [];
  if (!config.settings.theme) config.settings.theme = 'night';
  document.getElementById('kqlInput').value = DEFAULT_KQL;
  initTimeRange();
  renderSavedQueryList();
  updatePatStatus();
  initMermaid();
  initPromptCards();
  updateModelBadge();
  applyTheme(config.settings.theme);

  // Update notification
  window.addEventListener('bau-update-available', ({ detail }) => {
    const banner = document.getElementById('updateBanner');
    document.getElementById('updateMsg').textContent =
      `BauInspector v${detail.latest} is available (you have v${detail.current}).`;
    document.getElementById('updateLink').href =
      'https://github.com/carlabarintos/bau-helper/releases/latest';
    document.getElementById('updateLink').addEventListener('click', (e) => {
      e.preventDefault();
      window.open('https://github.com/carlabarintos/bau-helper/releases/latest');
    });
    document.getElementById('updateDismiss').addEventListener('click', () => {
      banner.style.display = 'none';
    });
    banner.style.display = 'flex';
  });

  const cfgPath = await api.config.getPath();
  document.getElementById('configPath').textContent = cfgPath;
})();

// ─── Login ────────────────────────────────────────────────────────────────────
async function checkLogin() {
  const res = await api.az.checkLogin();
  updateLoginStatus(res);
}

function updateLoginStatus({ ok, user, subscription }) {
  const el = document.getElementById('loginStatus');
  if (ok) {
    el.textContent = `${user}  |  ${subscription}`;
    el.className = 'login-status ok';
    document.getElementById('btnLogin').textContent = 'Signed in';
  } else {
    el.textContent = 'Not signed in';
    el.className = 'login-status';
    document.getElementById('btnLogin').textContent = 'Sign in';
  }
}

document.getElementById('btnLogin').addEventListener('click', async () => {
  showLoading('Opening Azure sign-in…');
  const res = await api.az.login();
  hideLoading();
  updateLoginStatus(res);
  if (!res.ok) alert('Sign-in failed: ' + (res.error || 'unknown error'));
});

// ─── Workspaces ───────────────────────────────────────────────────────────────
function renderWorkspaceList() {
  const ul = document.getElementById('workspaceList');
  ul.innerHTML = '';

  if (!config.workspaces.length) {
    const li = document.createElement('li');
    li.className = 'workspace-item';
    li.style.justifyContent = 'center';
    li.innerHTML = '<span style="color:var(--text-muted);font-size:12px">No workspaces yet</span>';
    ul.appendChild(li);
    return;
  }

  config.workspaces.forEach((ws, idx) => {
    const li = document.createElement('li');
    li.className = 'workspace-item';
    const color     = ws.color || WS_COLORS[idx % WS_COLORS.length];
    const hasCustom = !!(ws.customKql && ws.customKql.trim());
    li.innerHTML = `
      <span class="ws-color-dot" style="background:${color}" title="${escHtml(ws.name)}"></span>
      <input type="checkbox" id="ws-${idx}" ${ws.enabled ? 'checked' : ''} />
      <label for="ws-${idx}" class="ws-info sidebar-label" style="cursor:pointer">
        <div class="ws-name">${escHtml(ws.name)}</div>
        <div class="ws-id">${escHtml(ws.id.slice(0, 18))}…</div>
      </label>
      <button class="ws-kql-btn sidebar-label${hasCustom ? ' has-custom' : ''}" data-idx="${idx}" title="${hasCustom ? 'Custom KQL set — click to edit' : 'Set custom KQL for this workspace'}">&#9998;</button>
      <button class="ws-delete sidebar-label" data-idx="${idx}" title="Remove workspace">&times;</button>
    `;
    ul.appendChild(li);

    li.querySelector('input').addEventListener('change', (e) => {
      config.workspaces[idx].enabled = e.target.checked;
      saveConfig();
    });
    li.querySelector('.ws-kql-btn').addEventListener('click', (e) => {
      e.stopPropagation();
      openWsKqlModal(idx);
    });
    li.querySelector('.ws-delete').addEventListener('click', (e) => {
      e.stopPropagation();
      config.workspaces.splice(idx, 1);
      saveConfig();
      renderWorkspaceList();
    });
  });
}

function getSelectedWorkspaces() {
  return config.workspaces.filter(ws => ws.enabled);
}

// ─── Saved Queries ────────────────────────────────────────────────────────────
function renderSavedQueryList() {
  const ul = document.getElementById('savedQueryList');
  ul.innerHTML = '';
  const queries = config.savedQueries || [];

  if (!queries.length) {
    const li = document.createElement('li');
    li.className = 'saved-query-item';
    li.style.justifyContent = 'center';
    li.innerHTML = '<span style="color:var(--text-muted);font-size:12px">No saved queries</span>';
    ul.appendChild(li);
    return;
  }

  queries.forEach((q, idx) => {
    const li = document.createElement('li');
    li.className = 'saved-query-item';
    li.title = q.kql;
    li.innerHTML = `
      <div class="sq-info">
        <div class="sq-name">${escHtml(q.name)}</div>
        <div class="sq-preview">${escHtml(q.kql.replace(/\n/g, ' '))}</div>
        <div class="sq-timespan">${timespanLabel(q.timespan)}</div>
      </div>
      <button class="sq-delete" data-idx="${idx}" title="Delete">&times;</button>
    `;

    li.addEventListener('click', (e) => {
      if (e.target.classList.contains('sq-delete')) return;
      document.getElementById('kqlInput').value = q.kql;
      // Restore saved timespan: if it's a duration string convert to absolute dates
      if (q.timespan && q.timespan.includes('/')) {
        const [s, e] = q.timespan.split('/');
        setTimeRangeInputs(new Date(s), new Date(e));
      } else if (q.timespan) {
        const hrs = { PT1H:1, PT6H:6, PT12H:12, P1D:24, P7D:168, P30D:720 }[q.timespan] || 24;
        applyPreset(hrs);
      }
    });

    li.querySelector('.sq-delete').addEventListener('click', (e) => {
      e.stopPropagation();
      config.savedQueries.splice(idx, 1);
      saveConfig();
      renderSavedQueryList();
    });

    ul.appendChild(li);
  });
}

function timespanLabel(ts) {
  const map = { PT1H: 'Last 1 h', PT6H: 'Last 6 h', PT12H: 'Last 12 h', P1D: 'Last 24 h', P7D: 'Last 7 d', P30D: 'Last 30 d' };
  return map[ts] || ts;
}

// ─── Workspace custom KQL modal ───────────────────────────────────────────────
let _wsKqlIdx = null;

function openWsKqlModal(idx) {
  _wsKqlIdx = idx;
  const ws = config.workspaces[idx];
  document.getElementById('wsKqlTitle').textContent  = ws.name;
  document.getElementById('wsKqlInput').value        = ws.customKql || '';
  openModal('modalWsKql');
  setTimeout(() => document.getElementById('wsKqlInput').focus(), 100);
}

document.getElementById('btnSaveWsKql').addEventListener('click', () => {
  if (_wsKqlIdx === null) return;
  config.workspaces[_wsKqlIdx].customKql = document.getElementById('wsKqlInput').value.trim();
  saveConfig();
  renderWorkspaceList();
  closeModal('modalWsKql');
});

document.getElementById('btnClearWsKql').addEventListener('click', () => {
  document.getElementById('wsKqlInput').value = '';
});

document.getElementById('btnSaveQuery').addEventListener('click', () => {
  const kql = document.getElementById('kqlInput').value.trim();
  if (!kql) { alert('Nothing to save — write a query first.'); return; }
  document.getElementById('saveQueryName').value = '';
  document.getElementById('saveQueryPreview').textContent = kql;
  openModal('modalSaveQuery');
  setTimeout(() => document.getElementById('saveQueryName').focus(), 100);
});

document.getElementById('btnConfirmSaveQuery').addEventListener('click', () => {
  const name = document.getElementById('saveQueryName').value.trim();
  if (!name) { alert('Enter a name for the query.'); return; }
  const kql      = document.getElementById('kqlInput').value.trim();
  const timespan = getTimespan();
  if (!config.savedQueries) config.savedQueries = [];
  config.savedQueries.push({ name, kql, timespan });
  saveConfig();
  renderSavedQueryList();
  closeModal('modalSaveQuery');
});

document.getElementById('saveQueryName').addEventListener('keydown', (e) => {
  if (e.key === 'Enter') document.getElementById('btnConfirmSaveQuery').click();
});

// ─── Add workspace modal ───────────────────────────────────────────────────────
document.getElementById('btnAddWorkspace').addEventListener('click', () => {
  document.getElementById('wsName').value  = '';
  document.getElementById('wsId').value    = '';
  document.getElementById('wsSubId').value = '';
  renderColorPicker(nextColor());
  openModal('modalAddWorkspace');
});

let selectedColor = WS_COLORS[0];

function renderColorPicker(preselect) {
  selectedColor = preselect;
  const container = document.getElementById('wsColorPicker');
  container.innerHTML = '';
  WS_COLORS.forEach(color => {
    const swatch = document.createElement('div');
    swatch.className = 'color-swatch' + (color === preselect ? ' active' : '');
    swatch.style.background = color;
    swatch.title = color;
    swatch.addEventListener('click', () => {
      container.querySelectorAll('.color-swatch').forEach(s => s.classList.remove('active'));
      swatch.classList.add('active');
      selectedColor = color;
    });
    container.appendChild(swatch);
  });
}

document.getElementById('btnSaveWorkspace').addEventListener('click', () => {
  const name  = document.getElementById('wsName').value.trim();
  const id    = document.getElementById('wsId').value.trim();
  const subId = document.getElementById('wsSubId').value.trim();

  if (!name || !id) {
    alert('Please fill in Name and Workspace ID.');
    return;
  }

  config.workspaces.push({ name, id, subId, color: selectedColor, enabled: true });
  saveConfig();
  renderWorkspaceList();
  closeModal('modalAddWorkspace');
});

// ─── Discover workspaces ───────────────────────────────────────────────────────
document.getElementById('btnDiscoverWs').addEventListener('click', async () => {
  openModal('modalDiscover');
  const sel = document.getElementById('discoverSubSelect');
  sel.innerHTML = '<option value="">Loading subscriptions…</option>';

  const res = await api.az.listSubscriptions();
  if (res.ok) {
    sel.innerHTML = '<option value="">— All subscriptions —</option>';
    res.data.forEach(sub => {
      const opt = document.createElement('option');
      opt.value = sub.id;
      opt.textContent = `${sub.name} (${sub.id.slice(0, 8)}…)`;
      sel.appendChild(opt);
    });
  } else {
    sel.innerHTML = '<option value="">Sign in first</option>';
  }
});

document.getElementById('btnDiscoverFetch').addEventListener('click', async () => {
  const subId = document.getElementById('discoverSubSelect').value;
  const resultsDiv = document.getElementById('discoverResults');
  resultsDiv.innerHTML = '<div style="color:var(--text-dim);font-size:12px">Fetching workspaces…</div>';

  const res = await api.az.listWorkspaces(subId || null);
  if (!res.ok) {
    resultsDiv.innerHTML = `<div style="color:var(--error);font-size:12px">${escHtml(res.error)}</div>`;
    return;
  }

  if (!res.data.length) {
    resultsDiv.innerHTML = '<div style="color:var(--text-muted);font-size:12px">No workspaces found.</div>';
    return;
  }

  resultsDiv.innerHTML = '';
  res.data.forEach(ws => {
    const alreadyAdded = config.workspaces.some(w => w.id === ws.customerId);
    const div = document.createElement('div');
    div.className = 'discover-item';
    div.innerHTML = `
      <div class="discover-item-info">
        <div class="discover-name">${escHtml(ws.name)}</div>
        <div class="discover-id">${escHtml(ws.customerId || ws.id || '')}</div>
      </div>
      <button class="btn-primary btn-sm" ${alreadyAdded ? 'disabled' : ''} data-ws-name="${escAttr(ws.name)}" data-ws-id="${escAttr(ws.customerId || ws.id)}" data-ws-sub="${escAttr(subId)}">
        ${alreadyAdded ? 'Added' : 'Add'}
      </button>
    `;
    div.querySelector('button').addEventListener('click', (e) => {
      const btn = e.currentTarget;
      config.workspaces.push({ name: btn.dataset.wsName, id: btn.dataset.wsId, subId: btn.dataset.wsSub, enabled: true });
      saveConfig();
      renderWorkspaceList();
      btn.textContent = 'Added';
      btn.disabled = true;
    });
    resultsDiv.appendChild(div);
  });
});

// ─── Query & Run ──────────────────────────────────────────────────────────────
document.getElementById('btnRun').addEventListener('click', runQuery);
document.getElementById('btnClear').addEventListener('click', () => {
  document.getElementById('kqlInput').value = '';
  document.getElementById('queryMeta').textContent = '';
});

document.getElementById('kqlInput').addEventListener('keydown', (e) => {
  if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') runQuery();
});
document.getElementById('kqlInput').addEventListener('input', checkKqlLimit);

async function runQuery() {
  const rawDefaultKql = document.getElementById('kqlInput').value.trim();
  if (!rawDefaultKql) { alert('Enter a KQL query.'); return; }

  const selected = getSelectedWorkspaces();
  if (!selected.length) { alert('Select at least one workspace from the sidebar.'); return; }

  const workspaceIds = selected.map(ws => {
    const raw = (ws.customKql && ws.customKql.trim()) ? ws.customKql.trim() : rawDefaultKql;
    return { id: ws.id, name: ws.name, kql: stripKqlLimits(raw) };
  });

  const timespan = getTimespan();
  const meta     = document.getElementById('queryMeta');
  meta.textContent = '';
  currentPage = 0;
  totalChunks = 0;
  totalRows   = 0;

  const wsLabel = `${selected.length} workspace${selected.length > 1 ? 's' : ''}`;
  showLoading(`Querying ${wsLabel}…`);

  await api.cache.clear();

  const t0  = Date.now();
  const res = await api.az.query({ workspaceIds, timespan });
  const elapsed = ((Date.now() - t0) / 1000).toFixed(1);
  hideLoading();

  if (res.errors && res.errors.length) {
    document.getElementById('tableErrors').textContent =
      `Errors: ${res.errors.map(e => `${e.workspace}: ${e.error}`).join(' | ')}`;
  } else {
    document.getElementById('tableErrors').textContent = '';
  }

  // Save all rows to chunked cache files
  showLoading('Saving to cache…');
  const saved = await api.cache.save(res.rows);
  hideLoading();

  totalChunks = saved.chunks || 1;
  totalRows   = res.rows.length;
  meta.textContent = `${totalRows} rows · ${wsLabel} · ${elapsed}s`;

  await loadAndShowPage(0);
}

async function loadAndShowPage(page) {
  showLoading(`Loading page ${page + 1}…`);
  const result = await api.cache.loadChunk(page);
  hideLoading();
  currentPage = page;
  lastRows    = result.ok ? result.rows : [];
  renderTable(lastRows);
}

// ─── Table rendering ──────────────────────────────────────────────────────────
let wsColorMap = {}; // workspace name → color, rebuilt before each render

function buildWsColorMap() {
  wsColorMap = {};
  config.workspaces.forEach((ws, idx) => {
    wsColorMap[ws.name] = ws.color || WS_COLORS[idx % WS_COLORS.length];
  });
}

function renderTable(rows) {
  const emptyState   = document.getElementById('emptyState');
  const table        = document.getElementById('resultsTable');
  const thead        = document.getElementById('resultsHead');
  const tbody        = document.getElementById('resultsBody');
  const rowCountEl   = document.getElementById('rowCount');

  buildWsColorMap();

  if (!rows.length) {
    emptyState.style.display = 'flex';
    emptyState.querySelector('div:last-child').textContent = 'Query returned no results.';
    table.style.display = 'none';
    rowCountEl.textContent = '0 rows';
    lastCols = [];
    renderPagination();
    return;
  }

  emptyState.style.display = 'none';
  table.style.display = 'table';
  rowCountEl.textContent = `${rows.length} rows`;

  // Collect columns — put _workspace first
  const colSet = new Set();
  rows.forEach(r => Object.keys(r).forEach(k => colSet.add(k)));
  const allCols = ['_workspace', ...Array.from(colSet).filter(c => c !== '_workspace')];

  // Rebuild column structure only when the column set changes
  const colsChanged = JSON.stringify(allCols) !== JSON.stringify(lastCols);
  if (colsChanged) {
    lastCols  = allCols;
    sortState = { col: null, dir: 'asc' };
    hiddenCols = new Set();
    const anyMatch = allCols.some(c => DEFAULT_VISIBLE.has(c.toLowerCase()));
    if (anyMatch) {
      allCols.forEach(c => { if (!DEFAULT_VISIBLE.has(c.toLowerCase())) hiddenCols.add(c); });
    } else {
      allCols.slice(5).forEach(c => hiddenCols.add(c));
    }
    buildColPanel(allCols);
    renderHeader(allCols, thead);
  }

  renderRows(rows, allCols, tbody);
  renderPagination();
}

function renderPagination() {
  const bar = document.getElementById('pagination');
  bar.innerHTML = '';
  if (totalChunks <= 1) return;

  const prev = document.createElement('button');
  prev.className = 'btn-page';
  prev.textContent = '‹ Prev';
  prev.disabled = currentPage === 0;
  prev.addEventListener('click', () => loadAndShowPage(currentPage - 1));
  bar.appendChild(prev);

  // Page number buttons — show a window of 7 around current
  for (let p = 0; p < totalChunks; p++) {
    if (p !== 0 && p !== totalChunks - 1 && Math.abs(p - currentPage) > 3) {
      if (p === 1 || p === totalChunks - 2) {
        const dots = document.createElement('span');
        dots.className = 'page-info';
        dots.textContent = '…';
        bar.appendChild(dots);
      }
      continue;
    }
    const btn = document.createElement('button');
    btn.className = 'btn-page' + (p === currentPage ? ' active' : '');
    btn.textContent = p + 1;
    const pg = p;
    btn.addEventListener('click', () => loadAndShowPage(pg));
    bar.appendChild(btn);
  }

  const next = document.createElement('button');
  next.className = 'btn-page';
  next.textContent = 'Next ›';
  next.disabled = currentPage === totalChunks - 1;
  next.addEventListener('click', () => loadAndShowPage(currentPage + 1));
  bar.appendChild(next);

  const info = document.createElement('span');
  info.className = 'page-info';
  info.textContent = `Page ${currentPage + 1} / ${totalChunks}  ·  ${totalRows} total rows`;
  bar.appendChild(info);
}

function visibleCols(allCols) {
  return allCols.filter(c => !hiddenCols.has(c));
}

function renderHeader(allCols, thead) {
  thead.innerHTML = '';
  const tr = document.createElement('tr');
  visibleCols(allCols).forEach(col => {
    const th = document.createElement('th');
    th.textContent = col === '_workspace' ? 'Workspace' : col;
    th.dataset.col = col;
    if (sortState.col === col) th.classList.add(`sorted-${sortState.dir}`);
    th.addEventListener('click', () => sortTable(col));
    tr.appendChild(th);
  });
  thead.appendChild(tr);
}

function buildColPanel(allCols) {
  const panel = document.getElementById('colPanel');
  panel.innerHTML = '';

  const header = document.createElement('div');
  header.className = 'col-panel-header';
  header.innerHTML = '<span>Show / Hide Columns</span>';
  const resetBtn = document.createElement('button');
  resetBtn.className = 'col-panel-reset';
  resetBtn.textContent = 'Reset';
  resetBtn.addEventListener('click', () => {
    const anyMatch = allCols.some(c => DEFAULT_VISIBLE.has(c.toLowerCase()));
    hiddenCols = new Set();
    if (anyMatch) {
      allCols.forEach(c => { if (!DEFAULT_VISIBLE.has(c.toLowerCase())) hiddenCols.add(c); });
    }
    buildColPanel(allCols);
    renderHeader(allCols, document.getElementById('resultsHead'));
    renderTable(lastRows);
  });
  header.appendChild(resetBtn);
  panel.appendChild(header);

  allCols.forEach(col => {
    const item = document.createElement('label');
    item.className = 'col-panel-item';
    const cb = document.createElement('input');
    cb.type = 'checkbox';
    cb.checked = !hiddenCols.has(col);
    cb.addEventListener('change', () => {
      if (cb.checked) hiddenCols.delete(col);
      else hiddenCols.add(col);
      renderHeader(allCols, document.getElementById('resultsHead'));
      renderTable(lastRows);
    });
    const lbl = document.createElement('span');
    lbl.textContent = col === '_workspace' ? 'Workspace' : col;
    item.appendChild(cb);
    item.appendChild(lbl);
    panel.appendChild(item);
  });
}

// Column panel open/close
document.getElementById('btnColumns').addEventListener('click', (e) => {
  e.stopPropagation();
  document.getElementById('colPanel').classList.toggle('open');
});
document.addEventListener('click', () => {
  document.getElementById('colPanel').classList.remove('open');
});
document.getElementById('colPanel').addEventListener('click', e => e.stopPropagation());

// ─── Severity helpers ─────────────────────────────────────────────────────────
const SEV_MAP = {
  error:    { cls: 'sev-error',   label: 'ERR'  },
  critical: { cls: 'sev-error',   label: 'CRIT' },
  fatal:    { cls: 'sev-error',   label: 'FATAL'},
  warning:  { cls: 'sev-warning', label: 'WARN' },
  warn:     { cls: 'sev-warning', label: 'WARN' },
  information: { cls: 'sev-info', label: 'INFO' },
  info:     { cls: 'sev-info',    label: 'INFO' },
  verbose:  { cls: 'sev-verbose', label: 'VERB' },
  debug:    { cls: 'sev-verbose', label: 'DBG'  },
};
const SEV_COLS = ['severitylevel', 'level', 'severity', 'itemtype', 'type'];

function getSeverity(row) {
  for (const key of SEV_COLS) {
    const val = row[Object.keys(row).find(k => k.toLowerCase() === key)];
    if (val) {
      const norm = String(val).toLowerCase().trim();
      if (SEV_MAP[norm]) return SEV_MAP[norm];
      // numeric Azure severity: 0=critical,1=error,2=warning,3=info,4=verbose
      if (norm === '0') return { cls: 'sev-verbose', label: 'VERB' };
      if (norm === '1') return { cls: 'sev-info',    label: 'INFO' };
      if (norm === '2') return { cls: 'sev-warning', label: 'WARN' };
      if (norm === '3') return { cls: 'sev-error',   label: 'ERR'  };
      if (norm === '4') return { cls: 'sev-error',   label: 'CRIT' };
    }
  }
  return null;
}

function renderRows(rows, allCols, tbody) {
  const cols = visibleCols(allCols);
  tbody.innerHTML = '';
  const frag = document.createDocumentFragment();

  rows.forEach(row => {
    const wsName  = row['_workspace'] || '';
    const wsColor = wsColorMap[wsName] || 'var(--accent)';
    const sev     = getSeverity(row);

    // ── Summary row ──────────────────────────────────────────────────────────
    const tr = document.createElement('tr');
    tr.className = 'data-row';
    tr.style.borderLeft = `3px solid ${wsColor}`;
    if (sev) tr.classList.add(sev.cls + '-row');

    // Badge goes in workspace cell when visible; otherwise in the first data cell
    const wsVisible   = cols.includes('_workspace');
    let   badgeNeeded = sev && !wsVisible; // only inject into data cell if ws col hidden

    cols.forEach(col => {
      const td  = document.createElement('td');
      const val = row[col];

      if (col === '_workspace') {
        td.className = 'ws-col';
        td.innerHTML = `<span class="ws-row-dot" style="background:${wsColor}"></span><span style="color:${wsColor}">${escHtml(wsName)}</span>`;
        if (sev) td.innerHTML += ` <span class="sev-badge ${sev.cls}">${sev.label}</span>`;
        td.title = wsName;
      } else {
        const text = val === null || val === undefined ? '' : String(val);
        if (badgeNeeded) {
          td.innerHTML  = `<span class="sev-badge ${sev.cls}">${sev.label}</span> ${escHtml(text)}`;
          td.title      = text;
          badgeNeeded   = false;
        } else {
          td.textContent = text;
          td.title       = text;
        }
      }
      tr.appendChild(td);
    });

    // ── Expand row (hidden detail panel) ─────────────────────────────────────
    const expandTr = document.createElement('tr');
    expandTr.className = 'expand-row';
    const expandTd = document.createElement('td');
    expandTd.colSpan = cols.length;
    expandTd.className = 'expand-cell';

    // Build detail table of all fields
    const detail = document.createElement('div');
    detail.className = 'expand-detail';
    Object.entries(row).forEach(([k, v]) => {
      if (v === null || v === undefined || v === '') return;
      const item = document.createElement('div');
      item.className = 'expand-field';
      item.innerHTML = `<span class="expand-key">${escHtml(k)}</span><span class="expand-val">${escHtml(String(v))}</span>`;
      detail.appendChild(item);
    });
    expandTd.appendChild(detail);
    expandTr.appendChild(expandTd);

    // Toggle expand on row click
    tr.style.cursor = 'pointer';
    tr.addEventListener('click', () => {
      const open = expandTr.classList.toggle('open');
      tr.classList.toggle('expanded', open);
    });

    frag.appendChild(tr);
    frag.appendChild(expandTr);
  });
  tbody.appendChild(frag);
}

function sortTable(col) {
  if (sortState.col === col) {
    sortState.dir = sortState.dir === 'asc' ? 'desc' : 'asc';
  } else {
    sortState.col = col;
    sortState.dir = 'asc';
  }

  // Update header classes
  document.querySelectorAll('#resultsHead th').forEach(th => {
    th.classList.remove('sorted-asc', 'sorted-desc');
    if (th.dataset.col === col) th.classList.add(`sorted-${sortState.dir}`);
  });

  lastRows = [...lastRows].sort((a, b) => {
    const av = a[col] ?? '';
    const bv = b[col] ?? '';
    const cmp = String(av).localeCompare(String(bv), undefined, { numeric: true });
    return sortState.dir === 'asc' ? cmp : -cmp;
  });

  renderTable(lastRows);
}

// ─── AI Summary ───────────────────────────────────────────────────────────────
document.getElementById('btnSummarize').addEventListener('click', () => summarize(false));
document.getElementById('btnSummarizeAll').addEventListener('click', () => summarize(true));

async function summarize(allRows) {
  const input     = document.getElementById('aiPrompt');
  const promptType = activePromptType;
  const prompt    = promptType === 'dashboard'
    ? (input.dataset.dashboardKey || 'health')
    : (input.value.trim() || 'Summarize this log data. Highlight errors, anomalies, and patterns.');

  const resultEl = document.getElementById('aiResult');
  resultEl.className = 'ai-result loading';
  resultEl.textContent = 'Calling GitHub Copilot / Models API…';

  let logContext = null;
  const SAMPLE_CHAR_LIMIT = 12000;
  const knownTotal = totalRows || lastRows.length;
  if (knownTotal > 0) {
    let rows;
    if (allRows && totalChunks > 1) {
      rows = [];
      for (let i = 0; i < totalChunks; i++) {
        const chunk = await api.cache.loadChunk(i);
        if (chunk.ok) rows.push(...chunk.rows);
      }
    } else {
      rows = allRows ? lastRows : lastRows.slice(0, 100);
    }

    // Build sample and count how many rows actually fit in the char budget
    const fullSample = rows.map(r => JSON.stringify(r)).join('\n');
    const truncated  = fullSample.length > SAMPLE_CHAR_LIMIT;
    const sample     = fullSample.slice(0, SAMPLE_CHAR_LIMIT);
    const rowsFit    = truncated
      ? sample.split('\n').length - 1   // last line may be cut mid-row
      : rows.length;

    logContext = { count: knownTotal, sample };

    // Show context note before response arrives
    const contextNote = [
      `<div class="ai-context-note">`,
      `Sending <strong>${rowsFit.toLocaleString()}</strong> of <strong>${knownTotal.toLocaleString()}</strong> rows as context`,
      truncated
        ? ` &nbsp;<span class="ai-context-warn">&#9888; sample truncated at ${SAMPLE_CHAR_LIMIT.toLocaleString()} characters — AI sees a partial view</span>`
        : knownTotal > rowsFit
        ? ` &nbsp;<span class="ai-context-warn">&#9888; remaining ${(knownTotal - rowsFit).toLocaleString()} rows not included</span>`
        : '',
      `</div>`,
    ].join('');
    resultEl.innerHTML = contextNote + '<div class="ai-loading-msg">Calling GitHub Copilot / Models API…</div>';
  }

  const res = await api.ai.summarize({ prompt, logContext, promptType });

  // Preserve the context note div if it was injected
  const existingNote = resultEl.querySelector('.ai-context-note');
  const noteHtml = existingNote ? existingNote.outerHTML : '';

  if (res.ok) {
    resultEl.className = 'ai-result';

    if (res.promptType === 'dashboard') {
      renderDashboard(res.content, resultEl);
      if (noteHtml) resultEl.insertAdjacentHTML('afterbegin', noteHtml);
    } else if (res.promptType === 'diagram') {
      resultEl.innerHTML = noteHtml;
      resultEl.appendChild(await renderMermaidBlock(res.content));
    } else {
      await renderMarkdownAsync(res.content, resultEl);
      if (noteHtml) resultEl.insertAdjacentHTML('afterbegin', noteHtml);
    }

    document.getElementById('btnSaveImage').style.display = '';
  } else {
    resultEl.className = 'ai-result error';
    resultEl.innerHTML = noteHtml + `<div>Error: ${escHtml(res.error)}</div>`;
    document.getElementById('btnSaveImage').style.display = 'none';
  }
}

// ─── Dashboard renderer ───────────────────────────────────────────────────────
function renderDashboard(raw, container) {
  // Strip possible markdown code fences the model may add despite instructions
  const cleaned = raw.replace(/^```json?\s*/i, '').replace(/```\s*$/, '').trim();
  let data;
  try {
    data = JSON.parse(cleaned);
  } catch {
    // Fallback: render as markdown if JSON parse fails
    renderMarkdownAsync(raw, container);
    return;
  }

  const STATUS_COLOR = { ok: 'var(--success)', warn: 'var(--warning)', error: 'var(--error)', info: 'var(--accent)' };

  const dash = document.createElement('div');
  dash.className = 'dashboard-view';

  if (data.title) {
    const h = document.createElement('div');
    h.className = 'dash-title';
    h.textContent = data.title;
    dash.appendChild(h);
  }

  (data.sections || []).forEach(section => {
    if (section.type === 'metrics' && Array.isArray(section.cards)) {
      const grid = document.createElement('div');
      grid.className = 'dash-metrics';
      section.cards.forEach(card => {
        const el = document.createElement('div');
        el.className = 'dash-metric-card';
        const color = STATUS_COLOR[card.status] || 'var(--accent)';
        el.style.borderTopColor = color;
        el.innerHTML = `
          <div class="dash-metric-value" style="color:${color}">${escHtml(card.value)}</div>
          <div class="dash-metric-label">${escHtml(card.label)}</div>
          ${card.sublabel ? `<div class="dash-metric-sub">${escHtml(card.sublabel)}</div>` : ''}
        `;
        grid.appendChild(el);
      });
      dash.appendChild(grid);
    } else if (section.type === 'table' && Array.isArray(section.rows)) {
      const wrap = document.createElement('div');
      wrap.className = 'dash-section';
      if (section.title) {
        const t = document.createElement('div');
        t.className = 'dash-section-title';
        t.textContent = section.title;
        wrap.appendChild(t);
      }
      const tbl = document.createElement('table');
      tbl.className = 'dash-table';
      if (section.headers?.length) {
        const thead = document.createElement('thead');
        thead.innerHTML = '<tr>' + section.headers.map(h => `<th>${escHtml(h)}</th>`).join('') + '</tr>';
        tbl.appendChild(thead);
      }
      const tbody = document.createElement('tbody');
      section.rows.forEach(row => {
        const tr = document.createElement('tr');
        tr.innerHTML = (Array.isArray(row) ? row : [row]).map(c => `<td>${escHtml(String(c ?? ''))}</td>`).join('');
        tbody.appendChild(tr);
      });
      tbl.appendChild(tbody);
      wrap.appendChild(tbl);
      dash.appendChild(wrap);
    } else if (section.type === 'list' && Array.isArray(section.items)) {
      const wrap = document.createElement('div');
      wrap.className = 'dash-section';
      if (section.title) {
        const t = document.createElement('div');
        t.className = 'dash-section-title';
        t.textContent = section.title;
        wrap.appendChild(t);
      }
      const ul = document.createElement('ul');
      ul.className = 'dash-list';
      section.items.forEach(item => {
        const li = document.createElement('li');
        li.textContent = item;
        ul.appendChild(li);
      });
      wrap.appendChild(ul);
      dash.appendChild(wrap);
    } else if (section.type === 'text' && section.body) {
      const wrap = document.createElement('div');
      wrap.className = 'dash-section dash-text';
      if (section.title) {
        const t = document.createElement('div');
        t.className = 'dash-section-title';
        t.textContent = section.title;
        wrap.appendChild(t);
      }
      const p = document.createElement('p');
      p.textContent = section.body;
      wrap.appendChild(p);
      dash.appendChild(wrap);
    }
  });

  container.innerHTML = '';
  container.appendChild(dash);
}

// ─── Mermaid ──────────────────────────────────────────────────────────────────
function initMermaid() {
  if (typeof mermaid === 'undefined') return;
  mermaid.initialize({
    startOnLoad: false,
    theme: 'neutral',
    securityLevel: 'loose',
    fontSize: 14,
    xychart: { width: 800, height: 400, useWidth: 800 },
    pie:     { textPosition: 0.8 },
    timeline: { useWidth: 800 },
  });
}

// Sanitise AI-generated Mermaid code before rendering.
// Models often add %%{init}%%, markdown fences, or use wrong diagram keywords.
function sanitiseMermaid(raw) {
  let code = raw.trim();
  // Strip %%{init: ...}%% front-matter blocks
  code = code.replace(/%%\{[\s\S]*?\}%%\s*/gm, '').trim();
  // Strip any wrapping ```mermaid ... ``` fences that slipped through
  code = code.replace(/^```mermaid\s*/i, '').replace(/^```\s*/i, '').replace(/```\s*$/i, '').trim();
  // AI sometimes outputs a bare "bar" keyword — upgrade to xychart-beta
  if (/^bar\b/.test(code)) {
    code = 'xychart-beta\n' + code.replace(/^bar\b/, '').trim();
  }
  return code;
}

let mermaidId = 0;
async function renderMermaidBlock(rawCode) {
  const code = sanitiseMermaid(rawCode);

  const wrap = document.createElement('div');
  wrap.className = 'mermaid-block';

  const copyBtn = document.createElement('button');
  copyBtn.className = 'mermaid-copy-btn';
  copyBtn.textContent = 'Copy diagram code';
  copyBtn.addEventListener('click', () => {
    navigator.clipboard.writeText(code);
    copyBtn.textContent = 'Copied!';
    setTimeout(() => { copyBtn.textContent = 'Copy diagram code'; }, 1500);
  });

  // mermaid.render(id, text, container) passes code as a plain string (no HTML
  // encoding), and creates its temp rendering element INSIDE `offscreen`, so
  // D3-based charts (xychart-beta) can measure the 900 px parent width correctly.
  const offscreen = document.createElement('div');
  offscreen.style.cssText = 'position:fixed;left:-9999px;top:0;width:900px;background:#fff;';
  document.body.appendChild(offscreen);

  try {
    const id = `mermaid-${++mermaidId}`;
    const { svg } = await mermaid.render(id, code, offscreen);
    wrap.innerHTML = svg;

    const svgEl = wrap.querySelector('svg');
    if (svgEl) {
      svgEl.style.maxWidth = '100%';
      svgEl.style.display  = 'block';
      svgEl.style.margin   = '0 auto';
    }
    wrap.prepend(copyBtn);
  } catch (err) {
    console.warn('Mermaid render failed:', err);
    wrap.className = 'mermaid-error-wrap';
    const errNote = document.createElement('div');
    errNote.className   = 'mermaid-err-msg';
    errNote.textContent = `⚠ Diagram render failed: ${err?.message || String(err)}`;
    const raw = document.createElement('pre');
    raw.className = 'mermaid-raw';
    raw.textContent = code;
    wrap.appendChild(errNote);
    wrap.appendChild(copyBtn);
    wrap.appendChild(raw);
  } finally {
    if (offscreen.parentNode) document.body.removeChild(offscreen);
  }
  return wrap;
}

// Markdown renderer — splits on ```mermaid blocks, renders each async
async function renderMarkdownAsync(text, container) {
  const parts  = text.split(/(```mermaid[\s\S]*?```)/g);
  const frag   = document.createDocumentFragment();

  for (const part of parts) {
    const mermaidMatch = part.match(/^```mermaid\n?([\s\S]*?)```$/);
    if (mermaidMatch) {
      const code = mermaidMatch[1].trim();
      const el   = await renderMermaidBlock(code);
      frag.appendChild(el);
    } else if (part.trim()) {
      const div  = document.createElement('div');
      div.innerHTML = renderMarkdownText(part);
      frag.appendChild(div);
    }
  }

  container.innerHTML = '';
  container.appendChild(frag);
}

function renderMarkdownText(text) {
  let html = escHtml(text);
  html = html.replace(/^### (.+)$/gm, '<h3>$1</h3>');
  html = html.replace(/^## (.+)$/gm,  '<h2>$1</h2>');
  html = html.replace(/^# (.+)$/gm,   '<h1>$1</h1>');
  html = html.replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>');
  html = html.replace(/`([^`]+)`/g, '<code>$1</code>');
  html = html.replace(/^[-*] (.+)$/gm, '<li>$1</li>');
  html = html.replace(/(<li>[\s\S]*?<\/li>\n?)+/g, m => `<ul>${m}</ul>`);
  html = html.replace(/^\|(.+)\|$/gm, row => {
    if (/^[\s|:-]+$/.test(row)) return ''; // separator row
    const cells = row.split('|').filter((_, i, a) => i > 0 && i < a.length - 1);
    return '<tr>' + cells.map(c => `<td>${c.trim()}</td>`).join('') + '</tr>';
  });
  html = html.replace(/(<tr>[\s\S]*?<\/tr>\n?)+/g, m => `<table class="ai-table">${m}</table>`);
  html = html.replace(/\n\n+/g, '</p><p>').replace(/\n/g, '<br>');
  return `<p>${html}</p>`;
}

// ─── Model selector ───────────────────────────────────────────────────────────
function updateModelBadge() {
  const sel = document.getElementById('modelSelect');
  if (!sel) return;
  const model = config.settings.model || 'gpt-4o-mini';
  // Select matching option, or add a custom one if not in the list
  const exists = Array.from(sel.options).some(o => o.value === model);
  if (!exists) {
    const opt = document.createElement('option');
    opt.value = model; opt.textContent = model;
    sel.appendChild(opt);
  }
  sel.value = model;
}

document.getElementById('modelSelect').addEventListener('change', async () => {
  config.settings.model = document.getElementById('modelSelect').value;
  // Keep settings modal in sync
  const settingsSel = document.getElementById('settingsModel');
  if (settingsSel) settingsSel.value = config.settings.model;
  await saveConfig();
});

// ─── Prompt cards ─────────────────────────────────────────────────────────────
let activePromptType = 'summary';

function initPromptCards() {
  document.querySelectorAll('.prompt-card').forEach(card => {
    card.addEventListener('click', () => {
      document.querySelectorAll('.prompt-card').forEach(c => c.classList.remove('active'));
      card.classList.add('active');
      activePromptType = card.dataset.type || 'summary';
      const input = document.getElementById('aiPrompt');
      // dashboard cards use a key; others put the full prompt in the input
      if (activePromptType === 'dashboard') {
        input.value = '';
        input.placeholder = `Dashboard: ${card.textContent.trim()} (${card.dataset.prompt})`;
        input.dataset.dashboardKey = card.dataset.prompt;
      } else {
        input.value = card.dataset.prompt;
        input.placeholder = 'Ask about the logs, or click a template on the left…';
        delete input.dataset.dashboardKey;
      }
      input.focus();
    });
  });
}

// ─── Settings modal ───────────────────────────────────────────────────────────
document.getElementById('btnSettings').addEventListener('click', async () => {
  const cfg = await api.config.load();
  document.getElementById('settingsToken').value  = cfg.settings.githubToken || '';
  document.getElementById('settingsModel').value  = cfg.settings.model || 'gpt-4o-mini';
  document.getElementById('settingsAzPath').value = cfg.settings.azPath || '';
  openModal('modalSettings');
});

document.getElementById('btnSaveSettings').addEventListener('click', async () => {
  config.settings.githubToken = document.getElementById('settingsToken').value.trim();
  config.settings.model       = document.getElementById('settingsModel').value;
  // Keep inline selector in sync
  document.getElementById('modelSelect').value = config.settings.model;
  config.settings.azPath      = document.getElementById('settingsAzPath').value.trim();
  await saveConfig();
  updateModelBadge();
  closeModal('modalSettings');
});

// ─── Tabs ─────────────────────────────────────────────────────────────────────
document.querySelectorAll('.tab').forEach(tab => {
  tab.addEventListener('click', () => {
    document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
    document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
    tab.classList.add('active');
    document.getElementById(`tab-${tab.dataset.tab}`).classList.add('active');
  });
});

// ─── Modal helpers ────────────────────────────────────────────────────────────
function openModal(id)  { document.getElementById(id).classList.add('open'); }
function closeModal(id) { document.getElementById(id).classList.remove('open'); }

document.querySelectorAll('[data-close]').forEach(btn => {
  btn.addEventListener('click', () => closeModal(btn.dataset.close));
});
document.querySelectorAll('.modal-backdrop').forEach(bd => {
  bd.addEventListener('click', (e) => { if (e.target === bd) closeModal(bd.id); });
});

// ─── Loading ──────────────────────────────────────────────────────────────────
function showLoading(msg = 'Loading…') {
  document.getElementById('loadingMsg').textContent = msg;
  document.getElementById('loadingOverlay').classList.add('show');
}
function hideLoading() {
  document.getElementById('loadingOverlay').classList.remove('show');
}

// ─── Config save ──────────────────────────────────────────────────────────────
async function saveConfig() {
  await api.config.save(config);
}

// ─── Save as image ────────────────────────────────────────────────────────────

// Rasterise every inline SVG to a PNG <img> so html2canvas can capture it.
async function svgToCanvas(svg) {
  svg.setAttribute('xmlns', 'http://www.w3.org/2000/svg');

  // Prefer getBoundingClientRect (actual rendered px) — getAttribute may return "100%" which parseFloat
  // misreads as 100 px. This must be called while the SVG is in the live DOM.
  const rect = svg.getBoundingClientRect();
  const vb   = svg.viewBox?.baseVal;
  const attrW = parseFloat(svg.getAttribute('width'));
  const attrH = parseFloat(svg.getAttribute('height'));
  const w = Math.round((attrW > 1 ? attrW : null) || rect.width  || vb?.width  || 820);
  const h = Math.round((attrH > 1 ? attrH : null) || rect.height || vb?.height || 320);

  // Stamp explicit px dimensions so the serialised SVG has an intrinsic size
  svg.setAttribute('width',  w);
  svg.setAttribute('height', h);

  const svgStr  = new XMLSerializer().serializeToString(svg);
  // Blob URL is more permissive than data URL for SVGs with embedded <style> blocks
  const blob    = new Blob([svgStr], { type: 'image/svg+xml;charset=utf-8' });
  const blobUrl = URL.createObjectURL(blob);

  try {
    const tmpImg = new Image(w, h);
    tmpImg.src   = blobUrl;
    await new Promise((res, rej) => {
      tmpImg.onload  = res;
      tmpImg.onerror = rej;
      setTimeout(res, 3000); // fallback timeout
    });

    const cvs = document.createElement('canvas');
    cvs.width  = w * 2;
    cvs.height = h * 2;
    const ctx  = cvs.getContext('2d');
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, cvs.width, cvs.height);
    ctx.scale(2, 2);
    ctx.drawImage(tmpImg, 0, 0, w, h);
    return { pngUrl: cvs.toDataURL('image/png'), w, h };
  } finally {
    URL.revokeObjectURL(blobUrl);
  }
}

async function svgsToImages(el) {
  const backups = [];
  for (const svg of Array.from(el.querySelectorAll('svg'))) {
    try {
      const { pngUrl, w, h } = await svgToCanvas(svg);
      const img = document.createElement('img');
      img.src   = pngUrl;
      img.width  = w;
      img.height = h;
      img.style.cssText = `width:${w}px;height:${h}px;display:block;margin:auto;`;
      backups.push({ parent: svg.parentNode, next: svg.nextSibling, svg, img });
      svg.parentNode.replaceChild(img, svg);
    } catch (e) {
      console.warn('SVG rasterise failed:', e);
    }
  }
  return backups;
}

function restoreSvgs(backups) {
  for (const { parent, next, svg, img } of backups) {
    try { parent.insertBefore(svg, next || null); parent.removeChild(img); } catch { /* ignore */ }
  }
}

document.getElementById('btnSaveImage').addEventListener('click', async () => {
  const resultEl = document.getElementById('aiResult');
  const btn      = document.getElementById('btnSaveImage');

  btn.textContent = 'Capturing…';
  btn.disabled    = true;

  let captureEl = null;
  try {
    const isDayTheme = document.body.classList.contains('day');
    const bg         = isDayTheme ? '#f0f2f8' : '#1c1e2a';
    const textColor  = isDayTheme ? '#1a1d2e' : '#e2e4f0';

    // Clone into an off-screen fixed-width div so html2canvas sees a clean,
    // scrollable-content-sized element with no flex stretching (which caused
    // the right-side whitespace when ai-result was flex:1).
    captureEl = document.createElement('div');
    captureEl.style.cssText = [
      'position: fixed',
      'left: -9999px',
      'top: 0',
      `background: ${bg}`,
      `color: ${textColor}`,
      'font-family: system-ui, sans-serif',
      'font-size: 15px',
      'line-height: 1.7',
      'padding: 24px 28px',
      'width: 860px',          // fixed width — gives SVGs a reliable layout context
      'box-sizing: border-box',
      'white-space: pre-wrap',
    ].join(';');

    // Stamp explicit px dimensions on live SVGs BEFORE cloning —
    // getBoundingClientRect is accurate here since the element is in the real DOM.
    resultEl.querySelectorAll('svg').forEach(svg => {
      const rect = svg.getBoundingClientRect();
      const vb   = svg.viewBox?.baseVal;
      const attrW = parseFloat(svg.getAttribute('width'));
      const attrH = parseFloat(svg.getAttribute('height'));
      const w = Math.round((attrW > 1 ? attrW : null) || rect.width  || vb?.width  || 820);
      const h = Math.round((attrH > 1 ? attrH : null) || rect.height || vb?.height || 320);
      svg.setAttribute('width',  w);
      svg.setAttribute('height', h);
    });

    captureEl.innerHTML = resultEl.innerHTML;

    // Remove "Copy diagram code" buttons — they clutter the image
    captureEl.querySelectorAll('.mermaid-copy-btn').forEach(btn => btn.remove());
    document.body.appendChild(captureEl);

    // Replace inline SVGs with rasterised PNG <img> elements
    const svgBackups = await svgsToImages(captureEl);

    const canvas = await html2canvas(captureEl, {
      backgroundColor: bg,
      scale: 2,
      useCORS: false,
      logging: false,
      width:       captureEl.scrollWidth,
      height:      captureEl.scrollHeight,
      windowWidth: captureEl.scrollWidth,
      windowHeight: captureEl.scrollHeight,
    });

    restoreSvgs(svgBackups);

    const dataUrl   = canvas.toDataURL('image/png');
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
    const res = await api.image.save({ dataUrl, suggestedName: `bau-summary-${timestamp}.png` });

    if (res.ok) {
      btn.textContent = '✓ Saved';
      setTimeout(() => { btn.textContent = '📷 Save image'; btn.disabled = false; }, 2000);
    } else {
      btn.textContent = '📷 Save image';
      btn.disabled = false;
    }
  } catch (e) {
    console.error('Image save error:', e);
    btn.textContent = '📷 Save image';
    btn.disabled = false;
  } finally {
    if (captureEl) document.body.removeChild(captureEl);
  }
});

// ─── Theme ────────────────────────────────────────────────────────────────────
function applyTheme(theme) {
  const btn = document.getElementById('btnTheme');
  if (theme === 'day') {
    document.body.classList.add('day');
    if (btn) btn.textContent = '☀';
    btn?.setAttribute('title', 'Switch to night theme');
  } else {
    document.body.classList.remove('day');
    if (btn) btn.textContent = '☾';
    btn?.setAttribute('title', 'Switch to day theme');
  }
}

document.getElementById('btnTheme').addEventListener('click', async () => {
  const next = config.settings.theme === 'day' ? 'night' : 'day';
  config.settings.theme = next;
  applyTheme(next);
  await saveConfig();
});

// ─── Sidebar collapse ─────────────────────────────────────────────────────────
document.getElementById('btnCollapseSidebar').addEventListener('click', () => {
  document.getElementById('sidebar').classList.toggle('collapsed');
});

// ─── PAT (sidebar) ────────────────────────────────────────────────────────────
function updatePatStatus() {
  const el = document.getElementById('patStatus');
  el.textContent = config.settings.githubToken ? '✓ set' : '';
}

document.getElementById('btnSavePat').addEventListener('click', async () => {
  const val = document.getElementById('sidebarPat').value.trim();
  if (!val) return;
  config.settings.githubToken = val;
  await saveConfig();
  document.getElementById('sidebarPat').value = '';
  updatePatStatus();
});

document.getElementById('sidebarPat').addEventListener('keydown', (e) => {
  if (e.key === 'Enter') document.getElementById('btnSavePat').click();
});

// ─── Utils ────────────────────────────────────────────────────────────────────
function escHtml(str) {
  return String(str ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}
function escAttr(str) { return escHtml(str); }
