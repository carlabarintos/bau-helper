const { app, BrowserWindow, ipcMain, shell, dialog } = require('electron');
const { exec, spawn } = require('child_process');
const path = require('path');
const fs = require('fs');
const os = require('os');
const https = require('https');

// ─── Config persistence ──────────────────────────────────────────────────────

const CONFIG_PATH = path.join(app.getPath('userData'), 'bau-inspector-config.json');

function loadConfig() {
  try {
    return JSON.parse(fs.readFileSync(CONFIG_PATH, 'utf8'));
  } catch {
    return { workspaces: [], settings: { githubToken: '', model: 'gpt-4o-mini', azPath: '' } };
  }
}

function saveConfig(data) {
  fs.writeFileSync(CONFIG_PATH, JSON.stringify(data, null, 2), 'utf8');
}

// ─── Shell helper ────────────────────────────────────────────────────────────

function run(cmd, opts = {}) {
  return new Promise((resolve, reject) => {
    const env = { ...process.env, ...opts.env };
    exec(cmd, { timeout: 120000, maxBuffer: 50 * 1024 * 1024, env }, (err, stdout, stderr) => {
      if (err) {
        reject(new Error(stderr || err.message));
      } else {
        resolve(stdout.trim());
      }
    });
  });
}

// resolved once at first use
let _azBinResolved = null;

async function resolveAzBin() {
  if (_azBinResolved) return _azBinResolved;
  const cfg = loadConfig();
  if (cfg.settings.azPath) {
    _azBinResolved = cfg.settings.azPath;
    return _azBinResolved;
  }
  // ask Windows where az.cmd actually lives
  try {
    const found = await run('where az.cmd');
    _azBinResolved = found.split('\n')[0].trim();
    return _azBinResolved;
  } catch { /* fall through */ }
  try {
    const found = await run('where az');
    _azBinResolved = found.split('\n')[0].trim();
    return _azBinResolved;
  } catch { /* fall through */ }
  _azBinResolved = 'az.cmd'; // last resort
  return _azBinResolved;
}

// ─── Azure helpers ───────────────────────────────────────────────────────────

async function azRun(args) {
  const az = await resolveAzBin();
  return run(`"${az}" ${args} -o json`);
}

// KQL query — writes KQL to a .kql file, streams az stdout to a temp output file.
// Streaming to a file bypasses Node's exec maxBuffer limit for large result sets.
async function azQuery(workspaceId, kql, timespan) {
  const bin    = await resolveAzBin();
  const uid    = `${Date.now()}-${Math.random().toString(36).slice(2)}`;
  const tmpKql = path.join(os.tmpdir(), `bau-kql-${uid}.kql`);
  const tmpOut = path.join(os.tmpdir(), `bau-out-${uid}.json`);
  try {
    fs.writeFileSync(tmpKql, kql, 'utf8');
    const cmd = `"${bin}" monitor log-analytics query --workspace "${workspaceId}" --analytics-query "@${tmpKql}" --timespan "${timespan}" -o json`;
    await new Promise((resolve, reject) => {
      const child = spawn(cmd, { shell: true, windowsHide: true });
      const outStream = fs.createWriteStream(tmpOut);
      let stderr = '';
      child.stdout.pipe(outStream);
      child.stderr.on('data', d => { stderr += String(d); });
      child.on('close', code => {
        outStream.end();
        if (code !== 0) reject(new Error(stderr || `az exit code ${code}`));
        else resolve();
      });
      child.on('error', reject);
      setTimeout(() => { child.kill(); reject(new Error('az query timed out')); }, 300000);
    });
    return fs.readFileSync(tmpOut, 'utf8');
  } finally {
    try { fs.unlinkSync(tmpKql); } catch { /* ignore */ }
    try { fs.unlinkSync(tmpOut); } catch { /* ignore */ }
  }
}

ipcMain.handle('az:checkLogin', async () => {
  try {
    const out = await azRun('account show');
    const acct = JSON.parse(out);
    return { ok: true, user: acct.user?.name, subscription: acct.name };
  } catch {
    return { ok: false };
  }
});

ipcMain.handle('az:login', async () => {
  try {
    const bin = await resolveAzBin();
    await run(`"${bin}" login`);
    const out = await azRun('account show');
    const acct = JSON.parse(out);
    return { ok: true, user: acct.user?.name, subscription: acct.name };
  } catch (e) {
    return { ok: false, error: e.message };
  }
});

ipcMain.handle('az:listSubscriptions', async () => {
  try {
    const out = await azRun('account list');
    return { ok: true, data: JSON.parse(out) };
  } catch (e) {
    return { ok: false, error: e.message };
  }
});

ipcMain.handle('az:listWorkspaces', async (_e, subscriptionId) => {
  try {
    const sub = subscriptionId ? `--subscription "${subscriptionId}"` : '';
    const out = await azRun(`monitor log-analytics workspace list ${sub}`);
    return { ok: true, data: JSON.parse(out) };
  } catch (e) {
    return { ok: false, error: e.message };
  }
});

ipcMain.handle('az:query', async (_e, { workspaceIds, timespan }) => {
  const ts = timespan || 'P1D';

  // Each workspace carries its own kql (set by the renderer)
  const settled = await Promise.allSettled(
    workspaceIds.map(ws => azQuery(ws.id, ws.kql, ts).then(raw => ({ ws, raw })))
  );

  const results = [];
  const errors  = [];

  for (const outcome of settled) {
    if (outcome.status === 'rejected') {
      errors.push({ workspace: '(unknown)', error: outcome.reason?.message || String(outcome.reason) });
      continue;
    }
    const { ws, raw } = outcome.value;
    try {
      const parsed = JSON.parse(raw);
      let normalized;
      if (Array.isArray(parsed) && parsed.length > 0 && typeof parsed[0] === 'object' && !Array.isArray(parsed[0])) {
        normalized = parsed.map(r => ({ _workspace: ws.name, ...r }));
      } else {
        const table = Array.isArray(parsed) ? parsed[0] : parsed?.tables?.[0];
        const cols  = table?.columns?.map(c => c.name) || [];
        const rows  = table?.rows || [];
        normalized  = rows.map(row => {
          const obj = { _workspace: ws.name };
          cols.forEach((col, i) => { obj[col] = row[i]; });
          return obj;
        });
      }
      results.push(...normalized);
    } catch (e) {
      errors.push({ workspace: ws.name, error: e.message });
    }
  }

  // Merge and sort by the first recognised timestamp column, newest first
  const TS_COLS = ['TimeGenerated', 'timestamp', 'time', 'Timestamp', 'eventTimestamp'];
  const tsCol   = results.length ? TS_COLS.find(c => c in results[0]) : null;
  if (tsCol) {
    results.sort((a, b) => {
      const ta = a[tsCol] ? new Date(a[tsCol]).getTime() : 0;
      const tb = b[tsCol] ? new Date(b[tsCol]).getTime() : 0;
      return tb - ta; // newest first
    });
  }

  return { ok: true, rows: results, errors };
});

// ─── GitHub Copilot / Models AI ──────────────────────────────────────────────

const DASHBOARD_SCHEMA = `Return ONLY a JSON object — no markdown, no explanation, no code fences — matching this schema exactly:
{
  "title": "string",
  "sections": [
    { "type": "metrics", "cards": [{ "label": "string", "value": "string", "sublabel": "string", "status": "ok|warn|error|info" }] },
    { "type": "table",   "title": "string", "headers": ["string"], "rows": [["string"]] },
    { "type": "list",    "title": "string", "items": ["string"] },
    { "type": "text",    "title": "string", "body": "string" }
  ]
}`;

const DASHBOARD_PROMPTS = {
  health:      'Produce a health dashboard with: key metrics (total events, error count, error rate, warning count), top 5 errors table (message, count, workspace), activity by workspace table, and recommendations list.',
  performance: 'Produce a performance dashboard with: key metrics (slowest operation avg ms, timeout count, p95 estimate), slowest operations table (operation, avg duration, count), timeout patterns list, and optimisation recommendations.',
  errors:      'Produce an error breakdown dashboard with: key metrics (total errors, unique error types, most affected workspace), error frequency table (message, count, workspace, first seen), affected components list, and suggested fixes.',
  security:    'Produce a security overview dashboard with: key metrics (auth failures, suspicious IPs, unusual access count), suspicious events table (time, event, user, workspace), risk indicators list, and recommended actions.',
};

ipcMain.handle('ai:summarize', async (_e, { prompt, logContext, promptType }) => {
  const cfg = loadConfig();
  const token = cfg.settings.githubToken;
  if (!token) {
    return { ok: false, error: 'No GitHub token configured. Add it in Settings.' };
  }

  const model = cfg.settings.model || 'gpt-4o-mini';
  const isDashboard = promptType === 'dashboard';

  const isDiagram = promptType === 'diagram';

  const systemMsg = isDashboard
    ? `You are an expert Azure log analyst. Analyse the log data and produce a structured dashboard. ${DASHBOARD_SCHEMA}`
    : isDiagram
    ? `You are a Mermaid diagram generator. Output ONLY raw Mermaid syntax — no markdown fences, no \`\`\`, no %%{init}%% blocks, no explanations.

SYNTAX RULES (Mermaid v11):

Pie chart:
pie title "My Title"
    "Label A" : 42
    "Label B" : 58

Bar chart (xychart-beta):
xychart-beta
    title "My Title"
    x-axis ["LabelA", "LabelB", "LabelC"]
    y-axis "Count" 0 --> 500
    bar [120, 340, 80]

Timeline:
timeline
    title Key Events
    section 2026-03-31
        Job host stopped : 2026-03-31T18:27:06
        Service stopping : 2026-03-31T18:27:08
    section 2026-04-01
        Service restarted : 2026-04-01T09:00:00

Flowchart:
flowchart LR
    A[Start] --> B{Decision}
    B -- Yes --> C[End]
    B -- No --> D[Error]

IMPORTANT:
- Keep ALL labels under 20 characters — truncate workspace names if needed
- Use only real numeric values from the data
- Output nothing except the diagram code`
    : `You are an expert Azure log analyst and security investigator.
Analyse the provided log data concisely. Highlight anomalies, errors, patterns, and actionable insights.
Be precise and structured. Use bullet points and markdown tables where helpful.`;

  const resolvedPrompt = isDashboard ? (DASHBOARD_PROMPTS[prompt] || DASHBOARD_PROMPTS.health) : prompt;

  const userMsg = logContext
    ? `${resolvedPrompt}\n\nLog data (sample, ${logContext.count} total rows):\n\`\`\`\n${logContext.sample}\n\`\`\``
    : resolvedPrompt;

  const body = JSON.stringify({
    model,
    messages: [
      { role: 'system', content: systemMsg },
      { role: 'user', content: userMsg },
    ],
    max_tokens: 1500,
    temperature: 0.3,
  });

  return new Promise((resolve) => {
    const options = {
      hostname: 'models.inference.ai.azure.com',
      path: '/chat/completions',
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${token}`,
        'Content-Length': Buffer.byteLength(body),
      },
    };

    const req = https.request(options, (res) => {
      let data = '';
      res.on('data', chunk => { data += chunk; });
      res.on('end', () => {
        try {
          const json = JSON.parse(data);
          if (json.error) {
            resolve({ ok: false, error: json.error.message || JSON.stringify(json.error) });
          } else {
            const content = json.choices?.[0]?.message?.content || '';
            resolve({ ok: true, content, model, promptType: isDashboard ? 'dashboard' : promptType });
          }
        } catch (e) {
          resolve({ ok: false, error: `Parse error: ${e.message}\nRaw: ${data.slice(0, 200)}` });
        }
      });
    });

    req.on('error', e => resolve({ ok: false, error: e.message }));
    req.setTimeout(60000, () => { req.destroy(); resolve({ ok: false, error: 'Request timed out' }); });
    req.write(body);
    req.end();
  });
});

// ─── Results cache (chunked folder) ─────────────────────────────────────────

const CACHE_DIR  = path.join(app.getPath('userData'), 'bau-cache');
const CHUNK_SIZE = 500; // rows per chunk file

function cacheClear() {
  try {
    if (fs.existsSync(CACHE_DIR)) {
      for (const f of fs.readdirSync(CACHE_DIR)) {
        try { fs.unlinkSync(path.join(CACHE_DIR, f)); } catch { /* ignore */ }
      }
    }
  } catch { /* ignore */ }
}

ipcMain.handle('cache:save', (_e, rows) => {
  try {
    cacheClear();
    fs.mkdirSync(CACHE_DIR, { recursive: true });
    const chunks = Math.ceil(rows.length / CHUNK_SIZE) || 1;
    for (let i = 0; i < chunks; i++) {
      const slice = rows.slice(i * CHUNK_SIZE, (i + 1) * CHUNK_SIZE);
      fs.writeFileSync(
        path.join(CACHE_DIR, `chunk-${String(i).padStart(5, '0')}.json`),
        JSON.stringify(slice), 'utf8'
      );
    }
    fs.writeFileSync(
      path.join(CACHE_DIR, 'meta.json'),
      JSON.stringify({ total: rows.length, chunks }), 'utf8'
    );
    return { ok: true, chunks, total: rows.length };
  } catch (e) {
    return { ok: false, error: e.message };
  }
});

ipcMain.handle('cache:clear', () => {
  cacheClear();
  return { ok: true };
});

ipcMain.handle('cache:info', () => {
  try {
    const meta = JSON.parse(fs.readFileSync(path.join(CACHE_DIR, 'meta.json'), 'utf8'));
    return { ok: true, ...meta };
  } catch {
    return { ok: false, chunks: 0, total: 0 };
  }
});

ipcMain.handle('cache:loadChunk', (_e, index) => {
  try {
    const file = path.join(CACHE_DIR, `chunk-${String(index).padStart(5, '0')}.json`);
    const rows = JSON.parse(fs.readFileSync(file, 'utf8'));
    return { ok: true, rows };
  } catch (e) {
    return { ok: false, rows: [], error: e.message };
  }
});

// ─── Config IPC ─────────────────────────────────────────────────────────────

ipcMain.handle('config:load', () => loadConfig());

ipcMain.handle('config:save', (_e, data) => {
  saveConfig(data);
  _azBinResolved = null; // re-resolve on next call in case az path changed
  return { ok: true };
});

ipcMain.handle('config:getPath', () => CONFIG_PATH);

// ─── Image save ──────────────────────────────────────────────────────────────

ipcMain.handle('image:save', async (_e, { dataUrl, suggestedName }) => {
  const { filePath, canceled } = await dialog.showSaveDialog({
    title: 'Save Summary as Image',
    defaultPath: suggestedName || `bau-summary-${Date.now()}.png`,
    filters: [
      { name: 'PNG Image', extensions: ['png'] },
      { name: 'JPEG Image', extensions: ['jpg', 'jpeg'] },
    ],
  });

  if (canceled || !filePath) return { ok: false };

  const isJpeg  = filePath.toLowerCase().endsWith('.jpg') || filePath.toLowerCase().endsWith('.jpeg');
  const base64  = dataUrl.replace(/^data:image\/\w+;base64,/, '');
  fs.writeFileSync(filePath, Buffer.from(base64, 'base64'));

  return { ok: true, filePath };
});

// ─── Update check ────────────────────────────────────────────────────────────

function checkForUpdates(win) {
  const current = app.getVersion();
  const options = {
    hostname: 'api.github.com',
    path: '/repos/carlabarintos/bau-helper/releases/latest',
    headers: { 'User-Agent': 'BauInspector' },
  };
  https.get(options, (res) => {
    let data = '';
    res.on('data', c => { data += c; });
    res.on('end', () => {
      try {
        const latest = JSON.parse(data).tag_name?.replace(/^v/, '');
        if (latest && latest !== current) {
          win.webContents.send('update-available', { current, latest });
        }
      } catch { /* ignore parse errors */ }
    });
  }).on('error', () => { /* no network, skip silently */ });
}

// ─── Window ──────────────────────────────────────────────────────────────────

function createWindow() {
  const win = new BrowserWindow({
    width: 1400,
    height: 900,
    minWidth: 900,
    minHeight: 600,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
    },
    title: 'BauInspector',
    backgroundColor: '#0f1117',
    show: false,
  });

  win.loadFile(path.join(__dirname, 'renderer', 'index.html'));
  win.once('ready-to-show', () => {
    win.show();
    setTimeout(() => checkForUpdates(win), 3000); // check after UI settles
  });

  // Open external links in browser
  win.webContents.setWindowOpenHandler(({ url }) => {
    shell.openExternal(url);
    return { action: 'deny' };
  });
}

app.whenReady().then(createWindow);
app.on('window-all-closed', () => { if (process.platform !== 'darwin') app.quit(); });
app.on('activate', () => { if (BrowserWindow.getAllWindows().length === 0) createWindow(); });
