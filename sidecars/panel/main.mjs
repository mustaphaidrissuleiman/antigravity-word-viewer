import { existsSync, mkdirSync, readdirSync, readFileSync, statSync, writeFileSync } from 'node:fs';
import { homedir } from 'node:os';
import { basename, dirname, extname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { Response, SidecarApp } from 'sidecar_sdk';

const HERE = dirname(fileURLToPath(import.meta.url));
const DATA_DIR = process.env.ANTIGRAVITY_EXECUTABLE_DATA_DIR || join(HERE, '.data');
const STATE_FILE = join(DATA_DIR, 'state.json');

const app = new SidecarApp();

// ---------------------------------------------------------------------------
// Static Files & Page Routes
// ---------------------------------------------------------------------------

const readLocal = (name) => readFileSync(join(HERE, name), 'utf8');

const serveIndex = () => readLocal('index.html');
const serveAppJs = () => readLocal('app.js');
const serveStyles = () => new Response(readLocal('styles.css'), { contentType: 'text/css' });

let cachedJsZip = null;
let cachedDocxPreview = null;
let cachedMammoth = null;

const serveJsZip = () => {
  if (!cachedJsZip) cachedJsZip = readLocal('lib/jszip.min.js');
  return cachedJsZip;
};

const serveDocxPreview = () => {
  if (!cachedDocxPreview) cachedDocxPreview = readLocal('lib/docx-preview.min.js');
  return cachedDocxPreview;
};

const serveMammoth = () => {
  if (!cachedMammoth) cachedMammoth = readLocal('lib/mammoth.browser.min.js');
  return cachedMammoth;
};

// HTML Pages
app.page('/', serveIndex);
app.page('/index.html', serveIndex);
app.page('/open', serveIndex);
app.page('/viewer', serveIndex);
app.page('/panel', serveIndex);
app.page('/panel/', serveIndex);
app.page('/panel/index.html', serveIndex);

// Client App Script & CSS
app.page('/app.js', serveAppJs);
app.page('/panel/app.js', serveAppJs);

app.api('/styles.css', serveStyles, 'GET');
app.api('/panel/styles.css', serveStyles, 'GET');

// Vendored Libraries
app.page('/lib/jszip.min.js', serveJsZip);
app.page('/panel/lib/jszip.min.js', serveJsZip);

app.page('/lib/docx-preview.min.js', serveDocxPreview);
app.page('/panel/lib/docx-preview.min.js', serveDocxPreview);

app.page('/lib/mammoth.browser.min.js', serveMammoth);
app.page('/panel/lib/mammoth.browser.min.js', serveMammoth);

// ---------------------------------------------------------------------------
// Path Resolution Helpers
// ---------------------------------------------------------------------------

const SUPPORTED_EXTS = new Set(['.docx', '.doc', '.dotx', '.dotm', '.docm']);

function isWordDoc(filename) {
  if (!filename || typeof filename !== 'string') return false;
  const ext = extname(filename).toLowerCase();
  return SUPPORTED_EXTS.has(ext);
}

function searchForDocByName(dirPath, targetName, maxDepth = 4, currentDepth = 0) {
  if (currentDepth > maxDepth || !existsSync(dirPath)) return null;
  try {
    const entries = readdirSync(dirPath, { withFileTypes: true });
    for (const entry of entries) {
      if (entry.name.startsWith('.') || entry.name === 'node_modules' || entry.name === 'vendor') continue;
      const full = join(dirPath, entry.name);
      if (entry.isDirectory()) {
        const found = searchForDocByName(full, targetName, maxDepth, currentDepth + 1);
        if (found) return found;
      } else if (entry.isFile() && entry.name.toLowerCase() === targetName.toLowerCase()) {
        return full;
      }
    }
  } catch {}
  return null;
}

function resolveDocPath(inputPath, workspaceUris = []) {
  if (!inputPath || typeof inputPath !== 'string') return null;

  let cleaned = inputPath.trim();

  // Strip wrapping quotes (double, single, backticks)
  if ((cleaned.startsWith('"') && cleaned.endsWith('"')) ||
      (cleaned.startsWith("'") && cleaned.endsWith("'")) ||
      (cleaned.startsWith('`') && cleaned.endsWith('`'))) {
    cleaned = cleaned.slice(1, -1).trim();
  }

  // Handle file:// URIs
  if (cleaned.startsWith('file://')) {
    try {
      cleaned = fileURLToPath(cleaned);
    } catch {
      cleaned = cleaned.replace(/^file:\/\//, '');
    }
  }

  // Handle tilde ~
  if (cleaned === '~' || cleaned.startsWith('~/')) {
    cleaned = join(homedir(), cleaned.slice(2));
  }

  // 1. Check raw path directly
  if (existsSync(cleaned) && statSync(cleaned).isFile()) {
    return cleaned;
  }

  // 2. Check URL-decoded path
  try {
    const decoded = decodeURIComponent(cleaned);
    if (existsSync(decoded) && statSync(decoded).isFile()) {
      return decoded;
    }
  } catch {}

  // 3. Candidate base directories
  const candidateBases = [];

  if (Array.isArray(workspaceUris)) {
    for (const uri of workspaceUris) {
      try {
        const p = uri.startsWith('file://') ? fileURLToPath(uri) : decodeURIComponent(uri);
        if (existsSync(p)) candidateBases.push(p);
      } catch {}
    }
  }

  candidateBases.push(process.cwd());
  candidateBases.push(join(homedir(), '.gemini/antigravity/scratch'));
  candidateBases.push(join(homedir(), 'Downloads'));
  candidateBases.push(join(homedir(), 'Documents'));
  candidateBases.push(join(homedir(), 'Desktop'));

  // Test joining with candidate bases
  for (const base of candidateBases) {
    if (!existsSync(base)) continue;
    try {
      const candidate = join(base, cleaned);
      if (existsSync(candidate) && statSync(candidate).isFile()) {
        return candidate;
      }
      const decodedCandidate = join(base, decodeURIComponent(cleaned));
      if (existsSync(decodedCandidate) && statSync(decodedCandidate).isFile()) {
        return decodedCandidate;
      }
    } catch {}
  }

  // 4. Fuzzy fallback: search by basename in candidate bases
  const targetBaseName = basename(cleaned);
  if (targetBaseName && isWordDoc(targetBaseName)) {
    for (const base of candidateBases) {
      if (!existsSync(base)) continue;
      const found = searchForDocByName(base, targetBaseName, 4);
      if (found) return found;
    }
  }

  return null;
}

function findWordDocs(dirPath, maxDepth = 4, currentDepth = 0) {
  const docs = [];
  if (currentDepth > maxDepth || !existsSync(dirPath)) return docs;

  try {
    const entries = readdirSync(dirPath, { withFileTypes: true });
    for (const entry of entries) {
      if (entry.name.startsWith('.') || entry.name === 'node_modules' || entry.name === 'vendor') {
        continue;
      }
      const fullPath = join(dirPath, entry.name);
      if (entry.isDirectory()) {
        docs.push(...findWordDocs(fullPath, maxDepth, currentDepth + 1));
      } else if (entry.isFile() && isWordDoc(entry.name)) {
        const s = statSync(fullPath);
        docs.push({
          name: entry.name,
          path: fullPath,
          size: s.size,
          mtime: s.mtimeMs,
          extension: extname(entry.name).toLowerCase(),
        });
      }
    }
  } catch {}
  return docs;
}

// ---------------------------------------------------------------------------
// APIs
// ---------------------------------------------------------------------------

app.api('/api/status', () => ({ ok: true, version: '1.0.0', type: 'word-viewer' }), 'GET');

app.api('/api/get-state', () => {
  if (!existsSync(STATE_FILE)) return { lastOpenedPath: null };
  try {
    return JSON.parse(readFileSync(STATE_FILE, 'utf8'));
  } catch {
    return { lastOpenedPath: null };
  }
}, 'GET');

app.api('/api/save-state', (data) => {
  mkdirSync(DATA_DIR, { recursive: true });
  const state = { lastOpenedPath: data?.lastOpenedPath || null, savedAt: Date.now() };
  writeFileSync(STATE_FILE, JSON.stringify(state, null, 2), 'utf8');
  return state;
});

app.api('/api/workspace-docs', (data) => {
  const uris = Array.isArray(data?.workspaceUris) ? data.workspaceUris : [];
  const folders = [];

  for (const uri of uris) {
    try {
      const folderPath = uri.startsWith('file://') ? fileURLToPath(uri) : decodeURIComponent(uri);
      if (existsSync(folderPath)) folders.push(folderPath);
    } catch {}
  }

  // Include known project & scratch folders
  folders.push(join(homedir(), '.gemini/antigravity/scratch'));
  folders.push(join(homedir(), 'Documents'));
  folders.push(join(homedir(), 'Downloads'));

  const results = [];
  const checkedFolders = new Set();
  for (const folder of folders) {
    if (checkedFolders.has(folder) || !existsSync(folder)) continue;
    checkedFolders.add(folder);
    results.push(...findWordDocs(folder, 3, 0));
  }

  const seen = new Set();
  const unique = [];
  for (const doc of results) {
    if (!seen.has(doc.path)) {
      seen.add(doc.path);
      unique.push(doc);
    }
  }
  unique.sort((a, b) => b.mtime - a.mtime);
  return { docs: unique };
});

app.api('/api/raw-doc', (data) => {
  const resolved = resolveDocPath(data?.path, data?.workspaceUris);
  if (!resolved) {
    return new Response(JSON.stringify({ error: `File not found: ${data?.path || ''}` }), {
      contentType: 'application/json',
      status: 404,
    });
  }

  const buffer = readFileSync(resolved);
  const ext = extname(resolved).toLowerCase();
  const contentType = ext === '.docx'
    ? 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    : 'application/msword';

  return new Response(buffer, {
    contentType,
    headers: {
      'Content-Disposition': `inline; filename="${basename(resolved)}"`,
    },
  });
}, 'GET');

app.api('/api/load-doc', (data) => {
  const inputPath = data?.path;
  if (!inputPath) return { success: false, error: 'No path specified' };

  try {
    const resolved = resolveDocPath(inputPath, data?.workspaceUris);
    if (!resolved) {
      return { success: false, error: `File not found: ${inputPath}` };
    }

    const stat = statSync(resolved);
    if (!stat.isFile()) {
      return { success: false, error: 'Path is not a regular file' };
    }

    // Embed base64 for files up to 20MB
    let base64 = null;
    if (stat.size <= 20 * 1024 * 1024) {
      const buffer = readFileSync(resolved);
      base64 = buffer.toString('base64');
    }

    return {
      success: true,
      name: basename(resolved),
      path: resolved,
      extension: extname(resolved).toLowerCase(),
      size: stat.size,
      mtime: stat.mtimeMs,
      base64,
    };
  } catch (err) {
    return { success: false, error: err.message };
  }
});

app.api('/api/save-doc', (data) => {
  const inputPath = data?.path;
  const base64 = data?.base64;
  if (!inputPath || !base64) {
    return { success: false, error: 'Path and base64 document content are required' };
  }

  try {
    const resolved = resolveDocPath(inputPath, data?.workspaceUris) || inputPath;
    const buffer = Buffer.from(base64, 'base64');
    writeFileSync(resolved, buffer);
    const stat = statSync(resolved);

    return {
      success: true,
      path: resolved,
      size: stat.size,
      mtime: stat.mtimeMs,
    };
  } catch (err) {
    return { success: false, error: err.message };
  }
});

app.run();
