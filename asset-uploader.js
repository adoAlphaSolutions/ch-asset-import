// ============================================================================
// Asset Uploader + Linker — Content Hub External Component
// ----------------------------------------------------------------------------
// Upload a .zip of images. For each image the tool:
//   1. parses the product SKU from the filename,
//   2. finds the M.PCM.Product by SKU (FullText search + exact SKU verify),
//   3. (real run) uploads the image to create an M.Asset,
//   4. (real run) links the new asset to the product:
//        • always adds it to PCMProductToAllAssets (all assets of the product),
//        • if the product has NO master/cover yet, also sets it as the single
//          PCMProductToMasterAsset (cover image). Only the first asset per
//          product becomes the cover; the rest are all-assets only.
//   5. (real run) sets the asset's own type via AssetTypeToAsset (asset is child
//      of the M.AssetType taxonomy): cover asset -> M.AssetType.Thumbnail,
//      every other asset -> M.AssetType.PIMProduct.Stock.
//
// Two buttons:
//   • Validate (dry run) — reads the zip, resolves each SKU to a product, and
//     reports matches/misses. NO writes to Content Hub.
//   • Upload & link — does the real uploads + relations for matched images.
//
// Results view is built for large batches (100s of images):
//   • summary counts at the top (matched / not found / ambiguous / error),
//   • a filter (All · Problems only · Matched only) + filename/SKU search,
//   • "Download results" exports the full per-image outcome as an Excel file,
//   so you never have to scroll a long log to find the misses.
//
// Uses the authenticated SDK client the External Component receives
// (context.client). The upload call (client.uploads.uploadAsync) is the one
// piece that may need adjusting to your instance — its result/error is logged.
//
// Configuration (JSON), optional:
//   {
//     "skuSeparator": "",                        // if set, SKU = stem before FIRST occurrence of this
//     "stripSequenceSuffix": true,               // also try stem with a trailing -1/_2/ 3/(4) removed
//     "allAssetsRelation": "PCMProductToAllAssets",
//     "masterAssetRelation": "PCMProductToMasterAsset",
//     "setMasterIfEmpty": true,                  // set cover only when product has none
//     "assetTypeRelation": "AssetTypeToAsset",   // M.Asset is child of M.AssetType
//     "stockAssetType": "M.AssetType.PIMProduct.Stock",  // non-cover assets
//     "coverAssetType": "M.AssetType.Thumbnail",         // the cover asset
//     "lifecycleRelation": "FinalLifeCycleStatusToAsset",
//     "lifecycleStatus": "M.Final.LifeCycle.Status.Approved",
//     "setLifecycle": true,
//     "uploadConfiguration": "AssetUploadConfiguration",
//     "uploadAction": "NewAsset",
//     "attachFile": true                          // also push the image binary (needs M.UploadConfiguration Read)
//   }
// ============================================================================

const BUILD_VERSION = 'v3.3 · 2026-08-26';   // bump on every change; shown in the footer
const CH_HOST = window.location.origin;
const JSZIP_URL = 'https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js';
const SHEETJS_URL = 'https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js';

const DEFAULTS = {
  skuSeparator: '',
  stripSequenceSuffix: true,
  allAssetsRelation: 'PCMProductToAllAssets',
  masterAssetRelation: 'PCMProductToMasterAsset',
  setMasterIfEmpty: true,
  assetTypeRelation: 'AssetTypeToAsset',
  stockAssetType: 'M.AssetType.PIMProduct.Stock',
  coverAssetType: 'M.AssetType.Thumbnail',
  lifecycleRelation: 'FinalLifeCycleStatusToAsset',
  lifecycleStatus: 'M.Final.LifeCycle.Status.Approved',
  setLifecycle: true,
  uploadConfiguration: 'AssetUploadConfiguration',
  uploadAction: 'NewAsset',
  attachFile: true
};
const IMAGE_EXT = new Set(['png', 'jpg', 'jpeg', 'gif', 'bmp', 'tif', 'tiff', 'webp', 'svg', 'heic', 'heif']);

const CSS = `
  .a-wrap  { font-family: "Segoe UI", sans-serif; padding: 24px; max-width: 960px; }
  .a-title { font-size: 20px; font-weight: 600; margin-bottom: 2px; }
  .a-sub   { font-size: 13px; color: #555; margin-bottom: 18px; }
  .a-drop  { border: 2px dashed #aaa; border-radius: 8px; padding: 32px; text-align: center; cursor: pointer; color: #555; margin-bottom: 12px; }
  .a-drop.a-hover { border-color: #2b6cb0; background: #f0f6ff; color: #2b6cb0; }
  .a-row   { display: flex; gap: 10px; align-items: center; margin-bottom: 12px; flex-wrap: wrap; }
  .a-btn   { padding: 8px 16px; border: none; border-radius: 6px; cursor: pointer; font-size: 14px; }
  .a-btn:disabled { opacity: .5; cursor: not-allowed; }
  .a-dry   { background: #edf2f7; color: #2d3748; }
  .a-go    { background: #c53030; color: #fff; }
  .a-note  { font-size: 12px; color: #718096; }

  .a-chips { display: none; gap: 8px; flex-wrap: wrap; margin: 4px 0 12px; }
  .a-chip  { font-size: 12px; font-weight: 600; padding: 5px 11px; border-radius: 999px; }
  .a-chip.total { background: #edf2f7; color: #2d3748; }
  .a-chip.ok    { background: #c6f6d5; color: #22543d; }
  .a-chip.miss  { background: #fed7d7; color: #742a2a; }
  .a-chip.amb   { background: #feebc8; color: #7b341e; }
  .a-chip.err   { background: #fbb6ce; color: #702459; }

  .a-tools { display: none; gap: 10px; align-items: center; flex-wrap: wrap; margin-bottom: 8px; }
  .a-tools select, .a-tools input { padding: 6px 8px; font-size: 13px; border: 1px solid #cbd5e0; border-radius: 6px; }
  .a-tools input { min-width: 220px; }
  .a-shown { font-size: 12px; color: #718096; }

  .a-results { display: none; border: 1px solid #e2e8f0; border-radius: 8px; max-height: 460px; overflow: auto; margin-bottom: 12px; }
  .a-results table { border-collapse: collapse; width: 100%; font-size: 13px; }
  .a-results th { position: sticky; top: 0; background: #f7fafc; text-align: left; padding: 8px 10px; border-bottom: 1px solid #e2e8f0; font-weight: 600; color: #4a5568; }
  .a-results td { padding: 7px 10px; border-bottom: 1px solid #edf2f7; vertical-align: top; }
  .a-results tr.ok   td:first-child { color: #2f855a; }
  .a-results tr.bad  td:first-child { color: #c53030; }
  .a-results tr.bad  { background: #fff5f5; }
  .a-results .fn  { font-family: monospace; }
  .a-results .msg { color: #718096; }
  .a-results .bad .msg { color: #c53030; font-weight: 600; }

  .a-log   { background: #1a202c; color: #e2e8f0; font-family: monospace; font-size: 12px; padding: 12px 14px; border-radius: 6px; margin-top: 4px; max-height: 200px; overflow: auto; white-space: pre-wrap; display: none; }
  .a-ok { color: #68d391; } .a-skip { color: #cbd5e0; } .a-err { color: #fc8181; } .a-info { color: #90cdf4; }
  .a-foot  { margin-top: 16px; font-size: 11px; color: #a0aec0; text-align: right; }
`;

function loadScript(url, globalName) {
  return new Promise((resolve, reject) => {
    if (window[globalName]) return resolve(window[globalName]);
    const s = document.createElement('script');
    s.src = url;
    s.onload = () => (window[globalName] ? resolve(window[globalName]) : reject(new Error(`${globalName} loaded but missing.`)));
    s.onerror = () => reject(new Error(`Could not load ${globalName} (check CSP / network).`));
    document.head.appendChild(s);
  });
}

function baseName(p) { return String(p || '').split('/').pop().split('\\').pop(); }
function stripExt(n) { const i = n.lastIndexOf('.'); return i > 0 ? n.slice(0, i) : n; }
function extOf(n) { const i = n.lastIndexOf('.'); return i >= 0 ? n.slice(i + 1).toLowerCase() : ''; }
function ts() { const d = new Date(), p = n => String(n).padStart(2, '0'); return `${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}_${p(d.getHours())}${p(d.getMinutes())}`; }
function downloadBlob(blob, name) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a'); a.href = url; a.download = name;
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 5000);
}

function readProp(item, name) {
  const p = item && item.properties ? item.properties[name] : undefined;
  if (p == null) return '';
  if (typeof p === 'object') { const k = Object.keys(p); return k.length ? String(p[k[0]]) : ''; }
  return String(p);
}
function itemId(it) {
  if (it && it.id != null) return it.id;
  if (it && it.self && it.self.href) return Number(it.self.href.split('/').pop());
  return null;
}

const MIME = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', tif: 'image/tiff', tiff: 'image/tiff', webp: 'image/webp', svg: 'image/svg+xml', heic: 'image/heic', heif: 'image/heif' };
function mimeFor(name) { return MIME[extOf(name)] || 'application/octet-stream'; }

// Locate a real SDK constructor if the runtime exposes it (window / globalThis).
function findCtor(names) {
  const scopes = [typeof window !== 'undefined' ? window : null, typeof globalThis !== 'undefined' ? globalThis : null];
  for (const sc of scopes) { if (!sc) continue; for (const n of names) { if (typeof sc[n] === 'function') return sc[n]; } }
  return null;
}

// ---- REST relation helpers (bypass SDK lazy-load option classes) -----------
// The relation endpoint /api/entities/{id}/relations/{name} supports GET and PUT.
// We read the current members, append/replace, and PUT — trying the body shapes
// Content Hub accepts (hrefs / ids / {id} objects) until one works.
function entHref(id) { return `${CH_HOST}/api/entities/${id}`; }
function memberIds(arr) {
  if (!Array.isArray(arr)) return [];
  const out = [];
  for (const c of arr) {
    if (c == null) continue;
    if (typeof c === 'number') { out.push(c); continue; }
    if (c.id != null) { out.push(Number(c.id)); continue; }
    const h = c.href || (c.self && c.self.href) || (typeof c === 'string' ? c : '');
    const m = String(h).match(/(\d+)(?:[/?#].*)?$/);
    if (m) out.push(Number(m[1]));
  }
  return out;
}
let relPutOkLabel = '';   // which PUT body shape works on this instance (log once)
async function restGetRelation(entityId, name) {
  const res = await fetch(`${CH_HOST}/api/entities/${entityId}/relations/${encodeURIComponent(name)}`, { credentials: 'include', headers: { Accept: 'application/json' } });
  if (!res.ok) { const t = await res.text().catch(() => ''); throw new Error(`GET relation ${name} HTTP ${res.status}${t ? ' — ' + t.slice(0, 160) : ''}`); }
  return await res.json();
}
async function restGetEntity(id) {
  const res = await fetch(`${CH_HOST}/api/entities/${id}`, { credentials: 'include', headers: { Accept: 'application/json' } });
  if (!res.ok) { const t = await res.text().catch(() => ''); throw new Error(`GET entity ${id} HTTP ${res.status}${t ? ' — ' + t.slice(0, 160) : ''}`); }
  return await res.json();
}
async function restPutEntity(id, body) {
  const res = await fetch(`${CH_HOST}/api/entities/${id}`, {
    method: 'PUT', credentials: 'include',
    headers: { 'Content-Type': 'application/json', Accept: 'application/json' },
    body: JSON.stringify(body)
  });
  if (!res.ok) { const t = await res.text().catch(() => ''); throw new Error(`HTTP ${res.status}${t ? ' — ' + t.slice(0, 300) : ''}`); }
  return true;
}
// Update a relation side via ENTITY-level PUT (/api/entities/{id}) with the
// version + a relations block. Tries the member shapes CH may accept.
async function restSetMembers(entityId, name, side, ids, dbg) {
  const ent = await restGetEntity(entityId);
  const ver = ent.version != null ? ent.version : (ent.entity && ent.entity.version);
  const mk = (members) => { const b = { relations: { [name]: { [side]: members } } }; if (ver != null) b.version = ver; return b; };
  const variants = [
    [`${side}:hrefs`, mk(ids.map(id => ({ href: entHref(id) })))],
    [`${side}:ids`, mk(ids.slice())],
    [`${side}:id-objs`, mk(ids.map(id => ({ id })))]
  ];
  const errs = [];
  for (const [label, body] of variants) {
    try { await restPutEntity(entityId, body); if (dbg && relPutOkLabel !== label) { relPutOkLabel = label; dbg(`relation write format that works: ${label}`, 'a-ok'); } return; }
    catch (e) { errs.push(`${label}→${e.message}`); }
  }
  throw new Error(`entity PUT for ${name} failed all formats: ${errs.join(' | ')}`);
}
// Add a child to a parent-side relation (product -> assets). Skips duplicates.
async function restAddChild(parentId, name, childId, dbg) {
  const rel = await restGetRelation(parentId, name);
  if (dbg) dbg(`GET ${name}: ${JSON.stringify(rel).slice(0, 400)}`, 'a-info');
  const ids = memberIds(rel.children);
  if (ids.indexOf(Number(childId)) >= 0) return;
  ids.push(Number(childId));
  await restSetMembers(parentId, name, 'children', ids, dbg);
}
async function restChildCount(parentId, name) {
  const rel = await restGetRelation(parentId, name);
  return memberIds(rel.children).length;
}
// Set the single parent of a child-side relation (asset -> AssetType / lifecycle).
async function restSetParent(childId, name, parentId, dbg) {
  await restSetMembers(childId, name, 'parents', [Number(parentId)], dbg);
}

function safeJson(v) { try { return JSON.stringify(v).slice(0, 240); } catch (e) { return String(v); } }
// The upload response carries the new asset URL in a Location header (mirrors the
// working C# LinkHelper.IdFromEntityAsync(response.Headers.Location)).
function locationFromResponse(r) {
  if (!r) return null;
  try { if (r.headers && typeof r.headers.get === 'function') { const l = r.headers.get('Location') || r.headers.get('location'); if (l) return l; } } catch (_) { /* ignore */ }
  try { if (r.headers) { const l = r.headers.Location || r.headers.location; if (l) return l; } } catch (_) { /* ignore */ }
  return r.location || r.Location || (r.self && r.self.href) || null;
}
function idFromLocation(loc) {
  if (!loc) return null;
  const m = String(loc).match(/(\d+)(?:[/?#].*)?$/);
  return m ? Number(m[1]) : null;
}
// Pull as much detail as possible out of an SDK/HTTP error for diagnostics.
function errDetail(e) {
  if (!e) return '';
  const parts = [e.message || String(e)];
  try { if (e.statusCode != null) parts.push(`statusCode=${e.statusCode}`); } catch (_) { /* ignore */ }
  // responseMessage is a structured object ({responseHeaders, responseBody, ...}).
  // Dig for the actual server body instead of dumping the (truncated) whole thing.
  try {
    const rm = e.responseMessage;
    if (rm != null) {
      if (typeof rm === 'object') {
        parts.push(`responseMessage.keys=[${Object.keys(rm).join(',')}]`);
        for (const k of ['responseBody', 'body', 'content', 'data', 'message', 'error', 'detail', 'title']) {
          if (rm[k] != null) parts.push(`responseMessage.${k}=${typeof rm[k] === 'object' ? safeJson(rm[k]) : String(rm[k]).slice(0, 500)}`);
        }
      } else {
        parts.push(`responseMessage=${String(rm).slice(0, 500)}`);
      }
    }
  } catch (_) { /* ignore */ }
  for (const k of ['parameterName', 'argumentName', 'paramName', 'name', 'failures', 'errors', 'validationErrors', 'messages', 'detail', 'title', 'body', 'responseText', 'data', 'error', 'innerException']) {
    try { if (e[k] != null) parts.push(`${k}=${typeof e[k] === 'object' ? safeJson(e[k]) : String(e[k]).slice(0, 300)}`); } catch (_) { /* ignore */ }
  }
  try { if (typeof e.getErrors === 'function') { const ge = e.getErrors(); if (ge) parts.push(`getErrors=${safeJson(ge)}`); } } catch (_) { /* ignore */ }
  try { const ks = Object.keys(e); if (ks.length) parts.push(`ekeys=[${ks.join(',')}]`); } catch (_) { /* ignore */ }
  try { if (e.stack) parts.push(`stack=${String(e.stack).split('\n').slice(0, 2).join(' | ').slice(0, 220)}`); } catch (_) { /* ignore */ }
  return parts.join(' ; ');
}
function extractAssetId(r) {
  if (r == null) return null;
  if (typeof r === 'number') return r;
  if (typeof r === 'string' && /^\d+$/.test(r)) return Number(r);
  return r.assetId || r.entityId || r.id || (r.entity && r.entity.id) || r.targetId ||
    (r.asset && (r.asset.id || r.asset.entityId)) ||
    (Array.isArray(r) && r[0] && (r[0].id || r[0].entityId)) || null;
}

// Duck-typed IUploadSource: what client.uploads.uploadAsync reads off the source.
// Exposes the bytes in several shapes so the SDK can consume whichever it wants.
function makeDuckSource(blob, buffer, fileName) {
  const readable = {
    contentType: blob.type, mimeType: blob.type,
    length: blob.size, size: blob.size, fileSize: blob.size, contentLength: blob.size,
    blob, buffer,
    getBlob: () => blob,
    arrayBuffer: () => blob.arrayBuffer(),
    getStreamAsync: async () => (blob.stream ? blob.stream() : blob),
    stream: () => (blob.stream ? blob.stream() : blob)
  };
  return {
    fileName, name: fileName,
    fileSize: blob.size, size: blob.size, length: blob.size, fileLength: blob.size, contentLength: blob.size,
    contentType: blob.type, mimeType: blob.type,
    getReadableSourceAsync: async () => readable,
    getReadableSource: () => readable,
    blob, buffer
  };
}

// Find the M.PCM.Product whose SKU matches (FullText narrows, exact SKU verifies).
async function findProductBySku(sku) {
  const chql = `Definition.Name=='M.PCM.Product' and FullText=='${String(sku).replace(/'/g, "''")}'`;
  const url = `${CH_HOST}/api/entities/query?query=${encodeURIComponent(chql)}&members=TB.PCM.Product.SKU,TB.PCM.ProductName&take=50`;
  const res = await fetch(url, { credentials: 'include', headers: { Accept: 'application/json' } });
  if (!res.ok) { const b = await res.text().catch(() => ''); throw new Error(`product query HTTP ${res.status}${b ? ' — ' + b : ''}`); }
  const items = ((await res.json()) || {}).items || [];
  const exact = items.filter(it => readProp(it, 'TB.PCM.Product.SKU').trim().toLowerCase() === String(sku).trim().toLowerCase());
  if (exact.length === 1) return { id: itemId(exact[0]), name: readProp(exact[0], 'TB.PCM.ProductName'), status: 'ok' };
  if (exact.length > 1) return { status: 'multiple', count: exact.length };
  return { status: 'none', candidates: items.length };
}

// ---------------------------------------------------------------------------
export default function createExternalRoot(rootElement) {
  return {
    render(context) {
      const client = context && context.client;
      const cfg = Object.assign({}, DEFAULTS, (context && context.config) || {});

      const style = document.createElement('style'); style.textContent = CSS;
      const wrap = document.createElement('div'); wrap.className = 'a-wrap';
      wrap.innerHTML = `
        <div class="a-title">🖼️ Asset Uploader + Linker</div>
        <div class="a-sub">Upload a <b>.zip</b> of images named by product SKU. Validate first to confirm each
          SKU resolves to a product, then Upload &amp; link to create each M.Asset, add it to
          <code>${cfg.allAssetsRelation}</code>, and — if the product has no cover yet — set it as the
          <code>${cfg.masterAssetRelation}</code> (cover).</div>
        <div class="a-drop" id="a-drop">📦 Drop your .zip here, or click to browse</div>
        <input type="file" id="a-file" accept=".zip,application/zip" style="display:none" />
        <div class="a-row">
          <button class="a-btn a-dry" id="a-dry" disabled>🔍 Validate (dry run)</button>
          <button class="a-btn a-go"  id="a-go"  disabled>⬆ Upload &amp; link (writes to Content Hub)</button>
          <span id="a-status" class="a-note"></span>
        </div>

        <div class="a-chips" id="a-chips"></div>
        <div class="a-tools" id="a-tools">
          <label class="a-note">Show
            <select id="a-filter">
              <option value="all">All</option>
              <option value="problems">Problems only</option>
              <option value="matched">Matched only</option>
            </select>
          </label>
          <input id="a-search" type="text" placeholder="filter by filename or SKU…" />
          <button class="a-btn a-dry" id="a-dl">⬇ Download results</button>
          <span class="a-shown" id="a-shown"></span>
        </div>
        <div class="a-results" id="a-results"></div>

        <div class="a-log" id="a-log"></div>
        <div class="a-foot" id="a-foot">Asset Uploader + Linker · ${BUILD_VERSION}</div>
      `;
      rootElement.innerHTML = ''; rootElement.appendChild(style); rootElement.appendChild(wrap);

      const drop = wrap.querySelector('#a-drop'), input = wrap.querySelector('#a-file');
      const dryBtn = wrap.querySelector('#a-dry'), goBtn = wrap.querySelector('#a-go');
      const status = wrap.querySelector('#a-status'), logEl = wrap.querySelector('#a-log');
      const chipsEl = wrap.querySelector('#a-chips'), toolsEl = wrap.querySelector('#a-tools');
      const resultsEl = wrap.querySelector('#a-results'), shownEl = wrap.querySelector('#a-shown');
      const filterEl = wrap.querySelector('#a-filter'), searchEl = wrap.querySelector('#a-search');
      const dlBtn = wrap.querySelector('#a-dl');
      let currentFile = null;
      let results = [];       // [{ name, sku, ok, statusLabel, productId, productName, message }]
      let lastMode = 'dry';
      let uploadDiagDone = false;          // log the upload-API surface once per run set
      let uploadDiagText = '';             // cached surface string, appended to first error
      let winningForm = '';                // which upload call form succeeded (logged once)
      const assetTypeIdCache = new Map();  // M.AssetType identifier -> entity id

      function log(msg, cls) {
        logEl.style.display = 'block';
        const line = document.createElement('div'); if (cls) line.className = cls;
        line.textContent = msg; logEl.appendChild(line); logEl.scrollTop = logEl.scrollHeight;
      }
      function clearAll() {
        logEl.innerHTML = ''; logEl.style.display = 'none';
        results = []; resultsEl.style.display = 'none'; resultsEl.innerHTML = '';
        chipsEl.style.display = 'none'; chipsEl.innerHTML = '';
        toolsEl.style.display = 'none'; shownEl.textContent = '';
      }
      function onFile(f) { if (!f) return; clearAll(); currentFile = f; status.textContent = `${f.name} — ready`; dryBtn.disabled = false; goBtn.disabled = false; }

      drop.addEventListener('click', () => input.click());
      input.addEventListener('change', e => onFile(e.target.files[0]));
      drop.addEventListener('dragover', e => { e.preventDefault(); drop.classList.add('a-hover'); });
      drop.addEventListener('dragleave', () => drop.classList.remove('a-hover'));
      drop.addEventListener('drop', e => { e.preventDefault(); drop.classList.remove('a-hover'); onFile(e.dataTransfer.files[0]); });

      // Build ordered SKU candidates from a filename. We try the most specific
      // first (whole stem), then progressively strip a trailing sequence suffix
      // (-1, _2, " 3", "(4)") so "sku-1.png"/"sku-2.png" both resolve to "sku"
      // WITHOUT truncating SKUs that legitimately contain "-" or "_".
      function skuCandidates(name) {
        const stem = stripExt(baseName(name)).trim();
        const cands = [];
        const push = s => { s = (s || '').trim(); if (s && cands.indexOf(s) < 0) cands.push(s); };

        // Optional hard rule: everything before the first configured separator.
        if (cfg.skuSeparator) { const i = stem.indexOf(cfg.skuSeparator); if (i > 0) push(stem.slice(0, i)); }

        push(stem); // whole stem — matches SKUs that contain "-"/"_" as-is

        // Progressively drop a trailing sequence suffix. Requires a separator
        // (or parentheses) before the digits, so "SKU100" is never shortened.
        if (cfg.stripSequenceSuffix) {
          let s = stem;
          for (let k = 0; k < 6; k++) {
            const t = s.replace(/\s*\(\d+\)$/, '').replace(/[-_ ]\d+$/, '');
            if (t === s || !t) break;
            push(t); s = t;
          }
        }
        return cands.length ? cands : [stem];
      }

      // Try each candidate SKU until one resolves to exactly one product.
      // Returns the product result plus the SKU that actually matched.
      async function resolveProduct(name, cache) {
        const cands = skuCandidates(name);
        let lastNonOk = null;
        for (const c of cands) {
          let r = cache.get(c);
          if (!r) { try { r = await findProductBySku(c); } catch (e) { r = { status: 'error', message: e.message }; } cache.set(c, r); }
          if (r.status === 'ok') return Object.assign({}, r, { sku: c });
          lastNonOk = Object.assign({}, r, { sku: c });
        }
        return lastNonOk || { status: 'none', candidates: 0, sku: cands[0] };
      }

      async function readImages() {
        const JSZip = await loadScript(JSZIP_URL, 'JSZip');
        const zip = await JSZip.loadAsync(currentFile);
        const imgs = [];
        zip.forEach((path, entry) => {
          if (entry.dir) return;
          if (path.indexOf('__MACOSX/') === 0 || baseName(path).indexOf('._') === 0) return;
          if (IMAGE_EXT.has(extOf(baseName(path)))) imgs.push({ name: baseName(path), entry });
        });
        return imgs;
      }

      // Log the upload-API surface once, so we can confirm the exact shape this
      // instance expects (constructors present? uploadAsync arity? sub-keys?).
      function computeUploadDiag() {
        let keys = '?'; try { keys = Object.keys(client.uploads).join(','); } catch (e) { /* ignore */ }
        let proto = '?'; try { proto = Object.getOwnPropertyNames(Object.getPrototypeOf(client.uploads)).filter(n => n !== 'constructor').join(','); } catch (e) { /* ignore */ }
        let ckeys = '?'; try { ckeys = Object.keys(client).join(','); } catch (e) { /* ignore */ }
        const mgrMethods = m => { try { return Object.getOwnPropertyNames(Object.getPrototypeOf(client[m])).filter(n => n !== 'constructor').join(','); } catch (e) { return '?'; } };
        const has = n => !!findCtor([n]);
        uploadDiagText = `uploads.methods=[${proto}] uploadAsync.arity=${client.uploads.uploadAsync.length} ` +
          `assets.methods=[${mgrMethods('assets')}] settings.methods=[${mgrMethods('settings')}] ` +
          `client.keys=[${ckeys}] ` +
          `real={UploadRequest:${has('UploadRequest')},BlobUploadSource:${has('BlobUploadSource')},ArrayBufferUploadSource:${has('ArrayBufferUploadSource')}}`;
        return uploadDiagText;
      }
      // Try to fetch the real UploadConfiguration object (carries chunk size etc.)
      // from whichever manager the client exposes. Returns the object or null.
      async function fetchUploadConfig(name) {
        const managers = ['uploadConfigurations', 'uploadConfiguration', 'settings'];
        for (const m of managers) {
          const mgr = client[m];
          if (mgr && typeof mgr.getAsync === 'function') {
            try { const c = await mgr.getAsync(name); if (c) return c; } catch (e) { /* try next */ }
          }
        }
        return null;
      }
      function logUploadDiag() {
        if (uploadDiagDone) return; uploadDiagDone = true;
        log('upload API — ' + computeUploadDiag(), 'a-info');
      }

      // Upload one image -> new asset id. Prefers the real SDK classes if the
      // runtime exposes them; otherwise sends a duck-typed upload source.
      async function uploadAsset(img) {
        if (!client || !client.uploads || typeof client.uploads.uploadAsync !== 'function') {
          throw new Error('client.uploads.uploadAsync not available — the upload API differs on this instance.');
        }
        logUploadDiag();

        const buffer = await img.entry.async('arraybuffer');
        const blob = new Blob([buffer], { type: mimeFor(img.name) });
        const file = (typeof File === 'function') ? new File([blob], img.name, { type: blob.type }) : blob;

        // Mirror the working C# uploader:
        //   var src = new ByteArrayUploadSource(bytes, fileName);
        //   var req = new UploadRequest(src, "AssetUploadConfiguration", "NewAsset");
        //   var resp = await client.Uploads.UploadAsync(req);
        //   id = LinkHelper.IdFromEntityAsync(resp.Headers.Location);
        // i.e. uploadAsync(request) with the source EMBEDDED, id from Location.
        const up = client.uploads;
        const source = makeDuckSource(blob, buffer, img.name);
        // The SDK reads the source from request.source (confirmed by diagnostics),
        // so set that primarily; keep uploadSource as an alias just in case.
        const reqEmbedded = { source: source, uploadSource: source, uploadConfiguration: cfg.uploadConfiguration, actionName: cfg.uploadAction, fileName: img.name, fileSize: blob.size };

        // uploadFileAsync is the sanctioned browser method (chunks a File and reaches
        // the server). A 500 here is most likely a PERMISSION issue: the signed-in
        // user needs Read on the M.UploadConfiguration definition (per Sitecore docs).
        // uploadAsync(request) is kept as a fallback but needs SDK classes this
        // runtime doesn't expose, so it can't be fully satisfied by a hand-built object.
        const attempts = [
          ['uploadFileAsync(file, {fileName,fileSize,uploadConfiguration,actionName})', () => up.uploadFileAsync(file, { fileName: img.name, fileSize: blob.size, uploadConfiguration: cfg.uploadConfiguration, actionName: cfg.uploadAction })],
          ['uploadAsync({source,uploadConfiguration,actionName})', () => up.uploadAsync(reqEmbedded)]
        ];

        // Try each form; an upload that 500s before finalization creates nothing,
        // so it's safe to fall through to the next form and collect all errors.
        let result, usedForm;
        const errs = [];
        for (const [name, fn] of attempts) {
          try { result = await fn(); usedForm = name; break; }
          catch (e) { errs.push(`[${name}] ${errDetail(e)}`); }
        }
        if (usedForm == null) {
          const hint = /500/.test(errs.join(' ')) ? '  ||  HINT: a 500 on uploadFileAsync usually means the signed-in user lacks Read permission on the M.UploadConfiguration definition.' : '';
          throw new Error(`${errs.join('  ||  ')}${hint}  ||  ${uploadDiagText || computeUploadDiag()}`);
        }
        if (winningForm !== usedForm) { winningForm = usedForm; log(`✓ upload form that works: ${usedForm}`, 'a-ok'); }
        log(`upload response: ${safeJson(result)} ; location=${locationFromResponse(result) || '(none)'}`, 'a-info');

        // Prefer the Location header (like the C# LinkHelper), then fall back to
        // reading an id off the result body.
        let aid = idFromLocation(locationFromResponse(result));
        if (aid == null && client.linkHelper) {
          const loc = locationFromResponse(result);
          try {
            if (loc && typeof client.linkHelper.idFromEntityAsync === 'function') aid = await client.linkHelper.idFromEntityAsync(loc);
            else if (loc && typeof client.linkHelper.getIdFromEntity === 'function') aid = client.linkHelper.getIdFromEntity(loc);
          } catch (e) { /* ignore, fall through */ }
        }
        if (aid == null) aid = extractAssetId(result);
        if (aid == null) throw new Error(`upload ok via ${usedForm} but no asset id (result: ${safeJson(result)} ; location: ${locationFromResponse(result)})`);
        return aid;
      }

      // Force-load a relation via the SDK. The key is MemberLoadOption.LazyLoading
      // (the 3rd arg of getRelationAsync) — without it the SDK does no IO and
      // reports "not loaded". We try the likely numeric enum values, then plain.
      async function loadRel(entity, name) {
        let lastErr = null;
        if (typeof entity.getRelationAsync === 'function') {
          const argSets = [[name, undefined, 1], [name, undefined, 2], [name, undefined, 0], [name]];
          for (const args of argSets) {
            try { const r = await entity.getRelationAsync.apply(entity, args); if (r) return r; }
            catch (e) { lastErr = e; }
          }
        }
        try { const r = entity.getRelation(name); if (r) return r; } catch (e) { lastErr = lastErr || e; }
        if (lastErr) throw new Error(`loadRel('${name}'): ${lastErr.message || lastErr}`);
        return null;
      }
      function relIds(rel) {
        try { if (typeof rel.getIds === 'function') { const v = rel.getIds(); if (Array.isArray(v)) return v.slice(); } } catch (e) { /* ignore */ }
        return relGetIds(rel);
      }
      function relSet(rel, ids) {
        if (typeof rel.setIds === 'function') { rel.setIds(ids); return; }
        relSetIds(rel, ids);
      }

      // Get a relation, lazy-loading it if the entity was fetched without it.
      // Surfaces the real error (instead of hiding it) so we can see WHY a
      // relation won't resolve, and also tries loadRelationsAsync.
      async function getRel(entity, name) {
        // 1. already loaded?
        try { const r0 = entity.getRelation(name); if (r0) return r0; } catch (e) { /* not loaded yet */ }
        let firstErr = null;
        // 2. lazy-load the single relation (let a real error propagate)
        if (typeof entity.getRelationAsync === 'function') {
          try { const r = await entity.getRelationAsync(name); if (r) return r; }
          catch (e) { firstErr = e; }
        }
        // 3. explicit load, then sync get (try a couple of arg shapes)
        if (typeof entity.loadRelationsAsync === 'function') {
          for (const arg of [name, [name]]) {
            try { await entity.loadRelationsAsync(arg); const r2 = entity.getRelation(name); if (r2) return r2; }
            catch (e) { firstErr = firstErr || e; }
          }
        }
        if (firstErr) throw new Error(`getRelationAsync('${name}') failed: ${firstErr.message || firstErr}`);
        return null;
      }

      // Read the ids currently in a relation, across possible SDK shapes.
      function relGetIds(rel) {
        if (!rel) return [];
        if (typeof rel.getIds === 'function') { try { const v = rel.getIds(); if (Array.isArray(v)) return v; } catch (e) { /* ignore */ } }
        if (Array.isArray(rel.children)) return rel.children.slice();
        if (Array.isArray(rel.ids)) return rel.ids.slice();
        return [];
      }
      function relAdd(rel, id) {
        if (typeof rel.add === 'function') { rel.add(id); return; }
        if (Array.isArray(rel.children)) { rel.children.push(id); return; }
        if (typeof rel.setIds === 'function') { rel.setIds([...relGetIds(rel), id]); return; }
        throw new Error('could not add id to relation (unknown relation API)');
      }
      // Replace the ids on a relation (used to set the asset's single type parent).
      function relSetIds(rel, ids) {
        if (typeof rel.setIds === 'function') { rel.setIds(ids); return; }
        if (Array.isArray(rel.parents)) { rel.parents.length = 0; for (const i of ids) rel.parents.push(i); return; }
        if (Array.isArray(rel.children)) { rel.children.length = 0; for (const i of ids) rel.children.push(i); return; }
        if (typeof rel.setParents === 'function') { rel.setParents(ids); return; }
        if (typeof rel.add === 'function') { for (const i of ids) rel.add(i); return; }
        throw new Error('could not set ids on relation (unknown relation API)');
      }

      // Resolve an M.AssetType taxonomy identifier -> entity id (cached per session).
      async function assetTypeId(identifier) {
        if (assetTypeIdCache.has(identifier)) return assetTypeIdCache.get(identifier);
        const ent = await client.entities.getAsync(identifier);
        const id = ent && (ent.id != null ? ent.id : (typeof ent.getId === 'function' ? ent.getId() : null));
        if (id == null) throw new Error(`asset type '${identifier}' not found`);
        assetTypeIdCache.set(identifier, id);
        return id;
      }

      // Create the M.Asset ENTITY using the same proven path as the Division
      // staging rows (entityFactory.createAsync + entities.saveAsync). This is a
      // metadata write and does NOT need the upload pipeline. Returns the new id.
      async function createAssetEntity(img) {
        const culture = context && context.culture;
        const attempts = [
          () => client.entityFactory.createAsync('M.Asset'),
          () => client.entityFactory.createAsync('M.Asset', culture),
          () => client.entityFactory.createAsync('M.Asset', [culture])
        ];
        let asset, lastErr;
        for (const a of attempts) { try { asset = await a(); if (asset) break; } catch (e) { lastErr = e; } }
        if (!asset) throw new Error(`createAsync(M.Asset) failed: ${lastErr && lastErr.message ? lastErr.message : lastErr}`);
        try { asset.setPropertyValue('FileName', img.name); } catch (e) { /* ignore */ }
        try { asset.setPropertyValue('Title', stripExt(img.name)); } catch (e) { /* ignore */ }
        const saved = await client.entities.saveAsync(asset);
        const id = (typeof saved === 'number') ? saved : ((saved && (saved.id || saved.Id)) || asset.id || extractAssetId(saved));
        if (id == null) throw new Error('asset entity created but no id returned');
        return id;
      }

      // Set the asset's type (AssetTypeToAsset, asset is CHILD of M.AssetType) and,
      // optionally, its final lifecycle status. Cover -> coverAssetType (Thumbnail);
      // otherwise -> stockAssetType (Stock).
      async function setAssetType(assetId, isCover) {
        const identifier = isCover ? cfg.coverAssetType : cfg.stockAssetType;
        const typeId = await assetTypeId(identifier);
        const asset = await client.entities.getAsync(assetId);
        // AssetTypeToAsset: asset is CHILD of M.AssetType -> set the single parent.
        const rel = await loadRel(asset, cfg.assetTypeRelation);
        if (!rel) throw new Error(`relation ${cfg.assetTypeRelation} could not be loaded on asset`);
        relSet(rel, [typeId]);

        if (cfg.setLifecycle && cfg.lifecycleStatus) {
          try {
            const lcId = await assetTypeId(cfg.lifecycleStatus);   // resolves identifier -> id (cached)
            const lc = await loadRel(asset, cfg.lifecycleRelation);
            if (lc) relSet(lc, [lcId]);
          } catch (e) { /* lifecycle is best-effort */ }
        }
        try { await client.entities.saveAsync(asset); }
        catch (e) { throw new Error(`asset save failed — ${errDetail(e)}`); }
        return identifier;
      }

      // Attach the image binary to an EXISTING asset via the upload pipeline
      // (action NewMainFile + AssetId), mirroring the C# entityId>0 branch. This
      // is the part that needs M.UploadConfiguration Read permission.
      async function attachFileToAsset(assetId, img) {
        const buffer = await img.entry.async('arraybuffer');
        const blob = new Blob([buffer], { type: mimeFor(img.name) });
        const file = (typeof File === 'function') ? new File([blob], img.name, { type: blob.type }) : blob;
        return await client.uploads.uploadFileAsync(file, {
          fileName: img.name, fileSize: blob.size,
          uploadConfiguration: cfg.uploadConfiguration,
          actionName: 'NewMainFile',
          actionParameters: { AssetId: assetId }
        });
      }

      // Link asset -> product. Always add to the all-assets relation; if the
      // product has no master/cover yet, also set this asset as the cover.
      // Returns { master: true } when the asset was set as the cover.
      async function linkAssetToProduct(assetId, productId) {
        const product = await client.entities.getAsync(productId);

        // 1) all-assets (product is parent, asset is child) — append, skip dups
        const all = await loadRel(product, cfg.allAssetsRelation);
        if (!all) throw new Error(`relation ${cfg.allAssetsRelation} could not be loaded on product`);
        const ids = relIds(all);
        if (ids.indexOf(assetId) < 0) { ids.push(assetId); relSet(all, ids); }

        // 2) master/cover — only if the product has none yet (single asset)
        let master = false;
        if (cfg.setMasterIfEmpty) {
          const cover = await loadRel(product, cfg.masterAssetRelation);
          if (cover && relIds(cover).length === 0) { relSet(cover, [assetId]); master = true; }
        }

        try { await client.entities.saveAsync(product); }
        catch (e) { throw new Error(`product save failed — ${errDetail(e)}`); }
        return { master };
      }

      // ---- results rendering (filter + search + counts) --------------------
      function counts() {
        const c = { total: results.length, ok: 0, miss: 0, amb: 0, err: 0 };
        for (const r of results) {
          if (r.ok) c.ok++;
          else if (r.statusLabel === 'Not found') c.miss++;
          else if (r.statusLabel === 'Ambiguous') c.amb++;
          else c.err++;
        }
        return c;
      }
      function renderChips() {
        const c = counts();
        const okLabel = lastMode === 'dry' ? 'Matched' : 'Linked';
        chipsEl.innerHTML =
          `<span class="a-chip total">${c.total} image(s)</span>` +
          `<span class="a-chip ok">✓ ${okLabel} ${c.ok}</span>` +
          (c.miss ? `<span class="a-chip miss">⚠ Not found ${c.miss}</span>` : '') +
          (c.amb ? `<span class="a-chip amb">⚠ Ambiguous ${c.amb}</span>` : '') +
          (c.err ? `<span class="a-chip err">✗ Error ${c.err}</span>` : '');
        chipsEl.style.display = 'flex';
      }
      function filteredRows() {
        const mode = filterEl.value;
        const q = (searchEl.value || '').trim().toLowerCase();
        return results.filter(r => {
          if (mode === 'problems' && r.ok) return false;
          if (mode === 'matched' && !r.ok) return false;
          if (q && r.name.toLowerCase().indexOf(q) < 0 && String(r.sku).toLowerCase().indexOf(q) < 0) return false;
          return true;
        });
      }
      function renderResults() {
        if (!results.length) { resultsEl.style.display = 'none'; toolsEl.style.display = 'none'; return; }
        toolsEl.style.display = 'flex';
        const rows = filteredRows();
        const CAP = 1000;
        const shown = rows.slice(0, CAP);
        const prodCol = lastMode === 'dry' ? 'Product' : 'Result';
        let html = `<table><thead><tr><th style="width:34px"></th><th>File</th><th>SKU</th><th>${prodCol}</th></tr></thead><tbody>`;
        for (const r of shown) {
          const cls = r.ok ? 'ok' : 'bad';
          const icon = r.ok ? '✓' : '✗';
          const last = r.ok
            ? `${(r.productName || '(no name)')} <span class="msg">(${r.resultNote || ('product ' + r.productId)})</span>`
            : `<span class="msg">${r.statusLabel}${r.message ? ' — ' + r.message : ''}</span>`;
          html += `<tr class="${cls}"><td>${icon}</td><td class="fn">${escapeHtml(r.name)}</td><td class="fn">${escapeHtml(r.sku)}</td><td>${last}</td></tr>`;
        }
        html += '</tbody></table>';
        resultsEl.innerHTML = html;
        resultsEl.style.display = 'block';
        shownEl.textContent = rows.length > CAP
          ? `showing ${CAP} of ${rows.length} (download for the full list)`
          : `showing ${rows.length} of ${results.length}`;
      }
      function escapeHtml(s) { return String(s).replace(/[&<>"]/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[ch])); }

      filterEl.addEventListener('change', renderResults);
      searchEl.addEventListener('input', renderResults);
      dlBtn.addEventListener('click', downloadResults);

      async function downloadResults() {
        if (!results.length) return;
        const XLSX = await loadScript(SHEETJS_URL, 'XLSX');
        const header = ['File', 'SKU', 'Status', 'Product ID', 'Product Name', 'Detail'];
        const rows = results.map(r => [r.name, r.sku, r.statusLabel, r.ok ? (r.productId || '') : '', r.ok ? (r.productName || '') : '', r.message || r.resultNote || '']);
        const ws = XLSX.utils.aoa_to_sheet([header, ...rows]);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Results');
        const arr = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
        downloadBlob(new Blob([arr], { type: 'application/octet-stream' }), `AssetUpload_${lastMode}_${ts()}.xlsx`);
      }

      async function run(dryRun) {
        clearAll(); dryBtn.disabled = true; goBtn.disabled = true;
        lastMode = dryRun ? 'dry' : 'real';
        log(dryRun ? '── VALIDATE (no writes) ──' : '── UPLOAD & LINK ──', 'a-info');
        try {
          const imgs = await readImages();
          if (!imgs.length) { log('No image files found in the zip.', 'a-err'); return; }
          log(`Found ${imgs.length} image(s). Resolving…`, 'a-info');

          const cache = new Map();  // candidate SKU -> product lookup result (dedupes CH queries)
          let done = 0;

          for (const img of imgs) {
            const prod = await resolveProduct(img.name, cache);
            const sku = prod.sku;

            const rec = { name: img.name, sku, ok: false, statusLabel: '', productId: null, productName: '', message: '', resultNote: '' };

            if (prod.status !== 'ok') {
              rec.statusLabel = prod.status === 'none' ? 'Not found' : prod.status === 'multiple' ? 'Ambiguous' : 'Lookup error';
              rec.message = prod.status === 'none' ? `no product for SKU "${sku}"`
                : prod.status === 'multiple' ? `${prod.count} products match SKU "${sku}"`
                : prod.message;
            } else if (dryRun) {
              rec.ok = true; rec.statusLabel = 'Matched'; rec.productId = prod.id; rec.productName = prod.name;
              rec.resultNote = 'product ' + prod.id;
            } else {
              try {
                // 1. Create the asset ENTITY (proven metadata write — no upload pipeline).
                const assetId = await createAssetEntity(img);
                // 2. Link to the product (all-assets + cover-if-empty).
                const link = await linkAssetToProduct(assetId, prod.id);
                // 3. Set asset type (+ lifecycle).
                const typeIdent = await setAssetType(assetId, link.master);
                const shortType = String(typeIdent).replace(/^M\.AssetType\./, '');
                // 4. Attach the binary file (separate upload pipeline — best effort).
                let fileNote = '';
                if (cfg.attachFile) {
                  try { await attachFileToAsset(assetId, img); fileNote = ', file attached'; }
                  catch (fe) { fileNote = `, ⚠ file NOT attached (${errDetail(fe)})`; }
                }
                rec.ok = true; rec.statusLabel = link.master ? 'Linked + cover' : 'Linked';
                rec.productId = prod.id; rec.productName = prod.name;
                rec.resultNote = `asset ${assetId} → product ${prod.id}${link.master ? ' (cover)' : ''}, type ${shortType}${fileNote}`;
              } catch (e) {
                rec.statusLabel = 'Write error'; rec.message = (e && e.message) ? e.message : String(e);
              }
            }
            results.push(rec);

            done++;
            if (done % 10 === 0 || done === imgs.length) status.textContent = `Processing ${done}/${imgs.length}…`;
          }

          renderChips();
          renderResults();
          const c = counts();
          status.textContent = `${currentFile.name} — done`;
          if (dryRun) log(`Validate done — ${c.ok} matched, ${c.miss} not found, ${c.amb} ambiguous, ${c.err} error(s) of ${c.total}.`, (c.miss + c.amb + c.err) ? 'a-err' : 'a-ok');
          else log(`Done — ${c.ok} uploaded & linked, ${c.miss + c.amb + c.err} not processed of ${c.total}.`, (c.miss + c.amb + c.err) ? 'a-err' : 'a-ok');
        } catch (e) {
          log(`✗ ${e && e.message ? e.message : e}`, 'a-err');
        } finally { dryBtn.disabled = false; goBtn.disabled = false; }
      }

      dryBtn.addEventListener('click', () => run(true));
      goBtn.addEventListener('click', () => run(false));
    },
    unmount() { rootElement.innerHTML = ''; }
  };
}
