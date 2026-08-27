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
//     "uploadConfiguration": "AssetUploadConfiguration",
//     "uploadAction": "NewAsset"
//   }
// ============================================================================

const BUILD_VERSION = 'v1.7 · 2026-08-26';   // bump on every change; shown in the footer
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
  uploadConfiguration: 'AssetUploadConfiguration',
  uploadAction: 'NewAsset'
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

function safeJson(v) { try { return JSON.stringify(v).slice(0, 240); } catch (e) { return String(v); } }
// Pull as much detail as possible out of an SDK/HTTP error for diagnostics.
function errDetail(e) {
  if (!e) return '';
  const parts = [e.message || String(e)];
  for (const k of ['detail', 'title', 'body', 'responseText', 'data', 'error', 'innerException']) {
    try { if (e[k] != null) parts.push(`${k}=${typeof e[k] === 'object' ? safeJson(e[k]) : String(e[k]).slice(0, 200)}`); } catch (_) { /* ignore */ }
  }
  try { const ks = Object.keys(e); if (ks.length) parts.push(`ekeys=[${ks.join(',')}]`); } catch (_) { /* ignore */ }
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
    length: blob.size, size: blob.size, fileSize: blob.size,
    blob, buffer,
    getBlob: () => blob,
    arrayBuffer: () => blob.arrayBuffer(),
    getStreamAsync: async () => (blob.stream ? blob.stream() : blob),
    stream: () => (blob.stream ? blob.stream() : blob)
  };
  return {
    fileName, name: fileName,
    fileSize: blob.size, size: blob.size, contentType: blob.type, mimeType: blob.type,
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
        const has = n => !!findCtor([n]);
        uploadDiagText = `uploads.keys=[${keys}] uploads.methods=[${proto}] uploadAsync.arity=${client.uploads.uploadAsync.length} ` +
          `real={UploadRequest:${has('UploadRequest')},BlobUploadSource:${has('BlobUploadSource')},ArrayBufferUploadSource:${has('ArrayBufferUploadSource')}}`;
        return uploadDiagText;
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
        const source = makeDuckSource(blob, buffer, img.name);

        // uploadFileAsync(file, request) is the correct browser form (confirmed:
        // it reaches the server). The SDK appends the request's own fields to the
        // multipart form, so the request must contain ONLY the fields the server
        // expects — extra keys cause a 500. Send a clean, minimal request.
        const reqFull = { fileName: img.name, fileSize: blob.size, uploadConfiguration: cfg.uploadConfiguration, actionName: cfg.uploadAction };
        const reqNoAction = { fileName: img.name, fileSize: blob.size, uploadConfiguration: cfg.uploadConfiguration };
        const up = client.uploads;

        const attempts = [
          ['uploadFileAsync(file, {fileName,fileSize,uploadConfiguration,actionName})', () => up.uploadFileAsync(file, reqFull)],
          ['uploadFileAsync(file, {fileName,fileSize,uploadConfiguration})', () => up.uploadFileAsync(file, reqNoAction)],
          ['uploadAsync({...,actionName}, source)', () => up.uploadAsync(reqFull, source)]
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
          throw new Error(`${errs.join('  ||  ')}  ||  ${uploadDiagText || computeUploadDiag()}`);
        }
        if (winningForm !== usedForm) { winningForm = usedForm; log(`✓ upload form that works: ${usedForm}`, 'a-ok'); }

        const aid = extractAssetId(result);
        if (aid == null) throw new Error(`upload ok via ${usedForm} but no asset id returned (result: ${safeJson(result)})`);
        return aid;
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

      // Set the asset's type via AssetTypeToAsset (asset is CHILD of M.AssetType).
      // Cover -> coverAssetType (Thumbnail); otherwise -> stockAssetType (Stock).
      async function setAssetType(assetId, isCover) {
        const identifier = isCover ? cfg.coverAssetType : cfg.stockAssetType;
        const typeId = await assetTypeId(identifier);
        const asset = await client.entities.getAsync(assetId);
        const rel = asset.getRelation(cfg.assetTypeRelation);
        if (!rel) throw new Error(`relation ${cfg.assetTypeRelation} not found on asset`);
        relSetIds(rel, [typeId]);   // single type; replace rather than append
        await client.entities.saveAsync(asset);
        return identifier;
      }

      // Link asset -> product. Always add to the all-assets relation; if the
      // product has no master/cover yet, also set this asset as the cover.
      // Returns { master: true } when the asset was set as the cover.
      async function linkAssetToProduct(assetId, productId) {
        const product = await client.entities.getAsync(productId);

        // 1) all-assets (product is parent, asset is child) — skip if already present
        const all = product.getRelation(cfg.allAssetsRelation);
        if (!all) throw new Error(`relation ${cfg.allAssetsRelation} not found on product`);
        if (relGetIds(all).indexOf(assetId) < 0) relAdd(all, assetId);

        // 2) master/cover — only if empty (single asset allowed)
        let master = false;
        if (cfg.setMasterIfEmpty) {
          const cover = product.getRelation(cfg.masterAssetRelation);
          if (cover && relGetIds(cover).length === 0) { relAdd(cover, assetId); master = true; }
        }

        await client.entities.saveAsync(product);
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
                const assetId = await uploadAsset(img);
                const link = await linkAssetToProduct(assetId, prod.id);
                let typeIdent;
                try { typeIdent = await setAssetType(assetId, link.master); }
                catch (te) { throw new Error(`linked but asset type not set: ${te.message}`); }
                const shortType = String(typeIdent).replace(/^M\.AssetType\./, '');
                rec.ok = true; rec.statusLabel = link.master ? 'Linked + cover' : 'Linked';
                rec.productId = prod.id; rec.productName = prod.name;
                rec.resultNote = `asset ${assetId} → product ${prod.id}${link.master ? ' (cover)' : ''}, type ${shortType}`;
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
