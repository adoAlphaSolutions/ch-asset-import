// ============================================================================
// Asset Uploader + Linker — Content Hub External Component
// ----------------------------------------------------------------------------
// Upload a .zip of images. For each image the tool:
//   1. parses the product SKU from the filename,
//   2. finds the M.PCM.Product by SKU (FullText search + exact SKU verify),
//   3. (real run) uploads the image to create an M.Asset,
//   4. (real run) links the new asset to the product via PCMProductToAsset.
//
// Two buttons:
//   • Validate (dry run) — reads the zip, resolves each SKU to a product, and
//     reports matches/misses. NO writes to Content Hub.
//   • Upload & link — does the real uploads + relations for matched images.
//
// Uses the authenticated SDK client the External Component receives
// (context.client). The upload call (client.uploads.uploadAsync) is the one
// piece that may need adjusting to your instance — its result/error is logged.
//
// Configuration (JSON), optional:
//   {
//     "skuSeparator": "_",              // SKU = filename stem before this (blank = whole stem)
//     "relationName": "PCMProductToAsset",
//     "uploadConfiguration": "AssetUploadConfiguration",
//     "uploadAction": "NewAsset"
//   }
// ============================================================================

const CH_HOST = window.location.origin;
const JSZIP_URL = 'https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js';

const DEFAULTS = {
  skuSeparator: '',
  relationName: 'PCMProductToAsset',
  uploadConfiguration: 'AssetUploadConfiguration',
  uploadAction: 'NewAsset'
};
const IMAGE_EXT = new Set(['png', 'jpg', 'jpeg', 'gif', 'bmp', 'tif', 'tiff', 'webp', 'svg', 'heic', 'heif']);

const CSS = `
  .a-wrap  { font-family: "Segoe UI", sans-serif; padding: 24px; max-width: 900px; }
  .a-title { font-size: 20px; font-weight: 600; margin-bottom: 2px; }
  .a-sub   { font-size: 13px; color: #555; margin-bottom: 18px; }
  .a-drop  { border: 2px dashed #aaa; border-radius: 8px; padding: 32px; text-align: center; cursor: pointer; color: #555; margin-bottom: 12px; }
  .a-drop.a-hover { border-color: #2b6cb0; background: #f0f6ff; color: #2b6cb0; }
  .a-row   { display: flex; gap: 10px; align-items: center; margin-bottom: 12px; flex-wrap: wrap; }
  .a-btn   { padding: 8px 16px; border: none; border-radius: 6px; cursor: pointer; font-size: 14px; }
  .a-btn:disabled { opacity: .5; cursor: not-allowed; }
  .a-dry   { background: #edf2f7; color: #2d3748; }
  .a-go    { background: #c53030; color: #fff; }
  .a-log   { background: #1a202c; color: #e2e8f0; font-family: monospace; font-size: 12px; padding: 14px; border-radius: 6px; margin-top: 14px; max-height: 380px; overflow: auto; white-space: pre-wrap; display: none; }
  .a-ok { color: #68d391; } .a-skip { color: #cbd5e0; } .a-err { color: #fc8181; } .a-info { color: #90cdf4; }
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
          SKU resolves to a product, then Upload &amp; link to create each M.Asset and attach it to its product
          via <code>${cfg.relationName}</code>.</div>
        <div class="a-drop" id="a-drop">📦 Drop your .zip here, or click to browse</div>
        <input type="file" id="a-file" accept=".zip,application/zip" style="display:none" />
        <div class="a-row">
          <button class="a-btn a-dry" id="a-dry" disabled>🔍 Validate (dry run)</button>
          <button class="a-btn a-go"  id="a-go"  disabled>⬆ Upload &amp; link (writes to Content Hub)</button>
          <span id="a-status" style="font-size:13px;color:#555"></span>
        </div>
        <div class="a-log" id="a-log"></div>
      `;
      rootElement.innerHTML = ''; rootElement.appendChild(style); rootElement.appendChild(wrap);

      const drop = wrap.querySelector('#a-drop'), input = wrap.querySelector('#a-file');
      const dryBtn = wrap.querySelector('#a-dry'), goBtn = wrap.querySelector('#a-go');
      const status = wrap.querySelector('#a-status'), logEl = wrap.querySelector('#a-log');
      let currentFile = null;

      function log(msg, cls) {
        logEl.style.display = 'block';
        const line = document.createElement('div'); if (cls) line.className = cls;
        line.textContent = msg; logEl.appendChild(line); logEl.scrollTop = logEl.scrollHeight;
      }
      function clearLog() { logEl.innerHTML = ''; logEl.style.display = 'none'; }
      function onFile(f) { if (!f) return; clearLog(); currentFile = f; status.textContent = `${f.name} — ready`; dryBtn.disabled = false; goBtn.disabled = false; }

      drop.addEventListener('click', () => input.click());
      input.addEventListener('change', e => onFile(e.target.files[0]));
      drop.addEventListener('dragover', e => { e.preventDefault(); drop.classList.add('a-hover'); });
      drop.addEventListener('dragleave', () => drop.classList.remove('a-hover'));
      drop.addEventListener('drop', e => { e.preventDefault(); drop.classList.remove('a-hover'); onFile(e.dataTransfer.files[0]); });

      function skuFromName(name) {
        const stem = stripExt(baseName(name));
        if (cfg.skuSeparator) { const i = stem.indexOf(cfg.skuSeparator); return i > 0 ? stem.slice(0, i) : stem; }
        return stem;
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

      // Upload one image -> new asset id. Best-effort against the SDK upload client.
      async function uploadAsset(img) {
        if (!client || !client.uploads || typeof client.uploads.uploadAsync !== 'function') {
          throw new Error('client.uploads.uploadAsync not available — the upload API differs on this instance.');
        }
        const buffer = await img.entry.async('arraybuffer');
        const request = {
          uploadSource: { data: buffer, fileName: img.name },
          uploadConfiguration: cfg.uploadConfiguration,
          actionName: cfg.uploadAction,
          fileName: img.name
        };
        const result = await client.uploads.uploadAsync(request);
        // Try to read the created asset id from a few likely shapes.
        const aid = (result && (result.assetId || result.entityId || result.id ||
          (result.entity && result.entity.id))) || null;
        if (aid == null) throw new Error('upload returned no asset id (shape: ' + JSON.stringify(result).slice(0, 200) + ')');
        return aid;
      }

      // Link asset -> product via the PCMProductToAsset relation (product is parent).
      async function linkAssetToProduct(assetId, productId) {
        const product = await client.entities.getAsync(productId);
        const rel = product.getRelation(cfg.relationName);
        if (!rel) throw new Error(`relation ${cfg.relationName} not found on product`);
        // Add the asset id to the relation's children (product is the parent side).
        if (typeof rel.add === 'function') rel.add(assetId);
        else if (Array.isArray(rel.children)) rel.children.push(assetId);
        else if (typeof rel.setIds === 'function') rel.setIds([...(rel.getIds ? rel.getIds() : []), assetId]);
        else throw new Error('could not add id to relation (unknown relation API)');
        await client.entities.saveAsync(product);
      }

      async function run(dryRun) {
        clearLog(); dryBtn.disabled = true; goBtn.disabled = true;
        log(dryRun ? '── VALIDATE (no writes) ──' : '── UPLOAD & LINK ──', 'a-info');
        try {
          const imgs = await readImages();
          if (!imgs.length) { log('No image files found in the zip.', 'a-err'); return; }
          log(`Found ${imgs.length} image(s).`, 'a-info');

          let matched = 0, uploaded = 0, linked = 0, errors = 0;
          // Cache SKU -> product result within a run.
          const cache = new Map();

          for (const img of imgs) {
            const sku = skuFromName(img.name);
            let prod = cache.get(sku);
            if (!prod) { try { prod = await findProductBySku(sku); } catch (e) { prod = { status: 'error', message: e.message }; } cache.set(sku, prod); }

            if (prod.status !== 'ok') {
              errors++;
              const why = prod.status === 'none' ? `⚠ PRODUCT NOT FOUND for SKU "${sku}"`
                : prod.status === 'multiple' ? `⚠ AMBIGUOUS — ${prod.count} products match SKU "${sku}"`
                : `⚠ LOOKUP ERROR (SKU "${sku}") — ${prod.message}`;
              log(`  ✗  ${img.name}  —  ${why}`, 'a-err');
              continue;
            }
            matched++;
            if (dryRun) { log(`  ✓  ${img.name}  —  SKU "${sku}"  →  ${prod.name || '(no name)'}  (product ${prod.id})`, 'a-ok'); continue; }

            // Real run: upload + link
            try {
              const assetId = await uploadAsset(img); uploaded++;
              await linkAssetToProduct(assetId, prod.id); linked++;
              log(`  ✓ ${img.name} → asset ${assetId} linked to product ${prod.id} (${prod.name || ''})`, 'a-ok');
            } catch (e) {
              errors++;
              log(`  ✗ ${img.name} → ${e && e.message ? e.message : e}`, 'a-err');
            }
          }

          log('──────────────────────────────────────────', 'a-info');
          if (dryRun) log(`Validate done — ${matched} of ${imgs.length} image(s) resolve to a product, ${errors} unresolved.`, errors ? 'a-err' : 'a-ok');
          else log(`Done — uploaded ${uploaded}, linked ${linked}, errors ${errors} (of ${imgs.length}).`, errors ? 'a-err' : 'a-ok');
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
