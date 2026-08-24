// ============================================================================
// Asset Import Generator — Content Hub External Component (SIMULATION)
// ----------------------------------------------------------------------------
// Upload a .zip of images; the tool reads every image entry and produces a
// Content Hub M.Asset import workbook (one sheet named "M.Asset") ready to run
// through the standard importer.
//
// Output columns (exact order Content Hub expects):
//   Title | FileName | FinalLifeCycleStatusToAsset | ContentRepositoryToAsset | File
//
// SIMULATION NOTE: the images are NOT uploaded anywhere yet. The "File" column
// is filled with a MOCK URL built from a base you configure. When the real
// pipeline is wired up, that base becomes the Azure Blob container URL (and each
// file would be uploaded there first), so the import fetches the real image.
//
// Configuration textarea (JSON), all optional:
//   {
//     "fileBaseUrl": "https://YOURACCOUNT.blob.core.windows.net/assets",
//     "lifecycleStatus": "M.Final.LifeCycle.Status.Approved",
//     "contentRepository": "M.Content.Repository.Standard",
//     "titleFrom": "filename"   // "filename" = name without extension; "full" = full filename
//   }
// ============================================================================

const SHEET_NAME = 'M.Asset';
const JSZIP_URL = 'https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js';
const SHEETJS_URL = 'https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js';

const DEFAULTS = {
  fileBaseUrl: 'https://YOURACCOUNT.blob.core.windows.net/assets',
  lifecycleStatus: 'M.Final.LifeCycle.Status.Approved',
  contentRepository: 'M.Content.Repository.Standard',
  titleFrom: 'filename'
};

const IMAGE_EXT = new Set(['png', 'jpg', 'jpeg', 'gif', 'bmp', 'tif', 'tiff', 'webp', 'svg', 'heic', 'heif']);

const CSS = `
  .a-wrap  { font-family: "Segoe UI", sans-serif; padding: 24px; max-width: 860px; }
  .a-title { font-size: 20px; font-weight: 600; margin-bottom: 2px; }
  .a-sub   { font-size: 13px; color: #555; margin-bottom: 18px; }
  .a-drop  { border: 2px dashed #aaa; border-radius: 8px; padding: 32px; text-align: center; cursor: pointer; color: #555; margin-bottom: 12px; }
  .a-drop.a-hover { border-color: #2b6cb0; background: #f0f6ff; color: #2b6cb0; }
  .a-row   { display: flex; gap: 10px; align-items: center; margin-bottom: 12px; flex-wrap: wrap; }
  .a-btn   { padding: 8px 16px; border: none; border-radius: 6px; cursor: pointer; font-size: 14px; }
  .a-btn:disabled { opacity: .5; cursor: not-allowed; }
  .a-dry   { background: #edf2f7; color: #2d3748; }
  .a-go    { background: #2f855a; color: #fff; }
  .a-log   { background: #1a202c; color: #e2e8f0; font-family: monospace; font-size: 12px; padding: 14px; border-radius: 6px; margin-top: 14px; max-height: 360px; overflow: auto; white-space: pre-wrap; display: none; }
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

function baseName(path) { return String(path || '').split('/').pop().split('\\').pop(); }
function stripExt(name) { const i = name.lastIndexOf('.'); return i > 0 ? name.slice(0, i) : name; }
function extOf(name) { const i = name.lastIndexOf('.'); return i >= 0 ? name.slice(i + 1).toLowerCase() : ''; }
function ts() { const d = new Date(), p = n => String(n).padStart(2, '0'); return `${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}_${p(d.getHours())}${p(d.getMinutes())}`; }
function downloadBlob(blob, name) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a'); a.href = url; a.download = name;
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 5000);
}

// ---------------------------------------------------------------------------
export default function createExternalRoot(rootElement) {
  return {
    render(context) {
      const cfg = Object.assign({}, DEFAULTS, (context && context.config) || {});

      const style = document.createElement('style'); style.textContent = CSS;
      const wrap = document.createElement('div'); wrap.className = 'a-wrap';
      wrap.innerHTML = `
        <div class="a-title">🖼️ Asset Import Generator <span style="font-size:12px;color:#a0aec0">(simulation)</span></div>
        <div class="a-sub">Upload a <b>.zip</b> of images. The tool reads every image and builds a Content Hub
          <b>${SHEET_NAME}</b> import workbook. In this simulation the <code>File</code> URL is a mock built from
          <code>${cfg.fileBaseUrl}</code> — no upload happens yet.</div>
        <div class="a-drop" id="a-drop">📦 Drop your .zip here, or click to browse</div>
        <input type="file" id="a-file" accept=".zip,application/zip" style="display:none" />
        <div class="a-row">
          <button class="a-btn a-dry" id="a-dry" disabled>🔍 List images (dry run)</button>
          <button class="a-btn a-go"  id="a-go"  disabled>⬇ Generate import file</button>
          <span id="a-status" style="font-size:13px;color:#555"></span>
        </div>
        <div style="font-size:12px;color:#888;margin-bottom:6px">
          Columns: Title · FileName · FinalLifeCycleStatusToAsset · ContentRepositoryToAsset · File
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

      function onFile(file) {
        if (!file) return;
        clearLog(); currentFile = file;
        status.textContent = `${file.name} — ready`;
        dryBtn.disabled = false; goBtn.disabled = false;
      }
      drop.addEventListener('click', () => input.click());
      input.addEventListener('change', e => onFile(e.target.files[0]));
      drop.addEventListener('dragover', e => { e.preventDefault(); drop.classList.add('a-hover'); });
      drop.addEventListener('dragleave', () => drop.classList.remove('a-hover'));
      drop.addEventListener('drop', e => { e.preventDefault(); drop.classList.remove('a-hover'); onFile(e.dataTransfer.files[0]); });

      // Read image entries from the zip -> [{ name, ext }]
      async function readImages() {
        const JSZip = await loadScript(JSZIP_URL, 'JSZip');
        const zip = await JSZip.loadAsync(currentFile);
        const imgs = [];
        zip.forEach((path, entry) => {
          if (entry.dir) return;
          if (path.indexOf('__MACOSX/') === 0 || baseName(path).indexOf('._') === 0) return; // mac junk
          const name = baseName(path);
          const ext = extOf(name);
          if (IMAGE_EXT.has(ext)) imgs.push({ name, ext });
        });
        // De-duplicate by filename (keep first)
        const seen = new Set(); const out = [];
        for (const im of imgs) { const k = im.name.toLowerCase(); if (!seen.has(k)) { seen.add(k); out.push(im); } }
        return out;
      }

      function rowFor(img) {
        const title = cfg.titleFrom === 'full' ? img.name : stripExt(img.name);
        const file = `${String(cfg.fileBaseUrl).replace(/\/+$/, '')}/${encodeURIComponent(img.name)}`;
        return [title, img.name, cfg.lifecycleStatus, cfg.contentRepository, file];
      }

      async function run(dryRun) {
        clearLog(); dryBtn.disabled = true; goBtn.disabled = true;
        log(dryRun ? '── LIST IMAGES (no file written) ──' : '── GENERATE IMPORT FILE ──', 'a-info');
        try {
          const imgs = await readImages();
          if (!imgs.length) { log('No image files found in the zip.', 'a-err'); return; }
          log(`Found ${imgs.length} image(s).`, 'a-info');
          imgs.slice(0, 200).forEach(im => log(`  ${im.name}`, 'a-skip'));
          if (imgs.length > 200) log(`  … and ${imgs.length - 200} more`, 'a-skip');

          if (dryRun) { log('Dry run complete — Generate to build the M.Asset workbook.', 'a-info'); return; }

          const XLSX = await loadScript(SHEETJS_URL, 'XLSX');
          const header = ['Title', 'FileName', 'FinalLifeCycleStatusToAsset', 'ContentRepositoryToAsset', 'File'];
          const rows = [header, ...imgs.map(rowFor)];
          const ws = XLSX.utils.aoa_to_sheet(rows);
          const wb = XLSX.utils.book_new();
          XLSX.utils.book_append_sheet(wb, ws, SHEET_NAME);
          const arr = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
          const fname = `ContentHub_AssetImport_${ts()}.xlsx`;
          downloadBlob(new Blob([arr], { type: 'application/octet-stream' }), fname);
          log(`✓ Generated ${fname} — ${imgs.length} asset row(s), sheet "${SHEET_NAME}".`, 'a-ok');
          log('⚠ Simulation: the File URLs are mock. Upload the images to Azure and set fileBaseUrl before importing.', 'a-err');
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
