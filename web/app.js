/*
 * Savvy Repair for Microsoft Office — browser implementation.
 *
 * Reimplements the four recovery methods from the original VB.NET app:
 *   1. Zip structure repair — scan for PK\x03\x04 signatures and re-pack
 *      every recoverable entry, using ImmortalInflate so truncated DEFLATE
 *      streams (a common failure mode in corrupt .docx/.xlsx/.pptx files
 *      that makes plain JSZip throw) still yield partial bytes.
 *   2. Strict XML — truncate each part at the first XML parse error.
 *   3. Lax XML — drop the most-broken element until what remains parses.
 *   4. Plain-text salvage — extract <w:t>/<a:t>/<t> runs into a .txt.
 *
 * Runs entirely client-side. No upload. Tolerant unzip lives in
 * ./immortal-inflate.js; JSZip is only used to assemble clean outputs.
 */
(function () {
  'use strict';

  const $ = (id) => document.getElementById(id);
  const drop = $('drop'), fileInput = $('file'), runBtn = $('run'),
        clearBtn = $('clear'), statusEl = $('status'),
        logEl = $('log'), resultsEl = $('results'), pickedEl = $('picked');

  let selectedFile = null;
  let blobUrls = [];

  function log(msg, cls) {
    const line = document.createElement('div');
    if (cls) line.className = cls;
    line.textContent = msg;
    logEl.appendChild(line);
    logEl.scrollTop = logEl.scrollHeight;
  }
  function setStatus(text) { statusEl.textContent = text || ''; }

  function clearAll() {
    selectedFile = null;
    pickedEl.textContent = '';
    runBtn.disabled = true;
    logEl.innerHTML = '';
    resultsEl.innerHTML = '';
    setStatus('');
    blobUrls.forEach(URL.revokeObjectURL);
    blobUrls = [];
    fileInput.value = '';
  }

  function pickFile(f) {
    if (!f) return;
    const name = f.name.toLowerCase();
    if (!/\.(docx|xlsx|pptx|zip)$/.test(name)) {
      pickedEl.textContent = 'Unsupported file type. Pick a .docx, .xlsx, or .pptx.';
      runBtn.disabled = true; return;
    }
    selectedFile = f;
    pickedEl.textContent = `${f.name} — ${(f.size / 1024).toFixed(1)} KB`;
    runBtn.disabled = false;
  }

  drop.addEventListener('click', () => fileInput.click());
  drop.addEventListener('dragover', (e) => { e.preventDefault(); drop.classList.add('hover'); });
  drop.addEventListener('dragleave', () => drop.classList.remove('hover'));
  drop.addEventListener('drop', (e) => {
    e.preventDefault(); drop.classList.remove('hover');
    if (e.dataTransfer.files && e.dataTransfer.files[0]) pickFile(e.dataTransfer.files[0]);
  });
  fileInput.addEventListener('change', () => pickFile(fileInput.files[0]));
  clearBtn.addEventListener('click', clearAll);

  function addResult(name, blob, meta) {
    const url = URL.createObjectURL(blob);
    blobUrls.push(url);
    const row = document.createElement('div');
    row.className = 'result';
    const left = document.createElement('div');
    const n = document.createElement('div'); n.className = 'name'; n.textContent = name;
    const m = document.createElement('div'); m.className = 'meta'; m.textContent = meta;
    left.appendChild(n); left.appendChild(m);
    const a = document.createElement('a');
    a.href = url; a.download = name; a.textContent = 'Download';
    row.appendChild(left); row.appendChild(a);
    resultsEl.appendChild(row);
  }

  function ext(name) { const m = name.match(/\.([^.]+)$/); return m ? m[1].toLowerCase() : ''; }
  function baseName(name) { return name.replace(/\.[^.]+$/, ''); }

  async function readBytes(file) {
    const buf = await file.arrayBuffer();
    return new Uint8Array(buf);
  }

  /* ---------- Tolerant read: prefer immortal scan, fall back to JSZip ---------- */
  async function readEntries(file) {
    const u8 = await readBytes(file);
    const scan = tolerantUnzip(u8);
    const count = Object.keys(scan.entries).length;
    if (count > 0) {
      log(`Scanned ${count} entries via ImmortalInflate (${scan.partial.size} partial)`, 'info');
      return scan;
    }
    // Last resort: a perfectly valid zip might not match the heuristic if it
    // uses unusual extra fields; fall through to JSZip.
    log('ImmortalInflate found no entries — falling back to JSZip', 'warn');
    const zip = await JSZip.loadAsync(u8, { checkCRC32: false });
    const entries = {}, partial = new Set();
    for (const path of Object.keys(zip.files)) {
      const entry = zip.files[path];
      if (entry.dir) continue;
      try { entries[path] = await entry.async('uint8array'); }
      catch (e) { partial.add(path); log(`  could not inflate ${path}: ${e.message}`, 'warn'); }
    }
    return { entries, partial };
  }

  async function rebuildZip(entries, fileType) {
    const out = new JSZip();
    for (const [name, data] of Object.entries(entries)) out.file(name, data);
    return out.generateAsync({
      type: 'blob',
      mimeType: fileType || 'application/zip',
      compression: 'DEFLATE'
    });
  }

  /* ---------- Method 1: zip structure repair ---------- */
  async function repairZip(file) {
    const { entries, partial } = await readEntries(file);
    const count = Object.keys(entries).length;
    if (!count) { log('Zip repair: no entries recoverable', 'err'); return null; }
    log(`Zip repair: ${count} entries (${partial.size} partial) re-packed`, 'ok');
    return rebuildZip(entries, file.type);
  }

  /* ---------- XML helpers ---------- */
  const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
  const decoder = new TextDecoder('utf-8');
  const encoder = new TextEncoder();

  function parseXmlError(text) {
    const doc = new DOMParser().parseFromString(text, 'application/xml');
    const err = doc.getElementsByTagName('parsererror')[0];
    return err ? err.textContent : null;
  }

  /* ---------- Method 2: strict truncation ---------- */
  function strictTruncate(xml) {
    if (!parseXmlError(xml)) return { text: xml, changed: false };
    const opens = [];
    let i = 0;
    while (i < xml.length) {
      const lt = xml.indexOf('<', i); if (lt < 0) break;
      const gt = xml.indexOf('>', lt); if (gt < 0) break;
      opens.push(gt + 1);
      i = gt + 1;
    }
    for (let p = opens.length - 1; p >= 0; p--) {
      const end = opens[p];
      const localStack = [];
      let j = 0;
      while (j < end) {
        const lt = xml.indexOf('<', j); if (lt < 0 || lt >= end) break;
        const gt = xml.indexOf('>', lt); if (gt < 0 || gt >= end) break;
        const tag = xml.slice(lt, gt + 1);
        const m = tag.match(/^<\s*\/?\s*([^\s\/>]+)/);
        if (m && !tag.startsWith('<?') && !tag.startsWith('<!')) {
          const isClose = tag.startsWith('</');
          const selfClose = /\/\s*>$/.test(tag);
          if (!isClose && !selfClose) localStack.push(m[1]);
          else if (isClose && localStack.length && localStack[localStack.length - 1] === m[1]) localStack.pop();
        }
        j = gt + 1;
      }
      let candidate = xml.slice(0, end);
      while (localStack.length) candidate += `</${localStack.pop()}>`;
      if (!parseXmlError(candidate)) return { text: candidate, changed: true };
    }
    return { text: xml, changed: false, failed: true };
  }

  /* ---------- Method 3: lax repair ---------- */
  function laxRepair(xml) {
    let working = xml;
    for (let pass = 0; pass < 30; pass++) {
      if (!parseXmlError(working)) return { text: working, changed: pass > 0 };
      const removed = removeFirstBrokenElement(working);
      if (!removed) break;
      working = removed;
    }
    const tr = strictTruncate(working);
    return { text: tr.text, changed: true, failed: !!tr.failed };
  }

  function removeFirstBrokenElement(xml) {
    const rootOpen = xml.search(/<[A-Za-z_][^?!]/);
    if (rootOpen < 0) return null;
    const rootEnd = xml.indexOf('>', rootOpen);
    if (rootEnd < 0) return null;
    let i = rootEnd + 1;
    while (i < xml.length) {
      const lt = xml.indexOf('<', i); if (lt < 0) break;
      if (xml[lt + 1] === '/') break;
      if (xml[lt + 1] === '!' || xml[lt + 1] === '?') { i = xml.indexOf('>', lt) + 1; continue; }
      const span = findElementSpan(xml, lt);
      if (!span) break;
      const fragment = xml.slice(span.start, span.end);
      if (parseXmlError('<wrap xmlns:x="x">' + fragment + '</wrap>')) {
        return xml.slice(0, span.start) + xml.slice(span.end);
      }
      i = span.end;
    }
    return null;
  }

  function findElementSpan(xml, start) {
    const open = xml.indexOf('>', start);
    if (open < 0) return null;
    const tagSrc = xml.slice(start, open + 1);
    const m = tagSrc.match(/^<\s*([^\s\/>]+)/);
    if (!m) return null;
    if (/\/\s*>$/.test(tagSrc)) return { start, end: open + 1 };
    const name = m[1];
    const openRe = new RegExp(`<\\s*${escapeRe(name)}\\b[^>]*?>`, 'g');
    const closeRe = new RegExp(`</\\s*${escapeRe(name)}\\s*>`, 'g');
    let depth = 1; let pos = open + 1;
    while (pos < xml.length) {
      openRe.lastIndex = pos; closeRe.lastIndex = pos;
      const o = openRe.exec(xml); const c = closeRe.exec(xml);
      if (!c) return null;
      if (o && o.index < c.index) { depth++; pos = openRe.lastIndex; }
      else { depth--; pos = closeRe.lastIndex; if (depth === 0) return { start, end: pos }; }
    }
    return null;
  }
  function escapeRe(s) { return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }

  /* ---------- Apply XML repair across every .xml/.rels entry ---------- */
  async function processXmlParts(file, mode) {
    const { entries, partial } = await readEntries(file);
    if (!Object.keys(entries).length) { log(`XML ${mode}: nothing to read`, 'err'); return null; }
    let touched = 0, fixed = 0, failed = 0;
    const repaired = { ...entries };
    for (const path of Object.keys(repaired)) {
      if (!/\.(xml|rels)$/i.test(path)) continue;
      let text;
      try { text = decoder.decode(repaired[path]); }
      catch { continue; }
      touched++;
      const err = parseXmlError(text);
      if (!err && !partial.has(path)) continue;
      const res = mode === 'strict' ? strictTruncate(text) : laxRepair(text);
      if (res.failed) { failed++; log(`  ${path}: still invalid after ${mode}`, 'warn'); continue; }
      if (res.changed) {
        let final = res.text;
        if (!/^\s*<\?xml/i.test(final)) final = XML_DECL + final;
        repaired[path] = encoder.encode(final);
        fixed++;
        log(`  ${path}: repaired (${mode})${partial.has(path) ? ' [from truncated stream]' : ''}`, 'ok');
      }
    }
    log(`XML ${mode}: ${fixed}/${touched} parts fixed, ${failed} unfixable`, fixed ? 'ok' : 'info');
    return rebuildZip(repaired, file.type);
  }

  /* ---------- Method 4: plain-text salvage ---------- */
  async function textSalvage(file) {
    const { entries } = await readEntries(file);
    if (!Object.keys(entries).length) { log('Text salvage: nothing to read', 'err'); return null; }
    const parts = [];
    const candidates = Object.keys(entries).filter((p) =>
      /word\/.*\.xml$/i.test(p) ||
      /xl\/.*\.xml$/i.test(p) ||
      /ppt\/slides\/.*\.xml$/i.test(p) ||
      /ppt\/notesSlides\/.*\.xml$/i.test(p)
    );
    for (const path of candidates) {
      let text;
      try { text = decoder.decode(entries[path]); } catch { continue; }
      const withBreaks = text
        .replace(/<\/w:p\s*>/g, '\n')
        .replace(/<\/a:p\s*>/g, '\n')
        .replace(/<\/si\s*>/g, '\n')
        .replace(/<\/row\s*>/g, '\n');
      const re = /<(?:w:t|a:t|t)\b[^>]*>([\s\S]*?)<\/(?:w:t|a:t|t)>/g;
      const chunks = []; let m;
      while ((m = re.exec(withBreaks)) !== null) {
        const s = decodeXmlEntities(m[1]).replace(/\s+/g, ' ').trim();
        if (s) chunks.push(s);
      }
      if (chunks.length) parts.push(`--- ${path} ---\n${chunks.join('\n')}\n`);
    }
    if (!parts.length) { log('Text salvage: no recognisable text runs found', 'warn'); return null; }
    log(`Text salvage: extracted text from ${parts.length} parts`, 'ok');
    return new Blob([parts.join('\n')], { type: 'text/plain;charset=utf-8' });
  }

  function decodeXmlEntities(s) {
    return s
      .replace(/&lt;/g, '<').replace(/&gt;/g, '>')
      .replace(/&quot;/g, '"').replace(/&apos;/g, "'")
      .replace(/&#x([0-9a-fA-F]+);/g, (_, h) => String.fromCodePoint(parseInt(h, 16)))
      .replace(/&#(\d+);/g, (_, d) => String.fromCodePoint(parseInt(d, 10)))
      .replace(/&amp;/g, '&');
  }

  /* ---------- Driver ---------- */
  runBtn.addEventListener('click', async () => {
    if (!selectedFile) return;
    runBtn.disabled = true; clearBtn.disabled = true;
    logEl.innerHTML = ''; resultsEl.innerHTML = '';
    blobUrls.forEach(URL.revokeObjectURL); blobUrls = [];

    const base = baseName(selectedFile.name);
    const e = ext(selectedFile.name) || 'docx';
    log(`Working on ${selectedFile.name} (${(selectedFile.size / 1024).toFixed(1)} KB)`, 'info');

    try {
      if ($('opt-zip').checked) {
        setStatus('Method 1: zip structure repair…');
        const blob = await repairZip(selectedFile);
        if (blob) addResult(`${base}_zip-repaired.${e}`, blob, `Method 1 · ${(blob.size / 1024).toFixed(1)} KB`);
      }
      if ($('opt-strict').checked) {
        setStatus('Method 2: strict XML truncation…');
        const blob = await processXmlParts(selectedFile, 'strict');
        if (blob) addResult(`${base}_xml-strict.${e}`, blob, `Method 2 · ${(blob.size / 1024).toFixed(1)} KB`);
      }
      if ($('opt-lax').checked) {
        setStatus('Method 3: lax XML repair…');
        const blob = await processXmlParts(selectedFile, 'lax');
        if (blob) addResult(`${base}_xml-lax.${e}`, blob, `Method 3 · ${(blob.size / 1024).toFixed(1)} KB`);
      }
      if ($('opt-text').checked) {
        setStatus('Method 4: plain-text salvage…');
        const blob = await textSalvage(selectedFile);
        if (blob) addResult(`${base}_salvaged.txt`, blob, `Method 4 · ${(blob.size / 1024).toFixed(1)} KB`);
      }
      setStatus(resultsEl.children.length
        ? 'Done. Pick a download above.'
        : 'No outputs produced — file may be unrecoverable.');
    } catch (err) {
      log(`Fatal: ${err.message}`, 'err');
      setStatus('Repair failed. See log.');
    } finally {
      runBtn.disabled = false; clearBtn.disabled = false;
    }
  });

  /* ---------- PWA install prompt ---------- */
  let deferredPrompt = null;
  const installLink = $('install');
  window.addEventListener('beforeinstallprompt', (e) => {
    e.preventDefault();
    deferredPrompt = e;
    installLink.style.display = 'inline-flex';
  });
  installLink.addEventListener('click', async (e) => {
    e.preventDefault();
    if (!deferredPrompt) return;
    deferredPrompt.prompt();
    await deferredPrompt.userChoice;
    deferredPrompt = null;
    installLink.style.display = 'none';
  });
  ['cta-chromeos', 'cta-android', 'cta-ios'].forEach((id) => {
    const el = $(id); if (!el) return;
    el.addEventListener('click', (ev) => {
      ev.preventDefault();
      if (deferredPrompt) { deferredPrompt.prompt(); return; }
      alert(
        id === 'cta-ios'
          ? "iOS: tap Share, then 'Add to Home Screen'."
          : id === 'cta-chromeos'
            ? 'ChromeOS: use the install icon in the address bar, or menu → Install.'
            : 'Android: open menu → Install app / Add to Home Screen.'
      );
    });
  });
})();
