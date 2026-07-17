/* ═══════════════════════════════════════════
   Kytes Adoption Updater
   Handles: MRN, PO, WO, GRN, DPR, TDS,
            Drawing, Vendor Invoice
   Auto-detects feature type from columns.
   Outputs values in exact Adoption Sheet order.
═══════════════════════════════════════════ */

const PGH_SHEETS = ['PGH 1', 'PGH 2', 'PGH 3', 'PGH 4'];

/* ── Feature definitions ──────────────────
   Each feature:
   - label:       display name
   - col:         Adoption Sheet column header
   - pidPattern:  regex to find Project ID column
   - namePattern: regex to find Project Name column (fallback for GRN)
   - useNameMatch:true if file has no Project ID, must match by name
────────────────────────────────────────── */
const FEATURES = [
  { label: 'MRN No.',           col: 'MRN No.',              pidPattern: /associated.*project|project.*id/i },
  { label: 'Vendor PO',         col: 'Vendor Po No.',        pidPattern: /project\s*no\.?$/i },
  { label: 'Vendor WO',         col: 'Vendor Wo No.',        pidPattern: /project\s*no\.?$/i },
  { label: 'GRN',               col: 'GRN No.',              pidPattern: /project\s*no\.?/i,   dateCol: 'GRN Last Updated date', dateSrcPattern: /grn.?date|created.?date/i },
  { label: 'DPR',               col: 'DPR No.',              pidPattern: /project\s*no\.?/i,   dateCol: 'DPR Last Updated date', dateSrcPattern: /created.?date|dpr.?for.?date/i },
  { label: 'TDS',               col: 'TDS TRACKER (Qty)',    pidPattern: /project\s*no\.?$/i },
  { label: 'Drawing',           col: 'Drawing Tracker (Qty)',pidPattern: /project\s*no\.?$/i },
  { label: 'Vendor Invoice',    col: 'Vendor Invoice',       pidPattern: /project\s*no\.?$/i },
];

/* ── State ── */
let adoptionOrder = null;  // { 'PGH 1': [{pid, title}, ...], ... }
let featureData   = {};    // { featureLabel: { pid: count } }
let featureDates  = {};    // { featureLabel: { pid: latestDate } }
let nameToId      = {};    // project name → pid (built from adoption sheet, for GRN matching)
let activePGH     = 'PGH 1';
let activeFeature = '';

/* ── Init ── */
document.addEventListener('DOMContentLoaded', () => {
  setupDZ('dz-adoption', 'fi-adoption', false, handleAdoptionFile);
  setupDZ('dz-exports',  'fi-exports',  true,  handleExportFiles);
});

/* ── Drop zone setup ── */
function setupDZ(dzId, fiId, multi, handler) {
  const dz = document.getElementById(dzId);
  const fi = document.getElementById(fiId);
  dz.addEventListener('dragover',  e => { e.preventDefault(); dz.classList.add('over'); });
  dz.addEventListener('dragleave', ()  => dz.classList.remove('over'));
  dz.addEventListener('drop', e => {
    e.preventDefault(); dz.classList.remove('over');
    handler(multi ? e.dataTransfer.files : e.dataTransfer.files[0]);
  });
  fi.addEventListener('change', () => {
    handler(multi ? fi.files : fi.files[0]);
    fi.value = '';
  });
}

/* ════════════════════════════════════════
   STEP 1 — Adoption Sheet
════════════════════════════════════════ */
function handleAdoptionFile(file) {
  if (!file || !file.name.endsWith('.xlsx')) { alert('Please upload an .xlsx file.'); return; }
  setStatus('adoption-status', 'loading', 'Reading adoption sheet…');

  readFile(file, ab => {
    try {
      const wb    = XLSX.read(ab, { type: 'array' });
      const order = {};
      nameToId    = {};
      let total   = 0;

      PGH_SHEETS.forEach(sheet => {
        if (!wb.SheetNames.includes(sheet)) return;
        const raw = XLSX.utils.sheet_to_json(wb.Sheets[sheet], { header: 1, defval: '' });
        if (raw.length < 2) return;

        // Find real header row (has "Project ID")
        let hRow = -1;
        for (let i = 0; i < Math.min(5, raw.length); i++) {
          if (raw[i].some(c => /project.?id/i.test(String(c)))) { hRow = i; break; }
        }
        if (hRow < 0) return;

        const headers  = raw[hRow];
        const pidIdx   = headers.findIndex(h => /project.?id/i.test(String(h)));
        const titleIdx = headers.findIndex(h => /project.?title|project.?name/i.test(String(h)));
        if (pidIdx < 0) return;

        order[sheet] = raw.slice(hRow + 1)
          .map(row => ({
            pid:   String(row[pidIdx]   || '').trim(),
            title: String(row[titleIdx] || '').trim()
          }))
          .filter(r => /^DP\d+/i.test(r.pid));

        // Build name→pid map for GRN matching
        order[sheet].forEach(r => {
          if (r.title) nameToId[r.title.toLowerCase().trim()] = r.pid;
        });

        total += order[sheet].length;
      });

      if (!total) {
        setStatus('adoption-status', 'error', 'No valid Project IDs found. Make sure the file has PGH 1–4 sheets with a "Project ID" column.');
        return;
      }

      adoptionOrder = order;
      markStep(1, 'done');
      markStep(2, 'active');

      const summary = PGH_SHEETS.filter(s => order[s]).map(s => `${s}: ${order[s].length}`).join(' · ');
      setStatus('adoption-status', 'ok', `${summary} projects loaded`, file.name);

      if (Object.keys(featureData).length) buildOutput();

    } catch (err) {
      setStatus('adoption-status', 'error', 'Error: ' + err.message);
      console.error(err);
    }
  });
}

/* ════════════════════════════════════════
   STEP 2 — Feature Exports (multiple files)
════════════════════════════════════════ */
function handleExportFiles(files) {
  if (!adoptionOrder) { alert('Please upload the Kytes Adoption Sheet first (Step 1).'); return; }
  if (!files || !files.length) return;

  Array.from(files).forEach(file => {
    if (!file.name.endsWith('.xlsx')) return;
    readFile(file, ab => parseFeatureFile(file.name, ab));
  });
}

function parseFeatureFile(filename, ab) {
  try {
    const wb   = XLSX.read(ab, { type: 'array', cellDates: true });
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: '', raw: true });
    if (!rows.length) return;

    const keys = Object.keys(rows[0]);

    // Auto-detect which feature this file is
    let detected = null;
    let pidKey   = null;
    let nameKey  = null;

    for (const feat of FEATURES) {
      pidKey = keys.find(k => feat.pidPattern.test(k));
      if (pidKey) {
        // Extra disambiguation for files that share same pid column pattern
        // WO files have "WO" in their values, PO files have "PO", etc.
        // Check filename first for disambiguation
        const fn = filename.toLowerCase();
        if (feat.label === 'MRN No.'        && (keys.some(k => /associated/i.test(k)) || fn.includes('list_'))) { detected = feat; break; }
        if (feat.label === 'Vendor PO'       && (fn.includes('po_') || fn.includes('po '))) { detected = feat; break; }
        if (feat.label === 'Vendor WO'       && (fn.includes('wo_') || fn.includes('wo '))) { detected = feat; break; }
        if (feat.label === 'GRN'             && (fn.includes('grn'))) { detected = feat; break; }
        if (feat.label === 'DPR'             && (fn.includes('dpr'))) { detected = feat; break; }
        if (feat.label === 'TDS'             && (fn.includes('tds'))) { detected = feat; break; }
        if (feat.label === 'Drawing'         && (fn.includes('drawing'))) { detected = feat; break; }
        if (feat.label === 'Vendor Invoice'  && (fn.includes('invoice') || fn.includes('vi'))) { detected = feat; break; }
      }
    }

    // Fallback: try to detect by unique column signatures
    if (!detected) {
      if (keys.some(k => /associated.*project/i.test(k)))        detected = FEATURES.find(f => f.label === 'MRN No.');
      else if (keys.some(k => /task.?type/i.test(k))) {
        // TDS or Drawing — check first value
        const firstType = String(rows[0][keys.find(k => /task.?type/i.test(k))] || '').toLowerCase();
        if (firstType.includes('drawing'))  detected = FEATURES.find(f => f.label === 'Drawing');
        else                                detected = FEATURES.find(f => f.label === 'TDS');
      }
      else if (keys.some(k => /grn.?no/i.test(k)) || keys.some(k => /grn.?date/i.test(k))) detected = FEATURES.find(f => f.label === 'GRN');
      else if (keys.some(k => /invoice.?no/i.test(k)))           detected = FEATURES.find(f => f.label === 'Vendor Invoice');
      else if (keys.some(k => /dpr.?for.?date|dpr.?no/i.test(k))) detected = FEATURES.find(f => f.label === 'DPR');
    }

    if (!detected) {
      updateFileChip(filename, 'warn', '⚠ Could not detect feature type');
      return;
    }

    // Re-resolve pidKey for detected feature
    pidKey  = keys.find(k => detected.pidPattern.test(k));
    nameKey = keys.find(k => /project.?name/i.test(k));

    // Count rows per project + track latest date if feature has one
    const counts = {};
    const dates  = {};
    let total    = 0;

    // Find date column in this file if the feature needs it
    const dateKey = detected.dateSrcPattern
      ? keys.find(k => detected.dateSrcPattern.test(String(k)))
      : null;

    rows.forEach(r => {
      let pid = '';

      if (false) {
        // legacy name match - no longer used
      } else {
        pid = String(r[pidKey] || '').trim();
      }

      if (!pid || !/^DP\d+/i.test(pid)) return;
      counts[pid] = (counts[pid] || 0) + 1;
      total++;

      // Track latest date
      if (dateKey) {
        const rawDate = r[dateKey];
        if (rawDate !== '' && rawDate !== null && rawDate !== undefined) {
          let d = null;
          if (rawDate instanceof Date) {
            d = rawDate;
          } else if (typeof rawDate === 'string') {
            d = new Date(rawDate);
          } else if (typeof rawDate === 'number') {
            d = new Date(Math.round((rawDate - 25569) * 86400 * 1000));
          }
          if (d && !isNaN(d.getTime())) {
            const existing = dates[pid] ? new Date(dates[pid].split('/').reverse().join('-')) : null;
            if (!existing || d > existing) {
              const dd   = String(d.getDate()).padStart(2,'0');
              const mm   = String(d.getMonth()+1).padStart(2,'0');
              const yyyy = d.getFullYear();
              dates[pid] = `${dd}/${mm}/${yyyy}`;
            }
          }
        }
      }
    });

    const unique = Object.keys(counts).length;
    featureData[detected.label] = counts;
    if (detected.dateCol && Object.keys(dates).length) {
      featureDates[detected.label] = dates;
    }

    updateFileChip(filename, 'ok', `${detected.label} · ${unique} projects · ${total - unique} dupes removed`);
    markStep(2, 'done');
    buildOutput();

  } catch (err) {
    updateFileChip(filename, 'err', 'Error: ' + err.message);
    console.error(err);
  }
}

/* ════════════════════════════════════════
   FILE LIST UI
════════════════════════════════════════ */
const fileChips = {};

function updateFileChip(filename, status, msg) {
  fileChips[filename] = { status, msg };
  renderFileList();
}

function renderFileList() {
  const el = document.getElementById('file-list');
  el.innerHTML = Object.entries(fileChips).map(([name, { status, msg }]) => `
    <div class="file-chip">
      <i class="ti ti-file-spreadsheet ${status === 'ok' ? 'icon-ok' : status === 'warn' ? 'icon-warn' : 'icon-err'}"></i>
      <span class="fc-name">${esc(name)}</span>
      <span class="badge ${status}">${esc(msg)}</span>
    </div>`).join('');
}

/* ════════════════════════════════════════
   BUILD OUTPUT
════════════════════════════════════════ */
function buildOutput() {
  if (!adoptionOrder || !Object.keys(featureData).length) return;

  const loadedFeatures = FEATURES.filter(f => featureData[f.label]);
  if (!loadedFeatures.length) return;

  // Set default active feature
  if (!loadedFeatures.find(f => f.label === activeFeature)) {
    activeFeature = loadedFeatures[0].label;
  }

  // Stats
  const statsRow = document.getElementById('stats-row');
  statsRow.innerHTML = `
    <div class="stat blue"><div class="val">${Object.values(adoptionOrder).flat().length}</div><div class="lbl">Projects in sheet</div></div>
    <div class="stat green"><div class="val">${loadedFeatures.length}</div><div class="lbl">Features loaded</div></div>
    <div class="stat"><div class="val">${Object.values(featureData).reduce((s, c) => s + Object.keys(c).length, 0)}</div><div class="lbl">Project-feature matches</div></div>`;

  // PGH tabs
  const availPGHs = PGH_SHEETS.filter(s => adoptionOrder[s] && adoptionOrder[s].length);
  if (!availPGHs.includes(activePGH)) activePGH = availPGHs[0];

  document.getElementById('tab-row').innerHTML = availPGHs.map(s => `
    <button class="tab-btn ${s === activePGH ? 'active' : ''}" onclick="switchPGH('${s}')">
      ${s} <span class="tab-count">${adoptionOrder[s].length}</span>
    </button>`).join('');

  // Feature tabs
  document.getElementById('feature-tab-row').innerHTML = loadedFeatures.map(f => `
    <button class="ftab-btn ${f.label === activeFeature ? 'active' : ''}" onclick="switchFeature('${f.label}')">
      ${f.label}
    </button>`).join('');

  // Build all PGH×Feature tables
  const container = document.getElementById('pgh-tables');
  container.innerHTML = availPGHs.map(sheet => {
    const projects = adoptionOrder[sheet] || [];
    return loadedFeatures.map(feat => {
      const counts  = featureData[feat.label] || {};
      const tableId = `tbl-${sheet.replace(' ','-')}-${feat.label.replace(/\W+/g,'-')}`;
      const visible = sheet === activePGH && feat.label === activeFeature;
      const rows    = projects.map((p, i) => {
        const val     = counts[p.pid] !== undefined ? counts[p.pid] : 0;
        const hasData = counts[p.pid] !== undefined;
        let dateVal = '';
        if (feat.dateCol && featureDates[feat.label] && featureDates[feat.label][p.pid]) {
          dateVal = featureDates[feat.label][p.pid];
        }
        const dateTd = feat.dateCol ? `<td class="td-date">${esc(dateVal)}</td>` : '';
        return `<tr class="${!hasData ? 'row-missing' : ''}">
          <td class="td-num">${i + 1}</td>
          <td class="td-pid">${esc(p.pid)}</td>
          <td class="td-title">${esc(p.title)}</td>
          <td class="td-mrn ${!hasData ? 'mrn-zero' : ''}">${val}</td>
          ${dateTd}
        </tr>`;
      }).join('');

      return `
        <div class="pgh-table ${visible ? '' : 'hidden'}" id="${tableId}">
          <table>
            <thead><tr>
              <th class="th-num">#</th>
              <th>Project ID</th>
              <th>Project Title</th>
              <th class="th-mrn">${esc(feat.col)} ← copy this</th>
              ${feat.dateCol ? `<th class="th-date">${esc(feat.dateCol)} ← copy this</th>` : ''}
            </tr></thead>
            <tbody>${rows}</tbody>
          </table>
        </div>`;
    }).join('');
  }).join('');

  updateCopyBtn();
  document.getElementById('card-output').style.display = 'block';
  document.getElementById('card-output').scrollIntoView({ behavior: 'smooth', block: 'start' });
}

/* ════════════════════════════════════════
   TAB SWITCHING
════════════════════════════════════════ */
function switchPGH(pgh) {
  activePGH = pgh;
  document.querySelectorAll('.tab-btn').forEach(b =>
    b.classList.toggle('active', b.textContent.trim().startsWith(pgh)));
  refreshTableVisibility();
  updateCopyBtn();
}

function switchFeature(feat) {
  activeFeature = feat;
  document.querySelectorAll('.ftab-btn').forEach(b =>
    b.classList.toggle('active', b.textContent.trim() === feat));
  refreshTableVisibility();
  updateCopyBtn();
}

function refreshTableVisibility() {
  document.querySelectorAll('.pgh-table').forEach(t => {
    const id = `tbl-${activePGH.replace(' ','-')}-${activeFeature.replace(/\W+/g,'-')}`;
    t.classList.toggle('hidden', t.id !== id);
  });
}

function updateCopyBtn() {
  const feat = FEATURES.find(f => f.label === activeFeature);
  document.getElementById('copy-btn').innerHTML =
    `<i class="ti ti-clipboard"></i> Copy ${activeFeature} count — ${activePGH}`;
  const dateBtn = document.getElementById('copy-date-btn');
  if (dateBtn) {
    const showDate = feat && feat.dateCol && featureDates[activeFeature];
    dateBtn.style.display = showDate ? '' : 'none';
    if (showDate) {
      dateBtn.innerHTML = `<i class="ti ti-calendar"></i> Copy ${feat.dateCol} — ${activePGH}`;
    }
  }
}

/* ════════════════════════════════════════
   COPY
════════════════════════════════════════ */
function copyCurrentColumn(copyDate = false) {
  if (!adoptionOrder) return;
  const projects = adoptionOrder[activePGH] || [];
  const counts   = featureData[activeFeature] || {};
  const dates    = featureDates[activeFeature] || {};
  let values;
  if (copyDate) {
    values = projects.map(p => dates[p.pid] || '');
  } else {
    values = projects.map(p => counts[p.pid] !== undefined ? counts[p.pid] : 0);
  }
  const text     = values.join('\n');

  navigator.clipboard.writeText(text)
    .then(() => flashBtn(values.length, copyDate))
    .catch(() => {
      const ta = document.createElement('textarea');
      ta.value = text; document.body.appendChild(ta); ta.select();
      document.execCommand('copy'); document.body.removeChild(ta);
      flashBtn(values.length, copyDate);
    });
}

function flashBtn(count, isDate = false) {
  const btnId = isDate ? 'copy-date-btn' : 'copy-btn';
  const btn = document.getElementById(btnId);
  if (!btn) return;
  btn.innerHTML = `<i class="ti ti-check"></i> Copied ${count} values!`;
  btn.classList.add('copied');
  setTimeout(() => { updateCopyBtn(); btn.classList.remove('copied'); }, 2500);
}

/* ════════════════════════════════════════
   DOWNLOAD ALL AS EXCEL
════════════════════════════════════════ */
function downloadExcel() {
  if (!adoptionOrder) return;
  const wb = XLSX.utils.book_new();
  const loadedFeatures = FEATURES.filter(f => featureData[f.label]);

  PGH_SHEETS.forEach(sheet => {
    const projects = adoptionOrder[sheet];
    if (!projects || !projects.length) return;

    const data = projects.map((p, i) => {
      const row = { '#': i + 1, 'Project ID': p.pid, 'Project Title': p.title };
      loadedFeatures.forEach(f => {
        row[f.col] = featureData[f.label][p.pid] !== undefined ? featureData[f.label][p.pid] : 0;
        if (f.dateCol) {
          row[f.dateCol] = featureDates[f.label] && featureDates[f.label][p.pid] ? featureDates[f.label][p.pid] : '';
        }
      });
      return row;
    });

    const ws = XLSX.utils.json_to_sheet(data);
    const cols = [{ wch: 5 }, { wch: 12 }, { wch: 42 }];
    loadedFeatures.forEach(f => {
      cols.push({ wch: 16 });
      if (f.dateCol) cols.push({ wch: 22 });
    });
    ws['!cols'] = cols;
    XLSX.utils.book_append_sheet(wb, ws, sheet);
  });

  const date = new Date().toISOString().split('T')[0];
  XLSX.writeFile(wb, `kytes_update_${date}.xlsx`);
}

/* ════════════════════════════════════════
   RESET
════════════════════════════════════════ */
function resetAll() {
  adoptionOrder = null; featureData = {}; featureDates = {}; nameToId = {};
  activePGH = 'PGH 1'; activeFeature = '';
  Object.keys(fileChips).forEach(k => delete fileChips[k]);
  ['adoption-status'].forEach(id => document.getElementById(id).innerHTML = '');
  document.getElementById('file-list').innerHTML = '';
  document.getElementById('card-output').style.display = 'none';
  document.getElementById('fi-adoption').value = '';
  document.getElementById('fi-exports').value  = '';
  markStep(1, 'active'); markStep(2, '');
}

/* ════════════════════════════════════════
   HELPERS
════════════════════════════════════════ */
function readFile(file, cb) {
  const r = new FileReader();
  r.onload = e => cb(e.target.result);
  r.readAsArrayBuffer(file);
}

function setStatus(elId, type, msg, filename) {
  const icons = { ok: 'ti-circle-check', error: 'ti-alert-circle', loading: 'ti-loader-2' };
  const cls   = { ok: 'ok', error: 'err', loading: 'loading' };
  document.getElementById(elId).innerHTML = `
    <div class="status-chip ${cls[type]}">
      <i class="ti ${icons[type]} ${type === 'loading' ? 'spin' : ''}"></i>
      <div>${filename ? `<strong>${esc(filename)}</strong><br>` : ''}${msg}</div>
    </div>`;
}

function markStep(n, state) {
  const el = document.getElementById('n' + n);
  if (!el) return;
  el.className = 'step-num' + (state ? ' ' + state : '');
  el.innerHTML = state === 'done' ? '<i class="ti ti-check" style="font-size:11px"></i>' : n;
}

function esc(s) {
  return String(s || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}
