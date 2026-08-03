const $ = (id) => document.getElementById(id);
const el = (tag, cls) => { const e = document.createElement(tag); if (cls) e.className = cls; return e; };

const STATE = {
  attFileName: null, detailed: null, summaryValues: null, attError: null,
  empFileName: null, mapping: null, empError: null,
};

async function loadWorkbook(file) {
  const buf = await file.arrayBuffer();
  const wb = new ExcelJS.Workbook();
  try {
    await wb.xlsx.load(buf);
  } catch (e) {
    throw new Error('Could not read this file as .xlsx. If it\'s a legacy Excel 97-2003 (.xls) file, open it in Excel and re-save as .xlsx first.');
  }
  return wb;
}

async function handleAttFile(files) {
  const file = files[0];
  STATE.attFileName = file.name;
  STATE.attError = null;
  STATE.detailed = null;
  STATE.summaryValues = null;
  try {
    const wb = await loadWorkbook(file);
    const ws = wb.getWorksheet('Detailed');
    if (!ws) throw new Error('This file does not look like an Attendance export — it must contain a "Detailed" tab.');
    STATE.detailed = parseDetailed(ws.getSheetValues());
    const summaryWs = wb.getWorksheet('Summary');
    if (summaryWs) STATE.summaryValues = summaryWs.getSheetValues();
  } catch (e) {
    STATE.attError = e.message;
  }
  renderAll();
}

async function handleEmpFile(files) {
  const file = files[0];
  STATE.empFileName = file.name;
  STATE.empError = null;
  STATE.mapping = null;
  try {
    const wb = await loadWorkbook(file);
    const ws = wb.worksheets[0];
    if (!ws) throw new Error('This file appears to have no sheets.');
    STATE.mapping = parseEmployeeMapping(ws.getSheetValues());
  } catch (e) {
    STATE.empError = e.message;
  }
  renderAll();
}

function currentPlan() {
  if (!STATE.detailed) return null;
  return buildPlan(STATE.detailed, STATE.mapping || new Map());
}

function clearAttFile() {
  STATE.attFileName = null;
  STATE.detailed = null;
  STATE.summaryValues = null;
  STATE.attError = null;
  renderAll();
}

function clearEmpFile() {
  STATE.empFileName = null;
  STATE.mapping = null;
  STATE.empError = null;
  renderAll();
}

function renderFileStatus(prefix, fileName, error, okText, onRemove) {
  const box = $(prefix + 'Status');
  box.innerHTML = '';
  if (!fileName) return;

  const item = el('div', 'fileitem');
  if (error) item.style.borderColor = 'var(--sev-error)';

  const name = el('span'); name.textContent = fileName;
  item.appendChild(name);

  const meta = el('span', 'meta');
  meta.textContent = error || okText;
  if (error) meta.style.color = 'var(--sev-error)';
  item.appendChild(meta);

  const rm = el('button', 'rm');
  rm.type = 'button';
  rm.title = 'Remove file';
  rm.setAttribute('aria-label', 'Remove file');
  rm.textContent = '✕';
  rm.onclick = onRemove;
  item.appendChild(rm);

  box.appendChild(item);
}

function renderAll() {
  renderFileStatus('att', STATE.attFileName, STATE.attError,
    STATE.detailed ? `${new Set(STATE.detailed.rows.map((r) => r.idNumber)).size} employees, ${STATE.detailed.rows.length} day-rows` : '',
    clearAttFile);
  renderFileStatus('emp', STATE.empFileName, STATE.empError,
    STATE.mapping ? `${STATE.mapping.size} Biometric ID mappings` : '',
    clearEmpFile);
  renderValidation();
  $('btnExportFlat').disabled = !STATE.detailed;
  $('btnExportBiologs').disabled = !(STATE.detailed && STATE.mapping);
}

function renderValidation() {
  const sum = $('valSummary'), box = $('valIssues');
  box.innerHTML = '';
  if (!STATE.detailed) { sum.textContent = 'Upload the Attendance Report to begin.'; return; }

  const plan = currentPlan();
  const rules = runValidation(plan);
  if (!STATE.mapping) {
    const idx = rules.findIndex((r) => r.title.includes('no Biometric ID match'));
    if (idx !== -1) {
      rules.splice(idx, 1, {
        sev: 'info',
        title: 'Upload the Employee List Report to enable Biometric ID matching (required for the Biologs template export)',
      });
    }
  }
  if (!STATE.summaryValues) {
    rules.push({ sev: 'info', title: 'No "Summary" sheet found in the Attendance Report — the flat file export will not include a Summary tab.' });
  }

  const nErr = rules.filter((r) => r.sev === 'error').length;
  const nWarn = rules.filter((r) => r.sev === 'warn').length;
  sum.innerHTML = (nErr || nWarn)
    ? `<span style="color:var(--sev-error)">${nErr} error(s)</span>, <span style="color:var(--sev-warn)">${nWarn} warning(s)</span>`
    : '<span class="val-ok">No blocking issues</span>';

  const rank = { error: 0, warn: 1, info: 2, ok: 3 };
  rules.sort((a, b) => rank[a.sev] - rank[b.sev]);

  const CAP = 25;
  for (const rule of rules) {
    const det = el('details', 'issue');
    const expandable = rule.items && rule.items.length;
    if (!expandable) det.classList.add('flat');
    const sm = el('summary');
    const dot = el('span', 'dot ' + rule.sev);
    const t = el('span'); t.textContent = rule.title;
    sm.append(dot, t);
    if (expandable) sm.appendChild(el('span', 'chev'));
    det.appendChild(sm);
    if (expandable) {
      const body = el('div', 'body');
      if (rule.hint) { const h = el('div', 'hint'); h.textContent = rule.hint; body.appendChild(h); }
      rule.items.slice(0, CAP).forEach((it) => {
        const row = el('div', 'vr'); row.textContent = it; body.appendChild(row);
      });
      if (rule.items.length > CAP) {
        const more = el('div', 'hint'); more.textContent = `…and ${rule.items.length - CAP} more`;
        body.appendChild(more);
      }
      det.appendChild(body);
    }
    box.appendChild(det);
  }
}

async function downloadWorkbook(wb, filename) {
  const buf = await wb.xlsx.writeBuffer();
  const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  a.click();
  URL.revokeObjectURL(a.href);
}

async function exportFlat() {
  const btn = $('btnExportFlat'); const label = btn.textContent;
  try {
    btn.disabled = true; btn.textContent = 'Exporting…';
    const wb = buildFlatFile(ExcelJS, currentPlan(), STATE.summaryValues);
    await downloadWorkbook(wb, 'Attendance_Flat.xlsx');
  } catch (e) {
    alert('Export failed: ' + e.message);
  } finally {
    btn.textContent = label;
    renderAll();
  }
}

async function exportBiologs() {
  const btn = $('btnExportBiologs'); const label = btn.textContent;
  try {
    btn.disabled = true; btn.textContent = 'Exporting…';
    const wb = buildBiologsTemplate(ExcelJS, currentPlan());
    await downloadWorkbook(wb, 'HRIS_Biologs.xlsx');
  } catch (e) {
    alert('Export failed: ' + e.message);
  } finally {
    btn.textContent = label;
    renderAll();
  }
}

function initDrop(dropId, inputId, handler) {
  const drop = $(dropId), input = $(inputId);
  drop.onclick = () => input.click();
  input.onchange = () => { if (input.files.length) handler(input.files); input.value = ''; };
  ['dragover', 'dragenter'].forEach((ev) => drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.add('over'); }));
  ['dragleave', 'drop'].forEach((ev) => drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.remove('over'); }));
  drop.addEventListener('drop', (e) => { if (e.dataTransfer.files.length) handler(e.dataTransfer.files); });
}

function init() {
  initDrop('dropAtt', 'attFile', handleAttFile);
  initDrop('dropEmp', 'empFile', handleEmpFile);
  $('btnExportFlat').onclick = exportFlat;
  $('btnExportBiologs').onclick = exportBiologs;
  renderAll();
}

document.addEventListener('DOMContentLoaded', init);
