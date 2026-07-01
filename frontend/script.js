const API = '';

// ── Shared helpers ─────────────────────────────────────────────────────────────
async function loadDesignations() {
  const res = await fetch(`${API}/api/designations`);
  return res.json();
}

async function populateDesignationDropdowns() {
  let list;
  try { list = await loadDesignations(); } catch { list = []; }
  const selects = document.querySelectorAll('.desig-select');
  selects.forEach(sel => {
    const current = sel.value;
    sel.innerHTML = '<option value="">-- Select Designation --</option>';
    list.forEach(d => {
      const opt = document.createElement('option');
      opt.value = d.name;
      opt.textContent = `${d.name} (₹${d.rate}/day)`;
      opt.dataset.rate = d.rate;
      sel.appendChild(opt);
    });
    if (current) sel.value = current;
  });
  return list;
}

function showError(id, msg) {
  document.getElementById(id).textContent = msg;
}

// ── Tab switching ─────────────────────────────────────────────────────────────
function switchTab(tab) {
  document.querySelectorAll('.tab-content').forEach(el => el.classList.add('hidden'));
  document.querySelectorAll('.tab-btn').forEach(el => el.classList.remove('active'));
  document.getElementById(`tab-${tab}`).classList.remove('hidden');
  event.target.classList.add('active');
  if (tab === 'designations' || tab === 'multi' || tab === 'manual') {
    populateDesignationDropdowns();
  }
}

// ── File name display ─────────────────────────────────────────────────────────
document.getElementById('scheduleFile').addEventListener('change', function () {
  const name = this.files[0] ? this.files[0].name : 'No file selected';
  document.getElementById('fileNameDisplay').textContent = name;
});

// ── Designation Management ─────────────────────────────────────────────────────
async function renderDesignations() {
  let list;
  try { list = await loadDesignations(); } catch { list = []; }
  const tbody = document.getElementById('designationsBody');
  tbody.innerHTML = '';
  list.forEach(d => {
    const tr = document.createElement('tr');
    tr.dataset.desigName = d.name;
    tr.innerHTML = `
      <td>${d.name}</td>
      <td><input type="number" class="desig-rate-input" value="${d.rate}" min="1"/></td>
      <td>
        <button class="small-btn save-btn">💾</button>
        <button class="small-btn del-btn">🗑</button>
      </td>`;
    tbody.appendChild(tr);
  });
}

async function addDesignation() {
  const name = document.getElementById('desigNameInput').value.trim();
  const rate = Number(document.getElementById('desigRateInput').value);
  if (!name) { showError('desigError', 'Enter a designation name.'); return; }
  if (!rate || rate <= 0) { showError('desigError', 'Enter a valid daily rate.'); return; }
  showError('desigError', '');
  try {
    const res = await fetch(`${API}/api/designations`, {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ name, rate })
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error);
    document.getElementById('desigNameInput').value = '';
    document.getElementById('desigRateInput').value = '';
    await renderDesignations();
    await populateDesignationDropdowns();
  } catch (err) { showError('desigError', err.message); }
}

// ── Parse Schedule ─────────────────────────────────────────────────────────────
let parsedPersons = [];

async function parseSchedule() {
  const fileInput = document.getElementById('scheduleFile');
  const errEl = document.getElementById('parseError');
  errEl.textContent = '';
  document.getElementById('personsSection').classList.add('hidden');
  document.getElementById('resultsSection').classList.add('hidden');

  if (!fileInput.files[0]) { errEl.textContent = 'Please select a .xlsx schedule file.'; return; }

  const btn = document.getElementById('parseBtn');
  btn.disabled = true; btn.textContent = 'Parsing...';

  try {
    const fd = new FormData();
    fd.append('schedule', fileInput.files[0]);
    const res = await fetch(`${API}/api/parse-schedule`, { method: 'POST', body: fd });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || 'Parse failed');

    parsedPersons = data.persons;
    renderPersonsTable(parsedPersons);
    document.getElementById('personsSection').classList.remove('hidden');
  } catch (err) {
    errEl.textContent = err.message;
  } finally {
    btn.disabled = false; btn.textContent = 'Parse Schedule';
  }
}

function renderPersonsTable(persons) {
  document.getElementById('personCount').textContent = persons.length;
  const tbody = document.getElementById('personsBody');
  tbody.innerHTML = '';
  persons.forEach((p, i) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${i + 1}</td>
      <td>${p.name}</td>
      <td>${p.designation || '-'}</td>
      <td class="dates-cell">${p.dates.join(', ')}</td>
      <td>${p.dates.length}</td>
      <td>₹${p.rate * p.dates.length}</td>`;
    tbody.appendChild(tr);
  });
}

// ── Generate All (bulk) ────────────────────────────────────────────────────────
async function generateAll() {
  const errEl = document.getElementById('generateError');
  errEl.textContent = '';
  document.getElementById('resultsSection').classList.add('hidden');

  if (!parsedPersons.length) { errEl.textContent = 'No persons to generate.'; return; }

  const btn = document.getElementById('generateAllBtn');
  btn.disabled = true; btn.textContent = `⏳ Generating ${parsedPersons.length} bills...`;

  try {
    const res = await fetch(`${API}/api/generate-all`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ persons: parsedPersons })
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || 'Generation failed');

    document.getElementById('combinedPdfBtn').href = `${API}${data.combinedPdfUrl}`;

    const tbody = document.getElementById('resultsBody');
    tbody.innerHTML = '';
    data.results.forEach((r, i) => {
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td>${i + 1}</td>
        <td>${r.name}</td>
        <td>${r.designation || '-'}</td>
        <td>${r.days}</td>
        <td>₹${r.amount}</td>
        <td><a class="dl-btn dl-xlsx" href="${API}${r.xlsxUrl}" download>⬇ Excel</a></td>
        <td><a class="dl-btn dl-pdf" href="${API}${r.pdfUrl}" download>⬇ PDF</a></td>`;
      tbody.appendChild(tr);
    });

    document.getElementById('resultsSection').classList.remove('hidden');
    document.getElementById('resultsSection').scrollIntoView({ behavior: 'smooth' });
  } catch (err) {
    errEl.textContent = err.message;
  } finally {
    btn.disabled = false; btn.textContent = '⚡ Generate All Bills';
  }
}

// ── Single Entry ──────────────────────────────────────────────────────────────
const datesContainer = document.getElementById('datesContainer');

document.getElementById('desigSingle').addEventListener('change', recalc);

function addDateRow(value = '') {
  const row = document.createElement('div');
  row.className = 'date-row';
  const input = document.createElement('input');
  input.type = 'date'; input.value = value;
  input.addEventListener('change', recalc);
  const rm = document.createElement('button');
  rm.textContent = '×'; rm.addEventListener('click', () => { row.remove(); recalc(); });
  row.appendChild(input); row.appendChild(rm);
  datesContainer.appendChild(row);
  recalc();
}

function getDates() {
  return Array.from(datesContainer.querySelectorAll('input[type="date"]'))
    .map(i => i.value).filter(Boolean);
}

function recalc() {
  const dates = getDates();
  document.getElementById('days').value = dates.length;
  const sel = document.getElementById('desigSingle');
  const opt = sel.selectedOptions[0];
  const rate = opt && opt.dataset.rate ? Number(opt.dataset.rate) : 0;
  const total = dates.length * rate;
  document.getElementById('amountDisplay').textContent = `₹${total} (${dates.length} days × ₹${rate})`;
}

async function generateSingle() {
  const errEl = document.getElementById('errorMsg');
  const dlLinks = document.getElementById('downloadLinks');
  errEl.textContent = ''; dlLinks.classList.add('hidden');

  const name = document.getElementById('name').value.trim();
  const dates = getDates();
  const designation = document.getElementById('desigSingle').value;
  const days = Number(document.getElementById('days').value);

  if (!name) { errEl.textContent = 'Please enter the name.'; return; }
  if (!designation) { errEl.textContent = 'Please select a designation.'; return; }
  if (dates.length === 0) { errEl.textContent = 'Please add at least one date.'; return; }

  const btn = document.getElementById('generateBtn');
  btn.disabled = true; btn.textContent = 'Generating...';

  try {
    const res = await fetch(`${API}/api/generate`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ name, dates, daysWorked: days, designation })
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || 'Failed');

    document.getElementById('amountDisplay').textContent = `₹${data.amount}`;
    document.getElementById('xlsxLink').href = `${API}${data.xlsxUrl}`;
    document.getElementById('pdfLink').href = `${API}${data.pdfUrl}`;
    dlLinks.classList.remove('hidden');
  } catch (err) {
    errEl.textContent = err.message;
  } finally {
    btn.disabled = false; btn.textContent = 'Generate Excel & PDF';
  }
}

// ── Multi-Person Entry ────────────────────────────────────────────────────────
let multiPersonCount = 0;

async function addMultiPerson(nameValue = '') {
  multiPersonCount++;
  const container = document.getElementById('multiPersonsContainer');
  const idx = multiPersonCount;

  const entry = document.createElement('div');
  entry.className = 'card multi-person-entry';
  entry.style.position = 'relative';

  const removeBtn = document.createElement('button');
  removeBtn.textContent = '× Remove';
  removeBtn.className = 'remove-person-btn';
  removeBtn.onclick = () => { entry.remove(); };

  const nameLabel = document.createElement('label');
  nameLabel.textContent = 'Name';
  const nameInput = document.createElement('input');
  nameInput.type = 'text';
  nameInput.placeholder = 'Enter full name (Surname first)';
  nameInput.value = nameValue;
  nameInput.className = 'mp-name';

  const desigLabel = document.createElement('label');
  desigLabel.textContent = 'Designation';
  const desigSelect = document.createElement('select');
  desigSelect.className = 'desig-select mp-desig';

  const datesLabel = document.createElement('label');
  datesLabel.textContent = 'Dates';
  const datesContainer = document.createElement('div');
  datesContainer.className = 'mp-dates';

  const daysLabel = document.createElement('label');
  daysLabel.textContent = 'Number of Days Worked';
  const daysInput = document.createElement('input');
  daysInput.type = 'number';
  daysInput.readOnly = true;
  daysInput.className = 'mp-days';

  const amountRow = document.createElement('div');
  amountRow.className = 'amount-row';
  const amountSpan = document.createElement('span');
  amountSpan.textContent = 'Amount: ₹0';

  function recalcMulti() {
    const dates = Array.from(datesContainer.querySelectorAll('input[type="date"]'))
      .map(i => i.value).filter(Boolean);
    daysInput.value = dates.length;
    const opt = desigSelect.selectedOptions[0];
    const rate = opt && opt.dataset.rate ? Number(opt.dataset.rate) : 0;
    amountSpan.textContent = `Amount: ₹${dates.length * rate} (${dates.length} days × ₹${rate})`;
  }

  desigSelect.addEventListener('change', recalcMulti);

  // Load designations into this select
  (async () => {
    try {
      const res = await fetch(`${API}/api/designations`);
      const desigList = await res.json();
      desigSelect.innerHTML = '<option value="">-- Select --</option>';
      desigList.forEach(d => {
        const opt = document.createElement('option');
        opt.value = d.name;
        opt.textContent = `${d.name} (₹${d.rate}/day)`;
        opt.dataset.rate = d.rate;
        desigSelect.appendChild(opt);
      });
    } catch {}
  })();

  const addDateBtn = document.createElement('button');
  addDateBtn.type = 'button';
  addDateBtn.className = 'secondary-btn';
  addDateBtn.textContent = '+ Add Date';
  addDateBtn.onclick = () => {
    const row = document.createElement('div');
    row.className = 'date-row';
    const inp = document.createElement('input');
    inp.type = 'date';
    inp.addEventListener('change', recalcMulti);
    const rm = document.createElement('button');
    rm.textContent = '×';
    rm.onclick = () => { row.remove(); recalcMulti(); };
    row.appendChild(inp);
    row.appendChild(rm);
    datesContainer.appendChild(row);
    recalcMulti();
  };

  amountRow.appendChild(amountSpan);

  entry.appendChild(removeBtn);
  entry.appendChild(nameLabel);
  entry.appendChild(nameInput);
  entry.appendChild(desigLabel);
  entry.appendChild(desigSelect);
  entry.appendChild(datesLabel);
  entry.appendChild(datesContainer);
  entry.appendChild(addDateBtn);
  entry.appendChild(daysLabel);
  entry.appendChild(daysInput);
  entry.appendChild(amountRow);

  container.appendChild(entry);
}

function getMultiPersons() {
  const entries = document.querySelectorAll('.multi-person-entry');
  const persons = [];
  entries.forEach(entry => {
    const name = entry.querySelector('.mp-name').value.trim();
    const designation = entry.querySelector('.mp-desig').value;
    const dates = Array.from(entry.querySelectorAll('.mp-dates input[type="date"]'))
      .map(i => i.value).filter(Boolean);
    if (name && designation && dates.length > 0) {
      persons.push({ name, designation, dates });
    }
  });
  return persons;
}

async function generateMulti() {
  const errEl = document.getElementById('multiError');
  const resultsSection = document.getElementById('multiResultsSection');
  errEl.textContent = '';
  resultsSection.classList.add('hidden');

  const persons = getMultiPersons();
  if (persons.length === 0) {
    errEl.textContent = 'Please add at least one person with name, designation, and dates.';
    return;
  }

  for (const p of persons) {
    if (!p.name) { errEl.textContent = 'All persons must have a name.'; return; }
    if (!p.designation) { errEl.textContent = `"${p.name}" has no designation selected.`; return; }
    if (p.dates.length === 0) { errEl.textContent = `"${p.name}" has no dates.`; return; }
  }

  const btn = document.getElementById('generateMultiBtn');
  btn.disabled = true;
  btn.textContent = `⏳ Generating ${persons.length} bills...`;

  try {
    const res = await fetch(`${API}/api/generate-all`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ persons })
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || 'Generation failed');

    document.getElementById('multiCombinedPdfBtn').href = `${API}${data.combinedPdfUrl}`;

    const tbody = document.getElementById('multiResultsBody');
    tbody.innerHTML = '';
    data.results.forEach((r, i) => {
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td>${i + 1}</td>
        <td>${r.name}</td>
        <td>${r.designation || '-'}</td>
        <td>${r.days}</td>
        <td>₹${r.amount}</td>
        <td><a class="dl-btn dl-xlsx" href="${API}${r.xlsxUrl}" download>⬇ Excel</a></td>
        <td><a class="dl-btn dl-pdf" href="${API}${r.pdfUrl}" download>⬇ PDF</a></td>`;
      tbody.appendChild(tr);
    });

    resultsSection.classList.remove('hidden');
    resultsSection.scrollIntoView({ behavior: 'smooth' });
  } catch (err) {
    errEl.textContent = err.message;
  } finally {
    btn.disabled = false;
    btn.textContent = '⚡ Generate All Bills';
  }
}

// ── Designation table event delegation ─────────────────────────────────────────
document.getElementById('designationsBody').addEventListener('click', async (e) => {
  const tr = e.target.closest('tr');
  if (!tr) return;
  const name = tr.dataset.desigName;
  if (!name) return;
  if (e.target.classList.contains('save-btn')) {
    const input = tr.querySelector('.desig-rate-input');
    const rate = Number(input.value);
    if (!rate || rate <= 0) return;
    try {
      const res = await fetch(`${API}/api/designations/${encodeURIComponent(name)}`, {
        method: 'PUT', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ rate })
      });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error);
      await renderDesignations();
      await populateDesignationDropdowns();
    } catch (err) { showError('desigError', err.message); }
  } else if (e.target.classList.contains('del-btn')) {
    if (!confirm(`Delete "${name}"?`)) return;
    try {
      const res = await fetch(`${API}/api/designations/${encodeURIComponent(name)}`, { method: 'DELETE' });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error);
      await renderDesignations();
      await populateDesignationDropdowns();
    } catch (err) { showError('desigError', err.message); }
  }
});

// ── Init ───────────────────────────────────────────────────────────────────────
addDateRow();

// Load designations dropdown on page load
populateDesignationDropdowns();
renderDesignations();
