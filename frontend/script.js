const datesContainer = document.getElementById('datesContainer');
const addDateBtn = document.getElementById('addDateBtn');
const daysInput = document.getElementById('days');
const amountDisplay = document.getElementById('amountDisplay');
const generateBtn = document.getElementById('generateBtn');
const errorMsg = document.getElementById('errorMsg');
const downloadLinks = document.getElementById('downloadLinks');
const xlsxLink = document.getElementById('xlsxLink');
const pdfLink = document.getElementById('pdfLink');

const API_BASE = ''; // same origin; change if backend runs elsewhere, e.g. 'http://localhost:4000'

function addDateRow(value = '') {
  const row = document.createElement('div');
  row.className = 'date-row';

  const input = document.createElement('input');
  input.type = 'date';
  input.value = value;
  input.addEventListener('change', recalc);

  const removeBtn = document.createElement('button');
  removeBtn.textContent = '×';
  removeBtn.addEventListener('click', () => {
    row.remove();
    recalc();
  });

  row.appendChild(input);
  row.appendChild(removeBtn);
  datesContainer.appendChild(row);
  recalc();
}

function getDates() {
  return Array.from(datesContainer.querySelectorAll('input[type="date"]'))
    .map(i => i.value)
    .filter(Boolean);
}

function recalc() {
  const dates = getDates();
  daysInput.value = dates.length;
  const amount = dates.length * 200;
  amountDisplay.textContent = amount;
}

addDateBtn.addEventListener('click', () => addDateRow());

generateBtn.addEventListener('click', async () => {
  errorMsg.textContent = '';
  downloadLinks.classList.add('hidden');

  const name = document.getElementById('name').value.trim();
  const dates = getDates();
  const days = Number(daysInput.value);

  if (!name) {
    errorMsg.textContent = 'Please enter the name of the Junior Supervisor.';
    return;
  }
  if (dates.length === 0) {
    errorMsg.textContent = 'Please add at least one date.';
    return;
  }

  generateBtn.disabled = true;
  generateBtn.textContent = 'Generating...';

  try {
    const res = await fetch(`${API_BASE}/api/generate`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ name, dates, daysWorked: days })
    });
    const data = await res.json();

    if (!res.ok) {
      throw new Error(data.error || 'Something went wrong.');
    }

    amountDisplay.textContent = data.amount;
    xlsxLink.href = `${API_BASE}${data.xlsxUrl}`;
    pdfLink.href = `${API_BASE}${data.pdfUrl}`;
    downloadLinks.classList.remove('hidden');
  } catch (err) {
    errorMsg.textContent = err.message;
  } finally {
    generateBtn.disabled = false;
    generateBtn.textContent = 'Generate Excel & PDF';
  }
});

// Start with one date row
addDateRow();
