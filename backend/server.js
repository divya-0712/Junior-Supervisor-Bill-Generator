const express = require('express');
const cors = require('cors');
const path = require('path');
const fs = require('fs');
const { execFile, spawn } = require('child_process');
const ExcelJS = require('exceljs');
const multer = require('multer');

const app = express();
const PORT = process.env.PORT || 4000;

const TEMPLATE_PATH = path.resolve(__dirname, 'templates', 'Junior_Supervisor.xlsx');
const OUTPUT_DIR = path.resolve(__dirname, 'generated');
const UPLOAD_DIR = path.resolve(__dirname, 'uploads');

[OUTPUT_DIR, UPLOAD_DIR].forEach(d => { if (!fs.existsSync(d)) fs.mkdirSync(d, { recursive: true }); });

// ── Startup checks ────────────────────────────────────────────────────────────
if (!fs.existsSync(TEMPLATE_PATH)) {
  console.error('❌  ERROR: Template file not found at:', TEMPLATE_PATH);
  console.error('    Please make sure Junior_Supervisor.xlsx is inside the backend/templates/ folder.');
  process.exit(1);
}
console.log('✅  Template found:', TEMPLATE_PATH);

const DESIGNATIONS_PATH = path.resolve(__dirname, 'designations.json');

function loadDesignations() {
  try {
    return JSON.parse(fs.readFileSync(DESIGNATIONS_PATH, 'utf8'));
  } catch { return []; }
}

function saveDesignations(data) {
  fs.writeFileSync(DESIGNATIONS_PATH, JSON.stringify(data, null, 2), 'utf8');
}

app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, '..', 'frontend')));

// ── LibreOffice path detection ────────────────────────────────────────────────
function findSofficeCommand() {
  if (process.env.SOFFICE_PATH && fs.existsSync(process.env.SOFFICE_PATH)) return process.env.SOFFICE_PATH;
  if (process.platform === 'win32') {
    const candidates = [
      'C:\\Program Files\\LibreOffice\\program\\soffice.exe',
      'C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe',
    ];
    for (const c of candidates) { if (fs.existsSync(c)) { console.log('✅  LibreOffice found:', c); return c; } }
    console.warn('⚠️   LibreOffice not found in default locations. PDF conversion may fail.');
    return 'soffice';
  }
  if (process.platform === 'darwin') {
    const m = '/Applications/LibreOffice.app/Contents/MacOS/soffice';
    if (fs.existsSync(m)) return m;
  }
  return 'soffice';
}
const SOFFICE_CMD = findSofficeCommand();

function convertToPdf(xlsxPath, outDir) {
  return new Promise((resolve, reject) => {
    console.log('  Converting to PDF:', path.basename(xlsxPath));
    execFile(
      SOFFICE_CMD,
      ['--headless', '--convert-to', 'pdf', '--outdir', outDir, xlsxPath],
      { timeout: 120000 },
      (error, stdout, stderr) => {
        if (error) {
          return reject(new Error(
            `LibreOffice PDF conversion failed.\n` +
            `Command: ${SOFFICE_CMD}\n` +
            `Error: ${error.message}\n` +
            `Stderr: ${stderr}\n\n` +
            `Make sure LibreOffice is installed from https://www.libreoffice.org/download/download/`
          ));
        }
        // LibreOffice outputs the PDF with same base name as input
        const expectedPdf = path.join(outDir, path.basename(xlsxPath, '.xlsx') + '.pdf');
        if (!fs.existsSync(expectedPdf)) {
          return reject(new Error(`PDF was not created at expected path: ${expectedPdf}\nLibreOffice stdout: ${stdout}`));
        }
        console.log('  ✅ PDF created:', path.basename(expectedPdf));
        resolve(stdout);
      }
    );
  });
}

// ── Detect Python command (python3 on Linux/Mac, python on Windows) ──────────
function findPythonCommand() {
  if (process.platform === 'win32') return 'python';
  return 'python3';
}
const PYTHON_CMD = findPythonCommand();

// ── Merge PDFs using Python / pypdf ──────────────────────────────────────────
function mergePdfs(pdfPaths, outputPath) {
  return new Promise((resolve, reject) => {
    const scriptContent = [
      'import sys',
      'try:',
      '    from pypdf import PdfWriter',
      'except ImportError:',
      '    try:',
      '        from PyPDF2 import PdfWriter',
      '    except ImportError:',
      '        print("ERROR: pypdf not installed. Run: pip install pypdf", file=sys.stderr)',
      '        sys.exit(1)',
      'w = PdfWriter()',
      'for p in sys.argv[1:-1]:',
      '    w.append(p)',
      'with open(sys.argv[-1], "wb") as f:',
      '    w.write(f)',
      'print("ok")',
    ].join('\n');

    const tmpScript = path.join(__dirname, '_merge_pdfs_tmp.py');
    fs.writeFileSync(tmpScript, scriptContent);

    const args = [tmpScript, ...pdfPaths, outputPath];
    console.log('  Merging', pdfPaths.length, 'PDFs →', path.basename(outputPath));

    const proc = spawn(PYTHON_CMD, args);
    let out = '', err = '';
    proc.stdout.on('data', d => out += d);
    proc.stderr.on('data', d => err += d);
    proc.on('error', spawnErr => {
      try { fs.unlinkSync(tmpScript); } catch {}
      reject(new Error(
        `Could not start Python ("${PYTHON_CMD}"). Make sure Python is installed.\n` +
        `On Windows: install from https://www.python.org/downloads/ and check "Add to PATH".\n` +
        `Error: ${spawnErr.message}`
      ));
    });
    proc.on('close', code => {
      try { fs.unlinkSync(tmpScript); } catch {}
      if (code !== 0) {
        return reject(new Error(
          `PDF merge failed (exit ${code}).\n${err}\n` +
          `Fix: run "pip install pypdf" in your terminal.`
        ));
      }
      console.log('  ✅ Combined PDF created:', path.basename(outputPath));
      resolve(out);
    });
  });
}

// ── Normalise date string: "11.6 E" → "11.6(E)", "13.6M" → "13.6(M)" ───────
function normaliseDate(raw) {
  const s = String(raw || '').trim();
  const m = s.match(/^([\d.]+)\s*([MEme])$/);
  if (m) return `${m[1]}(${m[2].toUpperCase()})`;
  return s;
}

// ── Parse uploaded schedule Excel ────────────────────────────────────────────
// Expected columns: A=Sr.No, B=Name, C=Designation, D onwards=Dates
async function parseSchedule(filePath) {
  const wb2 = new ExcelJS.Workbook();
  await wb2.xlsx.readFile(filePath);
  const ws = wb2.getWorksheet(1);
  const persons = [];
  const designations = loadDesignations();

  ws.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return; // skip header row
    const nameCell = row.getCell(2).value; // column B = Name
    if (!nameCell || String(nameCell).trim() === '') return;
    const name = String(nameCell).trim();
    // column C = Designation
    const desigCell = row.getCell(3).value;
    const designation = desigCell ? String(desigCell).trim() : '';
    const found = designations.find(d => d.name.toLowerCase() === designation.toLowerCase());
    const rate = found ? found.rate : 0;
    const dates = [];
    row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
      if (colNumber < 4) return; // skip A (1), B (2), C (3)
      const val = cell.value;
      if (val !== null && val !== undefined && String(val).trim() !== '') {
        dates.push(normaliseDate(String(val)));
      }
    });
    if (dates.length > 0 && rate > 0) {
      persons.push({ name, designation, rate, dates });
    }
  });

  return persons;
}

// ── Generate one bill (xlsx + pdf) using the official template ────────────────
async function generateBill(name, designation, dates, rate, id) {
  const days = dates.length;
  const amount = days * rate;

  console.log(`  Generating bill for: ${name} (${designation}) | ${days} days × ₹${rate} = ₹${amount}`);

  // Always load a fresh copy of the template for every bill
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(TEMPLATE_PATH);

  const sheet = workbook.getWorksheet('Sheet1');
  if (!sheet) throw new Error('Sheet "Sheet1" not found in template. Check the template file.');

  // Fit content to one A4 page
  sheet.pageSetup.fitToPage = true;
  sheet.pageSetup.fitToWidth = 1;
  sheet.pageSetup.fitToHeight = 1;
  sheet.pageSetup.orientation = 'portrait';
  sheet.pageSetup.paperSize = 9; // A4
  sheet.pageSetup.margins = {
    top: 0.4, bottom: 0.4, left: 0.3, right: 0.3, header: 0, footer: 0
  };

  const blackFont = { color: { argb: 'FF000000' } };

  const d11 = sheet.getCell('D11');
  d11.value = `Name of the ${designation} :  ${name}`;
  d11.font = { ...d11.font, ...blackFont };

  const d22 = sheet.getCell('D22');
  d22.value = `Dates :- ${dates.join(', ')}`;
  d22.font = { ...d22.font, ...blackFont };

  const i21 = sheet.getCell('I21');
  i21.value = days;
  i21.font = { ...i21.font, ...blackFont };

  const i22 = sheet.getCell('I22');
  i22.value = `X ${rate}`;
  i22.font = { ...i22.font, ...blackFont };

  const i29 = sheet.getCell('I29');
  i29.value = amount;
  i29.font = { ...i29.font, ...blackFont };

  const xlsxPath = path.resolve(OUTPUT_DIR, `${id}.xlsx`);
  await workbook.xlsx.writeFile(xlsxPath);
  console.log('  ✅ Excel saved:', path.basename(xlsxPath));

  await convertToPdf(xlsxPath, OUTPUT_DIR);

  const pdfPath = path.resolve(OUTPUT_DIR, `${id}.pdf`);
  return { xlsxPath, pdfPath, days, amount, designation };
}

// ── Multer setup ─────────────────────────────────────────────────────────────
const upload = multer({
  dest: UPLOAD_DIR,
  fileFilter: (req, file, cb) => {
    const ok = file.mimetype.includes('spreadsheet') ||
                file.mimetype.includes('excel') ||
                file.originalname.endsWith('.xlsx');
    ok ? cb(null, true) : cb(new Error('Only .xlsx files are allowed'));
  }
});

// ── Route: Parse schedule Excel ───────────────────────────────────────────────
app.post('/api/parse-schedule', upload.single('schedule'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'No file uploaded.' });
    console.log('Parsing schedule:', req.file.originalname);
    const persons = await parseSchedule(req.file.path);
    try { fs.unlinkSync(req.file.path); } catch {}
    console.log(`✅  Parsed ${persons.length} persons`);
    res.json({ success: true, persons });
  } catch (err) {
    console.error('Parse error:', err.message);
    res.status(500).json({ error: 'Failed to parse schedule.', details: err.message });
  }
});

// ── Route: Generate all bills ────────────────────────────────────────────────
app.post('/api/generate-all', async (req, res) => {
  try {
    const { persons } = req.body;
    if (!Array.isArray(persons) || persons.length === 0)
      return res.status(400).json({ error: 'No persons provided.' });

    const designations = loadDesignations();

    console.log(`\n=== Generating ${persons.length} bills ===`);
    const batchId = `BATCH_${Date.now()}`;
    const results = [];
    const pdfPaths = [];

    for (let i = 0; i < persons.length; i++) {
      const { name, dates, designation } = persons[i];
      // Look up rate from designations list
      const found = designations.find(d => d.name.toLowerCase() === (designation || '').toLowerCase());
      const rate = found ? found.rate : 200;
      const desigName = found ? found.name : (designation || 'Junior Supervisor');

      const safeName = name.replace(/[^a-zA-Z0-9]/g, '_');
      const id = `${batchId}_${String(i + 1).padStart(3, '0')}_${safeName}`;
      console.log(`[${i + 1}/${persons.length}] ${name} (${desigName})`);
      const bill = await generateBill(name, desigName, dates, rate, id);
      results.push({
        name,
        designation: desigName,
        dates,
        days: bill.days,
        amount: bill.amount,
        xlsxUrl: `/api/download/${path.basename(bill.xlsxPath)}`,
        pdfUrl: `/api/download/${path.basename(bill.pdfPath)}`
      });
      pdfPaths.push(bill.pdfPath);
    }

    // Merge all individual PDFs into one combined PDF
    const combinedPdfName = `${batchId}_COMBINED_ALL.pdf`;
    const combinedPdfPath = path.resolve(OUTPUT_DIR, combinedPdfName);
    await mergePdfs(pdfPaths, combinedPdfPath);

    console.log(`\n✅  All ${persons.length} bills generated successfully!\n`);
    res.json({ success: true, results, combinedPdfUrl: `/api/download/${combinedPdfName}` });
  } catch (err) {
    console.error('Generate-all error:', err.message);
    res.status(500).json({ error: 'Failed to generate bills.', details: err.message });
  }
});

// ── Route: Manual single bill ─────────────────────────────────────────────────
app.post('/api/generate', async (req, res) => {
  try {
    const { name, dates, daysWorked, designation } = req.body;
    if (!name || !name.trim()) return res.status(400).json({ error: 'Name is required.' });
    if (!Array.isArray(dates) || dates.length === 0) return res.status(400).json({ error: 'At least one date is required.' });
    const days = Number(daysWorked);
    if (!Number.isFinite(days) || days <= 0) return res.status(400).json({ error: 'Days worked must be a positive number.' });

    const designations = loadDesignations();
    const found = designations.find(d => d.name.toLowerCase() === (designation || '').toLowerCase());
    const rate = found ? found.rate : 200;
    const desigName = found ? found.name : (designation || 'Junior Supervisor');

    const id = `JS_${Date.now()}`;
    const bill = await generateBill(name.trim(), desigName, dates, rate, id);
    res.json({
      success: true,
      amount: bill.amount,
      designation: desigName,
      rate,
      xlsxUrl: `/api/download/${path.basename(bill.xlsxPath)}`,
      pdfUrl: `/api/download/${path.basename(bill.pdfPath)}`
    });
  } catch (err) {
    console.error('Generate error:', err.message);
    res.status(500).json({ error: 'Failed to generate documents.', details: err.message });
  }
});

// ── Designation CRUD ──────────────────────────────────────────────────────────
app.get('/api/designations', (req, res) => {
  res.json(loadDesignations());
});

app.post('/api/designations', (req, res) => {
  try {
    const { name, rate } = req.body;
    if (!name || !name.trim()) return res.status(400).json({ error: 'Designation name is required.' });
    const r = Number(rate);
    if (!Number.isFinite(r) || r <= 0) return res.status(400).json({ error: 'Rate must be a positive number.' });

    const list = loadDesignations();
    if (list.some(d => d.name.toLowerCase() === name.trim().toLowerCase()))
      return res.status(400).json({ error: 'Designation already exists.' });

    list.push({ name: name.trim(), rate: r });
    saveDesignations(list);
    res.json({ success: true, designations: list });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.put('/api/designations/:name', (req, res) => {
  try {
    const { rate } = req.body;
    const r = Number(rate);
    if (!Number.isFinite(r) || r <= 0) return res.status(400).json({ error: 'Rate must be a positive number.' });

    const list = loadDesignations();
    const idx = list.findIndex(d => d.name.toLowerCase() === decodeURIComponent(req.params.name).toLowerCase());
    if (idx === -1) return res.status(404).json({ error: 'Designation not found.' });

    list[idx].rate = r;
    saveDesignations(list);
    res.json({ success: true, designations: list });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.delete('/api/designations/:name', (req, res) => {
  try {
    const list = loadDesignations();
    const filtered = list.filter(d => d.name.toLowerCase() !== decodeURIComponent(req.params.name).toLowerCase());
    if (filtered.length === list.length) return res.status(404).json({ error: 'Designation not found.' });
    saveDesignations(filtered);
    res.json({ success: true, designations: filtered });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ── Route: Download file ──────────────────────────────────────────────────────
app.get('/api/download/:fileName', (req, res) => {
  const fileName = path.basename(req.params.fileName); // strip any path traversal
  const filePath = path.resolve(OUTPUT_DIR, fileName);
  if (!fs.existsSync(filePath)) return res.status(404).json({ error: 'File not found. It may have been cleaned up.' });
  res.download(filePath);
});

app.listen(PORT, () => {
  console.log(`\n🚀  Junior Supervisor Bill Generator`);
  console.log(`    Running at: http://localhost:${PORT}`);
  console.log(`    Open this URL in your browser.\n`);
});
