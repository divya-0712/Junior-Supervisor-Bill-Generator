const express = require('express');
const cors = require('cors');
const path = require('path');
const fs = require('fs');
const { execFile } = require('child_process');
const ExcelJS = require('exceljs');

const app = express();
const PORT = process.env.PORT || 4000;

const TEMPLATE_PATH = path.join(__dirname, 'templates', 'Junior_Supervisor.xlsx');
const OUTPUT_DIR = path.join(__dirname, 'generated');

if (!fs.existsSync(OUTPUT_DIR)) fs.mkdirSync(OUTPUT_DIR, { recursive: true });

app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, '..', 'frontend')));

// Locate the LibreOffice executable across platforms
function findSofficeCommand() {
  if (process.env.SOFFICE_PATH && fs.existsSync(process.env.SOFFICE_PATH)) {
    return process.env.SOFFICE_PATH;
  }
  if (process.platform === 'win32') {
    const candidates = [
      'C:\\Program Files\\LibreOffice\\program\\soffice.exe',
      'C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe'
    ];
    for (const c of candidates) {
      if (fs.existsSync(c)) return c;
    }
  }
  if (process.platform === 'darwin') {
    const macPath = '/Applications/LibreOffice.app/Contents/MacOS/soffice';
    if (fs.existsSync(macPath)) return macPath;
  }
  // Fallback: assume it's on PATH (Linux, or manually added on Windows/Mac)
  return 'soffice';
}

const SOFFICE_CMD = findSofficeCommand();

// Convert an xlsx file to pdf using headless LibreOffice
function convertToPdf(xlsxPath, outDir) {
  return new Promise((resolve, reject) => {
    execFile(
      SOFFICE_CMD,
      ['--headless', '--convert-to', 'pdf', '--outdir', outDir, xlsxPath],
      { timeout: 60000 },
      (error, stdout, stderr) => {
        if (error) {
          error.message += `\n\nLibreOffice ("soffice") was not found at "${SOFFICE_CMD}". ` +
            `Install LibreOffice from https://www.libreoffice.org/download/download/, ` +
            `or set the SOFFICE_PATH environment variable to its exact soffice.exe location.`;
          return reject(error);
        }
        resolve(stdout);
      }
    );
  });
}

app.post('/api/generate', async (req, res) => {
  try {
    const { name, dates, daysWorked } = req.body;

    if (!name || !name.trim()) {
      return res.status(400).json({ error: 'Name of the Junior Supervisor is required.' });
    }
    if (!Array.isArray(dates) || dates.length === 0) {
      return res.status(400).json({ error: 'At least one date is required.' });
    }
    const days = Number(daysWorked);
    if (!Number.isFinite(days) || days <= 0) {
      return res.status(400).json({ error: 'Number of days worked must be a positive number.' });
    }

    const amount = days * 200;

    // Load the template, preserving all existing formatting
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(TEMPLATE_PATH);
    const sheet = workbook.getWorksheet('Sheet1');

    // Replace the red placeholder cells with real values
    sheet.getCell('D11').value = `Name of the Junior supervisor :  ${name.trim()}`;
    sheet.getCell('D22').value = `Dates :- ${dates.join(', ')}`;
    sheet.getCell('I21').value = days;
    sheet.getCell('I29').value = amount;

    // Force text in the sheet to black (removes the red placeholder color)
    sheet.eachRow({ includeEmpty: false }, (row) => {
      row.eachCell({ includeEmpty: false }, (cell) => {
        const existingFont = cell.font;
        if (existingFont) {
          cell.font = { ...existingFont, color: { argb: 'FF000000' } };
        }
      });
    });

    // Force the sheet to print on a single page (the template has no print
    // area / scaling configured, which is why it was splitting into 4 pages)
    sheet.pageSetup.printArea = 'A1:L69';
    sheet.pageSetup.fitToPage = true;
    sheet.pageSetup.fitToWidth = 1;
    sheet.pageSetup.fitToHeight = 1;
    sheet.pageSetup.orientation = 'portrait';
    sheet.pageSetup.margins = {
      left: 0.3, right: 0.3, top: 0.3, bottom: 0.3, header: 0.2, footer: 0.2
    };

    const id = `JS_${Date.now()}`;
    const xlsxFileName = `${id}.xlsx`;
    const xlsxPath = path.join(OUTPUT_DIR, xlsxFileName);
    await workbook.xlsx.writeFile(xlsxPath);

    // Convert the generated xlsx to pdf
    await convertToPdf(xlsxPath, OUTPUT_DIR);
    const pdfFileName = `${id}.pdf`;

    res.json({
      success: true,
      amount,
      xlsxUrl: `/api/download/${xlsxFileName}`,
      pdfUrl: `/api/download/${pdfFileName}`
    });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Failed to generate documents.', details: err.message });
  }
});

app.get('/api/download/:fileName', (req, res) => {
  const fileName = req.params.fileName;
  // basic sanitation to prevent path traversal
  if (fileName.includes('..') || fileName.includes('/')) {
    return res.status(400).send('Invalid file name.');
  }
  const filePath = path.join(OUTPUT_DIR, fileName);
  if (!fs.existsSync(filePath)) {
    return res.status(404).send('File not found.');
  }
  res.download(filePath);
});

app.listen(PORT, () => {
  console.log(`Junior Supervisor backend running on http://localhost:${PORT}`);
});
