# Junior Supervisor Bill Generator

An exam portal bill generation system for generating payment bills for Junior Supervisors, Lab Assistants, and other exam-duty designations. Supports single entry, bulk upload via Excel, and multi-person manual entry with dynamic daily rates per designation.

## Features

- **Single Entry** — Generate one bill at a time with name, designation, and dates
- **Multi Entry** — Add multiple people manually and generate all bills at once
- **Bulk Upload** — Upload an Excel schedule with names, designations, and dates to batch-generate bills
- **Designation Management** — Add/edit/remove designations with custom daily rates
- **PDF & Excel Output** — Each bill is generated as both an Excel (.xlsx) and PDF file
- **Combined PDF** — Download all generated bills merged into a single PDF
- **One-Page Layout** — Each bill fits on one A4 page

## Tech Stack

- **Backend**: Node.js, Express, ExcelJS, LibreOffice (PDF conversion)
- **Frontend**: Vanilla HTML, CSS, JavaScript
- **PDF Merge**: Python + pypdf

## Setup

### Prerequisites

- [Node.js](https://nodejs.org/) (v16+)
- [Python](https://www.python.org/) (for PDF merging)
- [LibreOffice](https://www.libreoffice.org/download/download/) (for PDF conversion)

### Installation

```bash
# Install backend dependencies
cd backend
npm install

# Install Python PDF merge library
pip install pypdf

# Start the server
node server.js
```

Then open **http://localhost:4000** in your browser.

## Schedule Excel Format

When using the Bulk Upload tab, the Excel file should have this column layout:

| A (Sr.No) | B (Name) | C (Designation) | D onwards (Dates) |
|-----------|----------|-----------------|-------------------|
| 1 | John Doe | Junior Supervisor | 11.6(M), 12.6(E) |
| 2 | Jane Smith | Lab Assistant | 13.6(M), 14.6(E) |

Date format examples: `11.6(M)` (11 June Morning), `11.6(E)` (11 June Evening)

## Designations

Default designations and daily rates:

| Designation | Rate (₹/day) |
|-------------|-------------|
| Junior Supervisor | 200 |
| Lab Assistant | 150 |
| Senior Supervisor | 250 |

You can manage designations from the **Designations** tab in the app.
