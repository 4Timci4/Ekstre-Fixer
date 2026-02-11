const { ipcRenderer } = require('electron');
const XLSX = require('xlsx');
const ExcelJS = require('exceljs');
const dayjs = require('dayjs');
const path = require('path');
const fs = require('fs');

// --- State ---
let selectedFiles = [];
let activeTab = 'genel';
let selectedIndices = new Set();
let sourceDir = localStorage.getItem('sourceDir') || '';
let outputDir = localStorage.getItem('outputDir') || '';

// --- Constants ---
const COLS = {
  FIRMA_ADI: 'Firma Adı',
  BORCLU_ADI: 'Borçlu Adı',
  ISLEM_TUTARI: 'İşlem Tutarı',
  ISLEM_TARIHI: 'İşlemTarihi',
  ISLEM_TURU: 'İşlem Türü',
  DISPUTE: 'Dispute',
  BAKIYE: 'Bakiye',
  FATURA_TARIHI: 'FaturaTarihi',
  FATURA_VADESI: 'Fatura Vadesi',
  FATURA_NO: 'Fatura No',
  COLUMNS_TO_DROP: ['Firma No', 'Firma Adı', 'Ekno', 'Döviz Cinsi', 'Borçlu No', 'Borçlu Adı', 'Muhabir No'],
  COLLECTION_TYPES: ['Tahsilat', 'GeriDevir', 'Senet', 'Çek'],
  UNMATCHED_TYPES: ['EslenmemisTahsilat', 'EslenmemisCek', 'EslenmemisSenet'],
  NEW_ISLEM_TUTARI_PLUS: 'İşlem Tutarı (+)',
  NEW_ISLEM_TUTARI_MINUS: 'İşlem Tutarı (-)'
};

const CONFIG = {
  HEADER_ROW: 4,
  DATE_FORMAT: 'DD/MM/YYYY'
};

// --- DOM Elements ---
const selectBtn = document.getElementById('selectBtn');
const dropZone = document.getElementById('dropZone');
const fileList = document.getElementById('fileList');
const deleteBtn = document.getElementById('deleteBtn');
const clearBtn = document.getElementById('clearBtn');
const processBtn = document.getElementById('processBtn');
const progressFill = document.getElementById('progressFill');
const statusEl = document.getElementById('status');
const mergeCheckbox = document.getElementById('mergeFiles');
const startDateInput = document.getElementById('startDate');
const tabs = document.querySelectorAll('.tab-btn');
const tabContents = document.querySelectorAll('.tab-content');
const sourcePathEl = document.getElementById('sourcePath');
const outputPathEl = document.getElementById('outputPath');
const sourceBtn = document.getElementById('sourceBtn');
const outputBtn = document.getElementById('outputBtn');

// --- Folder Path UI ---
function updateFolderUI() {
  if (sourceDir) {
    sourcePathEl.textContent = sourceDir;
    sourcePathEl.classList.remove('empty');
  } else {
    sourcePathEl.textContent = 'Seçilmedi';
    sourcePathEl.classList.add('empty');
  }
  if (outputDir) {
    outputPathEl.textContent = outputDir;
    outputPathEl.classList.remove('empty');
  } else {
    outputPathEl.textContent = 'Seçilmedi';
    outputPathEl.classList.add('empty');
  }
}
updateFolderUI();

sourceBtn.addEventListener('click', async () => {
  const dir = await ipcRenderer.invoke('select-folder', { title: 'Kaynak Klasörü Seçin', defaultPath: sourceDir });
  if (dir) {
    sourceDir = dir;
    localStorage.setItem('sourceDir', dir);
    updateFolderUI();
  }
});

outputBtn.addEventListener('click', async () => {
  const dir = await ipcRenderer.invoke('select-folder', { title: 'Çıktı Klasörü Seçin', defaultPath: outputDir });
  if (dir) {
    outputDir = dir;
    localStorage.setItem('outputDir', dir);
    updateFolderUI();
  }
});

// --- Event Listeners ---
selectBtn.addEventListener('click', async () => {
  const files = await ipcRenderer.invoke('select-files', sourceDir || undefined);
  if (files.length) addFiles(files);
});

dropZone.addEventListener('dragover', (e) => {
  e.preventDefault();
  dropZone.classList.add('dragover');
});

dropZone.addEventListener('dragleave', () => {
  dropZone.classList.remove('dragover');
});

dropZone.addEventListener('drop', (e) => {
  e.preventDefault();
  dropZone.classList.remove('dragover');
  const files = Array.from(e.dataTransfer.files)
    .filter(f => f.name.match(/\.xlsx?$/i))
    .map(f => f.path);
  if (files.length) addFiles(files);
});

deleteBtn.addEventListener('click', deleteSelected);
clearBtn.addEventListener('click', clearAll);
processBtn.addEventListener('click', processAndSave);

tabs.forEach(tab => {
  tab.addEventListener('click', () => {
    // Update tab styles
    tabs.forEach(t => {
      t.classList.remove('active');
      t.setAttribute('aria-selected', 'false');
    });
    tab.classList.add('active');
    tab.setAttribute('aria-selected', 'true');

    // Update content visibility
    tabContents.forEach(c => c.classList.add('hidden'));
    activeTab = tab.dataset.tab;
    document.getElementById(activeTab).classList.remove('hidden');
  });
});

// --- File Management ---
function addFiles(newFiles) {
  const existing = new Set(selectedFiles);
  newFiles.forEach(f => existing.add(f));
  selectedFiles = Array.from(existing).sort();
  updateFileList();
}

function updateFileList() {
  fileList.innerHTML = '';
  selectedIndices.clear();
  if (selectedFiles.length === 0) {
    fileList.innerHTML = `
      <li class="placeholder">
        <svg fill="none" stroke="currentColor" viewBox="0 0 24 24" aria-hidden="true">
          <path stroke-linecap="round" stroke-linejoin="round" stroke-width="1.5" d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12"/>
        </svg>
        <span class="ph-text">Dosyaları sürükleyin</span>
        <span class="ph-sub">.xlsx veya .xls formatında</span>
      </li>
    `;
    deleteBtn.disabled = true;
    clearBtn.disabled = true;
    processBtn.disabled = true;
  } else {
    selectedFiles.forEach((f, i) => {
      const li = document.createElement('li');
      li.className = 'file-item';
      li.addEventListener('click', (e) => {
        if (e.target.closest('.remove-btn')) return;
        if (selectedIndices.has(i)) {
          selectedIndices.delete(i);
          li.classList.remove('selected');
        } else {
          selectedIndices.add(i);
          li.classList.add('selected');
        }
      });
      li.innerHTML = `
        <div class="file-info">
          <svg fill="none" stroke="currentColor" viewBox="0 0 24 24" aria-hidden="true">
            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"/>
          </svg>
          <span class="file-name">${path.basename(f)}</span>
        </div>
        <button class="remove-btn" data-index="${i}" aria-label="Dosyayı kaldır">
          <svg fill="none" stroke="currentColor" viewBox="0 0 24 24" aria-hidden="true">
            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M6 18L18 6M6 6l12 12"/>
          </svg>
        </button>
      `;
      li.querySelector('.remove-btn').addEventListener('click', () => {
        selectedFiles.splice(i, 1);
        updateFileList();
      });
      fileList.appendChild(li);
    });
    deleteBtn.disabled = false;
    clearBtn.disabled = false;
    processBtn.disabled = false;
  }
}

function deleteSelected() {
  if (selectedIndices.size === 0) return;
  // Remove from end to start to preserve indices
  const indices = Array.from(selectedIndices).sort((a, b) => b - a);
  indices.forEach(idx => {
    if (idx >= 0 && idx < selectedFiles.length) {
      selectedFiles.splice(idx, 1);
    }
  });
  selectedIndices.clear();
  updateFileList();
}

function clearAll() {
  selectedFiles = [];
  updateFileList();
}

// --- Number Normalization ---
function normalizeNumber(val) {
  if (val === null || val === undefined || val === '') return 0;
  if (typeof val === 'number') return Math.round(val * 100) / 100;

  const str = String(val).trim();
  if (str === '' || str === 'nan' || str === 'None') return 0;

  // Turkish format: 1.234,56 or 88,49
  if (str.includes(',')) {
    const parts = str.split(',');
    if (parts.length === 2 && parts[1].replace(/\./g, '').length <= 2) {
      const normalized = str.replace(/\./g, '').replace(',', '.');
      const num = parseFloat(normalized);
      return isNaN(num) ? 0 : Math.round(num * 100) / 100;
    }
  }

  const num = parseFloat(str);
  return isNaN(num) ? 0 : Math.round(num * 100) / 100;
}

// --- Date Handling ---
function parseExcelDate(val) {
  if (!val) return null;
  if (typeof val === 'number') {
    // Excel serial date
    const date = XLSX.SSF.parse_date_code(val);
    if (date) return new Date(date.y, date.m - 1, date.d);
  }
  const d = dayjs(val);
  return d.isValid() ? d.toDate() : null;
}

function formatDate(date) {
  if (!date) return '';
  return dayjs(date).format(CONFIG.DATE_FORMAT);
}

// --- Excel Processing ---
function readExcelFile(filePath) {
  try {
    const workbook = XLSX.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    const data = XLSX.utils.sheet_to_json(sheet, { range: CONFIG.HEADER_ROW, defval: '' });

    if (data.length === 0) {
      throw new Error(`Dosya boş veya geçersiz: ${path.basename(filePath)}`);
    }

    return data.map(row => {
      const newRow = {};
      Object.keys(row).forEach(key => {
        newRow[key.trim()] = row[key];
      });
      return newRow;
    }).filter(row => Object.values(row).some(v => v !== '' && v !== null && v !== undefined));
  } catch (err) {
    if (err.code === 'ENOENT') {
      throw new Error(`Dosya bulunamadı: ${path.basename(filePath)}`);
    }
    if (err.code === 'EACCES' || err.code === 'EPERM') {
      throw new Error(`Dosya erişim izni yok: ${path.basename(filePath)}`);
    }
    throw err;
  }
}

function extractMetadata(data) {
  if (!data.length) return { firmaAdi: '', borcluAdi: '' };
  return {
    firmaAdi: String(data[0][COLS.FIRMA_ADI] || ''),
    borcluAdi: String(data[0][COLS.BORCLU_ADI] || '')
  };
}

function performConfirmation(data) {
  let confirmationMsg = '';
  if (data.length >= 2 && data[0].hasOwnProperty(COLS.ISLEM_TUTARI) && data[0].hasOwnProperty(COLS.BAKIYE)) {
    const lastRow = data[data.length - 1];
    const secondLastRow = data[data.length - 2];
    const islemTutariLast = normalizeNumber(lastRow[COLS.ISLEM_TUTARI]);
    const bakiyeSecondLast = normalizeNumber(secondLastRow[COLS.BAKIYE]);

    if (islemTutariLast === bakiyeSecondLast) {
      data.pop();
      confirmationMsg = 'Teyit Başarılı';
    } else {
      confirmationMsg = `Alt toplam uyuşmuyor (Son İşlem: ${islemTutariLast}, Önceki Bakiye: ${bakiyeSecondLast})`;
    }
  }
  return { data, confirmationMsg };
}

function checkUnmatchedBalance(data) {
  const hasUnmatched = data.some(row => COLS.UNMATCHED_TYPES.includes(row[COLS.ISLEM_TURU]));
  return hasUnmatched ? 'Ekstrede eşlenmemiş bakiyeler var.' : '';
}

function prepareDataframe(data) {
  return data.map(row => {
    const newRow = { ...row };
    COLS.COLUMNS_TO_DROP.forEach(col => delete newRow[col]);
    if (newRow[COLS.DISPUTE] === undefined) newRow[COLS.DISPUTE] = 0;
    return newRow;
  });
}

function normalizeAndSplitAmounts(data) {
  return data.map(row => {
    const islemTutari = normalizeNumber(row[COLS.ISLEM_TUTARI]);
    row[COLS.BAKIYE] = normalizeNumber(row[COLS.BAKIYE]);
    row[COLS.ISLEM_TARIHI] = parseExcelDate(row[COLS.ISLEM_TARIHI]);

    if (islemTutari < 0) {
      row[COLS.DISPUTE] = islemTutari;
      row[COLS.ISLEM_TUTARI] = 0;
    } else {
      row[COLS.DISPUTE] = 0;
      row[COLS.ISLEM_TUTARI] = islemTutari;
    }
    return row;
  });
}

function aggregateCollections(data) {
  const collections = data.filter(row => COLS.COLLECTION_TYPES.includes(row[COLS.ISLEM_TURU]));
  const others = data.filter(row => !COLS.COLLECTION_TYPES.includes(row[COLS.ISLEM_TURU]));

  if (collections.length === 0) return data;

  // Group by date
  const grouped = {};
  collections.forEach(row => {
    const dateKey = row[COLS.ISLEM_TARIHI] ? row[COLS.ISLEM_TARIHI].getTime() : 'null';
    if (!grouped[dateKey]) {
      grouped[dateKey] = {
        [COLS.ISLEM_TARIHI]: row[COLS.ISLEM_TARIHI],
        [COLS.ISLEM_TUTARI]: 0,
        [COLS.DISPUTE]: 0,
        [COLS.ISLEM_TURU]: row[COLS.ISLEM_TURU],
        [COLS.BAKIYE]: row[COLS.BAKIYE],
        [COLS.FATURA_TARIHI]: null,
        [COLS.FATURA_VADESI]: null,
        [COLS.FATURA_NO]: ''
      };
    }
    grouped[dateKey][COLS.ISLEM_TUTARI] += row[COLS.ISLEM_TUTARI];
    grouped[dateKey][COLS.DISPUTE] += row[COLS.DISPUTE];
    grouped[dateKey][COLS.BAKIYE] = row[COLS.BAKIYE];
    grouped[dateKey][COLS.ISLEM_TURU] = row[COLS.ISLEM_TURU];
  });

  const aggregated = Object.values(grouped);
  const result = [...others, ...aggregated];

  // Sort by date
  result.sort((a, b) => {
    const dateA = a[COLS.ISLEM_TARIHI] ? a[COLS.ISLEM_TARIHI].getTime() : 0;
    const dateB = b[COLS.ISLEM_TARIHI] ? b[COLS.ISLEM_TARIHI].getTime() : 0;
    return dateA - dateB;
  });

  return result;
}

function formatFinalData(data) {
  return data.map(row => {
    // Round numbers
    [COLS.ISLEM_TUTARI, COLS.DISPUTE, COLS.BAKIYE].forEach(col => {
      if (row[col] !== undefined && row[col] !== '' && row[col] !== null) {
        row[col] = Math.round(parseFloat(row[col]) * 100) / 100;
      }
    });

    // Parse dates
    [COLS.FATURA_TARIHI, COLS.FATURA_VADESI].forEach(col => {
      row[col] = parseExcelDate(row[col]);
    });

    // Replace 0 with empty string
    if (row[COLS.ISLEM_TUTARI] === 0) row[COLS.ISLEM_TUTARI] = '';
    if (row[COLS.DISPUTE] === 0) row[COLS.DISPUTE] = '';

    return row;
  });
}

function processGeneralEkstre(data) {
  const { firmaAdi, borcluAdi } = extractMetadata(data);
  data = prepareDataframe(data);
  data = normalizeAndSplitAmounts(data);
  data = aggregateCollections(data);
  data = formatFinalData(data);
  return { data, firmaAdi, borcluAdi };
}

function processDatedEkstre(data, startDate) {
  const { firmaAdi, borcluAdi } = extractMetadata(data);

  // Parse dates and normalize
  data.forEach(row => {
    row[COLS.ISLEM_TARIHI] = parseExcelDate(row[COLS.ISLEM_TARIHI]);
    row[COLS.FATURA_TARIHI] = parseExcelDate(row[COLS.FATURA_TARIHI]);
    row[COLS.BAKIYE] = normalizeNumber(row[COLS.BAKIYE]);
    row[COLS.ISLEM_TUTARI] = normalizeNumber(row[COLS.ISLEM_TUTARI]);
  });

  const parts = startDate.split('-');
  const startDateObj = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));

  // Calculate devir from previous transactions
  const previousTransactions = data.filter(row =>
    row[COLS.ISLEM_TARIHI] && row[COLS.ISLEM_TARIHI] < startDateObj
  );

  let toplamDevir = 0;
  if (previousTransactions.length > 0) {
    previousTransactions.sort((a, b) => a[COLS.ISLEM_TARIHI] - b[COLS.ISLEM_TARIHI]);
    toplamDevir = previousTransactions[previousTransactions.length - 1][COLS.BAKIYE] || 0;
  }

  // Special invoices from previous year
  const startYear = startDateObj.getFullYear();
  const specialFaturalar = data.filter(row => {
    if (!row[COLS.ISLEM_TARIHI] || !row[COLS.FATURA_TARIHI]) return false;
    return row[COLS.ISLEM_TARIHI].getFullYear() === startYear &&
           row[COLS.FATURA_TARIHI].getFullYear() === startYear - 1 &&
           row[COLS.ISLEM_TURU] === 'FaturaGiris';
  });

  specialFaturalar.forEach(row => {
    toplamDevir += row[COLS.ISLEM_TUTARI];
  });

  // Filter data - remove special faturalar using object reference identity
  const specialSet = new Set(specialFaturalar);
  data = data.filter(row => !specialSet.has(row));
  data = data.filter(row => row[COLS.ISLEM_TARIHI] && row[COLS.ISLEM_TARIHI] >= startDateObj);

  // Add devir row
  const devirRow = {
    [COLS.FATURA_TARIHI]: null,
    [COLS.FATURA_VADESI]: null,
    [COLS.FATURA_NO]: '',
    [COLS.ISLEM_TURU]: 'Devir',
    [COLS.ISLEM_TARIHI]: startDateObj,
    [COLS.ISLEM_TUTARI]: toplamDevir,
    [COLS.DISPUTE]: '',
    [COLS.BAKIYE]: toplamDevir
  };

  data = [devirRow, ...data];

  // Process as general
  data = prepareDataframe(data);
  data = normalizeAndSplitAmounts(data);
  data = aggregateCollections(data);
  data = formatFinalData(data);

  return { data, firmaAdi, borcluAdi };
}

function processSingleFile(filePath, mode, startDate) {
  let data = readExcelFile(filePath);

  const { data: confirmedData, confirmationMsg } = performConfirmation(data);
  data = confirmedData;

  const unmatchedWarning = checkUnmatchedBalance(data);

  const warnings = [];
  if (unmatchedWarning) warnings.push(unmatchedWarning);
  if (confirmationMsg && !confirmationMsg.includes('Teyit Başarılı')) {
    warnings.push(confirmationMsg);
  }

  let result;
  if (mode === 'tarihli') {
    if (!startDate) throw new Error('Tarihli ekstre için başlangıç tarihi belirtilmelidir.');
    result = processDatedEkstre(data, startDate);
  } else {
    result = processGeneralEkstre(data);
  }

  // Rename columns
  result.data = result.data.map(row => {
    const newRow = { ...row };
    newRow[COLS.NEW_ISLEM_TUTARI_PLUS] = newRow[COLS.ISLEM_TUTARI];
    newRow[COLS.NEW_ISLEM_TUTARI_MINUS] = newRow[COLS.DISPUTE];
    delete newRow[COLS.ISLEM_TUTARI];
    delete newRow[COLS.DISPUTE];
    return newRow;
  });

  return { ...result, warnings };
}

// --- Excel Writing ---
async function writeExcelFile(filePath, results, merge = false) {
  const workbook = new ExcelJS.Workbook();

  if (merge) {
    const sheetNames = {};
    for (const result of results) {
      let baseName = (result.firmaAdi || 'Bilinmeyen').substring(0, 25);
      let sheetName = baseName;
      if (sheetNames[sheetName]) {
        sheetNames[sheetName]++;
        sheetName = `${baseName}_${sheetNames[sheetName]}`;
      } else {
        sheetNames[sheetName] = 1;
      }
      addSheetToWorkbook(workbook, result, sheetName);
    }
  } else {
    addSheetToWorkbook(workbook, results[0], 'Sayfa1');
  }

  await workbook.xlsx.writeFile(filePath);
}

function addSheetToWorkbook(workbook, result, sheetName) {
  const columns = [
    COLS.FATURA_TARIHI, COLS.FATURA_VADESI, COLS.FATURA_NO,
    COLS.ISLEM_TURU, COLS.ISLEM_TARIHI,
    COLS.NEW_ISLEM_TUTARI_PLUS, COLS.NEW_ISLEM_TUTARI_MINUS, COLS.BAKIYE
  ];

  const columnWidths = {
    [COLS.FATURA_TARIHI]: 13,
    [COLS.FATURA_VADESI]: 13,
    [COLS.FATURA_NO]: 17,
    [COLS.ISLEM_TURU]: 10,
    [COLS.ISLEM_TARIHI]: 13,
    [COLS.NEW_ISLEM_TUTARI_PLUS]: 15,
    [COLS.NEW_ISLEM_TUTARI_MINUS]: 15,
    [COLS.BAKIYE]: 12
  };

  const dateColumns = new Set([COLS.FATURA_TARIHI, COLS.FATURA_VADESI, COLS.ISLEM_TARIHI]);
  const numberColumns = new Set([COLS.NEW_ISLEM_TUTARI_PLUS, COLS.NEW_ISLEM_TUTARI_MINUS, COLS.BAKIYE]);

  const centerAlignment = { horizontal: 'center', vertical: 'middle' };
  const numberFormat = '#,##0.00';
  const dateFormat = 'DD.MM.YYYY';

  const sheet = workbook.addWorksheet(sheetName);

  // Set column widths
  columns.forEach((col, i) => {
    sheet.getColumn(i + 1).width = columnWidths[col] || 12;
  });

  // Row 1: Firma Adı
  sheet.getCell('A1').value = result.firmaAdi;
  // Row 2: Borçlu Adı
  sheet.getCell('A2').value = result.borcluAdi;
  // Row 3-4: Empty

  // Row 5: Headers
  const headerRow = sheet.getRow(5);
  columns.forEach((col, i) => {
    const cell = headerRow.getCell(i + 1);
    cell.value = col;
    cell.font = { bold: true };
    cell.alignment = centerAlignment;
    cell.border = {};
  });

  // Data rows starting at row 6
  result.data.forEach((row, rowIndex) => {
    const excelRow = sheet.getRow(rowIndex + 6);
    columns.forEach((col, colIndex) => {
      const cell = excelRow.getCell(colIndex + 1);
      const val = row[col];

      if (dateColumns.has(col)) {
        if (val instanceof Date) {
          // Write as UTC noon to prevent timezone offset shifting the date
          cell.value = new Date(Date.UTC(val.getFullYear(), val.getMonth(), val.getDate(), 12, 0, 0));
          cell.numFmt = dateFormat;
        } else if (val) {
          cell.value = val;
        } else {
          cell.value = '';
        }
      } else if (numberColumns.has(col)) {
        if (val !== '' && val !== undefined && val !== null) {
          cell.value = typeof val === 'number' ? val : parseFloat(val) || 0;
          cell.numFmt = numberFormat;
        }
      } else {
        cell.value = val !== undefined ? val : '';
      }
      cell.alignment = centerAlignment;
    });
  });

  // Add balance formulas (starting from row 7, formula references previous row)
  const bakiyeColIndex = columns.indexOf(COLS.BAKIYE) + 1;
  const plusColIndex = columns.indexOf(COLS.NEW_ISLEM_TUTARI_PLUS) + 1;
  const minusColIndex = columns.indexOf(COLS.NEW_ISLEM_TUTARI_MINUS) + 1;

  if (bakiyeColIndex > 0 && plusColIndex > 0 && minusColIndex > 0) {
    const bakiyeCol = String.fromCharCode(64 + bakiyeColIndex);
    const plusCol = String.fromCharCode(64 + plusColIndex);
    const minusCol = String.fromCharCode(64 + minusColIndex);

    for (let r = 7; r <= result.data.length + 5; r++) {
      const cell = sheet.getCell(`${bakiyeCol}${r}`);
      cell.value = { formula: `${bakiyeCol}${r - 1}+${plusCol}${r}+${minusCol}${r}` };
      cell.numFmt = numberFormat;
      cell.alignment = centerAlignment;
    }
  }
}

// --- Main Process ---
async function processAndSave() {
  if (selectedFiles.length === 0) {
    await ipcRenderer.invoke('show-message', {
      type: 'error',
      title: 'Hata',
      message: 'Lütfen önce bir veya daha fazla dosya seçin!'
    });
    return;
  }

  const mode = activeTab;
  const startDate = mode === 'tarihli' ? startDateInput.value : null;
  const merge = mergeCheckbox.checked && selectedFiles.length > 1;

  if (mode === 'tarihli' && !startDate) {
    await ipcRenderer.invoke('show-message', {
      type: 'error',
      title: 'Hata',
      message: 'Lütfen ekstre başlangıç tarihini seçin!'
    });
    return;
  }

  processBtn.disabled = true;
  setStatus('Dosyalar işleniyor...', 'info');
  setProgress(0);

  const results = [];
  const errors = [];
  const warnings = [];

  for (let i = 0; i < selectedFiles.length; i++) {
    const filePath = selectedFiles[i];
    const fileName = path.basename(filePath);
    setStatus(`İşleniyor (${i + 1}/${selectedFiles.length}): ${fileName}`, 'info');
    setProgress((i + 1) / selectedFiles.length * 100);

    // Yield to event loop to keep UI responsive
    await new Promise(resolve => setTimeout(resolve, 0));

    try {
      const result = processSingleFile(filePath, mode, startDate);
      results.push(result);

      if (result.warnings.length) {
        warnings.push(`${fileName}:\n  - ${result.warnings.join('\n  - ')}`);
      }
    } catch (err) {
      errors.push(`${fileName}: ${err.message}`);
    }
  }

  processBtn.disabled = false;

  if (results.length === 0) {
    setStatus('İşlem başarısız.', 'error');
    await ipcRenderer.invoke('show-message', {
      type: 'error',
      title: 'Hata',
      message: 'Hiçbir dosya işlenemedi.\n\n' + errors.join('\n')
    });
    return;
  }

  // Save files
  let savedOutputDir = null;

  if (results.length === 1 || merge) {
    const defaultName = merge ? 'Birlestirilmis_Ekstreler.xlsx' :
      `${path.basename(selectedFiles[0], path.extname(selectedFiles[0]))}_Ekstre.xlsx`;
    const savePath = await ipcRenderer.invoke('save-file', defaultName, outputDir || undefined);

    if (savePath) {
      await writeExcelFile(savePath, results, merge);
      savedOutputDir = path.dirname(savePath);
    }
  } else {
    savedOutputDir = await ipcRenderer.invoke('select-directory', outputDir || undefined);

    if (savedOutputDir) {
      for (let i = 0; i < results.length; i++) {
        const result = results[i];
        const originalName = path.basename(selectedFiles[i], path.extname(selectedFiles[i]));
        const outputPath = path.join(savedOutputDir, `${originalName}_Ekstre.xlsx`);
        await writeExcelFile(outputPath, [result], false);
      }
    }
  }

  if (!savedOutputDir) {
    setStatus('İşlem kullanıcı tarafından iptal edildi.', 'warning');
    return;
  }

  // Show summary
  const successCount = results.length;
  const errorCount = errors.length;

  let statusClass = 'success';
  if (errors.length) statusClass = 'error';
  else if (warnings.length) statusClass = 'warning';

  setStatus(`İşlem tamamlandı. ${successCount} başarılı, ${errorCount} hatalı.`, statusClass);

  let summaryMessage = `${successCount} dosya başarıyla işlendi.`;
  if (warnings.length) {
    summaryMessage += '\n\nBazı dosyalarda uyarılar bulundu:\n' + warnings.join('\n');
  }
  if (errors.length) {
    summaryMessage += `\n\n${errorCount} dosyada hata oluştu:\n` + errors.join('\n');
  }

  const msgType = errors.length ? 'warning' : 'info';
  await ipcRenderer.invoke('show-message', {
    type: msgType,
    title: errors.length ? 'İşlem Sonucu: Hatalar Var' : 'İşlem Tamamlandı',
    message: summaryMessage
  });

  // Ask to open folder
  const openFolder = await ipcRenderer.invoke('show-message', {
    type: 'question',
    title: 'Klasörü Aç',
    message: `İşlenen dosyaların bulunduğu klasörü açmak ister misiniz?\n${savedOutputDir}`,
    buttons: ['Evet', 'Hayır']
  });

  if (openFolder === 0) {
    await ipcRenderer.invoke('open-folder', savedOutputDir);
  }
}

// --- UI Helpers ---
function setProgress(percent) {
  progressFill.style.width = `${percent}%`;
  progressFill.setAttribute('aria-valuenow', Math.round(percent));
}

function setStatus(text, type = 'info') {
  statusEl.textContent = text;
  statusEl.className = 'status-text';

  if (type === 'success') statusEl.classList.add('s-success');
  else if (type === 'error') statusEl.classList.add('s-error');
  else if (type === 'warning') statusEl.classList.add('s-warning');
}

// Initialize
updateFileList();
