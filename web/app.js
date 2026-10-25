'use strict';

const NUEVA_EPS_ALLOWED = new Set([
  'FEV', 'HEV', 'EPI', 'PDX', 'DQX', 'RAN', 'TAP', 'FMO',
  'OPF', 'LDP', 'HAU', 'HAM', 'PDE',
]);
const NUEVA_EPS_REQUIRED = new Set(['FEV', 'OPF', 'PDE', 'LDP']);
const NUEVA_EPS_NOT_COUNTED = new Set(['LDP']);
const STATUS_COLORS = {
  green: '99FF99',
  blue: '99CCFF',
  red: 'FF9999',
  yellow: 'FFFF99',
  orange: 'FFB347',
  purple: 'DDA0DD',
  brown: 'D2B48C',
  pink: 'FFC0CB',
};

let workbook = null;
let worksheet = null;
let currentFileName = 'auditoria_resultado.xlsx';
let columns = [];
let gridData = [];
let rowColors = [];
let supportFiles = [];
let hasBeenAudited = false;
const columnFilters = new Map();

const element = (id) => document.getElementById(id);

function showToast(message, type = 'success') {
  const toast = document.createElement('div');
  toast.className = `toast toast-${type}`;
  toast.textContent = message;
  document.querySelector('.toast-container').appendChild(toast);
  window.setTimeout(() => toast.remove(), 4200);
}

function setStatus(message, isError = false) {
  const status = element('audit-status');
  status.textContent = message;
  status.classList.toggle('error-text', isError);
}

function cellText(value) {
  if (value === null || value === undefined) return '';
  if (value instanceof Date) return value.toLocaleDateString('es-CO');
  if (typeof value === 'object') {
    if (Array.isArray(value.richText)) return value.richText.map((part) => part.text).join('');
    if (value.text) return String(value.text);
    if ('result' in value) return cellText(value.result);
    if (value.hyperlink) return value.text || value.hyperlink;
    return '';
  }
  return String(value).trim();
}

function normalizedKey(value) {
  return cellText(value).trim();
}

function getHeaderColumns(sheet) {
  const headers = [];
  const firstRow = sheet.getRow(1);
  for (let index = 1; index <= firstRow.cellCount; index += 1) {
    headers.push(cellText(firstRow.getCell(index).value) || `COLUMNA_${index}`);
  }
  return headers;
}

function sheetToRows(sheet, headers) {
  const rows = [];
  for (let rowNumber = 2; rowNumber <= sheet.rowCount; rowNumber += 1) {
    const sourceRow = sheet.getRow(rowNumber);
    const row = {};
    headers.forEach((header, index) => {
      row[header] = cellText(sourceRow.getCell(index + 1).value);
    });
    rows.push(row);
  }
  return rows;
}

function findColumn(pattern) {
  return columns.find((column) => pattern.test(column.trim()));
}

function getColumnValues(column) {
  return [...new Set(gridData.map((row) => cellText(row[column])))].sort((a, b) => (
    a.localeCompare(b, 'es', { numeric: true, sensitivity: 'base' })
  ));
}

function filteredRowIndexes() {
  const term = element('table-search').value.trim().toLocaleLowerCase();
  return gridData.map((row, index) => ({ row, index })).filter(({ row }) => {
    const values = columns.map((column) => cellText(row[column]));
    if (term && !values.some((value) => value.toLocaleLowerCase().includes(term))) return false;
    for (const [column, allowedValues] of columnFilters) {
      if (!allowedValues.has(cellText(row[column]))) return false;
    }
    return true;
  });
}

function renderGrid() {
  const table = element('data-grid');
  const thead = table.querySelector('thead');
  const tbody = table.querySelector('tbody');
  thead.replaceChildren();
  tbody.replaceChildren();

  const header = document.createElement('tr');
  columns.forEach((column) => {
    const th = document.createElement('th');
    const headerContent = document.createElement('div');
    headerContent.className = 'column-heading';
    const title = document.createElement('span');
    title.textContent = column;
    const filterButton = document.createElement('button');
    filterButton.type = 'button';
    filterButton.className = `column-filter-button${columnFilters.has(column) ? ' is-active' : ''}`;
    filterButton.textContent = columnFilters.has(column) ? '▾ •' : '▾';
    filterButton.setAttribute('aria-label', `Filtrar columna ${column}`);
    filterButton.setAttribute('aria-haspopup', 'dialog');
    filterButton.setAttribute('aria-expanded', 'false');
    filterButton.addEventListener('click', (event) => {
      event.stopPropagation();
      openColumnFilter(column, filterButton);
    });
    headerContent.append(title, filterButton);
    th.appendChild(headerContent);
    header.appendChild(th);
  });
  thead.appendChild(header);

  const visibleRows = filteredRowIndexes();
  visibleRows.forEach(({ row, index: rowIndex }) => {
    const tr = document.createElement('tr');
    if (rowColors[rowIndex]) tr.className = `row-${rowColors[rowIndex]}`;
    columns.forEach((column) => {
      const td = document.createElement('td');
      td.textContent = cellText(row[column]);
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
  element('row-count').textContent = `${visibleRows.length.toLocaleString('es-CO')} de ${gridData.length.toLocaleString('es-CO')} filas`;
  element('results-section').hidden = gridData.length === 0;
  const activeCount = columnFilters.size;
  element('clear-filters').disabled = activeCount === 0 && !element('table-search').value;
  element('clear-filters').textContent = activeCount ? `Limpiar filtros (${activeCount})` : 'Limpiar filtros';
}

function renderFilterOptions(column, draft, query, optionsContainer, allCheckbox) {
  const values = getColumnValues(column);
  optionsContainer.replaceChildren();
  const normalizedQuery = query.toLocaleLowerCase();
  values.forEach((value) => {
    const label = document.createElement('label');
    label.className = 'filter-option';
    label.hidden = !value.toLocaleLowerCase().includes(normalizedQuery);
    const checkbox = document.createElement('input');
    checkbox.type = 'checkbox';
    checkbox.checked = draft.has(value);
    checkbox.addEventListener('change', () => {
      if (checkbox.checked) draft.add(value);
      else draft.delete(value);
      updateSelectAll();
    });
    const text = document.createElement('span');
    text.textContent = value || '(Vacías)';
    label.append(checkbox, text);
    optionsContainer.appendChild(label);
  });
  updateSelectAll();

  function updateSelectAll() {
    allCheckbox.checked = draft.size === values.length;
    allCheckbox.indeterminate = draft.size > 0 && draft.size < values.length;
  }
}

function openColumnFilter(column, anchor) {
  const menu = element('column-filter-menu');
  const allValues = getColumnValues(column);
  const draft = new Set(columnFilters.get(column) || allValues);
  menu.replaceChildren();
  menu.setAttribute('role', 'dialog');
  menu.setAttribute('aria-label', `Filtro de ${column}`);

  const title = document.createElement('strong');
  title.className = 'filter-menu-title';
  title.textContent = `Filtrar: ${column}`;
  const search = document.createElement('input');
  search.type = 'search';
  search.className = 'filter-menu-search';
  search.placeholder = 'Buscar valores…';
  search.setAttribute('aria-label', `Buscar valores de ${column}`);

  const selectAllLabel = document.createElement('label');
  selectAllLabel.className = 'filter-select-all';
  const selectAll = document.createElement('input');
  selectAll.type = 'checkbox';
  const selectAllText = document.createElement('span');
  selectAllText.textContent = '(Seleccionar todo)';
  selectAllLabel.append(selectAll, selectAllText);

  const options = document.createElement('div');
  options.className = 'filter-options';
  const buttons = document.createElement('div');
  buttons.className = 'filter-menu-actions';
  const cancel = document.createElement('button');
  cancel.type = 'button';
  cancel.className = 'filter-menu-cancel';
  cancel.textContent = 'Cancelar';
  const apply = document.createElement('button');
  apply.type = 'button';
  apply.className = 'filter-menu-apply';
  apply.textContent = 'Aplicar';
  buttons.append(cancel, apply);
  menu.append(title, search, selectAllLabel, options, buttons);

  renderFilterOptions(column, draft, '', options, selectAll);
  selectAll.addEventListener('change', () => {
    if (selectAll.checked) allValues.forEach((value) => draft.add(value));
    else draft.clear();
    renderFilterOptions(column, draft, search.value, options, selectAll);
  });
  search.addEventListener('input', () => renderFilterOptions(column, draft, search.value, options, selectAll));
  cancel.addEventListener('click', () => {
    menu.hidden = true;
    anchor.setAttribute('aria-expanded', 'false');
  });
  apply.addEventListener('click', () => {
    if (draft.size === allValues.length) columnFilters.delete(column);
    else columnFilters.set(column, new Set(draft));
    menu.hidden = true;
    anchor.setAttribute('aria-expanded', 'false');
    renderGrid();
  });

  const bounds = anchor.getBoundingClientRect();
  menu.hidden = false;
  anchor.setAttribute('aria-expanded', 'true');
  const menuBounds = menu.getBoundingClientRect();
  menu.style.left = `${Math.max(8, Math.min(bounds.left, window.innerWidth - menuBounds.width - 8))}px`;
  menu.style.top = `${Math.max(8, Math.min(bounds.bottom + 6, window.innerHeight - menuBounds.height - 8))}px`;
  search.focus();
}

document.addEventListener('click', (event) => {
  const menu = element('column-filter-menu');
  if (!menu.hidden && !menu.contains(event.target)) {
    menu.hidden = true;
    document.querySelectorAll('.column-filter-button[aria-expanded="true"]').forEach((button) => {
      button.setAttribute('aria-expanded', 'false');
    });
  }
});

document.addEventListener('keydown', (event) => {
  if (event.key === 'Escape') {
    const menu = element('column-filter-menu');
    menu.hidden = true;
    document.querySelectorAll('.column-filter-button[aria-expanded="true"]').forEach((button) => {
      button.setAttribute('aria-expanded', 'false');
    });
  }
});

function renderSummary() {
  const summary = element('summary');
  summary.replaceChildren();
  const counts = new Map();
  gridData.forEach((row) => {
    const result = row.RESULTADO_AUDITORIA || 'Sin auditar';
    const category = result.startsWith('SIN RADICAR') ? 'Completas'
      : result.includes('SOBRE COSTO') ? 'Sobre costo'
        : result === 'NO CARPETA' ? 'Sin carpeta'
          : result.includes('DUPLICADO') ? 'Duplicadas'
            : result.startsWith('FALTAN') || result.includes('FALTA') ? 'Incompletas'
              : result === 'Sin auditar' ? 'Sin auditar' : 'Revisar';
    counts.set(category, (counts.get(category) || 0) + 1);
  });
  ['Completas', 'Incompletas', 'Sin carpeta', 'Duplicadas', 'Sobre costo', 'Revisar', 'Sin auditar']
    .forEach((label) => {
      const card = document.createElement('div');
      card.className = 'summary-item';
      const number = document.createElement('strong');
      number.textContent = (counts.get(label) || 0).toLocaleString('es-CO');
      const caption = document.createElement('span');
      caption.textContent = label;
      card.append(number, caption);
      summary.appendChild(card);
    });
}

function renderResults() {
  renderSummary();
  renderGrid();
}

async function loadWorkbook(file) {
  if (!window.ExcelJS) throw new Error('No se pudo cargar ExcelJS. Revisa la conexión a Internet y vuelve a intentarlo.');
  const loadedWorkbook = new ExcelJS.Workbook();
  await loadedWorkbook.xlsx.load(await file.arrayBuffer());
  const sheet = loadedWorkbook.worksheets[0];
  if (!sheet) throw new Error('El archivo Excel no contiene hojas de cálculo.');
  return { loadedWorkbook, sheet };
}

element('excel-file').addEventListener('change', async (event) => {
  const file = event.target.files[0];
  if (!file) return;
  try {
    const loaded = await loadWorkbook(file);
    workbook = loaded.loadedWorkbook;
    worksheet = loaded.sheet;
    columns = getHeaderColumns(worksheet);
    gridData = sheetToRows(worksheet, columns);
    rowColors = Array(gridData.length).fill('');
    columnFilters.clear();
    element('table-search').value = '';
    currentFileName = file.name.replace(/\.xlsx$/i, '_auditado.xlsx');
    hasBeenAudited = false;
    element('results-title').textContent = `Listado: ${file.name}`;
    renderResults();
    setStatus(`${gridData.length.toLocaleString('es-CO')} filas listas para revisar.`);
    showToast('Excel cargado correctamente.');
  } catch (error) {
    setStatus(error.message, true);
    showToast(error.message, 'error');
  }
});

async function collectSupportFiles(fileList) {
  const result = [];
  for (const file of Array.from(fileList)) {
    const relativePath = file.webkitRelativePath || file.name;
    if (/\.zip$/i.test(file.name)) {
      if (!window.JSZip) throw new Error('No se pudo cargar el lector ZIP. Revisa la conexión a Internet.');
      try {
        const archive = await JSZip.loadAsync(file);
        for (const [memberPath, member] of Object.entries(archive.files)) {
          if (member.dir) continue;
          result.push({
            name: memberPath.split('/').pop(),
            path: `${relativePath}/${memberPath}`,
            isFile: true,
          });
        }
      } catch (error) {
        throw new Error(`No se pudo leer el archivo ZIP "${file.name}": ${error.message}`);
      }
      continue;
    }
    result.push({ name: file.name, path: relativePath, isFile: true });
  }
  return result;
}

element('support-folder').addEventListener('change', async (event) => {
  try {
    supportFiles = await collectSupportFiles(event.target.files);
    const pdfs = supportFiles.filter((file) => /\.pdf$/i.test(file.name)).length;
    setStatus(`Carpeta cargada: ${pdfs.toLocaleString('es-CO')} PDF y ${supportFiles.length.toLocaleString('es-CO')} archivos en total.`);
    showToast('Carpeta de soportes lista.');
  } catch (error) {
    supportFiles = [];
    setStatus(error.message, true);
    showToast(error.message, 'error');
  }
});

function extractInvoiceId(value) {
  const text = normalizedKey(value);
  const match = text.match(/(\d+)$/);
  return match ? (match[1].replace(/^0+/, '') || '0') : text;
}

function matchingFiles(invoiceId) {
  return supportFiles.filter((file) => file.path.toUpperCase().includes(invoiceId.toUpperCase()));
}

function candidateFolders(invoiceId, matches) {
  const folders = new Set();
  matches.forEach((file) => {
    const parts = file.path.split(/[\\/]/).filter(Boolean);
    for (let index = 0; index < parts.length - 1; index += 1) {
      if (parts[index].toUpperCase().includes(invoiceId.toUpperCase())) {
        folders.add(parts.slice(0, index + 1).join('/'));
      }
    }
  });
  return Array.from(folders);
}

function isWrongLocation(path) {
  const parts = path.split(/[\\/]/);
  const name = parts[parts.length - 1] || '';
  const hasLetters = /[a-z]/i.test(name);
  const upperParts = parts.map((part) => part.toUpperCase());
  return hasLetters
    ? !upperParts.some((part) => part.includes('PENDIENTE'))
    : !upperParts.some((part) => part.includes('CONSULTA EXTERNA'));
}

function pendingDescription(path, invoiceId) {
  const name = path.split(/[\\/]/).filter(Boolean).pop() || '';
  return name.toUpperCase().replace(invoiceId, '').replace(/[_-]/g, ' ').trim() || 'PENDIENTE';
}

function statusResult(message, color) {
  return { message, color };
}

function applyOverCost(result, totalValue) {
  const totalText = String(totalValue ?? '').replace(/[^\d,.-]/g, '').replace(/\.(?=\d{3}(?:\D|$))/g, '').replace(',', '.');
  if (Number(totalText) <= 5000000) return result;
  const color = result.color === 'purple' ? 'purple'
    : result.color === 'green' || result.color === 'blue' ? 'green'
      : result.color === 'yellow' || result.color === 'orange' ? 'yellow' : 'pink';
  return statusResult(`SOBRE COSTO - ${result.message}`, color);
}

function auditNuevaEps(files, invoiceId) {
  const validSupports = new Set();
  let invalidCount = 0;
  files.filter((file) => /\.pdf$/i.test(file.name)).forEach((file) => {
    const match = file.name.match(/^([A-Z]{3})_(\d+)_(\d+)\.pdf$/i);
    if (!match) {
      invalidCount += 1;
      return;
    }
    const supportType = match[1].toUpperCase();
    const fileInvoiceId = match[3].replace(/^0+/, '') || '0';
    if (!NUEVA_EPS_ALLOWED.has(supportType) || fileInvoiceId !== invoiceId || validSupports.has(supportType)) {
      invalidCount += 1;
      return;
    }
    validSupports.add(supportType);
  });
  const counted = [...validSupports].filter((type) => !NUEVA_EPS_NOT_COUNTED.has(type)).length;
  const missing = [...NUEVA_EPS_REQUIRED].filter((type) => !validSupports.has(type)).sort();
  if (invalidCount) return statusResult(`MAL SOPORTADO (${invalidCount} PDF)`, 'orange');
  if (missing.length || counted < 4) {
    const detail = missing.length ? `; faltan ${missing.join(', ')}` : '';
    return statusResult(`FALTAN SOPORTES (${counted}/4${detail})`, 'yellow');
  }
  return statusResult('SIN RADICAR', 'green');
}

function ruleForCompany(company, files, path, invoiceId) {
  const pdfs = files.filter((file) => /\.pdf$/i.test(file.name) && !/^LDP_/i.test(file.name));
  const pdfNames = pdfs.map((file) => file.name.toUpperCase());
  const xmlCount = files.filter((file) => /\.xml$/i.test(file.name)
    || (/^AD/i.test(file.name) && !/\.pdf$/i.test(file.name))).length;
  const count = pdfs.length;

  if (company === 'General') {
    const isPolicia = /POLICIA/i.test(path);
    const isPrevisora = /PREVISORA/i.test(path);
    const required = isPrevisora ? 3 : 4;
    let invalid = !isPolicia && invoiceId.length !== 7 ? required : 0;
    pdfNames.forEach((name) => {
      const stem = name.replace(/\.PDF$/i, '');
      if (!isPolicia && (stem.includes('__') || !stem.includes(invoiceId))) invalid += 1;
    });
    if (invalid) return statusResult(`MAL SOPORTADO (${invalid}/${required})`, 'orange');
    if (count >= required) return statusResult('SIN RADICAR', 'green');
    return statusResult(`FALTAN SOPORTES (${count}/${required})`, 'yellow');
  }
  if (company === 'ADRES') {
    return count >= 3 && xmlCount > 0
      ? statusResult('SIN RADICAR (+XML)', 'green')
      : statusResult(`FALTAN SOPORTES (${count} PDFs, ${xmlCount} XML)`, 'yellow');
  }
  if (company === 'COOSALUD / SANIDAD MILITAR') {
    return count ? statusResult('SIN RADICAR', 'green') : statusResult('VACÍO / FALTAN SOPORTES', 'yellow');
  }
  if (company.startsWith('GRUPO SOLIDARIA')) {
    return count === 2 ? statusResult('SIN RADICAR (2 docs)', 'green') : statusResult(`FALTAN SOPORTES (${count}/2)`, 'yellow');
  }
  if (company === 'SEGUROS BOLIVAR ARL') {
    if (count && pdfNames.some((name) => name.includes('FAC_'))) return statusResult('SIN RADICAR (FAC OK)', 'green');
    return count ? statusResult("FALTA PDF 'FAC_'", 'yellow') : statusResult('VACÍO', 'yellow');
  }
  if (company.startsWith('GRUPO ESTADO')) {
    return count === 1 ? statusResult('SIN RADICAR (1 solo doc)', 'green') : statusResult(`ERROR APILAMIENTO (${count})`, 'yellow');
  }
  if (company === 'AXA COLPATRIA SEGUROS ESCOLARES') {
    return count === 3 ? statusResult('SIN RADICAR (3 docs)', 'green') : statusResult(`FALTAN SOPORTES (${count}/3)`, 'yellow');
  }
  if (company === 'MUNDIAL DE SEGUROS') {
    return count === 1 ? statusResult('SIN RADICAR (1 solo doc)', 'green') : statusResult(`ERROR APILAMIENTO (${count})`, 'yellow');
  }
  if (company === 'AXA COLPATRIA SOAT / EQUIDAD SEGUROS') {
    return count >= 4 ? statusResult('SIN RADICAR (4 docs)', 'green') : statusResult(`FALTAN SOPORTES (${count}/4)`, 'yellow');
  }
  if (company === 'LA PREVISORA') {
    const hasTypes = ['FEV', 'EPI', 'PDX'].every((type) => pdfNames.some((name) => name.includes(type)));
    if (count === 3 && hasTypes) return statusResult('SIN RADICAR (3 docs)', 'green');
    if (count === 3) return statusResult('NOMENCLATURA INCORRECTA', 'yellow');
    return statusResult(`FALTAN SOPORTES (${count}/3)`, 'yellow');
  }
  if (company === 'POSITIVA ARL Y SEGUROS ESCOLARES') {
    return count >= 3 ? statusResult('SIN RADICAR (Positiva)', 'green') : statusResult(`FALTAN SOPORTES (${count})`, 'yellow');
  }
  if (company === 'SURA EPS / SURA ARL') {
    const hasNit = pdfNames.some((name) => name.includes('800218979')) || path.includes('800218979');
    if (!count) return statusResult('VACÍO / FALTAN SOPORTES', 'yellow');
    return hasNit ? statusResult('SIN RADICAR (SURA OK)', 'green') : statusResult('FORMATO SURA INCORRECTO', 'orange');
  }
  return statusResult(`REVISAR (${count})`, 'yellow');
}

function auditInvoice(invoiceValue, company, totalValue) {
  const finish = (result) => applyOverCost(result, totalValue);
  const invoiceId = extractInvoiceId(invoiceValue);
  if (!invoiceId) return finish(statusResult('FALTA FACTURA', 'red'));

  if (company === 'COLSANITAS (PREPAGADA)') {
    const id = invoiceId.replace(/^0+/, '').toUpperCase();
    const invoiceExists = supportFiles.some((file) => /(^|\/)FACTURAS?(\/|$)/i.test(file.path)
      && /\.pdf$/i.test(file.name) && file.name.toUpperCase().includes(id));
    const supportExists = supportFiles.some((file) => /(^|\/)SOPORTES?(\/|$)/i.test(file.path)
      && /\.pdf$/i.test(file.name) && file.name.toUpperCase().includes(id) && /SOP/i.test(file.name));
    if (invoiceExists && supportExists) return finish(statusResult('SIN RADICAR', 'green'));
    if (invoiceExists || supportExists) return finish(statusResult('FALTAN SOPORTES', 'yellow'));
    return finish(statusResult('NO CARPETA', 'red'));
  }

  const matches = matchingFiles(invoiceId);
  const folders = candidateFolders(invoiceId, matches);
  if (folders.length > 1) return finish(statusResult(`DUPLICADO (${folders.length} ubicaciones)`, 'purple'));
  const path = folders[0] || matches[0]?.path || '';
  if (!path) return finish(statusResult('NO CARPETA', 'red'));
  if (isWrongLocation(path)) return finish(statusResult('UBICACIÓN EQUIVOCADA', 'brown'));

  const folderFiles = path === matches[0]?.path
    ? matches
    : supportFiles.filter((file) => file.path.startsWith(`${path}/`) || file.path.startsWith(`${path}\\`));
  const name = path.split(/[\\/]/).filter(Boolean).pop() || '';
  if (company === 'NUEVA EPS' && /[a-z]/i.test(name)) {
    return finish(statusResult(pendingDescription(path, invoiceId), 'blue'));
  }
  if (company === 'NUEVA EPS') return finish(auditNuevaEps(folderFiles, invoiceId));
  if (/[a-z]/i.test(name) && !company.includes('SURA') && !name.toUpperCase().includes('FAC_')) {
    return finish(statusResult(pendingDescription(path, invoiceId), 'blue'));
  }

  return finish(ruleForCompany(company, folderFiles, path, invoiceId));
}

function colorCell(cell, color) {
  const argb = `FF${STATUS_COLORS[color] || 'FFFFFF'}`;
  cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb } };
  cell.font = { ...cell.font, color: { argb: color === 'red' || color === 'orange' ? 'FF9A3412' : 'FF111827' } };
}

function applyGridToWorksheet() {
  const headerRow = worksheet.getRow(1);
  let resultColumn = columns.findIndex((column) => column.trim().toUpperCase() === 'RESULTADO_AUDITORIA') + 1;
  if (!resultColumn) {
    resultColumn = columns.length + 1;
    columns.push('RESULTADO_AUDITORIA');
    headerRow.getCell(resultColumn).value = 'RESULTADO_AUDITORIA';
  }

  const headerIndexes = new Map();
  columns.forEach((column, index) => headerIndexes.set(column, index + 1));
  gridData.forEach((row, index) => {
    const excelRow = worksheet.getRow(index + 2);
    columns.forEach((column) => {
      if (column === 'RESULTADO_AUDITORIA') return;
      const cell = excelRow.getCell(headerIndexes.get(column));
      const current = cellText(cell.value);
      if (current !== cellText(row[column])) cell.value = row[column];
    });
    excelRow.getCell(resultColumn).value = row.RESULTADO_AUDITORIA || '';
    if (rowColors[index]) {
      for (let col = 1; col <= resultColumn; col += 1) {
        colorCell(excelRow.getCell(col), rowColors[index]);
        excelRow.getCell(col).font = {
          ...excelRow.getCell(col).font,
          color: { argb: row.RESULTADO_AUDITORIA.startsWith('SOBRE COSTO') ? 'FFFF0000' : 'FF000000' },
        };
      }
    }
  });
  worksheet.autoFilter = `A1:${columnLetter(resultColumn)}${Math.max(1, worksheet.rowCount)}`;
}

function columnLetter(columnNumber) {
  let remaining = columnNumber;
  let result = '';
  while (remaining > 0) {
    const remainder = (remaining - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    remaining = Math.floor((remaining - 1) / 26);
  }
  return result;
}

element('audit-all').addEventListener('click', async () => {
  if (!workbook || !worksheet || !gridData.length) {
    showToast('Primero carga el listado Excel.', 'error');
    return;
  }
  if (!supportFiles.length) {
    showToast('Selecciona una carpeta de soportes antes de iniciar la auditoría.', 'error');
    return;
  }
  const button = element('audit-all');
  button.disabled = true;
  setStatus('Auditando…');
  try {
    const company = element('company-select').value;
    const invoiceColumn = findColumn(/SFANUMFAC/i);
    if (!invoiceColumn) throw new Error("El Excel no contiene la columna 'SFANUMFAC'.");
    if (!columns.some((column) => column.trim().toUpperCase() === 'RESULTADO_AUDITORIA')) {
      columns.push('RESULTADO_AUDITORIA');
    }
    rowColors = [];
    const totalColumn = findColumn(/^TOTAL_FACTURADO$/i);
    gridData.forEach((row, index) => {
      if (worksheet.getRow(index + 2).hidden) {
        rowColors[index] = '';
        return;
      }
      const result = auditInvoice(row[invoiceColumn], company, totalColumn ? row[totalColumn] : '');
      row.RESULTADO_AUDITORIA = result.message;
      rowColors[index] = result.color;
    });
    applyGridToWorksheet();
    hasBeenAudited = true;
    renderResults();
    setStatus(`Auditoría completada: ${gridData.length.toLocaleString('es-CO')} filas.`);
    showToast('Auditoría completada.');
  } catch (error) {
    setStatus(error.message, true);
    showToast(error.message, 'error');
  } finally {
    button.disabled = false;
  }
});

function downloadWorkbook(targetWorkbook, fileName) {
  return targetWorkbook.xlsx.writeBuffer().then((buffer) => {
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = fileName;
    link.click();
    window.setTimeout(() => URL.revokeObjectURL(url), 1000);
  });
}

element('download-excel').addEventListener('click', async () => {
  if (!workbook || !worksheet) {
    showToast('No hay un Excel cargado para descargar.', 'error');
    return;
  }
  try {
    if (hasBeenAudited) applyGridToWorksheet();
    await downloadWorkbook(workbook, currentFileName);
    showToast('Excel descargado correctamente.');
  } catch (error) {
    showToast(`No se pudo generar el Excel: ${error.message}`, 'error');
  }
});

function findHeaderIndex(sheet, name) {
  const row = sheet.getRow(1);
  for (let col = 1; col <= row.cellCount; col += 1) {
    if (cellText(row.getCell(col).value).trim().toUpperCase() === name.toUpperCase()) return col;
  }
  return 0;
}

element('merge-excel').addEventListener('click', async () => {
  const baseFile = element('base-excel-file').files[0];
  const newFile = element('new-excel-file').files[0];
  if (!baseFile || !newFile) {
    showToast('Selecciona el Excel base y el Excel nuevo.', 'error');
    return;
  }
  const button = element('merge-excel');
  button.disabled = true;
  try {
    const base = await loadWorkbook(baseFile);
    const incoming = await loadWorkbook(newFile);
    const baseSheet = base.sheet;
    const newSheet = incoming.sheet;
    const baseKeyIndex = findHeaderIndex(baseSheet, 'SFANUMFAC');
    const newKeyIndex = findHeaderIndex(newSheet, 'SFANUMFAC');
    if (!baseKeyIndex || !newKeyIndex) throw new Error("Ambos Excel deben tener la columna 'SFANUMFAC'.");

    const baseHeaders = getHeaderColumns(baseSheet);
    const newHeaders = getHeaderColumns(newSheet);
    const headerIndexes = new Map();
    baseHeaders.forEach((header, index) => headerIndexes.set(header.trim().toUpperCase(), index + 1));
    newHeaders.forEach((header) => {
      if (!headerIndexes.has(header.trim().toUpperCase())) {
        const index = baseHeaders.length + 1;
        baseHeaders.push(header);
        headerIndexes.set(header.trim().toUpperCase(), index);
        baseSheet.getRow(1).getCell(index).value = header;
      }
    });

    const existing = new Set();
    for (let rowNumber = 2; rowNumber <= baseSheet.rowCount; rowNumber += 1) {
      const key = normalizedKey(baseSheet.getRow(rowNumber).getCell(baseKeyIndex).value);
      if (key) existing.add(key);
    }
    let addedCount = 0;
    const incomingKeyHeader = cellText(newSheet.getRow(1).getCell(newKeyIndex).value).trim().toUpperCase();
    const targetKeyIndex = headerIndexes.get(incomingKeyHeader);
    for (let rowNumber = 2; rowNumber <= newSheet.rowCount; rowNumber += 1) {
      const incomingRow = newSheet.getRow(rowNumber);
      const key = normalizedKey(incomingRow.getCell(newKeyIndex).value);
      if (key && existing.has(key)) continue;
      const targetRow = baseSheet.addRow([]);
      newHeaders.forEach((header, index) => {
        targetRow.getCell(headerIndexes.get(header.trim().toUpperCase())).value = incomingRow.getCell(index + 1).value;
      });
      if (key) existing.add(key);
      addedCount += 1;
    }
    if (!targetKeyIndex) throw new Error("No se pudo ubicar la columna 'SFANUMFAC' en el resultado.");
    const fileName = baseFile.name.replace(/\.xlsx$/i, '_consolidado.xlsx');
    await downloadWorkbook(base.loadedWorkbook, fileName);
    showToast(`Consolidación lista: se agregaron ${addedCount} filas nuevas.`);
  } catch (error) {
    showToast(`No se pudo unir los Excel: ${error.message}`, 'error');
  } finally {
    button.disabled = false;
  }
});

element('table-search').addEventListener('input', renderGrid);
element('clear-filters').addEventListener('click', () => {
  columnFilters.clear();
  element('table-search').value = '';
  renderGrid();
});
