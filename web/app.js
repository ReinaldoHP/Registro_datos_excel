// app.js – Lógica principal de la SPA Auditor de Facturas
// Carga de librerías: SheetJS (XLSX) y ExcelJS están incluidos vía CDN en index.html
// Esta implementación está pensada para ejecutarse en el navegador sin servidor.

// ---------------------------------------------------------------------
// UTILIDADES
// ---------------------------------------------------------------------
/**
 * Formatea un número a string con dos decimales y separador de miles.
 */
function formatNumber(num) {
  return Number(num).toLocaleString('es-CO', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

/**
 * Crea una celda <td> editable que dispara eventos de cambio.
 */
function createEditableCell(value, rowIndex, colKey) {
  const td = document.createElement('td');
  td.textContent = value;
  td.contentEditable = true;
  td.dataset.row = rowIndex;
  td.dataset.col = colKey;
  td.addEventListener('blur', onCellEdit);
  return td;
}

/**
 * Genera la fila de encabezado y la fila de filtros a partir de las columnas detectadas.
 */
function renderHeader(columns) {
  const headerRow = document.getElementById('header-row');
  const filterRow = document.getElementById('filter-row');
  headerRow.innerHTML = '';
  filterRow.innerHTML = '';
  columns.forEach(col => {
    const th = document.createElement('th');
    th.textContent = col;
    headerRow.appendChild(th);

    const filterTh = document.createElement('th');
    const input = document.createElement('input');
    input.type = 'text';
    input.placeholder = 'Filtrar';
    input.dataset.col = col;
    input.addEventListener('input', onFilterChange);
    filterTh.appendChild(input);
    filterRow.appendChild(filterTh);
  });
  // Columna de acciones (eliminar)
  const actionTh = document.createElement('th');
  actionTh.textContent = 'Acciones';
  headerRow.appendChild(actionTh);
  const emptyTh = document.createElement('th');
  filterRow.appendChild(emptyTh);
}

/**
 * Renderiza todas las filas del <tbody> usando `gridData` (array de objetos).
 */
function renderGrid() {
  const tbody = document.querySelector('#data-grid tbody');
  tbody.innerHTML = '';
  gridData.forEach((row, i) => {
    const tr = document.createElement('tr');
    columns.forEach(col => {
      const td = createEditableCell(row[col] ?? '', i, col);
      tr.appendChild(td);
    });
    // Acción eliminar
    const actionTd = document.createElement('td');
    const delBtn = document.createElement('button');
    delBtn.textContent = '🗑️';
    delBtn.dataset.row = i;
    delBtn.addEventListener('click', onDeleteRow);
    actionTd.appendChild(delBtn);
    tr.appendChild(actionTd);
    tbody.appendChild(tr);
  });
}

// ---------------------------------------------------------------------
// ESTADO GLOBAL
// ---------------------------------------------------------------------
let isAuthenticated = false;
let gridData = []; // Array of objects, cada objeto = fila
let columns = []; // Nombres de columnas detectados
let supportFilesMap = new Map(); // clave = nombre de factura, valor = array de rutas de PDFs

// ---------------------------------------------------------------------
// LOGIN / LOGOUT
// ---------------------------------------------------------------------
  if (stored && atob(stored.password) === pass) {
    isAuthenticated = true;
    sessionStorage.setItem('auth', 'true');
    document.getElementById('login-section').classList.add('hidden');
    document.getElementById('register-section').classList.add('hidden');
    document.getElementById('dashboard').classList.remove('hidden');
    loginMessage.textContent = '';
    showToast('Acceso exitoso.', 'success');
  } else {
    loginMessage.textContent = 'Credenciales incorrectas.';
  }
});

// Manejo de registro
const registerForm = document.getElementById('register-form');
const registerMessage = document.getElementById('register-message');
registerForm.addEventListener('submit', e => {
  e.preventDefault();
  const user = document.getElementById('register-username').value.trim();
  const pass = document.getElementById('register-password').value.trim();
  const confirm = document.getElementById('register-confirm').value.trim();
  if (pass !== confirm) {
    registerMessage.textContent = 'Las contraseñas no coinciden.';
    return;
  }
  const users = JSON.parse(localStorage.getItem('users') || '[]');
  if (users.some(u => u.username === user)) {
    registerMessage.textContent = 'El usuario ya existe.';
    return;
  }
  const encoded = btoa(pass);
  users.push({ username: user, password: encoded });
  localStorage.setItem('users', JSON.stringify(users));
  registerMessage.textContent = 'Registro exitoso. Ya puedes iniciar sesión.';
});

// Enlaces para cambiar entre login y registro
document.getElementById('show-register').addEventListener('click', e => {
  e.preventDefault();
  document.getElementById('login-section').classList.add('hidden');
  document.getElementById('register-section').classList.remove('hidden');
});

document.getElementById('show-login').addEventListener('click', e => {
  e.preventDefault();
  document.getElementById('register-section').classList.add('hidden');
  document.getElementById('login-section').classList.remove('hidden');
});

// Manejo de logout
document.getElementById('logout').addEventListener('click', () => {
  isAuthenticated = false;
  sessionStorage.removeItem('auth');
  document.getElementById('dashboard').classList.add('hidden');
  document.getElementById('login-section').classList.remove('hidden');
  showToast('Sesión cerrada.', 'success');
});

// Función para requerir autenticación antes de ejecutar acciones críticas
function requireAuth() {
  if (!isAuthenticated) {
    showToast('Debe iniciar sesión para usar esta función.', 'error');
    return false;
  }
  return true;
}ión para usar esta función.');
    return false;
  }
  return true;
}

// ---------------------------------------------------------------------
// CARGA DE EXCEL
// ---------------------------------------------------------------------
// CARGA DE EXCEL
document.getElementById('excel-file').addEventListener('change', async ev => {
  if (!requireAuth()) return;
  const file = ev.target.files[0];
  if (!file) return;
  const data = await file.arrayBuffer();
  const workbook = XLSX.read(data, { type: 'array' });
  const firstSheetName = workbook.SheetNames[0];
  const ws = workbook.Sheets[firstSheetName];
  const json = XLSX.utils.sheet_to_json(ws, { defval: '' });
  if (json.length === 0) {
    showToast('El archivo Excel está vacío o no se pudo leer.', 'error');
    return;
  }
  const headerKeys = Object.keys(json[0]);
  columns = headerKeys;
  gridData = json.map(row => ({ ...row }));
  renderHeader(columns);
  renderGrid();
});
  const file = ev.target.files[0];
  if (!file) return;
  const data = await file.arrayBuffer();
  const workbook = XLSX.read(data, { type: 'array' });
  const firstSheetName = workbook.SheetNames[0];
  const ws = workbook.Sheets[firstSheetName];
  const json = XLSX.utils.sheet_to_json(ws, { defval: '' });
  if (json.length === 0) {
    alert('El archivo Excel está vacío o no se pudo leer.');
    return;
  }
  // Detectar columnas, buscar la de factura (SFANUMFAC o FACTURA)
  const headerKeys = Object.keys(json[0]);
  const facturaCol = headerKeys.find(k => /SFANUMFAC|FACTURA/i.test(k));
  columns = headerKeys; // usamos todas para la grilla
  gridData = json.map(row => ({ ...row }));
  renderHeader(columns);
  renderGrid();
});

// ---------------------------------------------------------------------
// CARGA DE SOPORTES (carpeta)
// ---------------------------------------------------------------------
// CARGA DE SOPORTES (carpeta)
document.getElementById('support-folder').addEventListener('change', ev => {
  if (!requireAuth()) return;
  const files = ev.target.files;
  supportFilesMap.clear();
  for (const file of files) {
    const match = file.name.match(/(\d{6,})/);
    if (match) {
      const invoice = match[1];
      if (!supportFilesMap.has(invoice)) supportFilesMap.set(invoice, []);
      supportFilesMap.get(invoice).push(file);
    }
  }
});
  const files = ev.target.files;
  supportFilesMap.clear();
  for (const file of files) {
    // El nombre del PDF suele contener el número de factura. Extraemos los dígitos iniciales.
    const match = file.name.match(/(\d{6,})/);
    if (match) {
      const invoice = match[1];
      if (!supportFilesMap.has(invoice)) supportFilesMap.set(invoice, []);
      supportFilesMap.get(invoice).push(file);
    }
  }
});

// ---------------------------------------------------------------------
// MOTOR DE AUDITORÍA (replica la lógica de Python)
// ---------------------------------------------------------------------
function auditarFila(fila) {
  // Obtención del número de factura
  const facturaKey = columns.find(k => /SFANUMFAC|FACTURA/i.test(k));
  const numFactura = String(fila[facturaKey] || '').trim();
  if (!numFactura) {
    fila['RESULTADO_AUDITORIA'] = 'FALTA FACTURA';
    return;
  }
  const soportes = supportFilesMap.get(numFactura) || [];
  // Regla genérica: al menos 3 PDFs y 1 XML para ADRES, 3 PDFs para COLSANITAS, etc.
  // Para simplificar, usamos los criterios que el usuario indicó.
  // Aquí aplicamos la regla más restrictiva (ADRES) como ejemplo.
  const tieneXML = soportes.some(f => f.name.toLowerCase().endsWith('.xml'));
  const pdfCount = soportes.filter(f => f.name.toLowerCase().endsWith('.pdf')).length;
  if (pdfCount >= 3 && tieneXML) {
    fila['RESULTADO_AUDITORIA'] = 'SIN RADICAR (+XML)';
  } else if (pdfCount >= 3) {
    fila['RESULTADO_AUDITORIA'] = 'SIN RADICAR';
  } else {
    fila['RESULTADO_AUDITORIA'] = 'FALTA SOPORTE';
  }
}

// Auditar todas las filas
document.getElementById('audit-all').addEventListener('click', () => {
  if (!requireAuth()) return;
  if (gridData.length === 0) {
    showToast('Primero cargue un archivo Excel.', 'error');
    return;
  }
  gridData.forEach(auditarFila);
  if (!columns.includes('RESULTADO_AUDITORIA')) {
    columns.push('RESULTADO_AUDITORIA');
  }
  renderHeader(columns);
  renderGrid();
  showToast('Auditoría completada.', 'success');
});
  if (gridData.length === 0) {
    alert('Primero cargue un archivo Excel.');
    return;
  }
  gridData.forEach(auditarFila);
  // Añadir columna de resultado si no estaba
  if (!columns.includes('RESULTADO_AUDITORIA')) {
    columns.push('RESULTADO_AUDITORIA');
  }
  renderHeader(columns);
  renderGrid();
});

// ---------------------------------------------------------------------
// FILTRADO EN TIEMPO REAL
// ---------------------------------------------------------------------
function onFilterChange(e) {
  const col = e.target.dataset.col;
  const term = e.target.value.toLowerCase();
  // Filtrar los datos originales sin mutar gridData original
  const filtered = gridData.filter(row => {
    const cell = String(row[col] ?? '').toLowerCase();
    return cell.includes(term);
  });
  // Renderizamos solo el subconjunto
  const tbody = document.querySelector('#data-grid tbody');
  tbody.innerHTML = '';
  filtered.forEach((row, i) => {
    const tr = document.createElement('tr');
    columns.forEach(colKey => {
      const td = createEditableCell(row[colKey] ?? '', i, colKey);
      tr.appendChild(td);
    });
    const actionTd = document.createElement('td');
    const delBtn = document.createElement('button');
    delBtn.textContent = '🗑️';
    delBtn.dataset.row = i;
    delBtn.addEventListener('click', onDeleteRow);
    actionTd.appendChild(delBtn);
    tr.appendChild(actionTd);
    tbody.appendChild(tr);
  });
}

// ---------------------------------------------------------------------
// CRUD – Añadir / Eliminar fila
// ---------------------------------------------------------------------
function onDeleteRow(e) {
  const rowIndex = Number(e.target.dataset.row);
  // eliminar del array original
  gridData.splice(rowIndex, 1);
  renderHeader(columns);
  renderGrid();
}

// Botón para añadir fila vacía (se coloca en la barra de herramientas)
const addRowBtn = document.createElement('button');
addRowBtn.textContent = 'Añadir fila';
addRowBtn.id = 'add-row';
addRowBtn.addEventListener('click', () => {
  const empty = {};
  columns.forEach(col => (empty[col] = ''));
  gridData.push(empty);
  renderHeader(columns);
  renderGrid();
});

document.querySelector('.toolbar').appendChild(addRowBtn);

// ---------------------------------------------------------------------
// EXPORTAR A EXCEL CON ESTILOS (ExcelJS)
// ---------------------------------------------------------------------
// Exportar a Excel
document.getElementById('download-excel').addEventListener('click', async () => {
  if (!requireAuth()) return;
  if (gridData.length === 0) {
    showToast('No hay datos para exportar.', 'error');
    return;
  }
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Auditoria');
  sheet.addRow(columns);
  // Estilos de encabezado
  const headerRow = sheet.getRow(1);
  headerRow.eachCell(cell => {
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF334155' } };
    cell.font = { color: { argb: 'FFFFFFFF' }, bold: true };
  });
  // Filas de datos con colores
  gridData.forEach(row => {
    const values = columns.map(col => row[col]);
    const excelRow = sheet.addRow(values);
    const resultIdx = columns.indexOf('RESULTADO_AUDITORIA');
    if (resultIdx >= 0) {
      const result = row['RESULTADO_AUDITORIA'];
      if (result === 'SIN RADICAR (+XML)') excelRow.getCell(resultIdx + 1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF10B981' } };
      else if (result === 'SIN RADICAR') excelRow.getCell(resultIdx + 1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEF4444' } };
      else if (result === 'FALTA SOPORTE') excelRow.getCell(resultIdx + 1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF59E0B' } };
    }
  });
  const buf = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'auditoria_resultado.xlsx';
  a.click();
  URL.revokeObjectURL(url);
  showToast('Archivo Excel descargado.', 'success');
});
  if (gridData.length === 0) {
    alert('No hay datos para exportar.');
    return;
  }
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Auditoria');
  // Encabezados
  sheet.addRow(columns);
  // Estilos de encabezado
  const headerRow = sheet.getRow(1);
  headerRow.eachCell(cell => {
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF334155' } };
    cell.font = { color: { argb: 'FFFFFFFF' }, bold: true };
  });
  // Filas de datos
  gridData.forEach(row => {
    const values = columns.map(col => row[col]);
    const excelRow = sheet.addRow(values);
    // Aplicar color según RESULTADO_AUDITORIA
    const resultIdx = columns.indexOf('RESULTADO_AUDITORIA');
    if (resultIdx >= 0) {
      const result = row['RESULTADO_AUDITORIA'];
      if (result === 'SIN RADICAR (+XML)') {
        excelRow.getCell(resultIdx + 1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF10B981' } }; // verde
      } else if (result === 'SIN RADICAR') {
        excelRow.getCell(resultIdx + 1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEF4444' } }; // rojo
      } else if (result === 'FALTA SOPORTE') {
        excelRow.getCell(resultIdx + 1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF59E0B' } }; // naranja
      }
    }
  });
  // Generar blob y descargar
  const buf = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'auditoria_resultado.xlsx';
  a.click();
  URL.revokeObjectURL(url);
});

// ---------------------------------------------------------------------
// EDICIÓN EN CELDA (actualiza gridData)
// ---------------------------------------------------------------------
function onCellEdit(e) {
  const td = e.target;
  const rowIdx = Number(td.dataset.row);
  const colKey = td.dataset.col;
  const newValue = td.textContent.trim();
  gridData[rowIdx][colKey] = newValue;
}

// ---------------------------------------------------------------------
// INICIALIZACIÓN
// ---------------------------------------------------------------------
// Si la página recarga y ya había datos en sessionStorage, podríamos restaurarlos (opcional).
// Por ahora, iniciamos en estado limpio.

