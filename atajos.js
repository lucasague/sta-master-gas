/**
 * Se ejecuta automáticamente al abrir el Google Sheet.
 * Crea los menús personalizados en la barra de herramientas.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  // ⚡ Menú Navegación
  ui.createMenu('Desarrollo')
    .addItem('Actualizar índice de hojas', 'crearIndice')
    .addItem('Crear matriz de dependencias', 'buildDependencyMatrix')
    .addToUi();

  // 🚀 Menú Atajos
  const menu = ui.createMenu('🚀 Atajos');
  menu.addItem('Generar factura detalle', 'generarFacturaDetalle');
  menu.addItem('Generar factura agrupada', 'generarFacturaAgrupada');
  menu.addSeparator();
  menu.addItem('Añadir maquilas a existencias', 'anadirMaquilasExistencias');
  menu.addSeparator();
  menu.addItem('Ocultar hojas auxiliares (_)', 'ocultarHojasAuxiliares');
  menu.addToUi();
}

function buildDependencyMatrix() {
  const ss = SpreadsheetApp.getActive();

  // Configuración
  const HEADER_ROW = 1;
  const DATA_START_ROW = 2;
  const OUT_SHEET_NAME = 'Matriz';
  const SKIP_PREFIXES = ['_', 'v', 'V']; // ignora hojas que empiecen por "_" o "v" (min/may)

  const t0 = Date.now();
  const toast = (msg) => ss.toast(msg, 'Matriz dependencias', 5);

  toast('Iniciando: inventario de hojas/columnas...');

  const sheetsAll = ss.getSheets();
  const sheets = sheetsAll.filter(sh => {
    const name = sh.getName();
    if (name === OUT_SHEET_NAME) return false;
    return !SKIP_PREFIXES.some(p => name.startsWith(p));
  });

  toast(`Hojas incluidas: ${sheets.length} / ${sheetsAll.length}`);

  // --- 1) Inventario de columnas: id -> metadata
  // id = "Hoja!Encabezado"
  const cols = [];
  const colIndexBySheetCol = new Map(); // key: "Hoja|A" -> global index

  sheets.forEach((sh, si) => {
    if ((si + 1) % 5 === 0) toast(`Inventariando columnas... (${si + 1}/${sheets.length})`);

    const name = sh.getName();
    const lastCol = sh.getLastColumn();
    if (lastCol < 1) return;

    const headers = sh.getRange(HEADER_ROW, 1, 1, lastCol).getValues()[0];

    for (let c = 1; c <= lastCol; c++) {
      const header = (headers[c - 1] || '').toString().trim();
      if (!header) continue;

      const colLetter = columnToLetter(c);
      const id = `${name}!${header}`;

      const idx = cols.length;
      cols.push({
        sheetName: name,
        sheet: sh,
        header,
        col: c,
        colLetter,
        id
      });

      colIndexBySheetCol.set(`${name}|${colLetter}`, idx);
    }
  });

  if (cols.length === 0) {
    throw new Error('No se encontraron columnas con encabezado (fila 1) en hojas incluidas.');
  }

  toast(`Columnas detectadas: ${cols.length}. Extrayendo fórmulas...`);

  // --- 2) Obtener una fórmula representativa por columna
  // Estrategia: desde DATA_START_ROW hacia abajo hasta encontrar la primera celda con fórmula.
  const formulaByColIdx = new Array(cols.length).fill('');

  cols.forEach((meta, idx) => {
    if ((idx + 1) % 50 === 0) toast(`Extrayendo fórmulas... (${idx + 1}/${cols.length})`);

    const sh = meta.sheet;
    const lastRow = sh.getLastRow();
    const maxRow = Math.max(lastRow, DATA_START_ROW);
    const numRows = maxRow - DATA_START_ROW + 1;
    if (numRows <= 0) return;

    const range = sh.getRange(DATA_START_ROW, meta.col, numRows, 1);
    const formulas = range.getFormulas(); // 2D

    for (let r = 0; r < formulas.length; r++) {
      const f = formulas[r][0];
      if (f && f.startsWith('=')) {
        formulaByColIdx[idx] = f;
        break;
      }
    }
  });

  toast('Analizando dependencias...');

  // --- 3) Construir matriz booleana deps[i][j] = i usa j
  const n = cols.length;
  const deps = Array.from({ length: n }, () => Array(n).fill(false));

  // Regex para capturar referencias a columnas por letra:
  // - A:A, $A:$A
  // - A2:A, $A$2:$A, A2:A1000
  // - Hoja!A:A, 'Hoja con espacios'!A:A, etc.
  const colRefRegex = /(?:(?:'([^']+)'|([A-Za-z0-9_]+))!)?\$?([A-Z]{1,3})\s*:\s*\$?([A-Z]{1,3})/g;
  const cellRangeRegex = /(?:(?:'([^']+)'|([A-Za-z0-9_]+))!)?\$?([A-Z]{1,3})\$?\d+\s*:\s*\$?([A-Z]{1,3})\$?\d*/g;
  const singleCellRegex = /(?:(?:'([^']+)'|([A-Za-z0-9_]+))!)?\$?([A-Z]{1,3})\$?\d+/g;

  for (let i = 0; i < n; i++) {
    if ((i + 1) % 50 === 0) toast(`Analizando dependencias... (${i + 1}/${n})`);

    const f = formulaByColIdx[i];
    if (!f) continue;

    const refs = new Set();
    collectRefs(f, colRefRegex, refs);
    collectRefs(f, cellRangeRegex, refs);
    collectRefs(f, singleCellRegex, refs);

    refs.forEach(ref => {
      const targetSheet = ref.sheetName || cols[i].sheetName;
      const key = `${targetSheet}|${ref.colLetter}`;
      const j = colIndexBySheetCol.get(key);
      if (j !== undefined) deps[i][j] = true;
    });

    deps[i][i] = false; // sin auto-dependencia
  }

  toast('Escribiendo matriz...');

  // --- 4) Volcar a hoja "Matriz"
  let out = ss.getSheetByName(OUT_SHEET_NAME);
  if (!out) out = ss.insertSheet(OUT_SHEET_NAME);
  out.clear();

  const labels = cols.map(c => c.id);

  const values = Array.from({ length: n + 1 }, () => Array(n + 1).fill(''));
  values[0][0] = '';

  for (let j = 0; j < n; j++) values[0][j + 1] = labels[j];
  for (let i = 0; i < n; i++) values[i + 1][0] = labels[i];

  for (let i = 0; i < n; i++) {
    for (let j = 0; j < n; j++) {
      values[i + 1][j + 1] = deps[i][j] ? 'X' : '';
    }
  }

  out.getRange(1, 1, n + 1, n + 1).setValues(values);
  out.setFrozenRows(1);
  out.setFrozenColumns(1);

  // Opcional: ancho razonable (autoResize masivo puede ser lento)
  out.setColumnWidth(1, 260);
  const maxAuto = Math.min(n, 30);
  if (maxAuto > 0) out.autoResizeColumns(2, maxAuto);

  const elapsed = ((Date.now() - t0) / 1000).toFixed(1);
  toast(`Listo. Columnas: ${n}. Tiempo: ${elapsed}s`);
}

function collectRefs(formula, regex, outSet) {
  regex.lastIndex = 0;
  let m;
  while ((m = regex.exec(formula)) !== null) {
    const sheetQuoted = m[1];
    const sheetPlain = m[2];
    const colLetter = m[3];
    const sheetName = sheetQuoted || sheetPlain || '';
    if (colLetter) outSet.add({ sheetName, colLetter });
  }
}

function columnToLetter(column) {
  let temp = '';
  let letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}


function collectRefs(formula, regex, outSet) {
  regex.lastIndex = 0;
  let m;
  while ((m = regex.exec(formula)) !== null) {
    const sheetQuoted = m[1];
    const sheetPlain = m[2];
    const colLetter = m[3]; // para singleCellRegex: m[3]; para ranges también
    const sheetName = sheetQuoted || sheetPlain || '';
    if (colLetter) outSet.add({ sheetName, colLetter });
  }
}

function columnToLetter(column) {
  let temp = '';
  let letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}


/**
 * Crea o actualiza una hoja llamada "ÍNDICE" con enlaces
 * internos (HYPERLINK) a todas las demás hojas del libro.
 */
function crearIndice() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const nombreHojaIndice = "ÍNDICE"; // El nombre de la hoja índice
  let hojaIndice = ss.getSheetByName(nombreHojaIndice);

  // 1. Si no existe la hoja índice, la creamos al principio
  if (!hojaIndice) {
    hojaIndice = ss.insertSheet(nombreHojaIndice, 0);
  } else {
    hojaIndice.clear(); // Limpiamos contenido anterior
    hojaIndice.activate();
    ss.moveActiveSheet(1); // La movemos a la posición 1 (después de la primera hoja)
  }

  // 2. Obtener todas las hojas y sus IDs
  const hojas = ss.getSheets();
  const datosLinks = [];

  hojas.forEach(hoja => {
    const nombre = hoja.getName();
    const id = hoja.getSheetId();
    
    // No nos enlazamos a nosotros mismos
    if (nombre !== nombreHojaIndice) {
      // Fórmula mágica: HYPERLINK interno usando el ID de la hoja (#gid=...)
      const formula = `=HYPERLINK("#gid=${id}"; "${nombre}")`;
      datosLinks.push({
        nombre: nombre, // Guardamos el nombre original
        formula: formula
      });
    }
  });

  // 3. Lógica de Ordenación PERSONALIZADA (omitiendo prefijos) 🚀
  datosLinks.sort((a, b) => {
    const nombreA = a.nombre;
    const nombreB = b.nombre;

    /**
     * Devuelve el nombre de la hoja sin los prefijos de omisión.
     * Los prefijos ignorados son: 
     * - Uno o dos guiones bajos al inicio (ej: _b, __c).
     * - 'v' seguido de una mayúscula al inicio (ej: vH, vJ).
     * @param {string} nombre El nombre de la hoja.
     * @returns {string} El nombre limpio y en minúsculas para ordenar.
     */
    const getNombreLimpio = (nombre) => {
      let limpio = nombre;
      
      // 1. Eliminar prefijos de guiones bajos (_ o __)
      limpio = limpio.replace(/^(_{1,2})/, '');
      
      // 2. Eliminar prefijo 'v' seguido de mayúscula (ej: vH)
      // El patrón debe aplicarse al inicio de la cadena limpia resultante
      limpio = limpio.replace(/^(v[A-Z])/, (match) => match.substring(1));

      return limpio.toLowerCase();
    };

    const limpioA = getNombreLimpio(nombreA);
    const limpioB = getNombreLimpio(nombreB);

    // Prioridad 1: Ordenar por el nombre limpio
    const comparacionLimpia = limpioA.localeCompare(limpioB);
    if (comparacionLimpia !== 0) {
      return comparacionLimpia;
    }

    // Prioridad 2: Si el nombre limpio es el mismo (ej. 'a' y '__a'), usar el nombre original para estabilidad
    return nombreA.localeCompare(nombreB);
  });
  
  // Convertir el resultado ordenado de nuevo a la matriz de fórmulas 2D
  const formulasOrdenadas = datosLinks.map(item => [item.formula]);

  // 4. Escribir los enlaces en la hoja
  if (formulasOrdenadas.length > 0) {
    // Título
    hojaIndice.getRange("A1").setValue("ÍNDICE ALFABÉTICO");
    hojaIndice.getRange("A1").setFontSize(14).setFontWeight("bold").setBackground("#efefef");
    
    // Pegar fórmulas
    const rango = hojaIndice.getRange(2, 1, formulasOrdenadas.length, 1);
    rango.setFormulas(formulasOrdenadas);
    
    // Formato visual de lista limpia
    rango.setFontColor("#1155cc"); // Color azul enlace clásico
    
    // Ajustar ancho
    hojaIndice.setColumnWidth(1, 400);
    
    // Ocultar líneas de cuadrícula para que parezca una app
    hojaIndice.setHiddenGridlines(true);
  }
}

// -----------------------
// FUNCIONES AUXILIARES
// -----------------------

/**
 * Oculta hojas cuyo nombre comience con uno o dos guiones bajos (_ o __).
 */
function ocultarHojasAuxiliares() {
  var ss = SpreadsheetApp.getActive();
  var hojas = ss.getSheets();
  var ocultadas = 0;
  
  // Expresión regular: /^_{1,2}/.test(nombre) busca nombres que empiecen con _ o __
  hojas.forEach(function(hoja) {
    if (hoja.isSheetHidden()) return; 
    var nombre = hoja.getName();
    if (/^_{1,2}/.test(nombre)) {
      hoja.hideSheet();
      ocultadas++;
    }
  });
  
  SpreadsheetApp.getUi().alert('Se ocultaron ' + ocultadas + ' hojas auxiliares.');
}

// ============================
// FACTURAS (Mantiene el formato de miles con punto y decimal con coma)
// ============================
function generarFacturaAgrupada() {
  const CONFIG = {
    EXPORT_FOLDER_ID: "1QKXWvjv6tQVAKuTz2tofidKREjNMRc6Z",

    // GIDs
    SHEET_FACTURAS_GID: 1019856549,     // Facturas (selección de IDs, col A)
    SHEET_DETALLE_GID: 333755606,       // _facturaDetalle (xlsx, valores)
    SHEET_AGRUPADA_GID: 1412036802,     // _facturaAgrupada (pdf)

    // Antes de exportar: escribir ID en A1 de esta hoja
    TARGET_A1_GID: 1412036802,          // _facturaAgrupada

    // Congelado fórmulas->valores (toda la hoja)
    ROW_BLOCK: 250,
    COL_BLOCK: 20,

    // PDF
    PDF_SCALE: 4,
    PDF_PORTRAIT: true,

    // XLSX
    DELETE_TEMP_FILES_AFTER: true,
  };

  ejecutarExport_(CONFIG, { doXlsx: true, doPdf: true, pdfGid: CONFIG.SHEET_AGRUPADA_GID });
}

function generarFacturaDetalle() {
  const CONFIG = {
    EXPORT_FOLDER_ID: "1QKXWvjv6tQVAKuTz2tofidKREjNMRc6Z",

    // GIDs
    SHEET_FACTURAS_GID: 1019856549,     // Facturas (selección de IDs, col A)
    SHEET_FACTURA_GID: 649064952,       // _factura (pdf)

    // Antes de exportar: escribir ID en A1 de esta hoja
    TARGET_A1_GID: 649064952,           // _factura

    // PDF
    PDF_SCALE: 4,
    PDF_PORTRAIT: true,

    // XLSX (no se usa aquí, pero se deja por compatibilidad del core)
    DELETE_TEMP_FILES_AFTER: true,
  };

  ejecutarExport_(CONFIG, { doXlsx: false, doPdf: true, pdfGid: CONFIG.SHEET_FACTURA_GID });
}

function ejecutarExport_(CONFIG, options) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ssId = ss.getId();
  const ui = SpreadsheetApp.getUi();

  // ===== helpers =====
  const getSheetByGid_ = (gid) => {
    if (gid == null) return null;
    for (const sh of ss.getSheets()) {
      if (sh.getSheetId() === gid) return sh;
    }
    return null;
  };

  const exportAsBlob_ = (filename, mime, url) => {
    const token = ScriptApp.getOAuthToken();
    const resp = UrlFetchApp.fetch(url, {
      method: "get",
      headers: { Authorization: `Bearer ${token}` },
      muteHttpExceptions: true,
    });
    const code = resp.getResponseCode();
    if (code !== 200) {
      const body = resp.getContentText();
      throw new Error(`Error exportando (${mime}) HTTP ${code}. Respuesta: ${body.slice(0, 800)}`);
    }
    return resp.getBlob().setName(filename);
  };

  const exportSheetPdf_ = (gid, filename) => {
    const base = `https://docs.google.com/spreadsheets/d/${encodeURIComponent(ssId)}/export`;
    const params = [
      "format=pdf",
      `gid=${encodeURIComponent(gid)}`,
      "exportFormat=pdf",
      "size=A4",
      `portrait=${CONFIG.PDF_PORTRAIT ? "true" : "false"}`,
      "fitw=true",
      `scale=${CONFIG.PDF_SCALE}`,
      "sheetnames=false",
      "printtitle=false",
      "pagenumbers=false",
      "gridlines=false",
      "fzr=false",
    ].join("&");
    return exportAsBlob_(filename, "pdf", `${base}?${params}`);
  };

  const exportSpreadsheetXlsx_ = (spreadsheetId, filename) => {
    const url = `https://docs.google.com/spreadsheets/d/${encodeURIComponent(spreadsheetId)}/export?format=xlsx`;
    return exportAsBlob_(filename, "xlsx", url);
  };

  const folder = CONFIG.EXPORT_FOLDER_ID
    ? DriveApp.getFolderById(CONFIG.EXPORT_FOLDER_ID)
    : DriveApp.getRootFolder();

  // ===== validar selección =====
  const facturasSheet = getSheetByGid_(CONFIG.SHEET_FACTURAS_GID);
  if (!facturasSheet) throw new Error("No existe la hoja Facturas (por GID).");

  const activeRange = ss.getActiveRange();
  const activeSheet = activeRange.getSheet();

  const startRow = activeRange.getRow();
  const startCol = activeRange.getColumn();
  const numCols = activeRange.getNumColumns();

  const ok =
    activeSheet.getSheetId() === CONFIG.SHEET_FACTURAS_GID &&
    startCol === 1 &&
    numCols === 1 &&
    startRow >= 2; // excluye A1 (y cualquier rango que empiece en fila 1)

  if (!ok) {
    ui.alert(
      "Selección inválida",
      "Debes seleccionar una o varias celdas de Facturas!A2:A (columna A, desde la fila 2). No se permite incluir A1 ni seleccionar otras columnas.",
      ui.ButtonSet.OK
    );
    return;
  }

  // IDs únicos no vacíos (en orden)
  const raw = activeRange.getValues().flat();
  const ids = [];
  const seen = new Set();
  for (const v of raw) {
    const s = String(v == null ? "" : v).trim();
    if (!s) continue;
    if (seen.has(s)) continue;
    seen.add(s);
    ids.push(s);
  }

  if (ids.length === 0) {
    ui.alert("No hay IDs", "La selección no contiene ningún ID válido.", ui.ButtonSet.OK);
    return;
  }

  if (ids.length > 1) {
    const resp = ui.alert(
      "Confirmación",
      `Has seleccionado ${ids.length} IDs.\n¿Quieres exportar TODOS los IDs seleccionados?`,
      ui.ButtonSet.YES_NO
    );
    if (resp !== ui.Button.YES) return;
  }

  // ===== preparar hojas implicadas =====
  const detalleSheet = options.doXlsx ? getSheetByGid_(CONFIG.SHEET_DETALLE_GID) : null;
  if (options.doXlsx && !detalleSheet) throw new Error("No existe la hoja _facturaDetalle (por GID).");

  const pdfSheet = getSheetByGid_(options.pdfGid);
  if (!pdfSheet) throw new Error("No existe la hoja del PDF (por GID).");

  const targetA1Sheet = getSheetByGid_(CONFIG.TARGET_A1_GID);
  if (!targetA1Sheet) throw new Error("No existe la hoja destino para escribir A1 (por GID).");

  // Guardar/forzar visibilidad
  const toRestore = [];
  const ensureVisible_ = (sh) => {
    const wasHidden = sh.isSheetHidden();
    toRestore.push([sh, wasHidden]);
    if (wasHidden) sh.showSheet();
  };

  // Temporales XLSX a limpiar
  const tempXlsxIds = [];

  try {
    ensureVisible_(pdfSheet);
    ensureVisible_(targetA1Sheet);
    if (detalleSheet) ensureVisible_(detalleSheet);

    SpreadsheetApp.flush();

    for (let i = 0; i < ids.length; i++) {
      const id = ids[i];

      // 1) Escribir A1 (sin formato)
      targetA1Sheet.getRange("A1").setValue(id);
      SpreadsheetApp.flush();

      // 2) PDF (nombre = ID)
      if (options.doPdf) {
        const pdfBlob = exportSheetPdf_(options.pdfGid, `${id}.pdf`);
        folder.createFile(pdfBlob);
      }

      // 3) XLSX (nombre = ID) con valores y sin columnas ocultas
      if (options.doXlsx) {
        // Hoja temporal dentro del original (mantiene referencias)
        const tempSheetName = `__TEMP_VALORES__${Date.now()}__${i + 1}`;
        const tempSheet = detalleSheet.copyTo(ss).setName(tempSheetName);

        try {
          SpreadsheetApp.flush();

          // Congelar fórmulas->valores en TODA la hoja por bloques
          const maxRows = tempSheet.getMaxRows();
          const maxCols = tempSheet.getMaxColumns();

          for (let r = 1; r <= maxRows; r += CONFIG.ROW_BLOCK) {
            const nr = Math.min(CONFIG.ROW_BLOCK, maxRows - r + 1);
            for (let c = 1; c <= maxCols; c += CONFIG.COL_BLOCK) {
              const nc = Math.min(CONFIG.COL_BLOCK, maxCols - c + 1);
              const block = tempSheet.getRange(r, c, nr, nc);
              block.copyTo(block, { contentsOnly: true });
            }
          }

          // Eliminar columnas ocultas
          for (let c = tempSheet.getMaxColumns(); c >= 1; c--) {
            if (tempSheet.isColumnHiddenByUser(c)) tempSheet.deleteColumn(c);
          }

          // Spreadsheet temporal para export
          const tempXlsxSs = SpreadsheetApp.create(`__TEMP_EXPORT__${Date.now()}__${i + 1}`);
          const tempXlsxId = tempXlsxSs.getId();
          tempXlsxIds.push(tempXlsxId);

          const defaultSheets = tempXlsxSs.getSheets();
          const copied = tempSheet.copyTo(tempXlsxSs).setName("Sheet1");
          copied.showSheet();
          defaultSheets.forEach(sh => {
            if (sh.getSheetId() !== copied.getSheetId()) tempXlsxSs.deleteSheet(sh);
          });

          const xlsxBlob = exportSpreadsheetXlsx_(tempXlsxId, `${id}.xlsx`);
          folder.createFile(xlsxBlob);

        } finally {
          // borrar hoja temporal del original
          try { ss.deleteSheet(tempSheet); } catch (_) {}
        }
      }
    }
  } finally {
    // restaurar ocultación
    try {
      toRestore.forEach(([sh, wasHidden]) => {
        try { if (wasHidden) sh.hideSheet(); } catch (_) {}
      });
    } catch (_) {}

    // borrar temporales xlsx
    try {
      if (CONFIG.DELETE_TEMP_FILES_AFTER) {
        tempXlsxIds.forEach(id => {
          try { DriveApp.getFileById(id).setTrashed(true); } catch (_) {}
        });
      }
    } catch (_) {}
  }
}

//-----------------------
// MAQUILAS
//-----------------------

/**
 * Mueve los registros de la hoja 'Maquilas' que aún no han sido
 * procesados a la primera fila libre en la hoja 'Existencias',
 * basándose en la coincidencia de encabezados.
 */
function anadirMaquilasExistencias() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const hojaMaq = ss.getSheetByName('Maquilas');
  const hojaEx = ss.getSheetByName('Existencias');

  if (!hojaMaq || !hojaEx) {
    ui.alert("Error: Asegúrate de que existen las hojas 'Maquilas' y 'Existencias'.");
    return;
  }

  // --- Leer encabezados ---
  const datosMaq = hojaMaq.getDataRange().getValues();
  if (datosMaq.length < 1) {
    ui.alert('La hoja Maquilas está vacía.');
    return;
  }
  const headersMaq = datosMaq[0];
  const headersEx  = hojaEx.getRange(1, 1, 1, hojaEx.getLastColumn()).getValues()[0];

  // Identificar columna de control en Maquilas
  const nombreColCheck = 'Añadida';
  const colCheckMaq = headersMaq.indexOf(nombreColCheck);
  if (colCheckMaq === -1) {
    ui.alert(`No se encontró la columna "${nombreColCheck}" en Maquilas.`);
    return;
  }
  const colInicioMaq = colCheckMaq + 1; 
  const colID_Maq = 0; 

  // --- Filtrar filas pendientes en Maquilas ---
  const filasMaq = datosMaq.slice(1); 
  const pendientes = [];
  const pendientesRowIdx = []; 

  for (let i = 0; i < filasMaq.length; i++) {
    const fila = filasMaq[i];
    const tieneID = fila[colID_Maq] !== '' && fila[colID_Maq] != null && String(fila[colID_Maq]).trim() !== '';
    const yaAñadida = fila[colCheckMaq] === true; 
    
    if (tieneID && !yaAñadida) {
      pendientes.push(fila);
      pendientesRowIdx.push(i + 2); 
    }
  }

  if (pendientes.length === 0) {
    ui.alert('No hay maquilas pendientes de añadir.');
    return;
  }

  // --- Mapeo de Columnas (Maquilas -> Existencias) ---
  const titulosCopiar = headersMaq.slice(colInicioMaq);
  const columnasMap = titulosCopiar.map((t, j) => ({
    titulo: t,
    idxMaq: colInicioMaq + j,
    idxEx: headersEx.indexOf(t)
  })).filter(m => m.idxEx !== -1); 

  if (columnasMap.length === 0) {
    ui.alert('Ninguna de las columnas a la derecha de "Añadida" en Maquilas existe en Existencias.');
    return;
  }

  // --- Detección de la primera fila libre en Existencias ---
  const idxExAncla = columnasMap[0].idxEx;
  const lastRowEx = hojaEx.getLastRow();
  
  const colAnclaValores = hojaEx.getRange(2, idxExAncla + 1, Math.max(lastRowEx - 1, 0), 1).getValues(); 
  let relPrimeraLibre = colAnclaValores.findIndex(r => r[0] === '' || r[0] == null);
  
  if (relPrimeraLibre === -1) {
    relPrimeraLibre = Math.max(lastRowEx - 1, 0); 
  }
  
  const filaPrimeraLibre = 2 + relPrimeraLibre; 
  
  // --- Pegar por columnas ---
  for (const m of columnasMap) {
    const colValues = pendientes.map(fila => [ fila[m.idxMaq] ]); 
    hojaEx
      .getRange(filaPrimeraLibre, m.idxEx + 1, pendientes.length, 1)
      .setValues(colValues);
  }

  // --- Marcar "Añadida" = true en Maquilas ---
  const rangoChecks = hojaMaq.getRange(2, colCheckMaq + 1, filasMaq.length, 1);
  const checks = rangoChecks.getValues(); 

  for (const r of pendientesRowIdx) {
    const rel = r - 2; 
    if (rel >= 0 && rel < checks.length) {
      checks[rel][0] = true;
    }
  }
  rangoChecks.setValues(checks); 

  // --- Aviso y columnas omitidas ---
  const omitidas = titulosCopiar.filter(t => headersEx.indexOf(t) === -1);
  let msg = `✅ Se añadieron ${pendientes.length} maquilas a Existencias.`;
  if (omitidas.length > 0) {
    msg += `\n\n(Omitidas por no existir en la hoja 'Existencias': ${omitidas.join(', ')})`;
  }
  ui.alert(msg);
}
