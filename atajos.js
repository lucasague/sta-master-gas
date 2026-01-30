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
  menu.addItem('Generar factura', 'generarFacturaPDF');
  menu.addItem('Generar factura agrupada', 'generarFacturaAgrupada');
  menu.addItem('Añadir maquilas a existencias', 'anadirMaquilasExistencias');
  menu.addItem('Ocultar hojas auxiliares (_)', 'ocultarHojasAuxiliares');
  menu.addSeparator();
  menu.addItem('__prueba xlsx__', 'exportFacturaDetalleAsXlsx');
  menu.addToUi();
}



// --------------------------------------------------------------------------------------------------------------
// - COMIENZO DE LA PRUEBA
// ------------------------------------------------------------------------------------------a--------------------
function exportFacturaDetalleAsXlsx() {
  // ======================
  // CONFIG (dentro función)
  // ======================
  const EXPORT_FOLDER_ID = "1QKXWvjv6tQVAKuTz2tofidKREjNMRc6Z";
  const DELETE_TEMP_FILES_AFTER = true;

  const SHEET_DETALLE_NAME = "_facturaDetalle";
  const SHEET_DETALLE_GID = 333755606;

  const SHEET_AGRUPADA_NAME = "_facturaAgrupada";
  const SHEET_AGRUPADA_GID = 1412036802;

  // Congelado de fórmulas -> valores (en hoja temporal dentro del original)
  const ROW_BLOCK = 250;
  const COL_BLOCK = 20;

  // PDF
  const PDF_SCALE = 4;          // 1..4 (4 = mejor calidad)
  const PDF_PORTRAIT = true;    // true retrato, false apaisado

  // ======================
  // Helpers (internos)
  // ======================
  const exportSpreadsheetIdAsBlob_ = (spreadsheetId, filename, mime, url) => {
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

  const exportSpreadsheetAsXlsxBlob_ = (spreadsheetId, filename) => {
    const url = `https://docs.google.com/spreadsheets/d/${encodeURIComponent(spreadsheetId)}/export?format=xlsx`;
    return exportSpreadsheetIdAsBlob_(spreadsheetId, filename, "xlsx", url);
  };

  const exportSheetAsPdfBlob_ = (spreadsheetId, gid, filename) => {
    const base = `https://docs.google.com/spreadsheets/d/${encodeURIComponent(spreadsheetId)}/export`;
    const params = [
      "format=pdf",
      `gid=${encodeURIComponent(gid)}`,
      "exportFormat=pdf",
      "size=A4",
      `portrait=${PDF_PORTRAIT ? "true" : "false"}`,
      "fitw=true",
      `scale=${PDF_SCALE}`,
      "sheetnames=false",
      "printtitle=false",
      "pagenumbers=false",
      "gridlines=false",
      "fzr=false",
      "top_margin=0.25",
      "bottom_margin=0.25",
      "left_margin=0.25",
      "right_margin=0.25",
      "horizontal_alignment=CENTER",
      "vertical_alignment=TOP",
    ].join("&");

    const url = `${base}?${params}`;
    return exportSpreadsheetIdAsBlob_(spreadsheetId, filename, "pdf", url);
  };

  // ======================
  // MAIN
  // ======================
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ssId = ss.getId();

  const sourceDetalle = ss.getSheetByName(SHEET_DETALLE_NAME);
  if (!sourceDetalle) throw new Error(`No existe la pestaña "${SHEET_DETALLE_NAME}".`);

  const sourceAgrupada = ss.getSheetByName(SHEET_AGRUPADA_NAME);
  if (!sourceAgrupada) throw new Error(`No existe la pestaña "${SHEET_AGRUPADA_NAME}".`);

  const folder = EXPORT_FOLDER_ID
    ? DriveApp.getFolderById(EXPORT_FOLDER_ID)
    : DriveApp.getRootFolder();

  // Guardar estados de ocultación para restaurarlos al final
  const wasDetalleHidden = sourceDetalle.isSheetHidden();
  const wasAgrupadaHidden = sourceAgrupada.isSheetHidden();

  // 1) Crear hoja temporal (dentro del original) para congelar fórmulas -> valores SIN romper referencias
  const tempSheetName = `${SHEET_DETALLE_NAME}__TEMP_VALORES__${Date.now()}`;
  let tempSheetInOriginal = null;

  // 2) Spreadsheet temporal para exportar SOLO detalle en XLSX
  let tempXlsxSsId = null;

  try {
    // Asegurar que ambas hojas objetivo estén visibles durante el proceso
    // (evita errores tipo "No se pueden eliminar todas las hojas visibles...")
    if (wasDetalleHidden) sourceDetalle.showSheet();
    if (wasAgrupadaHidden) sourceAgrupada.showSheet();

    SpreadsheetApp.flush();

    // Duplicar detalle dentro del original (la hoja creada es visible por defecto)
    tempSheetInOriginal = sourceDetalle.copyTo(ss).setName(tempSheetName);

    // Congelar fórmulas -> valores en TODA la hoja (por bloques)
    SpreadsheetApp.flush();
    const maxRows = tempSheetInOriginal.getMaxRows();
    const maxCols = tempSheetInOriginal.getMaxColumns();

    for (let r = 1; r <= maxRows; r += ROW_BLOCK) {
      const nr = Math.min(ROW_BLOCK, maxRows - r + 1);
      for (let c = 1; c <= maxCols; c += COL_BLOCK) {
        const nc = Math.min(COL_BLOCK, maxCols - c + 1);
        const block = tempSheetInOriginal.getRange(r, c, nr, nc);
        block.copyTo(block, { contentsOnly: true });
      }
    }

    // Eliminar columnas ocultas (derecha->izquierda)
    for (let c = tempSheetInOriginal.getMaxColumns(); c >= 1; c--) {
      if (tempSheetInOriginal.isColumnHiddenByUser(c)) {
        tempSheetInOriginal.deleteColumn(c);
      }
    }

    // 3) Crear Spreadsheet temporal para XLSX y copiar ahí la hoja ya plana
    const tempXlsxSs = SpreadsheetApp.create(`${ss.getName()}__${SHEET_DETALLE_NAME}__TEMP_EXPORT`);
    tempXlsxSsId = tempXlsxSs.getId();
    const defaultSheets = tempXlsxSs.getSheets();

    const copied = tempSheetInOriginal.copyTo(tempXlsxSs).setName(SHEET_DETALLE_NAME);
    tempXlsxSs.setActiveSheet(copied);

    // Asegurar que haya al menos 1 hoja visible antes de borrar las demás (caso extremo)
    copied.showSheet();

    // Borrar hojas por defecto del temp XLSX
    defaultSheets.forEach(sh => {
      if (sh.getSheetId() !== copied.getSheetId()) tempXlsxSs.deleteSheet(sh);
    });

    // 4) Exportar XLSX (detalle)
    const xlsxFilename = `${ss.getName()}__${SHEET_DETALLE_NAME}.xlsx`;
    const xlsxBlob = exportSpreadsheetAsXlsxBlob_(tempXlsxSsId, xlsxFilename);
    const xlsxFile = folder.createFile(xlsxBlob);

    // 5) Exportar PDF (agrupada) desde el spreadsheet original usando gid
    const pdfFilename = `${ss.getName()}__${SHEET_AGRUPADA_NAME}.pdf`;
    const pdfBlob = exportSheetAsPdfBlob_(ssId, SHEET_AGRUPADA_GID, pdfFilename);
    const pdfFile = folder.createFile(pdfBlob);

    Logger.log(`XLSX creado: ${xlsxFile.getName()} (ID: ${xlsxFile.getId()})`);
    Logger.log(`PDF creado:  ${pdfFile.getName()} (ID: ${pdfFile.getId()})`);

  } finally {
    // Limpieza hoja temporal del original
    try {
      if (tempSheetInOriginal) ss.deleteSheet(tempSheetInOriginal);
    } catch (_) {}

    // Restaurar estados de ocultación originales
    try {
      if (wasDetalleHidden) sourceDetalle.hideSheet();
      if (wasAgrupadaHidden) sourceAgrupada.hideSheet();
    } catch (_) {}

    // Limpieza spreadsheet temporal XLSX
    try {
      if (DELETE_TEMP_FILES_AFTER && tempXlsxSsId) DriveApp.getFileById(tempXlsxSsId).setTrashed(true);
    } catch (_) {}
  }
}
// --------------------------------------------------------------------------------------------------------------
// - FIN DE LA PRUEBA
// --------------------------------------------------------------------------------------------------------------



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

/**
 * Función base para generar un PDF (y opcionalmente un XLSX de detalle)
 * de una factura a partir de una plantilla, usando un ID de factura.
 * @param {object} config Configuración de la exportación.
 */
function generarFacturaBase(config) {
  // ID de la carpeta de Drive donde se guardarán los archivos
  const folder = DriveApp.getFolderById("1QKXWvjv6tQVAKuTz2tofidKREjNMRc6Z");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const sheet = ss.getSheetByName(config.hojaPDF);
  if (!sheet) {
    ui.alert("No se encontró la hoja: " + config.hojaPDF);
    return;
  }

  // 1. Pedir ID de factura
  const response = ui.prompt(config.tituloPrompt, config.mensajePrompt, ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;
  const facturaId = response.getResponseText().trim();
  if (!facturaId) return;

  // 2. Escribir ID y esperar a que la hoja recalcule (fundamental para PDF)
  sheet.getRange("A1").setValue(facturaId);
  SpreadsheetApp.flush(); // Fuerza la escritura
  Utilities.sleep(2000); // Espera 2 segundos para asegurar el recálculo
  
  // 3. Establecer formato numérico: punto como miles, coma como decimal.
  try {
     // Formato español: punto como separador de miles y coma como decimal.
     sheet.getRange("B:G").setNumberFormat("#.##0,00"); 
  } catch (e) {
     Logger.log("No se pudo aplicar formato numérico a B:G en " + config.hojaPDF + ": " + e.toString());
  }

  const spreadsheetId = ss.getId();
  const sheetId = sheet.getSheetId();
  const url_base = "https://docs.google.com/spreadsheets/d/" + spreadsheetId + "/export?";

  // 4. Exportar PDF
  const exportUrlPDF = url_base + [
    "format=pdf",
    "size=A4",
    "portrait=true", 
    "fitw=true", 
    "sheetnames=false",
    "printtitle=false",
    "pagenumbers=false",
    "gridlines=false",
    "fzr=FALSE",
    "top_margin=0.2",
    "bottom_margin=0",
    "left_margin=0.6",
    "right_margin=0.6",
    "gid=" + sheetId 
  ].join("&");

  const token = ScriptApp.getOAuthToken();
  try {
    const fetchResponsePDF = UrlFetchApp.fetch(exportUrlPDF, {
      headers: { 'Authorization': 'Bearer ' + token }
    });
    const pdfBlob = fetchResponsePDF.getBlob().setName(config.prefijoNombre + " " + facturaId + ".pdf");
    folder.createFile(pdfBlob);
  } catch (e) {
    ui.alert("Error al generar el PDF: " + e.toString());
    return;
  }
  
  // 5. Exportar XLSX (solo si se indicó hojaXLSX)
  if (config.hojaXLSX) {
    const hojaOriginal = ss.getSheetByName(config.hojaXLSX);
    if (!hojaOriginal) {
      ui.alert("No se encontró la hoja de detalle: " + config.hojaXLSX);
      return;
    }

    // Crear hoja temporal
    const tmpName = "__tmpExport";
    const hojaTmpExistente = ss.getSheetByName(tmpName);
    if (hojaTmpExistente) ss.deleteSheet(hojaTmpExistente); 
    const nuevaHoja = ss.insertSheet(tmpName);

    const ultimaFila = hojaOriginal.getLastRow();
    const ultimaCol = hojaOriginal.getLastColumn();
    
    if (ultimaFila < 1) {
      ss.deleteSheet(nuevaHoja);
      ui.alert("La hoja de detalle está vacía, se omite la exportación a XLSX.");
      return;
    }
    
    const rangoDatos = hojaOriginal.getRange(1, 1, ultimaFila, ultimaCol);

    // Copiar valores y formato
    rangoDatos.copyTo(nuevaHoja.getRange(1, 1), {contentsOnly:true}); 
    rangoDatos.copyTo(nuevaHoja.getRange(1, 1), {formatOnly:true}); 
    
    // Aplicar formato numérico a la hoja temporal para la exportación XLSX
    try {
       nuevaHoja.getRange("B:G").setNumberFormat("#.##0,00"); 
    } catch (e) { /* Ignorar error de formato */ }

    // Copiar anchos de columnas
    for (let col = 1; col <= ultimaCol; col++) {
      const ancho = hojaOriginal.getColumnWidth(col);
      nuevaHoja.setColumnWidth(col, ancho);
    }

    // Quitar columnas ocultas en la original
    for (let col = ultimaCol; col >= 1; col--) {
      if (hojaOriginal.isColumnHiddenByUser(col)) {
        nuevaHoja.deleteColumn(col);
      }
    }
    
    // Quitar notas/comentarios (limpieza)
    nuevaHoja.getDataRange().clearNote();

    // Exportar XLSX
    const exportUrlXLSX = url_base + [
      "format=xlsx",
      "gid=" + nuevaHoja.getSheetId()
    ].join("&");

    try {
      const fetchResponseXLSX = UrlFetchApp.fetch(exportUrlXLSX, {
        headers: { 'Authorization': 'Bearer ' + token }
      });
      const xlsxBlob = fetchResponseXLSX.getBlob().setName(config.prefijoNombre + " " + facturaId + " - Detalle.xlsx");
      folder.createFile(xlsxBlob);
    } catch (e) {
       ui.alert("Error al generar el XLSX: " + e.toString());
    } finally {
       // Siempre eliminar hoja temporal
       ss.deleteSheet(nuevaHoja);
    }
  }

  // 6. Aviso final
  ui.alert(config.alertaFinal.replace("{{facturaId}}", facturaId));
}

/**
 * Genera el PDF de una factura normal usando la plantilla "_factura".
 */
function generarFacturaPDF() {
  generarFacturaBase({
    hojaPDF: "_factura",
    tituloPrompt: "Facturación",
    mensajePrompt: "Pega el ID de la factura (ej. 2025-001):",
    prefijoNombre: "Factura",
    hojaXLSX: null, 
    alertaFinal: "Factura {{facturaId}} exportada en PDF."
  });
}

/**
 * Genera el PDF de una factura agrupada ("_facturaAgrupada")
 * y el detalle en XLSX ("_facturaDetalle").
 */
function generarFacturaAgrupada() {
  generarFacturaBase({
    hojaPDF: "_facturaAgrupada",
    tituloPrompt: "Facturación Agrupada",
    mensajePrompt: "Pega el ID de la factura agrupada (ej. AG-2025-10):",
    prefijoNombre: "Factura Agrupada",
    hojaXLSX: "_facturaDetalle", 
    alertaFinal: "Factura agrupada {{facturaId}} exportada:\n- Factura en PDF\n- Detalle en XLSX"
  });
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
