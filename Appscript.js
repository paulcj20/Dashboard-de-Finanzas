function normalizeKey_(s) {
    return s
      ? s.toString().toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '')
          .replace(/\./g, '')
          .replace(/\s+/g, ' ')
          .trim()
      : '';
}

function onEdit(e) {
    const sheet = e.range.getSheet();
    const row = e.range.getRow();
    if (row < 2) return;

    Logger.log(`=== onEdit INICIO: fila=${row}, columna=${e.range.getColumn()} ===`);

    try {
      const headers = sheet.getRange(1, 2, 1, sheet.getLastColumn() - 1)
        .getValues()[0].map(h => h.toString().trim().toLowerCase());

      const headersNorm = headers.map(normalizeKey_);
      const getCol = name => {
        const idx = headersNorm.indexOf(normalizeKey_(name));
        if (idx === -1) throw new Error(`Columna no encontrada: "${name}"`);
        return idx + 2;
      };

      const contenedorCol = getCol("contenedor");
      const finOpCol = getCol("fin op.");
      const enviadoCol = getCol("enviado");
      const carpetaCol = getCol("carpeta");
      const choferCol = getCol("chofer");
      const fleteroCol = getCol("fletero");
      const depositoDevCol = getCol("deposito dev.");
      const observacionesCol = getCol("peones/verif/etc");
      const inicioOpCol = getCol("inicio op.");
      const estadoCol = getCol("estado");
      const infoChoferCol = getCol("infochofer");
      const camionCol = getCol("camion");
      const aranaCol = getCol("araña/chata");
      const duaCol = getCol("dua");
      let tipoOpCol = -1;
      try { tipoOpCol = getCol("tipo de op"); } catch (e) { try { tipoOpCol = getCol("operador"); } catch (e2) {} }

      const lastRow = sheet.getLastRow();
      const data = sheet.getRange(2, 2, lastRow - 1, headers.length).getValues();

      if (tipoOpCol > 0 && e.range.getColumn() === tipoOpCol) {
        const tipoValEdit = (typeof e.value !== 'undefined') ? e.value : sheet.getRange(row, tipoOpCol).getValue();
        const tipoValNorm = normalizeKey_(tipoValEdit);
        if (tipoValNorm === 'carga suelta') {
          const carpetaVal = sheet.getRange(row, carpetaCol).getValue().toString().trim();
          const contActual = sheet.getRange(row, contenedorCol).getValue().toString().trim();
          if (!carpetaVal) {
            Logger.log(`CargaSuelta: Carpeta vacía en fila ${row}; no se puede asignar número.`);
          } else if (contActual) {
            Logger.log(`CargaSuelta: Ya hay un identificador en Contenedor ("${contActual}"), no se reasigna. Row=${row}`);
          } else {
            const opciones = getCargaSueltaOptions_(data, carpetaCol, tipoOpCol, contenedorCol, choferCol, carpetaVal);
            const siguiente = getNextCargaSueltaId_(opciones, carpetaVal);
            if (opciones.length === 0) {
              sheet.getRange(row, contenedorCol).setValue(siguiente);
              Logger.log(`CargaSuelta: No había existentes. Asignado automático: ${siguiente} (Row=${row})`);
            } else {
              sheet.getRange(row, contenedorCol).setNote('Usa menú Carga Suelta > Asignar N° para elegir o crear nuevo');
              Logger.log(`CargaSuelta: Existen opciones (${opciones.length}). Solicitar selección vía menú. Carpeta=${carpetaVal}`);
            }
          }
        }
      }

      if (e.range.getColumn() === carpetaCol && tipoOpCol > 0) {
        const tipoVal = normalizeKey_(sheet.getRange(row, tipoOpCol).getValue());
        if (tipoVal === 'carga suelta') {
          const carpetaVal = sheet.getRange(row, carpetaCol).getValue().toString().trim();
          const contActual = sheet.getRange(row, contenedorCol).getValue().toString().trim();
          if (carpetaVal && !contActual) {
            const opciones = getCargaSueltaOptions_(data, carpetaCol, tipoOpCol, contenedorCol, choferCol, carpetaVal);
            const siguiente = getNextCargaSueltaId_(opciones, carpetaVal);
            if (opciones.length === 0) {
              sheet.getRange(row, contenedorCol).setValue(siguiente);
              Logger.log(`CargaSuelta[CarpetaEdit]: Asignado automático ${siguiente}`);
            } else {
              sheet.getRange(row, contenedorCol).setNote('Usa menú Carga Suelta > Asignar N° para elegir o crear nuevo');
              Logger.log(`CargaSuelta[CarpetaEdit]: Opciones existentes (${opciones.length}). Solicitar selección vía menú.`);
            }
          }
        }
      }

      if (e.range.getColumn() === contenedorCol) {
        let val = e.range.getValue();
        if (typeof val === "string") val = val.trim().replace(/\s+/g, "").toUpperCase();
        if (/^([A-Z]{4})(\d{6})(\d)$/.test(val) && !/^([A-Z]{4})-\d{6}-\d$/.test(val)) {
          e.range.setValue(val.replace(/^([A-Z]{4})(\d{6})(\d)$/, "$1-$2-$3"));
        }
      }

      const columnasClave = [contenedorCol, duaCol];
      if (columnasClave.includes(e.range.getColumn())) {
        const carpetaActual = sheet.getRange(row, carpetaCol).getValue().toString().trim();
        if (!carpetaActual) {
          const valCont = sheet.getRange(row, contenedorCol).getValue().toString().trim().toLowerCase();
          const valDua = sheet.getRange(row, duaCol).getValue().toString().trim().toLowerCase();

          for (let i = 0; i < data.length; i++) {
            const filaData = data[i];
            const filaRow = i + 2;
            if (filaRow === row) continue;
            const compCont = filaData[contenedorCol - 2].toString().trim().toLowerCase();
            const compDua = filaData[duaCol - 2].toString().trim().toLowerCase();
            const carpeta = filaData[carpetaCol - 2];
            if ((valCont && valCont === compCont) || (valDua && valDua === compDua)) {
              if (carpeta) {
                sheet.getRange(row, carpetaCol).setValue(carpeta);
                break;
              }
            }
          }
        }
      }

      if (e.range.getColumn() === carpetaCol) {
        const nuevaCarpeta = e.range.getValue().toString().trim();
        if (nuevaCarpeta) {
          const valCont = sheet.getRange(row, contenedorCol).getValue().toString().trim().toLowerCase();
          const valDua = sheet.getRange(row, duaCol).getValue().toString().trim().toLowerCase();
          for (let i = 0; i < data.length; i++) {
            const filaData = data[i];
            const filaRow = i + 2;
            if (filaRow === row) continue;
            const compCont = filaData[contenedorCol - 2].toString().trim().toLowerCase();
            const compDua = filaData[duaCol - 2].toString().trim().toLowerCase();
            const carpetaFila = filaData[carpetaCol - 2].toString().trim();
            if (!carpetaFila && ((valCont && valCont === compCont) || (valDua && valDua === compDua))) {
              sheet.getRange(filaRow, carpetaCol).setValue(nuevaCarpeta);
            }
          }
        }
      }

      if (e.range.getColumn() === choferCol) {
        const choferVal = sheet.getRange(row, choferCol).getValue().toString().trim().toLowerCase();
        for (let i = 0; i < data.length; i++) {
          const filaData = data[i];
          const filaRow = i + 2;
          if (filaRow === row) continue;
          const infoChoferVal = filaData[infoChoferCol - 2].toString().trim().toLowerCase();
          if (choferVal && choferVal === infoChoferVal) {
            const camionVal = filaData[camionCol - 2];
            const aranaVal = filaData[aranaCol - 2];
            sheet.getRange(row, getCol("matricula 1")).setValue(camionVal);
            sheet.getRange(row, getCol("matricula 2")).setValue(aranaVal);
            break;
          }
        }
      }

      if (e.range.getColumn() === depositoDevCol) {
        const depositoVal = sheet.getRange(row, depositoDevCol).getValue().toString().trim().toLowerCase();
        if (depositoVal === "mayabel") {
          const obsCell = sheet.getRange(row, observacionesCol);
          const currentText = obsCell.getValue().toString();
          if (!currentText.toLowerCase().includes("devolucion")) {
            obsCell.setValue(currentText ? `${currentText} devolucion` : "devolucion");
            Logger.log(`onEdit: Agregado "devolucion" en Peones/Verif/Etc por Mayabel. Row=${row}`);
          }
        }
      }

      const editedCol = e.range.getColumn();
      const columnasRelevantesEstado = [choferCol, fleteroCol, inicioOpCol, finOpCol];
    
      if (columnasRelevantesEstado.includes(editedCol)) {
        const estadoCell = sheet.getRange(row, estadoCol);
        const estadoActual = estadoCell.getValue().toString().trim();
        const estadoActualNorm = normalizeKey_(estadoActual);
        const inicioOpVal = sheet.getRange(row, inicioOpCol).getValue().toString().trim();
        const finOpVal = sheet.getRange(row, finOpCol).getValue();
        const choferVal = sheet.getRange(row, choferCol).getValue().toString().trim();
        const fleteroVal = sheet.getRange(row, fleteroCol).getValue().toString().trim();
      
        Logger.log(`onEdit: Parte 5 - Columna relevante editada (${editedCol}). Verificando Estado...`);
      
        if (!isEstadoCompletado_(estadoActual)) {
          if (finOpVal) {
            estadoCell.setValue("Finalizado");
            Logger.log(`onEdit: Estado auto-> Finalizado por Fin Op. Row=${row}`);
          } else if (inicioOpVal) {
            estadoCell.setValue("En curso");
            Logger.log(`onEdit: Estado auto-> En curso por Inicio Op. Row=${row}`);
          } else if ((choferVal || fleteroVal) && !["asignado", "en curso", "finalizado"].includes(estadoActualNorm)) {
            estadoCell.setValue("Asignado");
            Logger.log(`onEdit: Estado auto-> Asignado por chofer/fletero. Row=${row}`);
          }
        } else {
          Logger.log(`onEdit: Estado es Completado ("${estadoActual}"), no se modifica. Row=${row}`);
        }
      }

      Logger.log(`onEdit: Verificando envío - columna=${e.range.getColumn()}, estadoCol=${estadoCol}, match=${e.range.getColumn() === estadoCol}`);
      if (e.range.getColumn() === estadoCol) {
        const estadoEdit = (typeof e.value !== 'undefined')
          ? e.value.toString().trim()
          : sheet.getRange(row, estadoCol).getValue().toString().trim();
        
        Logger.log(`onEdit: Estado editado="${estadoEdit}" en fila ${row}`);
        
        if (isEstadoCompletado_(estadoEdit)) {
          Logger.log(`onEdit: Estado reconocido como Completo/Completado. Iniciando envío en bloque.`);
          const opNumber = sheet.getRange(row, contenedorCol).getValue().toString().trim();
          if (!opNumber) {
            Logger.log(`onEdit: No hay N° de OP en Contenedor; no se puede enviar en bloque. Row=${row}`);
          } else {
            Logger.log(`onEdit: N° de OP encontrado: "${opNumber}". Buscando todas las filas con este N°...`);
            SpreadsheetApp.flush();
            
            const allRows = sheet.getLastRow();
            const rangeData = sheet.getRange(2, contenedorCol, allRows - 1, 1).getValues();
            const enviadoData = sheet.getRange(2, enviadoCol, allRows - 1, 1).getValues();
            
            const filasParaEnviar = [];
            const opLower = opNumber.toLowerCase();
            
            for (let i = 0; i < rangeData.length; i++) {
              const r = i + 2;
              const contVal = rangeData[i][0].toString().trim();
              if (!contVal) continue;
              if (contVal.toLowerCase() !== opLower) continue;
              const enviadoVal = enviadoData[i][0].toString().trim().toLowerCase();
              if (enviadoVal === 'enviado') {
                Logger.log(`onEdit: Fila ${r} ya marcada Enviado, se omite.`);
                continue;
              }
              filasParaEnviar.push(r);
            }
            
            Logger.log(`onEdit: Encontradas ${filasParaEnviar.length} filas para enviar.`);
            
            let count = 0;
            for (const r of filasParaEnviar) {
              try {
                Logger.log(`onEdit: Enviando fila ${r}...`);
                const res = sendRowToG_(sheet, r, getCol, headers);
                count++;
                Logger.log(`onEdit: ✓ Fila ${r} enviada a G.${sheet.getName()} (nueva fila ${res.filaFinal})`);
              } catch (errSend) {
                Logger.log(`onEdit: ✗ Error enviando fila ${r}: ${errSend.message}`);
                console.error(errSend);
              }
            }
            Logger.log(`onEdit: === RESUMEN: Enviadas ${count} filas para OP="${opNumber}" ===`);
          }
        } else {
          Logger.log(`onEdit: Estado "${estadoEdit}" no es Completo/Completado. No se envía.`);
        }
      }

    } catch (error) {
      Logger.log("Error onEdit: " + error.message);
      Logger.log("Error stack: " + error.stack);
      console.error(error);
    }
    
    Logger.log(`=== onEdit FIN: fila=${row} ===`);
}

function columnToLetter(column) {
    let temp = "";
    let col = column;
    while (col > 0) {
      let rem = (col - 1) % 26;
      temp = String.fromCharCode(65 + rem) + temp;
      col = Math.floor((col - 1) / 26);
    }
    return temp;
}

function resolveGSheet_(ss, nombreHoja) {
    const exact = ss.getSheetByName(`G. ${nombreHoja}`);
    if (exact) return exact;
    const noSpace = ss.getSheetByName(`G.${nombreHoja}`);
    if (noSpace) return noSpace;
    const noDot = ss.getSheetByName(`G ${nombreHoja}`);
    if (noDot) return noDot;
    const all = ss.getSheets();
    const prefix = `G. ${nombreHoja}`.toLowerCase();
    const prefixNoDot = `G ${nombreHoja}`.toLowerCase();
    let candidate = null;
    for (const sh of all) {
      const n = sh.getName().toString().toLowerCase();
      if (n === prefix) return sh;
      if (n === prefixNoDot) return sh;
      if (n.startsWith(prefix)) {
        candidate = candidate || sh;
      }
      if (!candidate && n.startsWith(prefixNoDot)) {
        candidate = sh;
      }
    }
    if (candidate) return candidate;
    const gSheets = all.filter(sh => sh.getName().toString().toLowerCase().startsWith('g. '));
    if (gSheets.length === 1) return gSheets[0];
    const gSheetsNoDot = all.filter(sh => sh.getName().toString().toLowerCase().startsWith('g '));
    if (gSheetsNoDot.length === 1) return gSheetsNoDot[0];
    return null;
}

function formatDateOnly_(val, tz) {
    if (!val) return '';
    try {
      const d = (val instanceof Date) ? val : new Date(val);
      if (isNaN(d)) return '';
      return Utilities.formatDate(d, tz || Session.getScriptTimeZone(), 'yyyy-MM-dd');
    } catch (e) {
      return '';
    }
}

function isEstadoCompletado_(s) {
    const t = normalizeKey_(s);
    return t === 'completado' || t === 'completo' || t === 'completada' || t === 'completa';
}

function getCargaSueltaOptions_(data, carpetaCol, tipoOpCol, contenedorCol, choferCol, carpetaVal) {
    const opts = [];
    const seen = {};
    const carpLower = carpetaVal.toString().trim().toLowerCase();
    for (let i = 0; i < data.length; i++) {
      const fila = data[i];
      const carp = (fila[carpetaCol - 2] || '').toString().trim().toLowerCase();
      if (carp !== carpLower) continue;
      const cont = (fila[contenedorCol - 2] || '').toString().trim();
      if (!cont) continue;
      const contLower = cont.toLowerCase();
      const pref = carpLower + '-';
      if (!contLower.startsWith(pref)) continue;
      const suf = cont.substring(carpetaVal.length + 1);
      if (!/^\d{3}$/.test(suf)) continue;
      if (!seen[cont]) {
        seen[cont] = true;
        const chofer = (fila[choferCol - 2] || '').toString().trim();
        opts.push({ id: cont, chofer: chofer });
      }
    }
    opts.sort((a, b) => {
      const na = parseInt(a.id.substring(carpetaVal.length + 1), 10) || 0;
      const nb = parseInt(b.id.substring(carpetaVal.length + 1), 10) || 0;
      return na - nb;
    });
    return opts;
}

function getNextCargaSueltaId_(opciones, carpetaVal) {
    let maxNum = 0;
    const prefLen = carpetaVal.length + 1;
    const carpLower = carpetaVal.toString().trim().toLowerCase() + '-';
    for (const op of opciones) {
      const idLower = op.id.toString().toLowerCase();
      if (idLower.startsWith(carpLower)) {
        const suf = op.id.substring(prefLen);
        const n = parseInt(suf, 10);
        if (!isNaN(n)) maxNum = Math.max(maxNum, n);
      }
    }
    const next = (maxNum + 1).toString().padStart(3, '0');
    return `${carpetaVal}-${next}`;
}

function showCargaSueltaDialog_(sheet, row, contenedorCol, carpetaVal, opciones, siguienteId) {
    const ui = SpreadsheetApp.getUi();
    const sheetName = sheet.getName();
    const itemsHtml = opciones.map(o => `
      <label class="opt">
        <input type="checkbox" name="opId" value="${htmlEscape_(o.id)}" />
        <span class="id">${htmlEscape_(o.id)}</span>
        <span class="meta">— Chofer: ${htmlEscape_(o.chofer || '-')}</span>
      </label>
    `).join('');
    const html = HtmlService.createHtmlOutput(`
      <html>
      <head>
        <meta charset="UTF-8" />
        <style>
          body { font-family: Arial, sans-serif; padding: 12px; }
          h2 { margin: 0 0 8px 0; font-size: 16px; }
          .desc { font-size: 12px; color: #444; margin-bottom: 10px; }
          .list { max-height: 220px; overflow: auto; border: 1px solid #ddd; padding: 8px; }
          .opt { display: block; margin: 6px 0; }
          .id { font-weight: 600; }
          .meta { color: #666; margin-left: 6px; }
          .actions { margin-top: 12px; display: flex; gap: 8px; }
          button { padding: 6px 12px; }
        </style>
      </head>
      <body>
        <h2>Asignar N° Operativa (Carga Suelta)</h2>
        <div class="desc">Carpeta: <b>${htmlEscape_(carpetaVal)}</b>. Seleccioná un número disponible o creá uno nuevo.</div>
        <div class="list">
          ${itemsHtml || '<div style="color:#777;">No hay números existentes en esta carpeta.</div>'}
          <label class="opt" style="margin-top:8px; border-top:1px dashed #ddd; padding-top:8px;">
            <input type="checkbox" name="opId" value="__new__" />
            <span class="id">Crear nuevo</span>
            <span class="meta">— se asignará ${htmlEscape_(siguienteId)}</span>
          </label>
        </div>
        <div class="actions">
          <button id="ok">Asignar</button>
          <button id="cancel">Cancelar</button>
        </div>
        <script>
          const inputs = Array.from(document.querySelectorAll('input[name="opId"]'));
          inputs.forEach(inp => {
            inp.addEventListener('change', () => {
              if (inp.checked) {
                inputs.forEach(x => { if (x !== inp) x.checked = false; });
              }
            });
          });
          document.getElementById('cancel').addEventListener('click', () => google.script.host.close());
          document.getElementById('ok').addEventListener('click', () => {
            const sel = inputs.find(i => i.checked);
            if (!sel) { alert('Seleccioná una opción.'); return; }
            let chosen = sel.value === '__new__' ? '${htmlEscape_(siguienteId)}' : sel.value;
            google.script.run
              .withSuccessHandler(() => google.script.host.close())
              .assignCargaSueltaId({ sheetName: '${htmlEscape_(sheetName)}', row: ${row}, column: ${contenedorCol}, id: chosen });
          });
        </script>
      </body>
      </html>
    `).setWidth(460).setHeight(420);
    ui.showModalDialog(html, 'Seleccionar N° Operativa');
}

function assignCargaSueltaId(payload) {
    try {
      if (!payload || !payload.sheetName || !payload.row || !payload.column || !payload.id) {
        throw new Error('Parámetros inválidos para assignCargaSueltaId');
      }
      const ss = SpreadsheetApp.getActive();
      const sh = ss.getSheetByName(payload.sheetName);
      if (!sh) throw new Error('Hoja no encontrada: ' + payload.sheetName);
      sh.getRange(payload.row, payload.column).setValue(payload.id);
      Logger.log(`assignCargaSueltaId: Fila=${payload.row}, Col=${payload.column}, ID=${payload.id}`);
    } catch (err) {
      Logger.log('assignCargaSueltaId Error: ' + err.message);
      throw err;
    }
}

function htmlEscape_(s) {
    return (s == null ? '' : String(s))
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
}

function setupInstallableOnEditTrigger() {
    const ss = SpreadsheetApp.getActive();
    ScriptApp.getProjectTriggers().forEach(t => {
      if (t.getHandlerFunction && t.getHandlerFunction() === 'onEdit') {
        ScriptApp.deleteTrigger(t);
      }
    });
    ScriptApp.newTrigger('onEdit').forSpreadsheet(ss).onEdit().create();
    Logger.log('Trigger instalable onEdit creado.');
}

function onOpen() {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('Carga Suelta')
      .addItem('Asignar N° Operativa fila activa', 'openCargaSueltaDialogForActiveRow_')
      .addToUi();
}

function openCargaSueltaDialogForActiveRow_() {
    const sheet = SpreadsheetApp.getActiveSheet();
    const range = sheet.getActiveRange();
    if (!range) { SpreadsheetApp.getUi().alert('Seleccioná una fila.'); return; }
    const row = range.getRow();
    if (row < 2) { SpreadsheetApp.getUi().alert('Seleccioná una fila de datos (>=2).'); return; }
    const headers = sheet.getRange(1, 2, 1, sheet.getLastColumn() - 1).getValues()[0].map(h => h.toString().trim().toLowerCase());
    const headersNorm = headers.map(normalizeKey_);
    const getCol = name => { const idx = headersNorm.indexOf(normalizeKey_(name)); if (idx === -1) throw new Error('Columna no encontrada: '+name); return idx + 2; };
    let tipoOpCol = -1; try { tipoOpCol = getCol('tipo de op'); } catch (e) { try { tipoOpCol = getCol('operador'); } catch (e2) {} }
    const carpetaCol = getCol('carpeta');
    const contenedorCol = getCol('contenedor');
    const choferCol = getCol('chofer');
    if (tipoOpCol <= 0) { SpreadsheetApp.getUi().alert('No existe la columna "Tipo de OP".'); return; }
    const tipoVal = normalizeKey_(sheet.getRange(row, tipoOpCol).getValue());
    if (tipoVal !== 'carga suelta') { SpreadsheetApp.getUi().alert('La fila activa no es Carga Suelta.'); return; }
    const carpetaVal = sheet.getRange(row, carpetaCol).getValue().toString().trim();
    if (!carpetaVal) { SpreadsheetApp.getUi().alert('La Carpeta está vacía.'); return; }
    const contActual = sheet.getRange(row, contenedorCol).getValue().toString().trim();
    const lastRow = sheet.getLastRow();
    const data = sheet.getRange(2, 2, lastRow - 1, headers.length).getValues();
    const opciones = getCargaSueltaOptions_(data, carpetaCol, tipoOpCol, contenedorCol, choferCol, carpetaVal);
    const siguiente = getNextCargaSueltaId_(opciones, carpetaVal);
    
    if (opciones.length === 0) {
      sheet.getRange(row, contenedorCol).setValue(siguiente);
      SpreadsheetApp.getUi().alert(`Asignado automáticamente: ${siguiente}`);
      return;
    }
    
    showCargaSueltaDialog_(sheet, row, contenedorCol, carpetaVal, opciones, siguiente);
}

function sendRowToG_(sheet, row, getCol, headers) {
    const ss = sheet.getParent();
    const nombreHoja = sheet.getName();
    const hojaCostos = resolveGSheet_(ss, nombreHoja);
    if (!hojaCostos) throw new Error('No se encontró hoja de costos G. ' + nombreHoja);

    const contenedorCol = getCol('contenedor');
    const enviadoCol = getCol('enviado');
    const carpetaCol = getCol('carpeta');
    const fleteroCol = getCol('fletero');
    const inicioOpCol = getCol('inicio op.');

    const fecha = sheet.getRange(row, 1).getValue();
    const carpeta = sheet.getRange(row, carpetaCol).getValue();
    const cliente = sheet.getRange(row, 4).getValue();
    const origen = sheet.getRange(row, 5).getValue();
    const destino = sheet.getRange(row, 6).getValue();
    const tarifa = safeGet_(sheet, row, getCol, 'tarifa');
    const moneda = safeGet_(sheet, row, getCol, 'moneda');
    const otrosServ = safeGet_(sheet, row, getCol, 'otros serv.');
    const costoFletero = safeGet_(sheet, row, getCol, 'costo');
    const fleteroVal = sheet.getRange(row, fleteroCol).getValue().toString().trim();
    const esTercerizado = fleteroVal ? 'Si' : 'No';
    const inicioOp = sheet.getRange(row, inicioOpCol).getValue();
    const contenedor = sheet.getRange(row, contenedorCol).getValue();
    const extrProveedores = safeGet_(sheet, row, getCol, 'extr. proveedores');
    const citacionChof = safeGet_(sheet, row, getCol, 'citación');
    const salidaChof = safeGet_(sheet, row, getCol, 'salida');
    let tipoOpVal = ''; try { tipoOpVal = sheet.getRange(row, getCol('tipo de op')).getValue(); } catch (e) { try { tipoOpVal = sheet.getRange(row, getCol('operador')).getValue(); } catch (e2) {} }
    const matriculaVal = safeGet_(sheet, row, getCol, 'matricula 1');
    const detExtrasVal = safeGet_(sheet, row, getCol, 'peones/verif/etc');
    
    let kmOp = '';
    let kmColFound = false;
    try {
      const kmCol = getCol('km promedio');
      Logger.log(`sendRowToG_: Columna 'KM Promedio' encontrada en posición ${kmCol} (fila ${row})`);
      const kmValue = sheet.getRange(row, kmCol).getValue();
      Logger.log(`sendRowToG_: Valor raw leído de KM Promedio: "${kmValue}" (tipo: ${typeof kmValue})`);
      if (kmValue !== '' && kmValue !== null && kmValue !== undefined) {
        const kmNum = parseFloat(kmValue);
        if (!isNaN(kmNum) && kmNum > 0) {
          kmOp = kmNum;
          kmColFound = true;
          Logger.log(`sendRowToG_: ✓ KM válido leído de 'KM Promedio': ${kmOp}`);
        } else {
          Logger.log(`sendRowToG_: ✗ KM en 'KM Promedio' no es válido o es 0: ${kmValue}`);
        }
      } else {
        Logger.log(`sendRowToG_: ✗ KM en 'KM Promedio' está vacío`);
      }
    } catch (e) {
      Logger.log(`sendRowToG_: ✗ No se encontró columna 'KM Promedio': ${e.message}`);
    }
    
    if (!kmColFound) {
      try {
        const kmCol = getCol('km');
        const kmValue = sheet.getRange(row, kmCol).getValue();
        if (kmValue !== '' && kmValue !== null) {
          const kmNum = parseFloat(kmValue);
          if (!isNaN(kmNum) && kmNum > 0) {
            kmOp = kmNum;
            Logger.log(`sendRowToG_: ✓ KM leído de 'KM': ${kmOp}`);
            kmColFound = true;
          }
        }
      } catch (e) {
        Logger.log(`sendRowToG_: Columna 'KM' no encontrada: ${e.message}`);
      }
    }
    
    if (!kmColFound) {
      try {
        const kmCol = getCol('kilometros');
        const kmValue = sheet.getRange(row, kmCol).getValue();
        if (kmValue !== '' && kmValue !== null) {
          const kmNum = parseFloat(kmValue);
          if (!isNaN(kmNum) && kmNum > 0) {
            kmOp = kmNum;
            Logger.log(`sendRowToG_: ✓ KM leído de 'Kilometros': ${kmOp}`);
            kmColFound = true;
          }
        }
      } catch (e) {
        Logger.log(`sendRowToG_: Columna 'Kilometros' no encontrada: ${e.message}`);
      }
    }
    
    Logger.log(`sendRowToG_: === KM FINAL para fila ${row}: "${kmOp}" (tipo: ${typeof kmOp}) ===`);
    
    const litrosCalc = (kmOp && !isNaN(kmOp) && kmOp > 0) ? (parseFloat(kmOp) / 2.8) : '';
    Logger.log(`sendRowToG_: Litros calculados=${litrosCalc} (fila ${row})`);
    const zonaHoraria = Session.getScriptTimeZone();
    let finOpHora = '';
    try {
      const finOpHeaderCol = getCol('hora final');
      const finOpCellVal = sheet.getRange(row, finOpHeaderCol).getValue();
      finOpHora = finOpCellVal ? Utilities.formatDate(new Date(finOpCellVal), zonaHoraria, 'HH:mm') : '';
    } catch (e) {
      try {
        const finOpCol = getCol('fin op.');
        const finOpVal = sheet.getRange(row, finOpCol).getValue();
        finOpHora = finOpVal ? Utilities.formatDate(new Date(finOpVal), zonaHoraria, 'HH:mm') : '';
      } catch (e2) {
        finOpHora = '';
      }
    }

    const headersCostos = hojaCostos.getRange(1, 1, 1, hojaCostos.getLastColumn()).getValues()[0].map(h => h.toString().trim().toLowerCase());
    const headersCostosNorm = headersCostos.map(normalizeKey_);
    const getColCostos = name => headersCostosNorm.indexOf(normalizeKey_(name)) + 1;
    const colFactura = getColCostos('factura');

    const nuevoRegistroArr = Array(headersCostos.length).fill('');
    const setIf = (colName, val) => { const c = getColCostos(colName); if (c > 0) nuevoRegistroArr[c - 1] = val; };
    setIf('carpeta', carpeta);
    setIf('factura', '');
    setIf('fecha de op.', fecha);
    setIf('nro. cont - op.', contenedor);
    setIf('tipo de op.', tipoOpVal);
    setIf('matricula', matriculaVal);
    setIf('cliente', cliente);
    setIf('origen', origen);
    setIf('destino', destino);
    setIf('terciarizado', esTercerizado);
    setIf('venta flete', '');
    setIf('costos extra', otrosServ);
    setIf('det. extras', detExtrasVal);
    setIf('costo fletero', costoFletero);
    setIf('extr. proveedores', extrProveedores);
    setIf('kilometros', kmOp);
    setIf('litros', litrosCalc);
    setIf('costo gasoil', '');
    setIf('citacion chof.', citacionChof);
    setIf('hora inicio', inicioOp);
    setIf('hora final', finOpHora);
    setIf('salida chof.', salidaChof);
    setIf('total de costos', '');
    setIf('total venta', '');
    setIf('margen', '');
    setIf('mcv', '');

    const lastRowCostos = hojaCostos.getLastRow();
    const filaFinal = lastRowCostos + 1;
    hojaCostos.getRange(filaFinal, 1, 1, nuevoRegistroArr.length).setValues([nuevoRegistroArr]);

    const colVentaFlete = getColCostos('venta flete');
    if (colVentaFlete > 0) {
      const monedaNorm = normalizeKey_(moneda);
      if ((monedaNorm === 'pesos' || monedaNorm === 'uyu') && tarifa) {
        hojaCostos.getRange(filaFinal, colVentaFlete).setFormula(`=${tarifa}/$AK$1`);
        Logger.log(`sendRowToG_: Venta Flete con fórmula (Pesos/UYU): =${tarifa}/$AK$1`);
      } else {
        hojaCostos.getRange(filaFinal, colVentaFlete).setValue(tarifa || '');
        Logger.log(`sendRowToG_: Venta Flete con valor directo: ${tarifa}`);
      }
    }

    const colCostoChofer = getColCostos('costo chofer');
    if (colCostoChofer > 0 && esTercerizado === 'No') {
      const costoChoferCalc = calcularCostoChofer_(citacionChof, inicioOp, finOpHora, salidaChof);
      if (costoChoferCalc !== null) {
        hojaCostos.getRange(filaFinal, colCostoChofer).setValue(costoChoferCalc);
        Logger.log(`sendRowToG_: Costo Chofer calculado=${costoChoferCalc} para fila ${row}`);
      }
    }

    const colCostoGasoil = getColCostos('costo gasoil');
    const colLitros = getColCostos('litros');
    if (colCostoGasoil > 0 && colLitros > 0) {
      const letraLitros = columnToLetter(colLitros);
      hojaCostos.getRange(filaFinal, colCostoGasoil).setFormula(`=IFS(${letraLitros}${filaFinal}="","",${letraLitros}${filaFinal}<>"",${letraLitros}${filaFinal}*$AL$1)`);
      Logger.log(`sendRowToG_: Fórmula Costo Gasoil aplicada en fila ${filaFinal}`);
    }

    const colTotalCostos = getColCostos('total de costos');
    if (colTotalCostos > 0) {
      const letraJ = columnToLetter(getColCostos('terciarizado'));
      const letraN = columnToLetter(getColCostos('costo fletero'));
      const letraO = columnToLetter(getColCostos('extr. proveedores'));
      const letraP = columnToLetter(getColCostos('costo chofer'));
      const letraS = columnToLetter(getColCostos('costo gasoil'));
      hojaCostos.getRange(filaFinal, colTotalCostos).setFormula(`=IFS(${letraJ}${filaFinal}="Si",${letraN}${filaFinal}+${letraO}${filaFinal},${letraJ}${filaFinal}="No",${letraP}${filaFinal}+${letraS}${filaFinal}+${letraO}${filaFinal})`);
      Logger.log(`sendRowToG_: Fórmula Total de Costos aplicada en fila ${filaFinal}`);
    }

    const colTotalVenta = getColCostos('total venta');
    if (colTotalVenta > 0) {
      const letraK = columnToLetter(getColCostos('venta flete'));
      const letraL = columnToLetter(getColCostos('costos extra'));
      hojaCostos.getRange(filaFinal, colTotalVenta).setFormula(`=SUM(${letraK}${filaFinal}+${letraL}${filaFinal})`);
      Logger.log(`sendRowToG_: Fórmula Total Venta aplicada en fila ${filaFinal}`);
    }

    const colMCV = getColCostos('mcv');
    if (colMCV > 0) {
      const letraY = columnToLetter(getColCostos('total venta'));
      const letraX = columnToLetter(getColCostos('total de costos'));
      hojaCostos.getRange(filaFinal, colMCV).setFormula(`=${letraY}${filaFinal}-${letraX}${filaFinal}`);
      Logger.log(`sendRowToG_: Fórmula MCV aplicada en fila ${filaFinal}`);
    }

    sheet.getRange(row, enviadoCol).setValue('Enviado');

    return { filaFinal };
}

function safeGet_(sheet, row, getCol, headerName) {
    try { return sheet.getRange(row, getCol(headerName)).getValue(); } catch (e) { return ''; }
}

function calcularCostoChofer_(citacion, horaInicio, horaFinal, salida) {
    try {
      let entrada = citacion || horaInicio;
      let salidaFinal = salida || horaFinal;
      
      Logger.log(`calcularCostoChofer_: entrada raw=${entrada} (tipo=${typeof entrada}), salida raw=${salidaFinal} (tipo=${typeof salidaFinal})`);
      
      if (!entrada || !salidaFinal) {
        Logger.log('calcularCostoChofer_: Falta entrada o salida, no se puede calcular.');
        return null;
      }

      const parseHora = (val) => {
        if (!val) return null;
        if (typeof val === 'number' && val >= 0 && val < 1) {
          return val * 24;
        }
        if (val instanceof Date) {
          return val.getHours() + (val.getMinutes() / 60);
        }
        const str = val.toString().trim();
        if (/^\d{1,2}:\d{2}$/.test(str)) {
          const [h, m] = str.split(':').map(Number);
          return h + (m / 60);
        }
        const d = new Date(val);
        if (!isNaN(d)) {
          return d.getHours() + (d.getMinutes() / 60);
        }
        return null;
      };

      const entradaHoras = parseHora(entrada);
      const salidaHoras = parseHora(salidaFinal);
      
      Logger.log(`calcularCostoChofer_: entradaHoras=${entradaHoras}, salidaHoras=${salidaHoras}`);
      
      if (entradaHoras === null || salidaHoras === null) {
        Logger.log('calcularCostoChofer_: No se pudo parsear entrada/salida.');
        return null;
      }

      let horasTrabajadas = salidaHoras - entradaHoras;
      if (horasTrabajadas < 0) horasTrabajadas += 24;
      if (horasTrabajadas < 0 || horasTrabajadas > 24) {
        Logger.log(`calcularCostoChofer_: Horas inválidas=${horasTrabajadas}`);
        return null;
      }

      Logger.log(`calcularCostoChofer_: Horas trabajadas=${horasTrabajadas.toFixed(2)}`);

      let costo = 0;
      if (horasTrabajadas < 8) {
        costo = horasTrabajadas * 10.25;
      } else {
        const horasBase = 8;
        const horasExtra = horasTrabajadas - 8;
        costo = (horasBase * 10.25) + (horasExtra * 15.04);
      }

      Logger.log(`calcularCostoChofer_: Costo calculado=${costo.toFixed(2)}`);
      return Math.round(costo * 100) / 100;
    } catch (err) {
      Logger.log('calcularCostoChofer_ Error: ' + err.message);
      return null;
    }
}

function actualizarKMEnGastos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaOperativa = ss.getSheets()[0];
  const nombreHoja = hojaOperativa.getName();
  const hojaGastos = resolveGSheet_(ss, "Nov 25");
  
  if (!hojaGastos) {
    Logger.log('ERROR: No se encontró la hoja de gastos G. ' + "Nov 25");
    return;
  }
  
  Logger.log(`=== Iniciando actualización de KM en ${hojaGastos.getName()} ===`);
  
  const headersOp = hojaOperativa.getRange(1, 2, 1, hojaOperativa.getLastColumn() - 1)
    .getValues()[0].map(h => h.toString().trim().toLowerCase());
  const headersOpNorm = headersOp.map(normalizeKey_);
  const getColOp = name => {
    const idx = headersOpNorm.indexOf(normalizeKey_(name));
    if (idx === -1) throw new Error(`Columna operativa no encontrada: "${name}"`);
    return idx + 2;
  };
  
  const headersGastos = hojaGastos.getRange(1, 1, 1, hojaGastos.getLastColumn())
    .getValues()[0].map(h => h.toString().trim().toLowerCase());
  const headersGastosNorm = headersGastos.map(normalizeKey_);
  const getColGastos = name => headersGastosNorm.indexOf(normalizeKey_(name)) + 1;
  
  const colOpCarpeta = getColOp('carpeta');
  const colOpContenedor = getColOp('contenedor');
  const colOpKM = getColOp('km promedio');
  
  const colGastosCarpeta = getColGastos('carpeta');
  const colGastosContenedor = getColGastos('nro. cont - op.');
  const colGastosKM = getColGastos('kilometros');
  const colGastosLitros = getColGastos('litros');
  
  Logger.log(`Columnas operativa: Carpeta=${colOpCarpeta}, Contenedor=${colOpContenedor}, KM=${colOpKM}`);
  Logger.log(`Columnas gastos: Carpeta=${colGastosCarpeta}, Contenedor=${colGastosContenedor}, KM=${colGastosKM}, Litros=${colGastosLitros}`);
  
  const lastRowOp = hojaOperativa.getLastRow();
  const dataOp = hojaOperativa.getRange(2, 2, lastRowOp - 1, hojaOperativa.getLastColumn() - 1).getValues();
  
  const lastRowGastos = hojaGastos.getLastRow();
  const dataGastos = hojaGastos.getRange(2, 1, lastRowGastos - 1, hojaGastos.getLastColumn()).getValues();
  
  let actualizados = 0;
  let sinKM = 0;
  
  for (let i = 0; i < dataGastos.length; i++) {
    const filaGastos = i + 2;
    const carpetaGastos = dataGastos[i][colGastosCarpeta - 1].toString().trim();
    const contenedorGastos = dataGastos[i][colGastosContenedor - 1].toString().trim();
    const kmActual = dataGastos[i][colGastosKM - 1];
    
    if (kmActual && parseFloat(kmActual) > 0) {
      continue;
    }
    
    let kmEncontrado = null;
    for (let j = 0; j < dataOp.length; j++) {
      const carpetaOp = dataOp[j][colOpCarpeta - 2].toString().trim();
      const contenedorOp = dataOp[j][colOpContenedor - 2].toString().trim();
      
      if (carpetaOp === carpetaGastos && contenedorOp === contenedorGastos) {
        const kmOp = dataOp[j][colOpKM - 2];
        if (kmOp && parseFloat(kmOp) > 0) {
          kmEncontrado = parseFloat(kmOp);
          break;
        }
      }
    }
    
    if (kmEncontrado) {
      hojaGastos.getRange(filaGastos, colGastosKM).setValue(kmEncontrado);
      const litros = kmEncontrado / 2.8;
      hojaGastos.getRange(filaGastos, colGastosLitros).setValue(litros);
      
      actualizados++;
      Logger.log(`✓ Fila ${filaGastos}: Carpeta=${carpetaGastos}, Contenedor=${contenedorGastos}, KM=${kmEncontrado}, Litros=${litros.toFixed(2)}`);
    } else {
      sinKM++;
      Logger.log(`✗ Fila ${filaGastos}: No se encontró KM en operativa (Carpeta=${carpetaGastos}, Contenedor=${contenedorGastos})`);
    }
  }
  
  Logger.log(`=== RESUMEN ===`);
  Logger.log(`Filas actualizadas con KM: ${actualizados}`);
  Logger.log(`Filas sin KM en operativa: ${sinKM}`);
  Logger.log(`Total filas procesadas: ${dataGastos.length}`);
  
  SpreadsheetApp.getUi().alert(`Actualización completada:\n\n✓ ${actualizados} filas actualizadas con KM\n✗ ${sinKM} filas sin KM en operativa\n\nTotal: ${dataGastos.length} filas procesadas`);
}

function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    const sheetParam = e.parameter.sheet || 'Oct 25';
    const sheetName = 'G. ' + sheetParam;
    Logger.log(`doGet: Parámetro recibido sheet="${sheetParam}", buscando hoja "${sheetName}"`);
    
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      Logger.log(`doGet: ERROR - Hoja "${sheetName}" no encontrada`);
      const errorData = JSON.stringify({ error: 'Hoja no encontrada: ' + sheetName });
      const callback = e.parameter.callback;
      
      if (callback) {
        return ContentService.createTextOutput(callback + '(' + errorData + ')')
          .setMimeType(ContentService.MimeType.JAVASCRIPT);
      }
      return ContentService.createTextOutput(errorData)
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const rows = data.slice(1);
    
    Logger.log(`doGet: Hoja "${sheetName}" encontrada, devolviendo ${rows.length} filas de datos`);
    
    const json = rows.map(row => {
      let obj = {};
      headers.forEach((header, i) => {
        obj[header] = row[i];
      });
      return obj;
    });
    
    const jsonString = JSON.stringify(json);
    
    const callback = e.parameter.callback;
    if (callback) {
      return ContentService.createTextOutput(callback + '(' + jsonString + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    
    return ContentService.createTextOutput(jsonString)
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    const errorData = JSON.stringify({ error: error.message });
    const callback = e.parameter.callback;
    
    if (callback) {
      return ContentService.createTextOutput(callback + '(' + errorData + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(errorData)
      .setMimeType(ContentService.MimeType.JSON);
  }
}