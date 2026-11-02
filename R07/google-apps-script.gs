/**
 * GOOGLE APPS SCRIPT PARA R07 - OLOR FRAGANTE
 * 
 * INSTRUCCIONES DE INSTALACIÓN:
 * 1. Abre tu cuenta de Google Drive / Google Sheets y crea (o abre) una hoja llamada "registros de r07".
 * 2. En el menú superior, ve a: Extensiones > Apps Script.
 * 3. Borra cualquier código que aparezca en el archivo "Código.gs" y pega todo este código.
 * 4. Guarda el proyecto (Ctrl + S o ícono de disquete).
 * 5. Haz clic en el botón azul arriba a la derecha: "Implementar" > "Nueva implementación".
 * 6. Haz clic en el ícono de engranaje junto a "Selecciona el tipo" y elige "Aplicación web".
 * 7. Configura los campos así:
 *    - Descripción: API Registro R07
 *    - Ejecutar como: "Yo" (tu correo de Google)
 *    - Quién tiene acceso: "Cualquier persona" (Importante para que la web pueda enviar datos sin bloqueo CORS)
 * 8. Haz clic en "Implementar", autoriza los permisos con tu cuenta de Google (Avanzado > Ir a proyecto).
 * 9. Copia la "URL de la aplicación web" (termina en /exec) y pégala en la configuración de tu app web R07 (ícono de engranaje arriba a la derecha) o en la constante DEFAULT_SCRIPT_URL de index.html.
 */

const SHEET_HEADERS = ['Fecha_Hora', 'Nombre', 'Tiempo_Destinado_Minutos', 'Dios_que_hablo'];

/**
 * Recibe las peticiones POST desde la aplicación web R07
 */
function doPost(e) {
  try {
    const sheet = getOrCreateSheet();
    ensureHeaders(sheet);

    let payload = {};
    if (e && e.postData && e.postData.contents) {
      payload = JSON.parse(e.postData.contents);
    } else if (e && e.parameter) {
      payload = e.parameter;
    }

    // Fecha y hora automática del registro en zona horaria local del script/hoja
    const now = new Date();
    const timezone = SpreadsheetApp.getActiveSpreadsheet()
      ? SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone()
      : Session.getScriptTimeZone();
    const fechaHora = Utilities.formatDate(now, timezone || 'America/Bogota', 'yyyy-MM-dd HH:mm:ss');

    const nombre = (payload.nombre || '').toString().trim();
    const tiempo = Number(payload.tiempo_destinado || payload.tiempo || 0);
    const diosQueHablo = (payload.Dios_que_hablo || payload.dios_que_hablo || '').toString().trim();

    // Buscar el índice exacto de cada columna por nombre para garantizar que "Dios_que_hablo" siempre caiga en su columna
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const rowData = new Array(headers.length).fill('');

    const idxFecha = findColumnIndex(headers, ['Fecha_Hora', 'Fecha', 'Timestamp']);
    const idxNombre = findColumnIndex(headers, ['Nombre', 'Persona']);
    const idxTiempo = findColumnIndex(headers, ['Tiempo_Destinado_Minutos', 'Tiempo_Destinado', 'Tiempo', 'Minutos']);
    const idxDios = findColumnIndex(headers, ['Dios_que_hablo', 'dios_que_hablo']);

    if (idxFecha !== -1) rowData[idxFecha] = fechaHora;
    if (idxNombre !== -1) rowData[idxNombre] = nombre;
    if (idxTiempo !== -1) rowData[idxTiempo] = tiempo;
    if (idxDios !== -1) rowData[idxDios] = diosQueHablo;

    sheet.appendRow(rowData);

    return ContentService
      .createTextOutput(JSON.stringify({
        status: 'success',
        message: 'Registro R07 guardado exitosamente',
        timestamp: fechaHora
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({
        status: 'error',
        message: error.toString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Permite probar en el navegador que el Web App está activo
 */
function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({
      status: 'ok',
      service: 'R07 - Olor Fragante Google Sheets Connector'
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Obtiene la hoja activa o busca/crea el archivo "registros de r07"
 */
function getOrCreateSheet() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    const files = DriveApp.getFilesByName('registros de r07');
    if (files.hasNext()) {
      ss = SpreadsheetApp.open(files.next());
    } else {
      ss = SpreadsheetApp.create('registros de r07');
    }
  }
  return ss.getSheets()[0];
}

/**
 * Verifica que la primera fila tenga los encabezados requeridos (incluyendo Dios_que_hablo)
 */
function ensureHeaders(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(SHEET_HEADERS);
    const headerRange = sheet.getRange(1, 1, 1, SHEET_HEADERS.length);
    headerRange.setFontWeight('bold');
    sheet.setFrozenRows(1);
    return;
  }

  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const currentHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => h.toString().trim());

  SHEET_HEADERS.forEach(requiredHeader => {
    const exists = currentHeaders.some(h => h.toLowerCase() === requiredHeader.toLowerCase());
    if (!exists) {
      const nextCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, nextCol).setValue(requiredHeader).setFontWeight('bold');
      currentHeaders.push(requiredHeader);
    }
  });
}

function findColumnIndex(headers, candidates) {
  for (let i = 0; i < headers.length; i++) {
    const h = headers[i].toString().trim().toLowerCase();
    for (let j = 0; j < candidates.length; j++) {
      if (h === candidates[j].toLowerCase()) {
        return i;
      }
    }
  }
  return -1;
}
