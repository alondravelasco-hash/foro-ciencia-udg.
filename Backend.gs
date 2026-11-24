const HEADERS = [
  'RegistroId',
  'Nombre',
  'Tipo',
  'Organizacion',
  'DiasRegistrados',
  'Conferencia',
  'Dia',
  'Hora',
  'Timestamp'
];

function doGet() {
  return jsonResponse_({ status: 'ok', message: 'Servicio de registro disponible' });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  let locked = false;

  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('No se recibieron datos.');
    }

    const record = JSON.parse(e.postData.contents);
    const type = String(record.tipoRegistro || '').trim().toLowerCase();
    const name = String(record.Nombre || '').trim();
    const recordId = String(record.RegistroId || Utilities.getUuid()).trim();

    if (!name) throw new Error('El QR no contiene un nombre.');
    if (!recordId) throw new Error('No se pudo identificar el registro.');
    if (type !== 'asistencia' && type !== 'actividad') {
      throw new Error('Tipo de registro no válido.');
    }

    lock.waitLock(15000);
    locked = true;

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!spreadsheet) throw new Error('No se encontró la hoja principal.');

    const attendance = getOrCreateSheet_(spreadsheet, 'Registro de asistencia');
    const activities = getOrCreateSheet_(spreadsheet, 'Registro de actividades');
    const attendanceHeaders = getHeaders_(attendance);
    const activityHeaders = getHeaders_(activities);
    const now = new Date();
    const day = type === 'asistencia'
      ? Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : String(record.Dia || '').trim();
    const activity = type === 'asistencia'
      ? 'Asistencia general'
      : String(record.Conferencia || '').trim();

    if (type === 'actividad' && (!day || !activity)) {
      throw new Error('Falta el nombre o la fecha de la actividad.');
    }

    const attendanceRows = getRows_(attendance);
    const attendanceNameColumn = columnIndex_(attendanceHeaders, 'Nombre');

    if (type === 'actividad' &&
        !attendanceRows.some(row => normalize_(row[attendanceNameColumn]) === normalize_(name))) {
      return jsonResponse_({
        status: 'error',
        code: 'attendance_required',
        message: 'Primero registra la asistencia general.'
      });
    }
 
    const target = type === 'asistencia' ? attendance : activities;
    const targetHeaders = type === 'asistencia' ? attendanceHeaders : activityHeaders;
    const targetRows = type === 'asistencia' ? attendanceRows : getRows_(activities);
    const idColumn = columnIndex_(targetHeaders, 'RegistroId');
    const nameColumn = columnIndex_(targetHeaders, 'Nombre');
    const activityColumn = columnIndex_(targetHeaders, 'Conferencia');
    const dayColumn = columnIndex_(targetHeaders, 'Dia');
    const duplicate = targetRows.some(row => {
      if (String(row[idColumn] || '').trim() === recordId) return true;
      if (normalize_(row[nameColumn]) !== normalize_(name)) return false;
      if (type === 'asistencia') {
        return normalize_(row[dayColumn]) === normalize_(day);
      }
      return normalize_(row[activityColumn]) === normalize_(activity) &&
        normalize_(row[dayColumn]) === normalize_(day);
    });

    if (duplicate) {
      return jsonResponse_({ status: 'duplicate', message: 'Este registro ya estaba guardado.' });
    }

    const values = targetHeaders.map(header => {
      switch (headerKey_(header)) {
        case 'registroid': return safeCell_(recordId);
        case 'nombre': return safeCell_(name);
        case 'tipo': return safeCell_(record.Tipo || '');
        case 'organizacion': return safeCell_(record.Organizacion || record.Organización || '');
        case 'diasregistrados':
          return safeCell_(record.DiasRegistrados || record['Dias registrados'] || '');
        case 'conferencia': return safeCell_(activity);
        case 'dia': return safeCell_(day);
        case 'hora':
          return type === 'asistencia'
            ? Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm:ss')
            : safeCell_(record.Hora || '');
        case 'timestamp': return now;
        default: return '';
      }
    });

    target.appendRow(values);
    return jsonResponse_({ status: 'ok', message: 'Registro guardado en Google Drive.' });
  } catch (error) {
    return jsonResponse_({
      status: 'error',
      message: String(error && error.message ? error.message : error)
    });
  } finally {
    if (locked) lock.releaseLock();
  }
}

function setupSheets() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) throw new Error('No se encontró la hoja principal.');
  getOrCreateSheet_(spreadsheet, 'Registro de asistencia');
  getOrCreateSheet_(spreadsheet, 'Registro de actividades');
}

function getOrCreateSheet_(spreadsheet, name) {
  const sheet = spreadsheet.getSheetByName(name) || spreadsheet.insertSheet(name);

  if (sheet.getLastRow() === 0 || !String(sheet.getRange(1, 1).getDisplayValue()).trim()) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    return sheet;
  }

  let headers = getHeaders_(sheet);
  HEADERS.forEach(header => {
    if (!headers.some(existing => headerKey_(existing) === headerKey_(header))) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue(header);
      headers = getHeaders_(sheet);
    }
  });

  return sheet;
}

function getHeaders_(sheet) {
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
}

function getRows_(sheet) {
  const lastRow = sheet.getLastRow();
  return lastRow < 2
    ? []
    : sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getDisplayValues();
}

function columnIndex_(headers, field) {
  const index = headers.findIndex(header => headerKey_(header) === headerKey_(field));
  if (index < 0) throw new Error('Falta la columna requerida: ' + field);
  return index;
}

function headerKey_(value) {
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9]/g, '');
}

function normalize_(value) {
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .trim()
    .replace(/\s+/g, ' ');
}

function safeCell_(value) {
  const text = String(value == null ? '' : value).trim();
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}

function jsonResponse_(value) {
  return ContentService.createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}
