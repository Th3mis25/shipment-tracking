const SHEET_NAME = 'Shipments';
const HEADERS = [
  'Executive','Trip','Trailer','Reference','Customer','Destination',
  'Status','Segment','TR-MX','TR-USA','Pickup Appointment','Arrival at Pickup',
  'Delivery Appointment','Arrival at Delivery','Comments','Documents','Tracking'
];
const UNIQUE_KEY = 'Trip';

function doGet(e) {
  try {
    const rows = readRows_();
    return buildJsonResponse_(rows);
  } catch (error) {
    return buildJsonResponse_({ success: false, error: error.message }, 500);
  }
}

function doPost(e) {
  let mode = '';
  try {
    const payload = parseBody_(e);
    mode = (payload.mode || '').toLowerCase();
    const record = payload.record || {};

    validateRecord_(record);

    let row = null;
    if (mode === 'add') {
      row = addRow_(record);
    } else if (mode === 'update') {
      row = updateRow_(payload.originalTrip, record);
    } else {
      throw new Error('Modo no soportado. Usa "add" o "update".');
    }

    if (!row) {
      throw new Error('No se pudo escribir el registro en la hoja.');
    }

    const message = mode === 'add'
      ? 'Registro creado correctamente.'
      : 'Registro actualizado correctamente.';

    return buildJsonResponse_({ success: true, mode, row, message });
  } catch (error) {
    const message = error && error.message ? error.message : 'Error desconocido.';
    return buildJsonResponse_({ success: false, mode, row: null, message }, 400);
  }
}

function doOptions() {
  return {
    statusCode: 204,
    headers: corsHeaders_(),
    body: ''
  };
}

function readRows_() {
  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();
  if (!values.length) {
    return [];
  }

  const [headerRow, ...rows] = values;
  const headerIndexes = HEADERS.map((header) => headerRow.indexOf(header));

  return rows
    .filter((row) => row.some((value) => value !== '' && value !== null))
    .map((row) => {
      const obj = {};
      HEADERS.forEach((key, idx) => {
        const columnIndex = headerIndexes[idx];
        const value = columnIndex > -1 ? row[columnIndex] : '';
        obj[key] = value == null ? '' : value.toString();
      });
      return obj;
    });
}

function addRow_(record) {
  const sheet = getSheet_();
  const sanitized = sanitiseRecord_(record);
  const values = HEADERS.map((key) => sanitized[key] || '');
  sheet.appendRow(values);
  const rowValues = sheet.getRange(sheet.getLastRow(), 1, 1, HEADERS.length).getValues()[0];
  return toObject_(rowValues);
}

function updateRow_(originalKey, record) {
  const sheet = getSheet_();
  const keyToFind = (originalKey || record[UNIQUE_KEY] || '').toString().trim();
  if (!keyToFind) {
    throw new Error('Se requiere el Trip original para actualizar.');
  }

  const range = sheet.getDataRange();
  const values = range.getValues();
  if (!values.length) {
    throw new Error('La hoja está vacía.');
  }

  const headerRow = values[0];
  const keyColumnIndex = headerRow.indexOf(UNIQUE_KEY);
  if (keyColumnIndex === -1) {
    throw new Error(`No se encontró la columna ${UNIQUE_KEY} en la hoja.`);
  }

  const sanitized = sanitiseRecord_(record);
  const rowValues = HEADERS.map((key) => sanitized[key] || '');

  for (let rowIndex = 1; rowIndex < values.length; rowIndex++) {
    const currentValue = values[rowIndex][keyColumnIndex];
    if (currentValue != null && currentValue.toString().trim() === keyToFind) {
      sheet.getRange(rowIndex + 1, 1, 1, HEADERS.length).setValues([rowValues]);
      const confirmed = sheet.getRange(rowIndex + 1, 1, 1, HEADERS.length).getValues()[0];
      return toObject_(confirmed);
    }
  }

  throw new Error(`No se encontró el Trip "${keyToFind}" para actualizar.`);
}

function parseBody_(e) {
  if (!e || !e.postData || !e.postData.contents) {
    throw new Error('No se recibió contenido en la petición.');
  }

  const raw = e.postData.contents;

  try {
    return JSON.parse(raw);
  } catch (error) {
    throw new Error('El cuerpo de la petición no es un JSON válido.');
  }
}

function validateRecord_(record) {
  const sanitized = sanitiseRecord_(record);
  if (!sanitized[UNIQUE_KEY]) {
    throw new Error('El campo Trip es obligatorio.');
  }
}

function sanitiseRecord_(record) {
  const result = {};
  HEADERS.forEach((key) => {
    let value = record && record[key];
    if (value === undefined || value === null) {
      value = '';
    }
    result[key] = value.toString();
  });
  return result;
}

function toObject_(rowValues) {
  const obj = {};
  HEADERS.forEach((key, idx) => {
    const value = rowValues[idx];
    obj[key] = value == null ? '' : value.toString();
  });
  return obj;
}

function getSheet_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  if (!sheet) {
    throw new Error(`No existe la hoja "${SHEET_NAME}".`);
  }
  return sheet;
}

function corsHeaders_() {
  return {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Access-Control-Allow-Methods': 'GET,POST,OPTIONS'
  };
}

function buildJsonResponse_(payload, statusCode) {
  return {
    statusCode: statusCode || 200,
    headers: Object.assign({ 'Content-Type': 'application/json' }, corsHeaders_()),
    body: JSON.stringify(payload)
  };
}
