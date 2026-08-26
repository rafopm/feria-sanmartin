const SHEET_NAME = 'Respuestas';
const EVENTS_SHEET_NAME = 'Eventos';
const PROP_KEY = 'LATEST_INGRESO_DISPLAY';
const RESPONSE_HEADERS = ['Fecha', 'ID', 'Nombres', 'Apellidos', 'Parentesco', 'Estudiante', 'Grado', 'Email', 'DNI', 'Estado', 'Hora'];

function doGet(e) {
  const action = e && e.parameter ? e.parameter.action : '';
  if (action === 'getEvents') return json({ events: getEvents() });
  if (action === 'checkEmail') {
    return json(checkEmail(e.parameter.email, e.parameter.eventoId));
  }
  if (action === 'check') return json(checkInvitationById(e.parameter.id));
  if (action === 'latest') return json(getLatestIngreso());

  const page = e && e.parameter ? e.parameter.page || 'Index' : 'Index';
  const file = page.toLowerCase() === 'display' ? 'Display' : page.toLowerCase() === 'scanner' ? 'Scanner' : 'Index';
  try {
    return HtmlService.createHtmlOutputFromFile(file)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    return json({ status: 'OK', message: 'API Feria San Martin activa' });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents || '{}');
    if (body.action === 'register') return json(registerAttendeeAPI(body));
    if (body.action === 'resend') return json(resendInvitation(body.email, body.eventoId));
    if (body.action === 'check') return json(checkInvitationById(body.id));
    return json({ status: 'ERROR', message: 'Accion no reconocida' });
  } catch (err) {
    return json({ status: 'ERROR', message: err.message });
  }
}

function doOptions() {
  return json({});
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getEvents() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(EVENTS_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return [];

  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(normalizeHeader);
  return values.slice(1).map(function (row) {
    return {
      id: cell(row, headers, ['id', 'eventoid']),
      nombre: cell(row, headers, ['nombre', 'evento']),
      descripcion: cell(row, headers, ['descripcion', 'descripción']),
      fecha: formatCell(cell(row, headers, ['fecha'])),
      mensaje: cell(row, headers, ['mensaje', 'mensajebienvenida']),
      activo: isActive(cell(row, headers, ['activo', 'estado']))
    };
  }).filter(function (event) {
    return event.id || event.nombre;
  });
}

function checkEmail(email, eventoId) {
  const normalizedEmail = String(email || '').trim().toLowerCase();
  if (!normalizedEmail) return { exists: false };

  const sheet = getResponsesSheet();
  const values = sheet.getDataRange().getValues();
  for (let r = 1; r < values.length; r++) {
    const rowEmail = String(values[r][7] || '').trim().toLowerCase();
    if (rowEmail === normalizedEmail) {
      return {
        exists: true,
        id: values[r][1],
        nombres: values[r][2],
        apellidos: values[r][3],
        estudiante: values[r][5],
        dni: values[r][8]
      };
    }
  }
  return { exists: false };
}

function checkInvitationById(invitationId) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    const sheet = getResponsesSheet();
    const values = sheet.getDataRange().getValues();
    const searchId = String(invitationId || '').trim();

    for (let r = 1; r < values.length; r++) {
      if (String(values[r][1]).trim() !== searchId) continue;

      const data = attendeeData(values[r], searchId);
      if (String(values[r][9]).trim().toUpperCase() === 'INGRESADO') {
        return { status: 'DUPLICADO', data: data };
      }

      const hora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HH:mm:ss');
      sheet.getRange(r + 1, 10).setValue('INGRESADO');
      sheet.getRange(r + 1, 11).setValue(hora);
      data.horaIngreso = hora;
      data.timestamp = Date.now();
      PropertiesService.getScriptProperties().setProperty(PROP_KEY, JSON.stringify({
        status: 'INGRESADO',
        id: searchId,
        nombres: data.nombres,
        apellidos: data.apellidos,
        parentesco: data.parentesco,
        estudiante: data.estudiante,
        dni: data.dni,
        eventoNombre: data.eventoNombre || '',
        eventoMensaje: data.eventoMensaje || '',
        horaIngreso: hora,
        timestamp: data.timestamp
      }));
      return { status: 'OK', data: data };
    }
    return { status: 'INVALIDO' };
  } finally {
    lock.releaseLock();
  }
}

function getLatestIngreso() {
  const value = PropertiesService.getScriptProperties().getProperty(PROP_KEY);
  if (!value) return { status: 'NONE' };
  const data = JSON.parse(value);
  if (Date.now() - data.timestamp > 20000) return { status: 'NONE' };
  return data;
}

function registerAttendeeAPI(data) {
  const required = ['nombres', 'apellidos', 'parentesco', 'nombreEstudiante', 'gradoSeccion', 'email', 'dni'];
  required.forEach(function (field) {
    if (!String(data[field] || '').trim()) throw new Error('Falta el campo: ' + field);
  });

  const dni = String(data.dni).trim();
  if (!/^\d{8}$/.test(dni)) throw new Error('El DNI debe tener 8 digitos');

  const sheet = getResponsesSheet();
  const existing = checkEmail(data.email, data.eventoId);
  if (existing.exists) return { status: 'CORREO_DUPLICADO', id: existing.id };

  const id = buildInvitationId(sheet);
  sheet.appendRow([
    new Date(), id, data.nombres, data.apellidos, data.parentesco,
    data.nombreEstudiante, data.gradoSeccion, data.email, dni, 'PENDIENTE', ''
  ]);

  const qrUrl = 'https://api.qrserver.com/v1/create-qr-code/?size=250x250&data=' + encodeURIComponent(id);
  sendInvitationEmail(data.email, data.nombres, data.apellidos, data.parentesco,
    data.nombreEstudiante, data.gradoSeccion, dni, id, qrUrl);
  return { status: 'OK', id: id, email: data.email };
}

function resendInvitation(email, eventoId) {
  const normalizedEmail = String(email || '').trim().toLowerCase();
  const sheet = getResponsesSheet();
  const values = sheet.getDataRange().getValues();

  for (let r = 1; r < values.length; r++) {
    if (String(values[r][7] || '').trim().toLowerCase() !== normalizedEmail) continue;
    const id = String(values[r][1]);
    const qrUrl = 'https://api.qrserver.com/v1/create-qr-code/?size=250x250&data=' + encodeURIComponent(id);
    sendInvitationEmail(values[r][7], values[r][2], values[r][3], values[r][4], values[r][5], values[r][6], values[r][8], id, qrUrl);
    return { status: 'OK', id: id };
  }
  return { status: 'ERROR', message: 'No encontramos una entrada con ese correo' };
}

function getResponsesSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('No existe la hoja ' + SHEET_NAME);
  if (sheet.getLastColumn() < RESPONSE_HEADERS.length) {
    sheet.getRange(1, 1, 1, RESPONSE_HEADERS.length).setValues([RESPONSE_HEADERS]);
  }
  return sheet;
}

function attendeeData(row, id) {
  return {
    id: id,
    nombres: row[2],
    apellidos: row[3],
    parentesco: row[4],
    estudiante: row[5],
    grado: row[6],
    dni: row[8]
  };
}

function buildInvitationId(sheet) {
  const last = sheet.getLastRow();
  if (last <= 1) return 'INV-1001';
  const values = sheet.getRange(2, 2, last - 1, 1).getValues();
  let max = 1000;
  values.forEach(function (row) {
    const match = String(row[0]).match(/INV-(\d+)/);
    if (match) max = Math.max(max, parseInt(match[1], 10));
  });
  return 'INV-' + (max + 1);
}

function sendInvitationEmail(email, nombres, apellidos, parentesco, estudiante, grado, dni, id, qrUrl) {
  const safe = [email, nombres, apellidos, parentesco, estudiante, grado, dni, id].map(escapeHtml);
  const htmlBody = '<div style="font-family:Arial">' +
    '<h2>Invitacion ' + safe[7] + '</h2>' +
    '<p>' + safe[1] + ' ' + safe[2] + '</p>' +
    '<p>Estudiante: ' + safe[4] + '</p>' +
    '<p>Grado: ' + safe[5] + '</p>' +
    '<p>DNI: ' + safe[6] + '</p>' +
    '<p>Para ingresar debe presentar este codigo QR y su DNI fisico.</p>' +
    '<img src="' + qrUrl + '"/></div>';
  const blob = UrlFetchApp.fetch(qrUrl).getBlob().setName(id + '.png');
  MailApp.sendEmail({
    to: email,
    subject: 'Invitacion Feria - QR ' + id,
    htmlBody: htmlBody,
    attachments: [blob]
  });
}

function normalizeHeader(value) {
  return String(value || '').toLowerCase().replace(/[áéíóúü]/g, function (letter) {
    return { 'á': 'a', 'é': 'e', 'í': 'i', 'ó': 'o', 'ú': 'u', 'ü': 'u' }[letter];
  }).replace(/[^a-z0-9]/g, '');
}

function cell(row, headers, names) {
  for (let i = 0; i < names.length; i++) {
    const index = headers.indexOf(normalizeHeader(names[i]));
    if (index >= 0) return row[index];
  }
  return '';
}

function formatCell(value) {
  return value instanceof Date ? Utilities.formatDate(value, Session.getScriptTimeZone(), 'dd/MM/yyyy') : value;
}

function isActive(value) {
  return value === true || ['true', 'si', 'sí', 'activo', '1'].indexOf(String(value).trim().toLowerCase()) >= 0;
}

function escapeHtml(value) {
  return String(value || '').replace(/[&<>\"']/g, function (character) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[character];
  });
}
