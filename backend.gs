const SHEET_NAME = 'Respuestas';
const EVENTS_SHEET_NAME = 'Eventos';
const PROJECTS_SHEET_NAME = 'Proyectos';
const VOTES_SHEET_NAME = 'Votos';
const PROP_KEY = 'LATEST_INGRESO_DISPLAY';
const VOTING_GRADES = [
  { key: '1', label: '1.° de Secundaria' },
  { key: '2', label: '2.° de Secundaria' },
  { key: '3', label: '3.° de Secundaria' }
];
const RESPONSE_HEADERS = ['Fecha', 'ID', 'Nombres', 'Apellidos', 'Parentesco', 'Estudiante', 'Grado', 'Email', 'DNI', 'Estado', 'Hora', 'Evento ID'];
const PROJECT_HEADERS = ['Proyecto ID', 'Grado', 'Título', 'Equipo', 'Foto URL', 'Activo'];
const VOTE_HEADERS = ['Fecha', 'Evento ID', 'Registro ID', 'Correo', 'Grado', 'Proyecto ID'];

function doGet(e) {
  const action = e && e.parameter ? e.parameter.action : '';
  if (action === 'getEvents') return json({ events: getEvents() });
  if (action === 'getVotingStatus') return json(getVotingStatus());
  if (action === 'checkEmail') {
    return json(checkEmail(e.parameter.email, e.parameter.eventoId));
  }
  if (action === 'check') return json(checkInvitationById(e.parameter.id));
  if (action === 'latest') return json(getLatestIngreso());

  const page = e && e.parameter ? e.parameter.page || 'Index' : 'Index';
  const pageName = page.toLowerCase();
  const file = pageName === 'display' ? 'display' : pageName === 'scanner' ? 'scanner' : pageName === 'admin' ? 'admin' : pageName === 'vote' ? 'vote' : 'index';
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
    if (body.action === 'checkVoter') return json(checkVoter(body.email, body.dni));
    if (body.action === 'castVote') return json(castVote(body));
    if (body.action === 'adminCheck') return json(checkAdminPin(body.pin));
    if (body.action === 'setVoting') return json(setVoting(body));
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
      id: cell(row, headers, ['idevento', 'id', 'eventoid']),
      nombre: cell(row, headers, ['nombre', 'evento']),
      descripcion: cell(row, headers, ['descripcion', 'descripción']),
      fecha: formatCell(cell(row, headers, ['fecha'])),
      mensaje: cell(row, headers, ['mensajebienvenida', 'mensaje']),
      activo: isActive(cell(row, headers, ['activo', 'estado']))
    };
  }).filter(function (event) {
    return event.id || event.nombre;
  });
}

function checkEmail(email, eventoId) {
  const normalizedEmail = String(email || '').trim().toLowerCase();
  if (!normalizedEmail) return { exists: false };

  const activeEvent = getActiveEvent();
  if (!activeEvent || String(eventoId || '').trim() !== String(activeEvent.id).trim()) return { exists: false };

  const sheet = getResponsesSheet();
  const values = sheet.getDataRange().getValues();
  for (let r = 1; r < values.length; r++) {
    const rowEmail = String(values[r][7] || '').trim().toLowerCase();
    const rowEventId = String(values[r][11] || '').trim();
    if (rowEmail === normalizedEmail && rowEventId === String(activeEvent.id).trim()) {
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
    const activeEvent = getActiveEvent();
    if (!activeEvent) return { status: 'NO_HAY_EVENTO' };

    for (let r = 1; r < values.length; r++) {
      if (String(values[r][1]).trim() !== searchId) continue;
      if (String(values[r][11]).trim() !== String(activeEvent.id).trim()) return { status: 'EVENTO_INACTIVO' };

      const data = attendeeData(values[r], searchId, activeEvent);
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
        eventoId: activeEvent.id,
        eventoNombre: activeEvent.nombre,
        eventoMensaje: activeEvent.mensaje,
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

function getVotingStatus() {
  const properties = PropertiesService.getScriptProperties();
  const activeEvent = getActiveEvent();
  return {
    grades: VOTING_GRADES.map(function (grade) {
      return {
        key: grade.key,
        label: grade.label,
        open: Boolean(activeEvent && properties.getProperty(votingStateKey(activeEvent.id, grade.key)) === 'true'),
        projects: getProjectsForGrade(grade.key)
      };
    })
  };
}

function checkVoter(email, dni) {
  const attendee = findEligibleAttendee(email, dni);
  if (!attendee) return { status: 'NO_AUTORIZADO' };
  return { status: 'OK', voterId: attendee.id, votes: getVoterVotes(attendee.id, attendee.eventId) };
}

function castVote(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const gradeKey = String(data.grade || '');
    const grade = VOTING_GRADES.find(function (item) { return item.key === gradeKey; });
    if (!grade) return { status: 'ERROR', message: 'Grado no válido' };
    const attendee = findEligibleAttendee(data.email, data.dni);
    if (!attendee) return { status: 'NO_AUTORIZADO' };
    if (PropertiesService.getScriptProperties().getProperty(votingStateKey(attendee.eventId, gradeKey)) !== 'true') {
      return { status: 'CERRADA', message: 'La votación de este grado está cerrada' };
    }

    const projectId = String(data.projectId || '').trim();
    const project = getProjectsForGrade(gradeKey).find(function (item) { return item.id === projectId; });
    if (!project) return { status: 'PROYECTO_INVALIDO' };

    const votesSheet = getVotingSheet();
    const votes = votesSheet.getDataRange().getValues();
    for (let row = 1; row < votes.length; row++) {
      if (String(votes[row][1]) === attendee.eventId && String(votes[row][2]) === attendee.id && String(votes[row][4]) === gradeKey) {
        return { status: 'YA_VOTO' };
      }
    }
    votesSheet.appendRow([new Date(), attendee.eventId, attendee.id, attendee.email, gradeKey, projectId]);
    return { status: 'OK', grade: grade.label };
  } finally {
    lock.releaseLock();
  }
}

function setVoting(data) {
  if (!checkAdminPin(data.pin).authorized) return { status: 'NO_AUTORIZADO' };

  const gradeKey = String(data.grade || '');
  if (!VOTING_GRADES.some(function (grade) { return grade.key === gradeKey; })) return { status: 'ERROR', message: 'Grado no válido' };
  const activeEvent = getActiveEvent();
  if (!activeEvent) return { status: 'EVENTO_INACTIVO', message: 'No hay un evento activo' };
  const open = data.open === true;
  PropertiesService.getScriptProperties().setProperty(votingStateKey(activeEvent.id, gradeKey), open ? 'true' : 'false');
  return { status: 'OK', grade: gradeKey, open: open };
}

function votingStateKey(eventId, gradeKey) {
  return 'VOTING_OPEN_' + encodeURIComponent(String(eventId)) + '_' + gradeKey;
}

function checkAdminPin(pin) {
  const expectedPin = String(PropertiesService.getScriptProperties().getProperty('VOTING_ADMIN_PIN') || '').trim();
  const submittedPin = String(pin || '').trim();
  return { authorized: Boolean(expectedPin) && submittedPin === expectedPin };
}

function findEligibleAttendee(email, dni) {
  const normalizedEmail = String(email || '').trim().toLowerCase();
  const normalizedDni = String(dni || '').trim();
  const activeEvent = getActiveEvent();
  if (!normalizedEmail || !/^\d{8}$/.test(normalizedDni) || !activeEvent) return null;

  const values = getResponsesSheet().getDataRange().getValues();
  for (let row = 1; row < values.length; row++) {
    const record = values[row];
    if (String(record[7] || '').trim().toLowerCase() !== normalizedEmail || String(record[8] || '').trim() !== normalizedDni) continue;
    if (String(record[11] || '').trim() !== String(activeEvent.id).trim() || String(record[9] || '').trim().toUpperCase() !== 'INGRESADO') continue;
    return { id: String(record[1]), email: normalizedEmail, eventId: String(activeEvent.id) };
  }
  return null;
}

function getVoterVotes(attendeeId, eventId) {
  const votes = getVotingSheet().getDataRange().getValues();
  const result = {};
  for (let row = 1; row < votes.length; row++) {
    if (String(votes[row][1]) === String(eventId) && String(votes[row][2]) === String(attendeeId)) {
      result[String(votes[row][4])] = String(votes[row][5]);
    }
  }
  return result;
}

function getProjectsForGrade(gradeKey) {
  const values = getProjectsSheet().getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0].map(normalizeHeader);
  return values.slice(1).map(function (row) {
    const active = cell(row, headers, ['activo', 'estado']);
    const projectGrade = normalizeGrade(cell(row, headers, ['grado', 'gradoseccion']));
    return {
      id: String(cell(row, headers, ['proyectoid', 'id', 'codigo']) || '').trim(),
      grade: projectGrade,
      title: String(cell(row, headers, ['titulo', 'proyecto', 'nombre']) || '').trim(),
      team: String(cell(row, headers, ['equipo', 'integrantes']) || '').trim(),
      photo: String(cell(row, headers, ['fotourl', 'foto', 'imagen', 'urlfoto']) || '').trim(),
      active: active === '' || isActive(active)
    };
  }).filter(function (project) {
    return project.id && project.title && project.grade === gradeKey && project.active;
  });
}

function normalizeGrade(value) {
  const normalized = String(value || '').toLowerCase().replace(/[°º.]/g, '').trim();
  const match = normalized.match(/[123]/);
  if (match) return match[0];
  if (normalized.indexOf('primero') >= 0 || normalized.indexOf('primer') >= 0) return '1';
  if (normalized.indexOf('segundo') >= 0) return '2';
  if (normalized.indexOf('tercero') >= 0 || normalized.indexOf('tercer') >= 0) return '3';
  return '';
}

function getProjectsSheet() {
  return getOrCreateSheet(PROJECTS_SHEET_NAME, PROJECT_HEADERS);
}

function getVotingSheet() {
  return getOrCreateSheet(VOTES_SHEET_NAME, VOTE_HEADERS);
}

function getOrCreateSheet(name, headers) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(name);
  if (!sheet) sheet = spreadsheet.insertSheet(name);
  if (sheet.getLastRow() === 0) sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  return sheet;
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

  const activeEvent = getActiveEvent();
  if (!activeEvent || String(data.eventoId || '').trim() !== String(activeEvent.id).trim()) {
    return { status: 'EVENTO_INACTIVO', message: 'El evento seleccionado ya no esta activo' };
  }

  const sheet = getResponsesSheet();
  const existing = checkEmail(data.email, data.eventoId);
  if (existing.exists) return { status: 'CORREO_DUPLICADO', id: existing.id };

  const id = buildInvitationId(sheet);
  sheet.appendRow([
    new Date(), id, data.nombres, data.apellidos, data.parentesco,
    data.nombreEstudiante, data.gradoSeccion, data.email, dni, 'PENDIENTE', '', activeEvent.id
  ]);

  const qrUrl = 'https://api.qrserver.com/v1/create-qr-code/?size=250x250&data=' + encodeURIComponent(id);
  try {
    sendInvitationEmail(data.email, data.nombres, data.apellidos, data.parentesco,
      data.nombreEstudiante, data.gradoSeccion, dni, id, qrUrl);
  } catch (err) {
    console.error('Error enviando invitacion ' + id + ': ' + err.message);
    return { status: 'EMAIL_ERROR', id: id, email: data.email, message: 'La entrada se registro, pero no se pudo enviar el correo: ' + err.message };
  }
  return { status: 'OK', id: id, email: data.email };
}

function resendInvitation(email, eventoId) {
  const normalizedEmail = String(email || '').trim().toLowerCase();
  const sheet = getResponsesSheet();
  const values = sheet.getDataRange().getValues();
  const activeEvent = getActiveEvent();
  if (!activeEvent || String(eventoId || '').trim() !== String(activeEvent.id).trim()) {
    return { status: 'ERROR', message: 'El evento ya no esta activo' };
  }

  for (let r = 1; r < values.length; r++) {
    if (String(values[r][7] || '').trim().toLowerCase() !== normalizedEmail || String(values[r][11] || '').trim() !== String(activeEvent.id).trim()) continue;
    const id = String(values[r][1]);
    const qrUrl = 'https://api.qrserver.com/v1/create-qr-code/?size=250x250&data=' + encodeURIComponent(id);
    try {
      sendInvitationEmail(values[r][7], values[r][2], values[r][3], values[r][4], values[r][5], values[r][6], values[r][8], id, qrUrl);
    } catch (err) {
      return { status: 'EMAIL_ERROR', id: id, message: 'No se pudo reenviar el correo: ' + err.message };
    }
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

function attendeeData(row, id, event) {
  return {
    id: id,
    nombres: row[2],
    apellidos: row[3],
    parentesco: row[4],
    estudiante: row[5],
    grado: row[6],
    dni: row[8],
    eventoId: event.id,
    eventoNombre: event.nombre,
    eventoMensaje: event.mensaje
  };
}

function getActiveEvent() {
  const events = getEvents().filter(function (event) { return event.activo; });
  return events.length ? events[0] : null;
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
  const votingUrl = ScriptApp.getService().getUrl() + '?page=vote';
  const htmlBody = '<div style="font-family:Arial">' +
    '<h2>Invitacion ' + safe[7] + '</h2>' +
    '<p>' + safe[1] + ' ' + safe[2] + '</p>' +
    '<p>Estudiante: ' + safe[4] + '</p>' +
    '<p>Grado: ' + safe[5] + '</p>' +
    '<p>DNI: ' + safe[6] + '</p>' +
    '<p>Para ingresar debe presentar este codigo QR y su DNI fisico.</p>' +
    '<p>El dia de la feria, despues de validar tu ingreso, podras votar desde tu celular: <a href="' + escapeHtml(votingUrl) + '">abrir votacion</a>.</p>' +
    '<img src="' + qrUrl + '"/></div>';
  const qrResponse = UrlFetchApp.fetch(qrUrl, { muteHttpExceptions: true });
  if (qrResponse.getResponseCode() !== 200) {
    throw new Error('No se pudo generar el codigo QR (HTTP ' + qrResponse.getResponseCode() + ')');
  }
  const blob = qrResponse.getBlob().setName(id + '.png');
  MailApp.sendEmail({
    to: email,
    subject: 'Invitacion Feria - QR ' + id,
    htmlBody: htmlBody,
    attachments: [blob]
  });
}

function autorizarCorreo() {
  MailApp.getRemainingDailyQuota();
  UrlFetchApp.fetch('https://www.google.com', { muteHttpExceptions: true });
  Logger.log('Autorizacion de correo y URL Fetch completada');
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
