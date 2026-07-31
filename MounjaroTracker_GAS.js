// =====================================================
// MounjaroTracker — Google Apps Script
// Copia-incolla questo codice nell'editor GAS,
// poi: Distribuisci > Nuova distribuzione > Tipo: App web
// Esegui come: Io | Accesso: Chiunque → Distribuisci
// =====================================================

const SHEET_MEAS = 'Misurazioni';
const SHEET_INJ  = 'Punture';
const SHEET_SYM  = 'Sintomi';

// ── GET: lettura e scrittura (URL-based) ──────────────
function doGet(e) {
  var result;
  try {
    var action = e.parameter.action;
    if      (action === 'getData')  result = getData();
    else if (action === 'saveAll')  result = saveAll(e.parameter);
    else result = { error: 'Azione non riconosciuta: ' + action };
  } catch(err) {
    result = { error: err.toString() };
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── POST: fallback per grandi volumi di dati ─────────
function doPost(e) {
  var result;
  try {
    var action = e.parameter.action;
    if (action === 'saveAll') result = saveAll(e.parameter);
    else result = { error: 'Azione non riconosciuta: ' + action };
  } catch(err) {
    result = { error: err.toString() };
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Crea il foglio se non esiste ─────────────────────
function sheet(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

// ── Lettura dati ─────────────────────────────────────
function getData() {
  var measurements = readSheet(SHEET_MEAS, function(r) {
    return { date: r[0], weight: r[1], chest: r[2] };
  });
  var injections = readSheet(SHEET_INJ, function(r) {
    return {
      date:     r[0],
      time:     r[1] || '',
      dose:     r[2] || '',
      site:     r[3] || '',
      symptoms: r[4] ? String(r[4]).split(',').map(function(s){ return s.trim(); }).filter(Boolean) : [],
      notes:    r[5] || ''
    };
  });
  var symptomLogs = readSheet(SHEET_SYM, function(r) {
    return {
      date:     r[0],
      symptoms: r[1] ? String(r[1]).split(',').map(function(s){ return s.trim(); }).filter(Boolean) : [],
      note:     r[2] || '',
      endDate:  r[3] || null
    };
  });
  return {
    measurements: measurements,
    injections:   injections,
    symptomLogs:  symptomLogs
  };
}

function readSheet(name, mapper) {
  var data = sheet(name).getDataRange().getValues();
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) rows.push(mapper(data[i]));
  }
  return rows;
}

// ── Salvataggio dati ─────────────────────────────────
function saveAll(params) {
  var measurements = JSON.parse(params.measurements || '[]');
  var injections   = JSON.parse(params.injections   || '[]');
  var symptomLogs  = JSON.parse(params.symptomLogs  || '[]');

  // Misurazioni
  var ms = sheet(SHEET_MEAS);
  ms.clearContents();
  ms.appendRow(['Data', 'Peso (kg)', 'Torace (cm)']);
  measurements.forEach(function(m) {
    ms.appendRow([m.date, m.weight || '', m.chest || '']);
  });

  // Punture
  var is = sheet(SHEET_INJ);
  is.clearContents();
  is.appendRow(['Data', 'Orario', 'Dosaggio', 'Sede', 'Sintomi', 'Note']);
  injections.forEach(function(i) {
    is.appendRow([
      i.date, i.time || '', i.dose || '', i.site || '',
      (i.symptoms || []).join(', '), i.notes || ''
    ]);
  });

  // Sintomi giornalieri
  var ss2 = sheet(SHEET_SYM);
  ss2.clearContents();
  ss2.appendRow(['Data', 'Sintomi', 'Nota', 'Fine Sintomo']);
  symptomLogs.forEach(function(s) {
    ss2.appendRow([
      s.date,
      (s.symptoms || []).join(', '),
      s.note    || '',
      s.endDate || ''
    ]);
  });

  return {
    ok:           true,
    measurements: measurements.length,
    injections:   injections.length,
    symptomLogs:  symptomLogs.length
  };
}
