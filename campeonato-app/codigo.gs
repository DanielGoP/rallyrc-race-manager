const CHAMP_SHEET_CONFIG = 'Config';
const CHAMP_SHEET_POINTS = 'Puntuacion';
const CHAMP_SHEET_RACES = 'Carreras';
const CHAMP_SHEET_RESULTS = 'ResultadosCampeonato';
const CHAMP_SCHEMA_VERSION = '1';
// Es el nombre de la Script Property; no sustituirlo por el token real.
const CHAMP_IMPORT_TOKEN_PROPERTY = 'CHAMPIONSHIP_TOKEN';

const CHAMP_CONFIG_ROWS = [
  ['key', 'value'],
  ['SCHEMA_VERSION', CHAMP_SCHEMA_VERSION],
  ['CAMPEONATO_NOMBRE', 'Campeonato Rally RC'],
  ['NUM_DESCARTES', '0']
];

const CHAMP_POINTS_ROWS = [
  ['posicion', 'puntos'],
  [1, 25],
  [2, 18],
  [3, 15],
  [4, 12],
  [5, 10],
  [6, 8],
  [7, 6],
  [8, 4],
  [9, 2],
  [10, 1]
];

const CHAMP_RACE_HEADERS = [
  'carreraId',
  'nombre',
  'fechaPublicacion',
  'totalParticipantes'
];

const CHAMP_RESULT_HEADERS = [
  'carreraId',
  'inscripcionId',
  'pilotoId',
  'piloto',
  'categoriaId',
  'categoria',
  'dorsal',
  'posicion',
  'ida1',
  'vuelta1',
  'ida2',
  'vuelta2',
  'ida3',
  'vuelta3',
  'descartada',
  'penalizaciones',
  'total',
  'gap',
  'completadas',
  'estado'
];

function doGet() {
  assertChampionshipReady_();

  return HtmlService
    .createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Rally RC - Campeonato')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function doPost(e) {
  try {
    setupChampionshipSheets();

    const request = parseChampionshipRequest_(e);
    validateChampionshipToken_(request.token);
    const result = importChampionshipRace_(request.payload);

    return championshipJsonResponse_({
      ok: true,
      status: result.status,
      carreraId: result.carreraId,
      message: result.message
    });
  } catch (error) {
    return championshipJsonResponse_({
      ok: false,
      status: 'ERROR',
      message: error && error.message ? error.message : 'Error importando la carrera'
    });
  }
}

function setupChampionshipSheets() {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const created = [];

    const config = ensureChampionshipSheet_(
      ss,
      CHAMP_SHEET_CONFIG,
      CHAMP_CONFIG_ROWS,
      CHAMP_CONFIG_ROWS[0]
    );
    if (config.created) {
      created.push(CHAMP_SHEET_CONFIG);
    }
    ensureChampionshipConfigRows_(config.sheet);
    validateChampionshipSchema_(config.sheet);

    const points = ensureChampionshipSheet_(
      ss,
      CHAMP_SHEET_POINTS,
      CHAMP_POINTS_ROWS,
      CHAMP_POINTS_ROWS[0]
    );
    if (points.created) {
      created.push(CHAMP_SHEET_POINTS);
    }

    const races = ensureChampionshipSheet_(
      ss,
      CHAMP_SHEET_RACES,
      [CHAMP_RACE_HEADERS],
      CHAMP_RACE_HEADERS
    );
    if (races.created) {
      created.push(CHAMP_SHEET_RACES);
    }

    const results = ensureChampionshipSheet_(
      ss,
      CHAMP_SHEET_RESULTS,
      [CHAMP_RESULT_HEADERS],
      CHAMP_RESULT_HEADERS
    );
    if (results.created) {
      created.push(CHAMP_SHEET_RESULTS);
    }

    formatChampionshipSheet_(config.sheet);
    formatChampionshipSheet_(points.sheet);
    formatChampionshipSheet_(races.sheet);
    formatChampionshipSheet_(results.sheet);

    points.sheet.getRange('A:B').setNumberFormat('0.###');
    races.sheet.getRange('D:D').setNumberFormat('0');
    results.sheet.getRange('H:N').setNumberFormat('0.###');
    results.sheet.getRange('P:S').setNumberFormat('0.###');

    return {
      status: 'OK',
      created: created,
      message: created.length
        ? 'Hojas creadas: ' + created.join(', ')
        : 'La base del campeonato ya estaba inicializada'
    };
  } finally {
    lock.releaseLock();
  }
}

function assertChampionshipReady_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const required = [
    CHAMP_SHEET_CONFIG,
    CHAMP_SHEET_POINTS,
    CHAMP_SHEET_RACES,
    CHAMP_SHEET_RESULTS
  ];
  const missing = required.filter(function(name) {
    return !ss.getSheetByName(name);
  });

  if (missing.length) {
    throw new Error('Ejecuta setupChampionshipSheets() antes de abrir la app');
  }

  validateChampionshipHeaders_(ss.getSheetByName(CHAMP_SHEET_CONFIG), CHAMP_CONFIG_ROWS[0]);
  validateChampionshipHeaders_(ss.getSheetByName(CHAMP_SHEET_POINTS), CHAMP_POINTS_ROWS[0]);
  validateChampionshipHeaders_(ss.getSheetByName(CHAMP_SHEET_RACES), CHAMP_RACE_HEADERS);
  validateChampionshipHeaders_(ss.getSheetByName(CHAMP_SHEET_RESULTS), CHAMP_RESULT_HEADERS);
  validateChampionshipSchema_(ss.getSheetByName(CHAMP_SHEET_CONFIG));
}

function ensureChampionshipSheet_(ss, name, initialRows, expectedHeaders) {
  let sheet = ss.getSheetByName(name);
  let created = false;

  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet
      .getRange(1, 1, initialRows.length, initialRows[0].length)
      .setValues(initialRows);
    created = true;
  } else if (sheet.getLastRow() === 0 || sheet.getLastColumn() === 0) {
    sheet
      .getRange(1, 1, initialRows.length, initialRows[0].length)
      .setValues(initialRows);
  } else {
    validateChampionshipHeaders_(sheet, expectedHeaders);
  }

  sheet.setFrozenRows(1);
  return { sheet: sheet, created: created };
}

function validateChampionshipHeaders_(sheet, expectedHeaders) {
  if (sheet.getLastColumn() < expectedHeaders.length) {
    throw new Error('La hoja ' + sheet.getName() + ' no tiene las cabeceras esperadas');
  }

  const current = sheet
    .getRange(1, 1, 1, expectedHeaders.length)
    .getValues()[0]
    .map(function(value) {
      return String(value || '').trim();
    });

  const valid = expectedHeaders.every(function(header, index) {
    return current[index] === header;
  });

  if (!valid) {
    throw new Error('Las cabeceras de ' + sheet.getName() + ' no coinciden con el esquema esperado');
  }
}

function ensureChampionshipConfigRows_(sheet) {
  const values = sheet.getDataRange().getValues();
  const existing = {};

  for (let i = 1; i < values.length; i++) {
    const key = String(values[i][0] || '').trim();
    if (key) {
      existing[key] = true;
    }
  }

  CHAMP_CONFIG_ROWS.slice(1).forEach(function(row) {
    if (!existing[row[0]]) {
      sheet.appendRow(row);
    }
  });
}

function validateChampionshipSchema_(sheet) {
  const config = readChampionshipConfig_(sheet);
  const versionText = String(config.SCHEMA_VERSION || '').trim();

  if (!/^\d+$/.test(versionText)) {
    throw new Error('SCHEMA_VERSION del campeonato no es válido');
  }

  const version = Number(versionText);
  const supported = Number(CHAMP_SCHEMA_VERSION);

  if (version > supported) {
    throw new Error('La hoja de campeonato usa un esquema más reciente que esta aplicación');
  }

  if (version < supported) {
    throw new Error('La hoja de campeonato requiere una migración de esquema');
  }
}

function formatChampionshipSheet_(sheet) {
  const lastColumn = sheet.getLastColumn();

  if (!lastColumn) {
    return;
  }

  sheet
    .getRange(1, 1, 1, lastColumn)
    .setFontWeight('bold')
    .setBackground('#f97316')
    .setFontColor('#111827');
  sheet.autoResizeColumns(1, lastColumn);
}

function parseChampionshipRequest_(e) {
  const raw = e && e.postData ? String(e.postData.contents || '') : '';

  if (!raw) {
    throw new Error('La petición no contiene datos');
  }

  let request;

  try {
    request = JSON.parse(raw);
  } catch (error) {
    throw new Error('El cuerpo de la petición no es JSON válido');
  }

  if (!request || typeof request !== 'object') {
    throw new Error('La petición no es válida');
  }

  return request;
}

function validateChampionshipToken_(receivedToken) {
  const configuredToken = PropertiesService
    .getScriptProperties()
    .getProperty(CHAMP_IMPORT_TOKEN_PROPERTY);

  if (!configuredToken) {
    throw new Error('CHAMPIONSHIP_TOKEN no está configurado en la app de campeonato');
  }

  if (String(receivedToken || '') !== String(configuredToken)) {
    throw new Error('Token de importación incorrecto');
  }
}

function importChampionshipRace_(payload) {
  const normalized = normalizeChampionshipPayload_(payload);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const racesSheet = ss.getSheetByName(CHAMP_SHEET_RACES);
    const resultsSheet = ss.getSheetByName(CHAMP_SHEET_RESULTS);
    const races = readChampionshipTable_(racesSheet);

    const existing = races.some(function(race) {
      return String(race.carreraId || '') === normalized.carreraId;
    });

    if (existing) {
      return {
        status: 'ALREADY_EXISTS',
        carreraId: normalized.carreraId,
        message: 'La carrera ya estaba publicada y no se ha modificado'
      };
    }

    deleteOrphanChampionshipRows_(resultsSheet, normalized.carreraId);

    const resultRows = normalized.resultados.map(function(row) {
      return CHAMP_RESULT_HEADERS.map(function(header) {
        return sanitizeChampionshipCell_(row[header]);
      });
    });

    if (resultRows.length) {
      resultsSheet
        .getRange(resultsSheet.getLastRow() + 1, 1, resultRows.length, CHAMP_RESULT_HEADERS.length)
        .setValues(resultRows);
    }

    racesSheet.appendRow([
      sanitizeChampionshipCell_(normalized.carreraId),
      sanitizeChampionshipCell_(normalized.nombre),
      normalized.fechaPublicacion,
      normalized.resultados.length
    ]);

    return {
      status: 'IMPORTED',
      carreraId: normalized.carreraId,
      message: 'Resultados publicados correctamente en el campeonato'
    };
  } finally {
    lock.releaseLock();
  }
}

function normalizeChampionshipPayload_(payload) {
  if (!payload || typeof payload !== 'object') {
    throw new Error('No se ha recibido la carrera');
  }

  if (Number(payload.schemaVersion) !== 1) {
    throw new Error('La versión del envío no es compatible');
  }

  const race = payload.carrera || {};
  const carreraId = String(race.carreraId || '').trim();
  const nombre = String(race.nombre || '').trim();
  const fechaPublicacion = String(payload.fechaPublicacion || '').trim();
  const resultados = Array.isArray(payload.resultados) ? payload.resultados : [];

  if (!/^[A-Za-z0-9][A-Za-z0-9_-]{0,99}$/.test(carreraId)) {
    throw new Error('El identificador de carrera no es válido');
  }

  if (!nombre || nombre.length > 200) {
    throw new Error('El nombre de carrera no es válido');
  }

  if (!fechaPublicacion || isNaN(new Date(fechaPublicacion).getTime())) {
    throw new Error('La fecha de publicación no es válida');
  }

  if (!resultados.length) {
    throw new Error('La carrera no contiene resultados de clasificación');
  }

  if (resultados.length > 1000) {
    throw new Error('La carrera supera el máximo de participantes permitido');
  }

  const normalizedRows = resultados.map(function(row) {
    return normalizeChampionshipResult_(carreraId, row);
  });
  validateChampionshipResultSet_(normalizedRows);

  return {
    carreraId: carreraId,
    nombre: nombre,
    fechaPublicacion: new Date(fechaPublicacion).toISOString(),
    resultados: normalizedRows
  };
}

function normalizeChampionshipResult_(carreraId, row) {
  const inscripcionId = String(row && row.inscripcionId || '').trim();
  const piloto = String(row && row.piloto || '').trim();
  const categoria = String(row && row.categoria || '').trim();
  const posicion = Number(row && row.posicion);

  if (!inscripcionId || !piloto || !categoria || !Number.isInteger(posicion) || posicion < 1) {
    throw new Error('Existe una fila de clasificación incompleta o inválida');
  }

  const completadas = Number(row && row.completadas);

  if (!Number.isInteger(completadas) || completadas < 0 || completadas > 6) {
    throw new Error('El número de pasadas completadas no es válido');
  }

  const normalized = {
    carreraId: carreraId,
    inscripcionId: inscripcionId,
    pilotoId: String(row.pilotoId || '').trim(),
    piloto: piloto,
    categoriaId: String(row.categoriaId || '').trim(),
    categoria: categoria,
    dorsal: String(row.dorsal || '').trim(),
    posicion: posicion,
    ida1: championshipNumberOrBlank_(row.ida1),
    vuelta1: championshipNumberOrBlank_(row.vuelta1),
    ida2: championshipNumberOrBlank_(row.ida2),
    vuelta2: championshipNumberOrBlank_(row.vuelta2),
    ida3: championshipNumberOrBlank_(row.ida3),
    vuelta3: championshipNumberOrBlank_(row.vuelta3),
    descartada: String(row.descartada || '').trim(),
    penalizaciones: championshipNumberOrBlank_(row.penalizaciones),
    total: championshipNumberOrBlank_(row.total),
    gap: row.gap === '-' ? '-' : championshipNumberOrBlank_(row.gap),
    completadas: completadas,
    estado: String(row.estado || '').trim()
  };

  const splitCount = [
    normalized.ida1,
    normalized.vuelta1,
    normalized.ida2,
    normalized.vuelta2,
    normalized.ida3,
    normalized.vuelta3
  ].filter(function(value) {
    return value !== '';
  }).length;

  if (splitCount !== completadas) {
    throw new Error('Las pasadas completadas no coinciden con los tiempos recibidos');
  }

  if (/^Completo/i.test(normalized.estado) && (completadas !== 6 || normalized.total === '')) {
    throw new Error('Un resultado completo debe contener sus seis pasadas y un total');
  }

  return normalized;
}

function validateChampionshipResultSet_(rows) {
  const inscriptions = {};
  const participants = {};
  const positions = {};

  rows.forEach(function(row) {
    const inscriptionKey = row.inscripcionId;
    const participantKey = championshipCategoryKey_(row) + '|' + championshipPilotKey_(row);
    const positionKey = championshipCategoryKey_(row) + '|' + row.posicion;

    if (inscriptions[inscriptionKey]) {
      throw new Error('La carrera contiene una inscripción duplicada');
    }

    if (participants[participantKey]) {
      throw new Error('La carrera contiene un piloto duplicado en una categoría');
    }

    if (positions[positionKey]) {
      throw new Error('La carrera contiene una posición duplicada en una categoría');
    }

    inscriptions[inscriptionKey] = true;
    participants[participantKey] = true;
    positions[positionKey] = true;
  });
}

function championshipNumberOrBlank_(value) {
  if (value === '' || value === null || value === undefined) {
    return '';
  }

  const number = Number(value);
  if (!isFinite(number)) {
    throw new Error('Existe un valor numérico no válido en la clasificación');
  }

  return number;
}

function sanitizeChampionshipCell_(value) {
  if (typeof value !== 'string') {
    return value;
  }

  const clean = value.slice(0, 500);
  return /^[=+\-@]/.test(clean) ? "'" + clean : clean;
}

function deleteOrphanChampionshipRows_(sheet, carreraId) {
  if (sheet.getLastRow() < 2) {
    return;
  }

  const rowCount = sheet.getLastRow() - 1;
  const columnCount = CHAMP_RESULT_HEADERS.length;
  const range = sheet.getRange(2, 1, rowCount, columnCount);
  const values = range.getValues();
  const retained = values.filter(function(row) {
    return String(row[0] || '') !== carreraId;
  });

  if (retained.length === values.length) {
    return;
  }

  range.clearContent();
  if (retained.length) {
    sheet.getRange(2, 1, retained.length, columnCount).setValues(retained);
  }
}

function championshipJsonResponse_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function getChampionshipData() {
  assertChampionshipReady_();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = readChampionshipConfig_(ss.getSheetByName(CHAMP_SHEET_CONFIG));
  const points = readChampionshipPoints_(ss.getSheetByName(CHAMP_SHEET_POINTS));
  const races = readChampionshipTable_(ss.getSheetByName(CHAMP_SHEET_RACES))
    .map(normalizeRaceForClient_)
    .sort(function(a, b) {
      return String(b.fechaPublicacion).localeCompare(String(a.fechaPublicacion));
    });
  const raceIds = {};
  races.forEach(function(race) {
    raceIds[race.carreraId] = true;
  });

  const results = readChampionshipTable_(ss.getSheetByName(CHAMP_SHEET_RESULTS))
    .filter(function(row) {
      return raceIds[String(row.carreraId || '')];
    })
    .map(normalizeResultForClient_);
  validateChampionshipClientResults_(results);
  recalculateChampionshipRacePositions_(results);
  const discardCount = parseChampionshipDiscards_(config.NUM_DESCARTES);

  return {
    campeonato: config.CAMPEONATO_NOMBRE || 'Campeonato Rally RC',
    descartes: discardCount,
    updatedAt: new Date().toISOString(),
    carreras: races,
    categorias: buildChampionshipCategories_(results),
    pilotos: buildChampionshipPilots_(results),
    resultados: results,
    general: buildChampionshipStandings_(races, results, points, discardCount)
  };
}

function readChampionshipConfig_(sheet) {
  const values = sheet.getDataRange().getValues();
  const config = {};

  for (let i = 1; i < values.length; i++) {
    const key = String(values[i][0] || '').trim();
    if (key) {
      config[key] = values[i][1];
    }
  }

  return config;
}

function readChampionshipPoints_(sheet) {
  const values = sheet.getDataRange().getValues();
  const points = {};

  for (let i = 1; i < values.length; i++) {
    const position = Number(values[i][0]);
    const score = Number(values[i][1]);

    if (Number.isInteger(position) && position > 0 && isFinite(score)) {
      points[position] = score;
    }
  }

  return points;
}

function readChampionshipTable_(sheet) {
  if (!sheet || sheet.getLastRow() < 2) {
    return [];
  }

  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(function(value) {
    return String(value || '').trim();
  });

  return values.slice(1).filter(function(row) {
    return row.some(function(value) {
      return value !== '';
    });
  }).map(function(row) {
    const item = {};
    headers.forEach(function(header, index) {
      item[header] = row[index];
    });
    return item;
  });
}

function normalizeRaceForClient_(race) {
  return {
    carreraId: String(race.carreraId || ''),
    nombre: String(race.nombre || ''),
    fechaPublicacion: championshipDateForClient_(race.fechaPublicacion),
    totalParticipantes: Number(race.totalParticipantes || 0)
  };
}

function normalizeResultForClient_(row) {
  return {
    carreraId: String(row.carreraId || ''),
    inscripcionId: String(row.inscripcionId || ''),
    pilotoId: String(row.pilotoId || ''),
    piloto: String(row.piloto || ''),
    categoriaId: String(row.categoriaId || ''),
    categoria: String(row.categoria || ''),
    dorsal: String(row.dorsal || ''),
    posicion: championshipRequiredNumber_(row.posicion, 'posición'),
    ida1: championshipClientNumber_(row.ida1),
    vuelta1: championshipClientNumber_(row.vuelta1),
    ida2: championshipClientNumber_(row.ida2),
    vuelta2: championshipClientNumber_(row.vuelta2),
    ida3: championshipClientNumber_(row.ida3),
    vuelta3: championshipClientNumber_(row.vuelta3),
    descartada: String(row.descartada || ''),
    penalizaciones: championshipClientNumber_(row.penalizaciones),
    total: championshipClientNumber_(row.total),
    gap: row.gap === '-' ? '-' : championshipClientNumber_(row.gap),
    completadas: championshipRequiredNumber_(row.completadas, 'pasadas completadas'),
    estado: String(row.estado || '')
  };
}

function championshipClientNumber_(value) {
  if (value === '' || value === null || value === undefined) {
    return '';
  }

  const number = Number(value);
  if (!isFinite(number)) {
    throw new Error('La hoja de resultados contiene un valor numérico no válido');
  }

  return number;
}

function championshipRequiredNumber_(value, label) {
  const number = Number(value);
  if (!isFinite(number)) {
    throw new Error('La hoja de resultados contiene una ' + label + ' no válida');
  }
  return number;
}

function parseChampionshipDiscards_(value) {
  const number = Number(value || 0);
  if (!Number.isInteger(number) || number < 0) {
    throw new Error('NUM_DESCARTES debe ser un entero mayor o igual que cero');
  }
  return number;
}

function validateChampionshipClientResults_(results) {
  const entries = {};

  results.forEach(function(row) {
    const key = row.carreraId + '|' + row.inscripcionId;
    if (entries[key]) {
      throw new Error('ResultadosCampeonato contiene una inscripción duplicada');
    }
    entries[key] = true;
  });
}

function recalculateChampionshipRacePositions_(results) {
  const groups = {};

  results.forEach(function(row) {
    const key = row.carreraId + '|' + championshipCategoryKey_(row);
    if (!groups[key]) {
      groups[key] = [];
    }
    groups[key].push(row);
  });

  Object.keys(groups).forEach(function(key) {
    const rows = groups[key];
    const complete = rows.filter(championshipResultIsComplete_)
      .sort(compareChampionshipRaceResult_);
    const incomplete = rows.filter(function(row) {
      return complete.indexOf(row) === -1;
    }).sort(compareChampionshipRaceResult_);
    const ordered = complete.concat(incomplete);

    ordered.forEach(function(row, index) {
      row.posicion = index + 1;

      if (complete.indexOf(row) === -1) {
        row.gap = '-';
      } else if (index === 0) {
        row.gap = '-';
      } else {
        row.gap = Math.round((Number(row.total) - Number(ordered[index - 1].total)) * 1000) / 1000;
      }
    });
  });
}

function compareChampionshipRaceResult_(a, b) {
  const aHasTotal = a.total !== '';
  const bHasTotal = b.total !== '';

  if (aHasTotal && bHasTotal && Number(a.total) !== Number(b.total)) {
    return Number(a.total) - Number(b.total);
  }
  if (aHasTotal !== bHasTotal) {
    return aHasTotal ? -1 : 1;
  }

  const dorsalA = Number(a.dorsal);
  const dorsalB = Number(b.dorsal);
  if (isFinite(dorsalA) && isFinite(dorsalB) && dorsalA !== dorsalB) {
    return dorsalA - dorsalB;
  }
  return String(a.dorsal).localeCompare(String(b.dorsal));
}

function championshipResultIsComplete_(row) {
  const splits = [
    row.ida1,
    row.vuelta1,
    row.ida2,
    row.vuelta2,
    row.ida3,
    row.vuelta3
  ];

  return /^Completo/i.test(String(row.estado || '')) &&
    Number(row.completadas) === 6 &&
    row.total !== '' &&
    splits.every(function(value) {
      return value !== '' && value !== null && value !== undefined && isFinite(Number(value));
    });
}

function championshipDateForClient_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]') {
    return value.toISOString();
  }

  return String(value || '');
}

function buildChampionshipCategories_(results) {
  const byId = {};

  results.forEach(function(row) {
    const id = championshipCategoryKey_(row);
    if (!byId[id]) {
      byId[id] = { categoriaId: id, nombre: row.categoria };
    }
  });

  return Object.keys(byId).map(function(id) {
    return byId[id];
  }).sort(function(a, b) {
    return a.nombre.localeCompare(b.nombre);
  });
}

function buildChampionshipPilots_(results) {
  const byId = {};

  results.forEach(function(row) {
    const id = championshipPilotKey_(row);
    if (!byId[id]) {
      byId[id] = { pilotoId: id, nombre: row.piloto };
    }
  });

  return Object.keys(byId).map(function(id) {
    return byId[id];
  }).sort(function(a, b) {
    return a.nombre.localeCompare(b.nombre);
  });
}

function buildChampionshipStandings_(races, results, points, discardValue) {
  const racesById = {};
  races.forEach(function(race) {
    racesById[race.carreraId] = race;
  });

  const groups = {};
  const categoryNames = {};

  results.forEach(function(row) {
    const categoryKey = championshipCategoryKey_(row);
    const pilotKey = championshipPilotKey_(row);
    const groupKey = categoryKey + '|' + pilotKey;
    const completed = championshipResultIsComplete_(row);
    const score = completed ? Number(points[row.posicion] || 0) : 0;
    categoryNames[categoryKey] = row.categoria;

    if (!groups[groupKey]) {
      groups[groupKey] = {
        categoriaId: categoryKey,
        categoria: row.categoria,
        pilotoId: pilotKey,
        piloto: row.piloto,
        resultados: []
      };
    }

    groups[groupKey].resultados.push({
      carreraId: row.carreraId,
      carrera: racesById[row.carreraId] ? racesById[row.carreraId].nombre : '',
      fechaPublicacion: racesById[row.carreraId]
        ? racesById[row.carreraId].fechaPublicacion
        : '',
      posicion: row.posicion,
      puntos: score,
      completo: completed,
      descartado: false
    });
  });

  const requestedDiscards = Math.max(0, Math.floor(Number(discardValue || 0)));
  Object.keys(groups).forEach(function(groupKey) {
    const group = groups[groupKey];
    group.categoria = categoryNames[group.categoriaId] || group.categoria;
  });
  const standings = Object.keys(groups).map(function(groupKey) {
    const group = groups[groupKey];
    const sortedForDiscard = group.resultados.slice().sort(function(a, b) {
      if (a.puntos !== b.puntos) {
        return a.puntos - b.puntos;
      }
      return String(a.fechaPublicacion).localeCompare(String(b.fechaPublicacion));
    });
    const discardCount = Math.min(requestedDiscards, Math.max(0, sortedForDiscard.length - 1));

    for (let i = 0; i < discardCount; i++) {
      sortedForDiscard[i].descartado = true;
    }

    const positionCounts = {};
    let total = 0;

    group.resultados.forEach(function(result) {
      if (result.completo) {
        positionCounts[result.posicion] = (positionCounts[result.posicion] || 0) + 1;
      }
      if (!result.descartado) {
        total += Number(result.puntos || 0);
      }
    });

    group.resultados.sort(function(a, b) {
      return String(b.fechaPublicacion).localeCompare(String(a.fechaPublicacion));
    });

    return {
      categoriaId: group.categoriaId,
      categoria: group.categoria,
      pilotoId: group.pilotoId,
      piloto: group.piloto,
      puntos: total,
      participaciones: group.resultados.length,
      victorias: Number(positionCounts[1] || 0),
      positionCounts: positionCounts,
      resultados: group.resultados,
      posicion: 0
    };
  });

  standings.sort(compareChampionshipStanding_);

  let category = '';
  let categoryIndex = 0;
  let previous = null;

  standings.forEach(function(row) {
    if (row.categoriaId !== category) {
      category = row.categoriaId;
      categoryIndex = 1;
      previous = null;
    } else {
      categoryIndex++;
    }

    row.posicion = previous && championshipStandingTie_(previous, row)
      ? previous.posicion
      : categoryIndex;
    previous = row;
  });

  return standings;
}

function compareChampionshipStanding_(a, b) {
  if (a.categoria !== b.categoria) {
    return a.categoria.localeCompare(b.categoria);
  }

  if (a.puntos !== b.puntos) {
    return b.puntos - a.puntos;
  }

  const maxPosition = Math.max(
    20,
    championshipMaxPosition_(a.positionCounts),
    championshipMaxPosition_(b.positionCounts)
  );

  for (let position = 1; position <= maxPosition; position++) {
    const diff = Number(b.positionCounts[position] || 0) - Number(a.positionCounts[position] || 0);
    if (diff !== 0) {
      return diff;
    }
  }

  return a.piloto.localeCompare(b.piloto);
}

function championshipStandingTie_(a, b) {
  if (a.categoriaId !== b.categoriaId || a.puntos !== b.puntos) {
    return false;
  }

  const maxPosition = Math.max(
    championshipMaxPosition_(a.positionCounts),
    championshipMaxPosition_(b.positionCounts)
  );

  for (let position = 1; position <= maxPosition; position++) {
    if (Number(a.positionCounts[position] || 0) !== Number(b.positionCounts[position] || 0)) {
      return false;
    }
  }

  return true;
}

function championshipMaxPosition_(counts) {
  return Object.keys(counts || {}).reduce(function(max, key) {
    return Math.max(max, Number(key || 0));
  }, 0);
}

function championshipCategoryKey_(row) {
  return String(row.categoriaId || '').trim() || 'categoria:' + championshipNormalize_(row.categoria);
}

function championshipPilotKey_(row) {
  return String(row.pilotoId || '').trim() || 'piloto:' + championshipNormalize_(row.piloto);
}

function championshipNormalize_(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');
}
