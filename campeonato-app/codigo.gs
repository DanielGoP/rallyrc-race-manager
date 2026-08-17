const CHAMP_SHEET_CONFIG = 'Config';
const CHAMP_SHEET_POINTS = 'Puntuacion';
const CHAMP_SHEET_RACES = 'Carreras';
const CHAMP_SHEET_RESULTS = 'ResultadosCampeonato';
const CHAMP_SHEET_STAGES = 'TramosCarrera';
const CHAMP_SHEET_CLASSIFICATIONS = 'ClasificacionesCampeonato';
const CHAMP_SHEET_TIMES = 'TiemposCampeonato';
const CHAMP_SCHEMA_VERSION = '2';
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

const CHAMP_RACE_BASE_HEADERS = [
  'carreraId',
  'nombre',
  'fechaPublicacion',
  'totalParticipantes'
];
const CHAMP_RACE_HEADERS = CHAMP_RACE_BASE_HEADERS.concat([
  'schemaVersion',
  'payloadHash'
]);

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

const CHAMP_STAGE_HEADERS = [
  'carreraId',
  'tramoId',
  'tramoNombre',
  'tramoOrden',
  'pasadaId',
  'pasadaLabel',
  'pasadaTipo',
  'pasadaNumero',
  'pasadaOrden',
  'totalPasadas',
  'numDescartes'
];

const CHAMP_CLASSIFICATION_HEADERS = [
  'carreraId',
  'inscripcionId',
  'pilotoId',
  'piloto',
  'categoriaId',
  'categoria',
  'posicion',
  'descartada',
  'penalizaciones',
  'total',
  'gap',
  'completadas',
  'previstas',
  'estado'
];

const CHAMP_TIME_HEADERS = [
  'carreraId',
  'inscripcionId',
  'pasadaId',
  'tiempo',
  'penalizacion',
  'total',
  'descartada'
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
    const request = parseChampionshipRequest_(e);
    validateChampionshipToken_(request.token);
    validateChampionshipPayloadVersion_(request.payload);
    setupChampionshipSheets();
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
    const existingConfig = ss.getSheetByName(CHAMP_SHEET_CONFIG);
    if (existingConfig && existingConfig.getLastRow() > 0 && existingConfig.getLastColumn() > 0) {
      validateChampionshipHeaders_(existingConfig, CHAMP_CONFIG_ROWS[0]);
      assertChampionshipSchemaNotFuture_(existingConfig);
    }

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
      CHAMP_RACE_BASE_HEADERS
    );
    if (races.created) {
      created.push(CHAMP_SHEET_RACES);
    }
    ensureChampionshipRaceColumns_(races.sheet);

    const results = ensureChampionshipSheet_(
      ss,
      CHAMP_SHEET_RESULTS,
      [CHAMP_RESULT_HEADERS],
      CHAMP_RESULT_HEADERS
    );
    if (results.created) {
      created.push(CHAMP_SHEET_RESULTS);
    }

    const stages = ensureChampionshipSheet_(
      ss,
      CHAMP_SHEET_STAGES,
      [CHAMP_STAGE_HEADERS],
      CHAMP_STAGE_HEADERS
    );
    if (stages.created) {
      created.push(CHAMP_SHEET_STAGES);
    }

    const classifications = ensureChampionshipSheet_(
      ss,
      CHAMP_SHEET_CLASSIFICATIONS,
      [CHAMP_CLASSIFICATION_HEADERS],
      CHAMP_CLASSIFICATION_HEADERS
    );
    if (classifications.created) {
      created.push(CHAMP_SHEET_CLASSIFICATIONS);
    }

    const times = ensureChampionshipSheet_(
      ss,
      CHAMP_SHEET_TIMES,
      [CHAMP_TIME_HEADERS],
      CHAMP_TIME_HEADERS
    );
    if (times.created) {
      created.push(CHAMP_SHEET_TIMES);
    }

    migrateChampionshipSchema_(config.sheet);

    formatChampionshipSheet_(config.sheet);
    formatChampionshipSheet_(points.sheet);
    formatChampionshipSheet_(races.sheet);
    formatChampionshipSheet_(results.sheet);
    formatChampionshipSheet_(stages.sheet);
    formatChampionshipSheet_(classifications.sheet);
    formatChampionshipSheet_(times.sheet);

    points.sheet.getRange('A:B').setNumberFormat('0.###');
    races.sheet.getRange('D:D').setNumberFormat('0');
    races.sheet.getRange('E:F').setNumberFormat('@');
    results.sheet.getRange('H:N').setNumberFormat('0.###');
    results.sheet.getRange('P:S').setNumberFormat('0.###');
    stages.sheet.getRange('D:K').setNumberFormat('0.###');
    classifications.sheet.getRange('G:M').setNumberFormat('0.###');
    times.sheet.getRange('D:F').setNumberFormat('0.###');

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
    CHAMP_SHEET_RESULTS,
    CHAMP_SHEET_STAGES,
    CHAMP_SHEET_CLASSIFICATIONS,
    CHAMP_SHEET_TIMES
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
  validateChampionshipHeaders_(ss.getSheetByName(CHAMP_SHEET_STAGES), CHAMP_STAGE_HEADERS);
  validateChampionshipHeaders_(ss.getSheetByName(CHAMP_SHEET_CLASSIFICATIONS), CHAMP_CLASSIFICATION_HEADERS);
  validateChampionshipHeaders_(ss.getSheetByName(CHAMP_SHEET_TIMES), CHAMP_TIME_HEADERS);
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

function ensureChampionshipRaceColumns_(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function(value) {
    return String(value || '').trim();
  });
  CHAMP_RACE_HEADERS.forEach(function(header) {
    if (headers.indexOf(header) === -1) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue(header);
      headers.push(header);
    }
  });
  validateChampionshipHeaders_(sheet, CHAMP_RACE_HEADERS);
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

function assertChampionshipSchemaNotFuture_(sheet) {
  const config = readChampionshipConfig_(sheet);
  const versionText = String(config.SCHEMA_VERSION || '').trim();
  if (!versionText) {
    return;
  }
  if (!/^\d+$/.test(versionText)) {
    throw new Error('SCHEMA_VERSION del campeonato no es válido');
  }
  if (Number(versionText) > Number(CHAMP_SCHEMA_VERSION)) {
    throw new Error('La hoja de campeonato usa un esquema más reciente que esta aplicación');
  }
}

function migrateChampionshipSchema_(sheet) {
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
  if (version < 1) {
    throw new Error('La hoja de campeonato usa un esquema no compatible');
  }
  if (version === supported) {
    return;
  }

  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0] || '').trim() === 'SCHEMA_VERSION') {
      sheet.getRange(i + 1, 2).setValue(CHAMP_SCHEMA_VERSION);
      return;
    }
  }
  sheet.appendRow(['SCHEMA_VERSION', CHAMP_SCHEMA_VERSION]);
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

function validateChampionshipPayloadVersion_(payload) {
  const version = Number(payload && payload.schemaVersion);
  if (version !== 1 && version !== 2) {
    throw new Error('La versión del envío no es compatible');
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
    const stagesSheet = ss.getSheetByName(CHAMP_SHEET_STAGES);
    const classificationsSheet = ss.getSheetByName(CHAMP_SHEET_CLASSIFICATIONS);
    const timesSheet = ss.getSheetByName(CHAMP_SHEET_TIMES);
    const races = readChampionshipTable_(racesSheet);

    const existing = races.find(function(race) {
      return String(race.carreraId || '') === normalized.carreraId;
    });

    if (existing) {
      return resolveChampionshipDuplicate_(existing, normalized);
    }

    deleteOrphanChampionshipRows_(resultsSheet, normalized.carreraId);
    deleteOrphanChampionshipRows_(stagesSheet, normalized.carreraId);
    deleteOrphanChampionshipRows_(classificationsSheet, normalized.carreraId);
    deleteOrphanChampionshipRows_(timesSheet, normalized.carreraId);

    if (normalized.schemaVersion === 1) {
      appendChampionshipObjects_(resultsSheet, CHAMP_RESULT_HEADERS, normalized.resultados);
    } else {
      appendChampionshipObjects_(stagesSheet, CHAMP_STAGE_HEADERS, normalized.definicion.pasadas.map(function(pass) {
        const stage = normalized.definicion.tramosById[pass.tramoId];
        return {
          carreraId: normalized.carreraId,
          tramoId: pass.tramoId,
          tramoNombre: stage.nombre,
          tramoOrden: stage.orden,
          pasadaId: pass.pasadaId,
          pasadaLabel: pass.label,
          pasadaTipo: pass.tipo,
          pasadaNumero: pass.numero,
          pasadaOrden: pass.orden,
          totalPasadas: normalized.definicion.totalPasadas,
          numDescartes: normalized.definicion.numDescartes
        };
      }));
      appendChampionshipObjects_(classificationsSheet, CHAMP_CLASSIFICATION_HEADERS, normalized.resultados);

      const timeRows = [];
      normalized.resultados.forEach(function(row) {
        row.pasadas.forEach(function(pass) {
          timeRows.push({
            carreraId: normalized.carreraId,
            inscripcionId: row.inscripcionId,
            pasadaId: pass.pasadaId,
            tiempo: pass.tiempo,
            penalizacion: pass.penalizacion,
            total: pass.total,
            descartada: pass.descartada ? 'SI' : ''
          });
        });
      });
      appendChampionshipObjects_(timesSheet, CHAMP_TIME_HEADERS, timeRows);
    }

    appendChampionshipObjects_(racesSheet, CHAMP_RACE_HEADERS, [{
      carreraId: normalized.carreraId,
      nombre: normalized.nombre,
      fechaPublicacion: normalized.fechaPublicacion,
      totalParticipantes: normalized.resultados.length,
      schemaVersion: normalized.schemaVersion,
      payloadHash: normalized.payloadHash
    }]);

    return {
      status: 'IMPORTED',
      carreraId: normalized.carreraId,
      message: 'Resultados publicados correctamente en el campeonato'
    };
  } finally {
    lock.releaseLock();
  }
}

function resolveChampionshipDuplicate_(existing, normalized) {
  if (normalized.schemaVersion === 2) {
    const existingHash = String(existing.payloadHash || '').trim().toLowerCase();
    if (!existingHash || existingHash !== normalized.payloadHash) {
      throw new Error('Conflicto de carrera: el carreraId ya existe con un payloadHash diferente');
    }
  }
  return {
    status: 'ALREADY_EXISTS',
    carreraId: normalized.carreraId,
    message: 'La carrera ya estaba publicada y no se ha modificado'
  };
}

function normalizeChampionshipPayload_(payload) {
  if (!payload || typeof payload !== 'object') {
    throw new Error('No se ha recibido la carrera');
  }

  const schemaVersion = Number(payload.schemaVersion);
  if (schemaVersion !== 1 && schemaVersion !== 2) {
    throw new Error('La versión del envío no es compatible');
  }

  const race = payload.carrera || {};
  const carreraId = String(race.carreraId || '').trim();
  const nombre = String(race.nombre || '').trim();
  const fechaPublicacion = String(payload.fechaPublicacion || '').trim();
  const payloadHash = schemaVersion === 2 ? String(payload.payloadHash || '').trim().toLowerCase() : '';
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

  if (schemaVersion === 2 && !/^[a-f0-9]{64}$/.test(payloadHash)) {
    throw new Error('payloadHash de la carrera no es válido');
  }
  if (schemaVersion === 2 && computeChampionshipPayloadHash_(payload) !== payloadHash) {
    throw new Error('payloadHash no coincide con el contenido de la carrera');
  }

  if (!resultados.length) {
    throw new Error('La carrera no contiene resultados de clasificación');
  }

  if (resultados.length > 1000) {
    throw new Error('La carrera supera el máximo de participantes permitido');
  }

  const definicion = schemaVersion === 2
    ? normalizeChampionshipDefinition_(race.definicion)
    : buildLegacyChampionshipDefinition_();
  const normalizedRows = resultados.map(function(row) {
    return schemaVersion === 2
      ? normalizeChampionshipV2Result_(carreraId, row, definicion)
      : normalizeChampionshipResult_(carreraId, row);
  });
  validateChampionshipResultSet_(normalizedRows);

  return {
    schemaVersion: schemaVersion,
    carreraId: carreraId,
    nombre: nombre,
    fechaPublicacion: new Date(fechaPublicacion).toISOString(),
    payloadHash: payloadHash,
    definicion: definicion,
    resultados: normalizedRows
  };
}

function computeChampionshipPayloadHash_(payload) {
  return championshipSha256Hex_(JSON.stringify({
    schemaVersion: payload.schemaVersion,
    carrera: payload.carrera,
    incompletos: payload.incompletos,
    resultados: payload.resultados
  }));
}

function championshipSha256Hex_(value) {
  return Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    value,
    Utilities.Charset.UTF_8
  ).map(function(byte) {
    return ('0' + ((byte + 256) % 256).toString(16)).slice(-2);
  }).join('');
}

function normalizeChampionshipDefinition_(definition) {
  if (!definition || typeof definition !== 'object') {
    throw new Error('La carrera no contiene una definición válida');
  }

  const rawStages = Array.isArray(definition.tramos) ? definition.tramos : [];
  let rawPasses = Array.isArray(definition.pasadas) ? definition.pasadas : [];
  if (!rawPasses.length) {
    rawPasses = [];
    rawStages.forEach(function(stage) {
      (Array.isArray(stage.pasadas) ? stage.pasadas : []).forEach(function(pass) {
        rawPasses.push(Object.assign({}, pass, {
          tramoId: pass.tramoId || stage.tramoId || stage.id
        }));
      });
    });
  }

  if (!rawStages.length || !rawPasses.length || rawStages.length > 100 || rawPasses.length > 500) {
    throw new Error('La definición de tramos y pasadas no es válida');
  }

  const stagesById = {};
  const stages = rawStages.map(function(stage, index) {
    const tramoId = championshipIdentifier_(stage && (stage.tramoId || stage.id), 'tramo');
    const nombre = championshipText_(stage && (stage.nombre || stage.label), 'nombre de tramo', 200);
    const orden = championshipPositiveInteger_(stage && stage.orden, index + 1, 'orden de tramo');
    if (stagesById[tramoId]) {
      throw new Error('La definición contiene un tramo duplicado');
    }
    const normalized = { tramoId: tramoId, nombre: nombre, orden: orden };
    stagesById[tramoId] = normalized;
    return normalized;
  }).sort(function(a, b) {
    return a.orden - b.orden;
  });

  const passIds = {};
  const passes = rawPasses.map(function(pass, index) {
    const pasadaId = championshipIdentifier_(pass && (pass.pasadaId || pass.id), 'pasada');
    const tramoId = championshipIdentifier_(pass && pass.tramoId, 'tramo de pasada');
    if (!stagesById[tramoId] || passIds[pasadaId]) {
      throw new Error('La definición contiene una pasada duplicada o sin tramo');
    }
    passIds[pasadaId] = true;
    return {
      pasadaId: pasadaId,
      tramoId: tramoId,
      label: championshipText_(pass && (pass.label || pass.nombre), 'nombre de pasada', 200),
      tipo: String(pass && pass.tipo || '').trim().slice(0, 100),
      numero: championshipPositiveInteger_(pass && (pass.numero || pass.num), index + 1, 'número de pasada'),
      orden: championshipPositiveInteger_(pass && pass.orden, index + 1, 'orden de pasada')
    };
  }).sort(function(a, b) {
    return a.orden - b.orden;
  });

  const usedStages = {};
  passes.forEach(function(pass) {
    usedStages[pass.tramoId] = true;
  });
  stages.forEach(function(stage) {
    if (!usedStages[stage.tramoId]) {
      throw new Error('La definición contiene un tramo sin pasadas');
    }
  });

  const totalPasadas = Number(definition.totalPasadas);
  const numDescartes = Number(definition.numDescartes || 0);
  if (!Number.isInteger(totalPasadas) || totalPasadas !== passes.length) {
    throw new Error('totalPasadas no coincide con las pasadas configuradas');
  }
  const expectedDiscards = totalPasadas === 1 ? 0 : 1;
  if (!Number.isInteger(numDescartes) || numDescartes !== expectedDiscards) {
    throw new Error('numDescartes no es válido');
  }

  return {
    tramos: stages,
    tramosById: stagesById,
    pasadas: passes,
    totalPasadas: totalPasadas,
    numDescartes: numDescartes
  };
}

function normalizeChampionshipV2Result_(carreraId, row, definition) {
  const base = normalizeChampionshipResultBase_(carreraId, row, true);
  const rawPasses = Array.isArray(row && row.pasadas) ? row.pasadas : [];
  const definitionsById = {};
  definition.pasadas.forEach(function(pass) {
    definitionsById[pass.pasadaId] = pass;
  });
  const seen = {};
  const passes = rawPasses.map(function(pass) {
    const pasadaId = championshipIdentifier_(pass && (pass.pasadaId || pass.id), 'pasada de resultado');
    if (!definitionsById[pasadaId] || seen[pasadaId]) {
      throw new Error('Existe un tiempo duplicado o para una pasada no configurada');
    }
    seen[pasadaId] = true;
    const definitionPass = definitionsById[pasadaId];
    const time = championshipNonnegativeNumber_(pass.tiempo, 'tiempo de pasada', false);
    const penalty = championshipNonnegativeNumber_(pass.penalizacion, 'penalización de pasada', true);
    const expectedTotal = championshipRound3_(time + penalty);
    const receivedTotal = championshipNonnegativeNumber_(pass.total, 'total de pasada', false);
    if (!championshipNumbersEqual_(receivedTotal, expectedTotal)) {
      throw new Error('El total de una pasada no coincide con tiempo más penalización');
    }
    return {
      pasadaId: pasadaId,
      tiempo: time,
      penalizacion: penalty,
      total: expectedTotal,
      orden: definitionPass.orden,
      label: definitionPass.label,
      descartada: pass.descartada === true || String(pass.descartada || '').toUpperCase() === 'SI'
    };
  });
  const completadas = Number(row && row.completadas);
  const previstas = Number(row && row.previstas);
  if (!Number.isInteger(completadas) || completadas < 0 || completadas > definition.totalPasadas || completadas !== passes.length) {
    throw new Error('El número de pasadas completadas no es válido');
  }
  if (!Number.isInteger(previstas) || previstas !== definition.totalPasadas) {
    throw new Error('Las pasadas previstas no coinciden con la carrera');
  }

  const complete = completadas === previstas;
  const expectedDiscarded = complete
    ? passes.slice().sort(function(a, b) {
      return b.total - a.total || b.orden - a.orden;
    }).slice(0, definition.numDescartes)
    : [];
  const expectedDiscardedIds = {};
  expectedDiscarded.forEach(function(pass) {
    expectedDiscardedIds[pass.pasadaId] = true;
  });
  passes.forEach(function(pass) {
    if (pass.descartada !== Boolean(expectedDiscardedIds[pass.pasadaId])) {
      throw new Error('Las pasadas descartadas no coinciden con la configuración de carrera');
    }
  });

  const expectedPenalties = championshipRound3_(passes.reduce(function(sum, pass) {
    return sum + pass.penalizacion;
  }, 0));
  const receivedPenalties = championshipNonnegativeNumber_(row.penalizaciones, 'penalizaciones', true);
  if (!championshipNumbersEqual_(receivedPenalties, expectedPenalties)) {
    throw new Error('Las penalizaciones no coinciden con la suma de las pasadas');
  }

  const expectedTotal = passes.length
    ? championshipRound3_(passes.reduce(function(sum, pass) {
      return sum + (expectedDiscardedIds[pass.pasadaId] ? 0 : pass.total);
    }, 0))
    : '';
  const receivedTotal = championshipNumberOrBlank_(row.total);
  if ((expectedTotal === '' && receivedTotal !== '') ||
      (expectedTotal !== '' && !championshipNumbersEqual_(receivedTotal, expectedTotal))) {
    throw new Error('El total de clasificación no coincide con las pasadas puntuables');
  }

  const expectedDiscardText = expectedDiscarded.map(function(pass) {
    return pass.label + ' - ' + pass.total;
  }).join(' · ');
  if (String(row.descartada || '').trim() !== expectedDiscardText) {
    throw new Error('La descripción de pasadas descartadas no coincide con los descartes');
  }

  const expectedState = !passes.length
    ? 'Sin resultados'
    : complete
      ? (expectedPenalties > 0 ? 'Completo con penalización' : 'Completo')
      : (expectedPenalties > 0 ? 'Pendiente con penalización' : 'Pendiente');
  if (String(row.estado || '').trim() !== expectedState) {
    throw new Error('El estado no coincide con las pasadas completadas');
  }

  const gap = row.gap === '-' ? '-' : championshipNumberOrBlank_(row.gap);
  if (gap !== '' && gap !== '-' && Number(gap) < 0) {
    throw new Error('El gap no puede ser negativo');
  }

  base.pasadas = passes.map(function(pass) {
    return {
      pasadaId: pass.pasadaId,
      tiempo: pass.tiempo,
      penalizacion: pass.penalizacion,
      total: pass.total,
      descartada: pass.descartada
    };
  });
  base.descartada = expectedDiscardText;
  base.penalizaciones = expectedPenalties;
  base.total = expectedTotal;
  base.gap = gap;
  base.completadas = completadas;
  base.previstas = previstas;
  base.estado = expectedState;
  return base;
}

function normalizeChampionshipResultBase_(carreraId, row, requireStableIds) {
  const inscripcionId = String(row && row.inscripcionId || '').trim();
  const piloto = String(row && row.piloto || '').trim();
  const categoria = String(row && row.categoria || '').trim();
  const posicion = Number(row && row.posicion);
  if (!inscripcionId || !piloto || !categoria || !Number.isInteger(posicion) || posicion < 1) {
    throw new Error('Existe una fila de clasificación incompleta o inválida');
  }
  const pilotoId = String(row.pilotoId || '').trim();
  const categoriaId = String(row.categoriaId || '').trim();
  if (requireStableIds && (!pilotoId || !categoriaId)) {
    throw new Error('Los resultados v2 requieren pilotoId y categoriaId');
  }
  return {
    carreraId: carreraId,
    inscripcionId: inscripcionId,
    pilotoId: pilotoId,
    piloto: piloto,
    categoriaId: categoriaId,
    categoria: categoria,
    posicion: posicion
  };
}

function normalizeChampionshipResult_(carreraId, row) {
  const base = normalizeChampionshipResultBase_(carreraId, row);
  const completadas = Number(row && row.completadas);
  if (!Number.isInteger(completadas) || completadas < 0 || completadas > 6) {
    throw new Error('El número de pasadas completadas no es válido');
  }
  const definition = buildLegacyChampionshipDefinition_();
  const splits = definition.pasadas.map(function(pass) {
    const value = championshipNonnegativeNumberOrBlank_(row[pass.pasadaId], 'pasada legacy');
    return Object.assign({}, pass, { total: value });
  });
  const completedSplits = splits.filter(function(pass) { return pass.total !== ''; });
  if (completedSplits.length !== completadas) {
    throw new Error('Las pasadas completadas no coinciden con los tiempos recibidos');
  }
  const penalties = championshipNonnegativeNumber_(row.penalizaciones, 'penalizaciones legacy', true);
  const complete = completadas === definition.totalPasadas;
  const discarded = complete
    ? completedSplits.slice().sort(function(a, b) {
      return Number(b.total) - Number(a.total) || b.orden - a.orden;
    })[0]
    : null;
  if (discarded) {
    validateLegacyDiscardText_(row.descartada, discarded);
  } else if (String(row.descartada || '').trim()) {
    throw new Error('Una clasificación legacy incompleta no puede contener descarte');
  }
  const expectedTotal = completedSplits.length
    ? championshipRound3_(completedSplits.reduce(function(sum, pass) {
      return sum + (discarded && pass.pasadaId === discarded.pasadaId ? 0 : Number(pass.total));
    }, 0))
    : '';
  const receivedTotal = championshipNonnegativeNumberOrBlank_(row.total, 'total legacy');
  if ((expectedTotal === '' && receivedTotal !== '') ||
      (expectedTotal !== '' && !championshipNumbersEqual_(expectedTotal, receivedTotal))) {
    throw new Error('El total legacy no coincide con las pasadas puntuables');
  }
  const expectedState = !completedSplits.length
    ? 'Sin resultados'
    : complete
      ? (penalties > 0 ? 'Completo con penalización' : 'Completo')
      : (penalties > 0 ? 'Pendiente con penalización' : 'Pendiente');
  if (String(row.estado || '').trim() !== expectedState) {
    throw new Error('El estado legacy no coincide con sus pasadas y penalizaciones');
  }
  const gap = row.gap === '-' ? '-' : championshipNonnegativeNumberOrBlank_(row.gap, 'gap legacy');
  const normalized = Object.assign(base, {
    dorsal: String(row.dorsal || '').trim(),
    ida1: splits[0].total,
    vuelta1: splits[1].total,
    ida2: splits[2].total,
    vuelta2: splits[3].total,
    ida3: splits[4].total,
    vuelta3: splits[5].total,
    descartada: discarded ? discarded.label + ' - ' + discarded.total : '',
    penalizaciones: penalties,
    total: expectedTotal,
    gap: gap,
    completadas: completadas,
    estado: expectedState,
    previstas: 6
  });
  return normalized;
}

function validateLegacyDiscardText_(value, discarded) {
  const text = String(value || '').trim();
  const prefix = discarded.label + ' - ';
  if (text.indexOf(prefix) !== 0) {
    throw new Error('La pasada descartada legacy no coincide con la peor pasada');
  }
  const discardedTotal = championshipNonnegativeNumber_(text.slice(prefix.length), 'descarte legacy', false);
  if (!championshipNumbersEqual_(discardedTotal, discarded.total)) {
    throw new Error('La pasada descartada legacy no coincide con la peor pasada');
  }
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

function championshipNonnegativeNumber_(value, label, blankAsZero) {
  if ((value === '' || value === null || value === undefined) && blankAsZero) {
    return 0;
  }
  const number = Number(value);
  if (!isFinite(number) || number < 0 || value === '' || value === null || value === undefined) {
    throw new Error('El valor de ' + label + ' no es válido');
  }
  return number;
}

function championshipNonnegativeNumberOrBlank_(value, label) {
  if (value === '' || value === null || value === undefined) {
    return '';
  }
  return championshipNonnegativeNumber_(value, label, false);
}

function championshipRound3_(value) {
  return Math.round((Number(value) + Number.EPSILON) * 1000) / 1000;
}

function championshipNumbersEqual_(left, right) {
  return Math.abs(Number(left) - Number(right)) < 0.0005;
}

function championshipIdentifier_(value, label) {
  const identifier = String(value || '').trim();
  if (!/^[A-Za-z0-9][A-Za-z0-9_.:-]{0,99}$/.test(identifier)) {
    throw new Error('El identificador de ' + label + ' no es válido');
  }
  return identifier;
}

function championshipText_(value, label, maxLength) {
  const text = String(value || '').trim();
  if (!text || text.length > maxLength) {
    throw new Error('El ' + label + ' no es válido');
  }
  return text;
}

function championshipPositiveInteger_(value, fallback, label) {
  const number = value === '' || value === null || value === undefined
    ? fallback
    : Number(value);
  if (!Number.isInteger(number) || number < 1) {
    throw new Error('El ' + label + ' no es válido');
  }
  return number;
}

function buildLegacyChampionshipDefinition_() {
  const stages = [
    { tramoId: 'ida', nombre: 'Ida', orden: 1 },
    { tramoId: 'vuelta', nombre: 'Vuelta', orden: 2 }
  ];
  const stagesById = { ida: stages[0], vuelta: stages[1] };
  const passes = [];
  for (let index = 1; index <= 3; index++) {
    passes.push({ pasadaId: 'ida' + index, tramoId: 'ida', label: 'Ida ' + index, tipo: 'ida', numero: index, orden: index * 2 - 1 });
    passes.push({ pasadaId: 'vuelta' + index, tramoId: 'vuelta', label: 'Vuelta ' + index, tipo: 'vuelta', numero: index, orden: index * 2 });
  }
  return {
    tramos: stages,
    tramosById: stagesById,
    pasadas: passes,
    totalPasadas: 6,
    numDescartes: 1
  };
}

function sanitizeChampionshipCell_(value) {
  if (typeof value !== 'string') {
    return value;
  }

  const clean = value.slice(0, 500);
  return /^[=+\-@]/.test(clean) ? "'" + clean : clean;
}

function appendChampionshipObjects_(sheet, headers, objects) {
  if (!objects.length) {
    return;
  }
  const rows = objects.map(function(item) {
    return headers.map(function(header) {
      return sanitizeChampionshipCell_(item[header]);
    });
  });
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
}

function deleteOrphanChampionshipRows_(sheet, carreraId) {
  if (sheet.getLastRow() < 2) {
    return;
  }

  const rowCount = sheet.getLastRow() - 1;
  const raceIds = sheet.getRange(2, 1, rowCount, 1).getValues();
  for (let index = raceIds.length - 1; index >= 0; index--) {
    if (String(raceIds[index][0] || '') === carreraId) {
      sheet.deleteRow(index + 2);
    }
  }
}

function championshipJsonResponse_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function getChampionshipData() {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    return buildChampionshipData_();
  } finally {
    lock.releaseLock();
  }
}

function buildChampionshipData_() {
  assertChampionshipReady_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = readChampionshipConfig_(ss.getSheetByName(CHAMP_SHEET_CONFIG));
  const points = readChampionshipPoints_(ss.getSheetByName(CHAMP_SHEET_POINTS));
  const rawRaces = readChampionshipTable_(ss.getSheetByName(CHAMP_SHEET_RACES));
  const raceIds = {};
  rawRaces.forEach(function(race) {
    raceIds[String(race.carreraId || '')] = true;
  });
  const stageRows = readChampionshipTable_(ss.getSheetByName(CHAMP_SHEET_STAGES))
    .filter(function(row) {
      return raceIds[String(row.carreraId || '')];
    });
  const definitions = buildChampionshipDefinitionsForClient_(stageRows);
  const races = rawRaces
    .map(function(race) {
      return normalizeRaceForClient_(race, definitions[String(race.carreraId || '')]);
    })
    .sort(function(a, b) {
      return String(b.fechaPublicacion).localeCompare(String(a.fechaPublicacion));
    });
  const legacyResults = readChampionshipTable_(ss.getSheetByName(CHAMP_SHEET_RESULTS))
    .filter(function(row) {
      return raceIds[String(row.carreraId || '')];
    })
    .map(normalizeResultForClient_);
  const timesByResult = buildChampionshipTimesForClient_(
    readChampionshipTable_(ss.getSheetByName(CHAMP_SHEET_TIMES)).filter(function(row) {
      return raceIds[String(row.carreraId || '')];
    }),
    definitions
  );
  const dynamicResults = readChampionshipTable_(ss.getSheetByName(CHAMP_SHEET_CLASSIFICATIONS))
    .filter(function(row) {
      return raceIds[String(row.carreraId || '')];
    })
    .map(function(row) {
      const key = String(row.carreraId || '') + '|' + String(row.inscripcionId || '');
      return normalizeDynamicResultForClient_(
        row,
        timesByResult[key] || [],
        definitions[String(row.carreraId || '')]
      );
    });
  const results = legacyResults.concat(dynamicResults);
  reconcileChampionshipIdentities_(results);
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

function normalizeRaceForClient_(race, definition) {
  return {
    carreraId: String(race.carreraId || ''),
    nombre: String(race.nombre || ''),
    fechaPublicacion: championshipDateForClient_(race.fechaPublicacion),
    totalParticipantes: Number(race.totalParticipantes || 0),
    definicion: definition || championshipDefinitionForClient_(buildLegacyChampionshipDefinition_())
  };
}

function normalizeResultForClient_(row) {
  const definition = buildLegacyChampionshipDefinition_();
  const passes = definition.pasadas.map(function(pass) {
    const value = championshipClientNumber_(row[pass.pasadaId]);
    return {
      pasadaId: pass.pasadaId,
      tramoId: pass.tramoId,
      label: pass.label,
      tipo: pass.tipo,
      numero: pass.numero,
      orden: pass.orden,
      tiempo: value,
      penalizacion: '',
      total: value,
      descartada: String(row.descartada || '').indexOf(pass.label + ' -') === 0
    };
  });
  return {
    carreraId: String(row.carreraId || ''),
    inscripcionId: String(row.inscripcionId || ''),
    pilotoId: String(row.pilotoId || ''),
    piloto: String(row.piloto || ''),
    categoriaId: String(row.categoriaId || ''),
    categoria: String(row.categoria || ''),
    posicion: championshipRequiredNumber_(row.posicion, 'posición'),
    pasadas: passes,
    descartada: String(row.descartada || ''),
    penalizaciones: championshipClientNumber_(row.penalizaciones),
    total: championshipClientNumber_(row.total),
    gap: row.gap === '-' ? '-' : championshipClientNumber_(row.gap),
    completadas: championshipRequiredNumber_(row.completadas, 'pasadas completadas'),
    previstas: 6,
    estado: String(row.estado || ''),
    sourceSchemaVersion: 1
  };
}

function normalizeDynamicResultForClient_(row, passes, definition) {
  if (!definition) {
    throw new Error('ClasificacionesCampeonato contiene una carrera sin definición');
  }
  return {
    carreraId: String(row.carreraId || ''),
    inscripcionId: String(row.inscripcionId || ''),
    pilotoId: String(row.pilotoId || ''),
    piloto: String(row.piloto || ''),
    categoriaId: String(row.categoriaId || ''),
    categoria: String(row.categoria || ''),
    posicion: championshipRequiredNumber_(row.posicion, 'posición'),
    pasadas: passes,
    descartada: String(row.descartada || ''),
    penalizaciones: championshipClientNumber_(row.penalizaciones),
    total: championshipClientNumber_(row.total),
    gap: row.gap === '-' ? '-' : championshipClientNumber_(row.gap),
    completadas: championshipRequiredNumber_(row.completadas, 'pasadas completadas'),
    previstas: definition.totalPasadas,
    estado: String(row.estado || ''),
    sourceSchemaVersion: 2
  };
}

function reconcileChampionshipIdentities_(results) {
  const pilotIdsByName = {};
  const categoryIdsByName = {};
  results.forEach(function(row) {
    if (Number(row.sourceSchemaVersion || 2) !== 2) {
      return;
    }
    registerChampionshipStableName_(pilotIdsByName, row.piloto, row.pilotoId);
    registerChampionshipStableName_(categoryIdsByName, row.categoria, row.categoriaId);
  });
  results.forEach(function(row) {
    if (!String(row.pilotoId || '').trim()) {
      const pilotId = pilotIdsByName[championshipNormalize_(row.piloto)];
      if (pilotId) {
        row.pilotoId = pilotId;
      }
    }
    if (!String(row.categoriaId || '').trim()) {
      const categoryId = categoryIdsByName[championshipNormalize_(row.categoria)];
      if (categoryId) {
        row.categoriaId = categoryId;
      }
    }
  });
  return results;
}

function registerChampionshipStableName_(map, name, id) {
  const normalizedName = championshipNormalize_(name);
  const stableId = String(id || '').trim();
  if (!normalizedName || !stableId) {
    return;
  }
  if (!Object.prototype.hasOwnProperty.call(map, normalizedName)) {
    map[normalizedName] = stableId;
  } else if (map[normalizedName] !== stableId) {
    map[normalizedName] = '';
  }
}

function buildChampionshipDefinitionsForClient_(rows) {
  const grouped = {};
  rows.forEach(function(row) {
    const raceId = String(row.carreraId || '');
    if (!raceId) {
      return;
    }
    if (!grouped[raceId]) {
      grouped[raceId] = { stages: {}, passes: [], totalPasadas: Number(row.totalPasadas), numDescartes: Number(row.numDescartes || 0) };
    }
    const group = grouped[raceId];
    const stageId = String(row.tramoId || '');
    if (!group.stages[stageId]) {
      group.stages[stageId] = {
        tramoId: stageId,
        nombre: String(row.tramoNombre || ''),
        orden: championshipRequiredNumber_(row.tramoOrden, 'orden de tramo')
      };
    }
    group.passes.push({
      pasadaId: String(row.pasadaId || ''),
      tramoId: stageId,
      label: String(row.pasadaLabel || ''),
      tipo: String(row.pasadaTipo || ''),
      numero: championshipRequiredNumber_(row.pasadaNumero, 'número de pasada'),
      orden: championshipRequiredNumber_(row.pasadaOrden, 'orden de pasada')
    });
  });

  const definitions = {};
  Object.keys(grouped).forEach(function(raceId) {
    const group = grouped[raceId];
    definitions[raceId] = {
      tramos: Object.keys(group.stages).map(function(id) { return group.stages[id]; }).sort(function(a, b) { return a.orden - b.orden; }),
      pasadas: group.passes.sort(function(a, b) { return a.orden - b.orden; }),
      totalPasadas: group.totalPasadas,
      numDescartes: group.numDescartes
    };
  });
  return definitions;
}

function buildChampionshipTimesForClient_(rows, definitions) {
  const passesByRace = {};
  Object.keys(definitions).forEach(function(raceId) {
    passesByRace[raceId] = {};
    definitions[raceId].pasadas.forEach(function(pass) {
      passesByRace[raceId][pass.pasadaId] = pass;
    });
  });
  const grouped = {};
  rows.forEach(function(row) {
    const raceId = String(row.carreraId || '');
    const passId = String(row.pasadaId || '');
    const definition = passesByRace[raceId] && passesByRace[raceId][passId];
    if (!definition) {
      throw new Error('TiemposCampeonato contiene una pasada no configurada');
    }
    const key = raceId + '|' + String(row.inscripcionId || '');
    if (!grouped[key]) {
      grouped[key] = [];
    }
    grouped[key].push(Object.assign({}, definition, {
      tiempo: championshipClientNumber_(row.tiempo),
      penalizacion: championshipClientNumber_(row.penalizacion),
      total: championshipClientNumber_(row.total),
      descartada: String(row.descartada || '').toUpperCase() === 'SI'
    }));
  });
  Object.keys(grouped).forEach(function(key) {
    grouped[key].sort(function(a, b) { return a.orden - b.orden; });
  });
  return grouped;
}

function championshipDefinitionForClient_(definition) {
  return {
    tramos: definition.tramos,
    pasadas: definition.pasadas,
    totalPasadas: definition.totalPasadas,
    numDescartes: definition.numDescartes
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

  const pilotComparison = String(a.piloto || '').localeCompare(String(b.piloto || ''));
  return pilotComparison || String(a.inscripcionId || '').localeCompare(String(b.inscripcionId || ''));
}

function championshipResultIsComplete_(row) {
  const hasDynamicPasses = Array.isArray(row.pasadas);
  const passes = hasDynamicPasses ? row.pasadas : [
    row.ida1, row.vuelta1, row.ida2, row.vuelta2, row.ida3, row.vuelta3
  ].map(function(total) { return { total: total }; });
  const expected = Number(row.previstas || (hasDynamicPasses ? 0 : 6));

  return /^Completo/i.test(String(row.estado || '')) &&
    Number.isInteger(expected) && expected > 0 &&
    Number(row.completadas) === expected &&
    row.total !== '' &&
    passes.length === expected &&
    passes.every(function(pass) {
      return pass.total !== '' && pass.total !== null && pass.total !== undefined && isFinite(Number(pass.total));
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
  reconcileChampionshipIdentities_(results);
  const racesById = {};
  races.forEach(function(race) {
    racesById[race.carreraId] = race;
  });

  const groups = {};
  const categoryNames = {};
  const pilotNames = {};
  const categoryRaces = {};

  results.forEach(function(row) {
    const categoryKey = championshipCategoryKey_(row);
    if (!categoryRaces[categoryKey]) {
      categoryRaces[categoryKey] = {};
    }
    categoryRaces[categoryKey][row.carreraId] = true;
  });

  results.forEach(function(row) {
    const categoryKey = championshipCategoryKey_(row);
    const pilotKey = championshipPilotKey_(row);
    const groupKey = categoryKey + '|' + pilotKey;
    const completed = championshipResultIsComplete_(row);
    const score = completed ? Number(points[row.posicion] || 0) : 0;
    categoryNames[categoryKey] = row.categoria;
    pilotNames[pilotKey] = row.piloto;

    if (!groups[groupKey]) {
      groups[groupKey] = {
        categoriaId: categoryKey,
        categoria: row.categoria,
        pilotoId: pilotKey,
        piloto: row.piloto,
        participaciones: 0,
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
      participado: true,
      descartado: false
    });
    groups[groupKey].participaciones++;
  });

  const requestedDiscards = Math.max(0, Math.floor(Number(discardValue || 0)));
  Object.keys(groups).forEach(function(groupKey) {
    const group = groups[groupKey];
    group.categoria = categoryNames[group.categoriaId] || group.categoria;
    group.piloto = pilotNames[group.pilotoId] || group.piloto;
    const participatedRaceIds = {};
    group.resultados.forEach(function(result) {
      participatedRaceIds[result.carreraId] = true;
    });
    Object.keys(categoryRaces[group.categoriaId] || {}).forEach(function(raceId) {
      if (participatedRaceIds[raceId]) {
        return;
      }
      const race = racesById[raceId] || {};
      group.resultados.push({
        carreraId: raceId,
        carrera: race.nombre || '',
        fechaPublicacion: race.fechaPublicacion || '',
        posicion: '',
        puntos: 0,
        completo: false,
        participado: false,
        descartado: false
      });
    });
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
      participaciones: group.participaciones,
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
