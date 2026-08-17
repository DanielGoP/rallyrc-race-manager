const SHEET_CONFIG = 'Config';
const SHEET_INSCRIPCIONES = 'Inscripciones';
const SHEET_RESULTADOS = 'Resultados';
const SHEET_CLASIFICACION = 'Clasificacion';
const SHEET_PILOTOS_DB = 'PilotosDB';
const SHEET_CATEGORIAS_DB = 'CategoriasDB';
const SHEET_TRAMOS = 'Tramos';
const SHEET_PASADAS = 'Pasadas';
const CONFIG_SCHEMA_VERSION = 'SCHEMA_VERSION';
const CONFIG_ENVIRONMENT = 'ENVIRONMENT';
const CONFIG_CARRERA_ID = 'CARRERA_ID';
const CONFIG_CARRERA_ESTADO = 'CARRERA_ESTADO';
const CONFIG_REVISION = 'CONFIG_REVISION';
const CONFIG_CAMPEONATO_PUBLICADA = 'CAMPEONATO_PUBLICADA';
const CONFIG_CAMPEONATO_PUBLICACION_INICIADA = 'CAMPEONATO_PUBLICACION_INICIADA';
const CHAMPIONSHIP_ENDPOINT_PROPERTY = 'CHAMPIONSHIP_ENDPOINT';
const CHAMPIONSHIP_TOKEN_PROPERTY = 'CHAMPIONSHIP_TOKEN';
const CHAMPIONSHIP_PUBLICATION_LEASE_MS = 10 * 60 * 1000;
const CURRENT_SCHEMA_VERSION = '5';
const MAX_TRAMOS = 100;
const MAX_PASADAS = 500;

const INSCRIPCIONES_HEADERS = ['inscripcionId', 'pilotoId', 'categoriaId', 'piloto', 'categoria'];
const RESULTADOS_HEADERS = [
  'resultadoId', 'timestamp', 'carreraId', 'inscripcionId', 'tramoId', 'pasadaId',
  'piloto', 'categoria', 'pasadaLabel', 'tiempo', 'penalizacion', 'total', 'juez'
];
const CLASIFICACION_HEADERS = [
  'inscripcionId', 'pilotoId', 'categoriaId', 'piloto', 'categoria', 'descartada',
  'penalizaciones', 'total', 'gap', 'completadas', 'previstas', 'estado'
];
const TRAMOS_HEADERS = ['tramoId', 'nombre', 'orden', 'numIdas', 'numVueltas'];
const PASADAS_HEADERS = ['pasadaId', 'tramoId', 'label', 'tipo', 'numero', 'orden'];
const LEGACY_PASADAS = [
  { pasadaId: 'P1', tramoId: 'T1', label: 'Ida 1', tipo: 'ida', numero: 1, orden: 1001 },
  { pasadaId: 'P2', tramoId: 'T1', label: 'Vuelta 1', tipo: 'vuelta', numero: 1, orden: 1002 },
  { pasadaId: 'P3', tramoId: 'T1', label: 'Ida 2', tipo: 'ida', numero: 2, orden: 1003 },
  { pasadaId: 'P4', tramoId: 'T1', label: 'Vuelta 2', tipo: 'vuelta', numero: 2, orden: 1004 },
  { pasadaId: 'P5', tramoId: 'T1', label: 'Ida 3', tipo: 'ida', numero: 3, orden: 1005 },
  { pasadaId: 'P6', tramoId: 'T1', label: 'Vuelta 3', tipo: 'vuelta', numero: 3, orden: 1006 }
];

function doGet() {
  setupSheets();
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('Rally RC - Resultados')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const current = parseSchemaVersion_(CURRENT_SCHEMA_VERSION);
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const existingConfig = ss.getSheetByName(SHEET_CONFIG);
    let stored = 0;
    if (existingConfig) {
      const configRows = existingConfig.getDataRange().getValues();
      const versionRow = configRows.slice(1).find(function(row) {
        return cellText_(row[0]) === CONFIG_SCHEMA_VERSION;
      });
      stored = parseSchemaVersion_(versionRow ? versionRow[1] : '0');
    }
    if (stored > current) throw new Error('La hoja usa un esquema más reciente que esta aplicación');

    createSheetIfNotExists_(ss, SHEET_CONFIG, [
      ['key', 'value'], ['PIN', 'CAMBIAR_PIN_ACCESO'], ['ADMIN_PIN', 'CAMBIAR_PIN_ADMIN'],
      ['CARRERA_ACTIVA', 'Rally RC'], [CONFIG_CARRERA_ID, Utilities.getUuid()],
      [CONFIG_CARRERA_ESTADO, 'CONFIGURACION'], [CONFIG_REVISION, '1'],
      [CONFIG_CAMPEONATO_PUBLICADA, 'NO'], [CONFIG_CAMPEONATO_PUBLICACION_INICIADA, ''],
      [CONFIG_ENVIRONMENT, 'PRODUCTION'], [CONFIG_SCHEMA_VERSION, '0']
    ]);
    createSheetIfNotExists_(ss, SHEET_PILOTOS_DB, [[
      'pilotoId', 'nombre', 'alias', 'activo', 'notas', 'createdAt', 'updatedAt'
    ]]);
    createSheetIfNotExists_(ss, SHEET_CATEGORIAS_DB, [[
      'categoriaId', 'nombre', 'activa', 'orden', 'notas', 'createdAt', 'updatedAt'
    ]]);
    createSheetIfNotExists_(ss, SHEET_TRAMOS, [TRAMOS_HEADERS]);
    createSheetIfNotExists_(ss, SHEET_PASADAS, [PASADAS_HEADERS]);
    createSheetIfNotExists_(ss, SHEET_INSCRIPCIONES, [INSCRIPCIONES_HEADERS]);
    createSheetIfNotExists_(ss, SHEET_RESULTADOS, [RESULTADOS_HEADERS]);
    createSheetIfNotExists_(ss, SHEET_CLASIFICACION, [CLASIFICACION_HEADERS]);

    if (stored < current) {
      ensureMigrationBackup_();
      migrateToNormalizedRace_();
      updateConfigValue_(CONFIG_SCHEMA_VERSION, CURRENT_SCHEMA_VERSION);
    }
  } finally {
    lock.releaseLock();
  }
}

function parseSchemaVersion_(value) {
  const text = cellText_(value || '0');
  if (!/^\d+$/.test(text)) throw new Error('SCHEMA_VERSION no es un entero válido');
  return Number(text);
}

function createSheetIfNotExists_(ss, name, rows) {
  let sheet = ss.getSheetByName(name);
  if (sheet) return sheet;
  sheet = ss.insertSheet(name);
  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  sheet.setFrozenRows(1);
  autoResize_(sheet);
  return sheet;
}

function autoResize_(sheet) {
  if (sheet.getLastColumn() > 0) sheet.autoResizeColumns(1, sheet.getLastColumn());
}

function getConfigValues_() {
  const rows = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CONFIG).getDataRange().getValues();
  const config = {};
  rows.slice(1).forEach(function(row) {
    const key = cellText_(row[0]);
    if (key) config[key] = cellText_(row[1]);
  });
  return config;
}

function getConfigValue_(key) {
  return getConfigValues_()[key] || '';
}

function updateConfigValue_(key, value) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CONFIG);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (cellText_(rows[i][0]) === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  sheet.appendRow([key, value]);
}

function validatePin_(pin) {
  return cellText_(pin) === getConfigValue_('PIN');
}

function validatePin(pin) {
  return validatePin_(pin);
}

function validateAdminPin_(pin) {
  return cellText_(pin) === getConfigValue_('ADMIN_PIN');
}

function validateAdminPin(pin) {
  return validateAdminPin_(pin);
}

function assertCredentials_(pin, adminPin) {
  if (!validatePin_(pin)) throw new Error('PIN de carrera incorrecto');
  if (!validateAdminPin_(adminPin)) throw new Error('PIN de administración incorrecto');
}

function getHeaderMap_(sheet) {
  return getHeaderMapFromHeaders_(sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]);
}

function getHeaderMapFromHeaders_(headers) {
  const map = {};
  headers.forEach(function(header, index) { map[cellText_(header)] = index; });
  return map;
}

function readObjects_(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return [];
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(cellText_);
  return values.slice(1).map(function(row) {
    const object = {};
    headers.forEach(function(header, index) { if (header) object[header] = row[index]; });
    return object;
  }).filter(function(row) {
    return headers.some(function(header) { return row[header] !== '' && row[header] !== null; });
  });
}

function writeObjects_(sheet, headers, objects) {
  const rows = objects.map(function(object) {
    return headers.map(function(header) { return object[header] === undefined ? '' : object[header]; });
  });
  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (rows.length) sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  sheet.setFrozenRows(1);
  autoResize_(sheet);
}

function appendObject_(sheet, object) {
  const map = getHeaderMap_(sheet);
  const row = new Array(sheet.getLastColumn()).fill('');
  Object.keys(object).forEach(function(key) { if (map[key] !== undefined) row[map[key]] = object[key]; });
  sheet.appendRow(row);
}

function cellText_(value) {
  return String(value === undefined || value === null ? '' : value).trim();
}

function normalizeText_(value) {
  return cellText_(value).toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
}

function parseNumber_(value) {
  if (typeof value === 'number') return value;
  return Number(cellText_(value).replace(',', '.'));
}

function numberOrZero_(value) {
  const number = parseNumber_(value);
  return isFinite(number) ? round3_(number) : 0;
}

function round3_(value) {
  return Math.round(Number(value) * 1000) / 1000;
}

function nowString_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}

function formatDateForClient_(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]') {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  }
  return String(value);
}

function secondsToTimeParts_(value) {
  const tenths = Math.round(Number(value || 0) * 10);
  return { minutos: Math.floor(tenths / 600), segundos: Math.floor((tenths % 600) / 10), decimas: tenths % 10 };
}

function getNextId_(prefix, ids) {
  let max = 0;
  ids.forEach(function(id) {
    const match = cellText_(id).match(new RegExp('^' + prefix + '(\\d+)$'));
    if (match) max = Math.max(max, Number(match[1]));
  });
  return prefix + String(max + 1).padStart(3, '0');
}

function ensureMigrationBackup_() {
  const config = getConfigValues_();
  if (config.MIGRATION_V5_BACKUP) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const suffix = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm-ss');
  [SHEET_CONFIG, SHEET_TRAMOS, SHEET_PASADAS, SHEET_INSCRIPCIONES, SHEET_RESULTADOS, SHEET_CLASIFICACION]
    .forEach(function(name) {
      const sheet = ss.getSheetByName(name);
      if (sheet) sheet.copyTo(ss).setName(uniqueBackupName_(ss, name + '_backup_v5_' + suffix));
    });
  updateConfigValue_('MIGRATION_V5_BACKUP', suffix);
}

function uniqueBackupName_(ss, base) {
  let name = base.slice(0, 99);
  let number = 2;
  while (ss.getSheetByName(name)) {
    const suffix = '_' + number++;
    name = base.slice(0, 99 - suffix.length) + suffix;
  }
  return name;
}

function readLegacyInscripciones_(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return [];
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(cellText_);
  const map = getHeaderMapFromHeaders_(headers);
  return values.slice(1).filter(function(row) { return row.some(function(value) { return value !== ''; }); })
    .map(function(row, index) {
      const oldFourColumns = map.pilotoId === undefined;
      const result = {
        inscripcionId: cellText_(row[map.inscripcionId === undefined ? 0 : map.inscripcionId]),
        pilotoId: oldFourColumns ? '' : cellText_(row[map.pilotoId]),
        categoriaId: map.categoriaId === undefined ? '' : cellText_(row[map.categoriaId]),
        piloto: cellText_(row[map.piloto === undefined ? 2 : map.piloto]),
        categoria: cellText_(row[map.categoria === undefined ? 3 : map.categoria])
      };
      if (!result.inscripcionId || !result.piloto || !result.categoria) {
        throw new Error('Inscripción legacy incompleta en la fila ' + (index + 2));
      }
      return result;
    });
}

function resolveLegacyMaster_(row, records, idKey, rowNameKey, recordNameKey, prefix, sheet, buildRow) {
  const suppliedId = cellText_(row[idKey]);
  let record = suppliedId ? records.find(function(item) { return item[idKey] === suppliedId; }) : null;
  if (suppliedId && !record) throw new Error('Identificador legacy desconocido: ' + suppliedId);
  if (!record) {
    const matches = records.filter(function(item) {
      return normalizeText_(item[recordNameKey]) === normalizeText_(row[rowNameKey]);
    });
    if (matches.length > 1) throw new Error('Nombre legacy ambiguo: ' + row[rowNameKey]);
    record = matches[0];
  }
  if (!record) {
    const id = getNextId_(prefix, records.map(function(item) { return item[idKey]; }));
    const object = buildRow(id, row[rowNameKey]);
    appendObject_(sheet, object);
    records.push(object);
    record = object;
  }
  return record;
}

function readLegacyResultados_(sheet, carreraId, inscriptionsById, definition) {
  if (!sheet || sheet.getLastRow() < 2) return [];
  const values = sheet.getDataRange().getValues();
  const map = getHeaderMapFromHeaders_(values[0]);
  const normalized = map.pasadaId !== undefined && map.tramoId !== undefined && map.carreraId !== undefined;
  return values.slice(1).filter(function(row) { return row.some(function(value) { return value !== ''; }); })
    .map(function(row, index) {
      const inscripcionId = cellText_(row[map.inscripcionId]);
      if (!inscriptionsById[inscripcionId]) throw new Error('Resultado sin inscripción en fila ' + (index + 2));
      let pass;
      if (normalized) {
        pass = definition.pasadas.find(function(item) { return item.pasadaId === cellText_(row[map.pasadaId]); });
      } else {
        const legacyNumber = Number(row[map.pasada]);
        pass = Number.isInteger(legacyNumber) && legacyNumber >= 1 && legacyNumber <= 6
          ? LEGACY_PASADAS[legacyNumber - 1] : null;
      }
      if (!pass) throw new Error('Pasada legacy inválida en fila ' + (index + 2));
      const sourceCarreraId = normalized ? cellText_(row[map.carreraId]) : '';
      if (sourceCarreraId && sourceCarreraId !== carreraId) {
        throw new Error('Resultado de otra carrera en fila ' + (index + 2));
      }
      return {
        resultadoId: cellText_(row[map.resultadoId]) || Utilities.getUuid(),
        timestamp: row[map.timestamp] || new Date(),
        carreraId: carreraId,
        inscripcionId: inscripcionId, tramoId: pass.tramoId, pasadaId: pass.pasadaId,
        piloto: cellText_(row[map.piloto]) || inscriptionsById[inscripcionId].piloto,
        categoria: cellText_(row[map.categoria]) || inscriptionsById[inscripcionId].categoria,
        pasadaLabel: normalized ? cellText_(row[map.pasadaLabel]) || pass.label : cellText_(row[map.pasadaLabel]) || pass.label,
        tiempo: numberOrZero_(row[map.tiempo]), penalizacion: numberOrZero_(row[map.penalizacion]),
        total: row[map.total] === '' || row[map.total] === undefined
          ? round3_(numberOrZero_(row[map.tiempo]) + numberOrZero_(row[map.penalizacion]))
          : numberOrZero_(row[map.total]),
        juez: cellText_(row[map.juez])
      };
    });
}

function migrateToNormalizedRace_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = getConfigValues_();
  const inscriptionSource = readLegacyInscripciones_(ss.getSheetByName(SHEET_INSCRIPCIONES));
  const pilots = getPilotosDB_();
  const categories = getCategoriasDB_();
  const now = nowString_();
  const inscriptions = inscriptionSource.map(function(row) {
    const pilot = resolveLegacyMaster_(row, pilots, 'pilotoId', 'piloto', 'nombre', 'P', ss.getSheetByName(SHEET_PILOTOS_DB),
      function(id, name) { return { pilotoId: id, nombre: name, alias: '', activo: 'SI', notas: '', createdAt: now, updatedAt: now }; });
    const categoryRow = { categoriaId: row.categoriaId, categoria: row.categoria };
    const category = resolveLegacyMaster_(categoryRow, categories, 'categoriaId', 'categoria', 'nombre', 'C', ss.getSheetByName(SHEET_CATEGORIAS_DB),
      function(id, name) { return { categoriaId: id, nombre: name, activa: 'SI', orden: categories.length + 1, notas: '', createdAt: now, updatedAt: now }; });
    return {
      inscripcionId: row.inscripcionId, pilotoId: pilot.pilotoId, categoriaId: category.categoriaId,
      piloto: pilot.nombre, categoria: category.nombre
    };
  });
  if (inscriptions.length !== inscriptionSource.length) throw new Error('La migración perdió inscripciones');
  const inscriptionIds = {};
  const inscriptionPairs = {};
  inscriptions.forEach(function(row) {
    if (inscriptionIds[row.inscripcionId]) throw new Error('inscripcionId duplicado en migración');
    const pair = row.pilotoId + '|' + row.categoriaId;
    if (inscriptionPairs[pair]) throw new Error('Piloto y categoría duplicados en migración: ' + pair);
    inscriptionIds[row.inscripcionId] = row;
    inscriptionPairs[pair] = true;
  });
  const carreraId = cellText_(config[CONFIG_CARRERA_ID]) || Utilities.getUuid();
  const resultSheet = ss.getSheetByName(SHEET_RESULTADOS);
  const sourceResultCount = resultSheet && resultSheet.getLastRow() > 1
    ? resultSheet.getDataRange().getValues().slice(1).filter(function(row) {
        return row.some(function(value) { return value !== ''; });
      }).length : 0;
  const hasStoredDefinition = ss.getSheetByName(SHEET_TRAMOS).getLastRow() > 1 &&
    ss.getSheetByName(SHEET_PASADAS).getLastRow() > 1;
  const definition = hasStoredDefinition ? getRaceDefinition_() : buildLegacyRaceDefinition_();
  const resultados = readLegacyResultados_(resultSheet, carreraId, inscriptionIds, definition);
  if (resultados.length !== sourceResultCount) throw new Error('La migración perdió resultados');
  const keys = {};
  resultados.forEach(function(row) {
    const key = row.carreraId + '|' + row.inscripcionId + '|' + row.pasadaId;
    if (keys[key]) throw new Error('Resultado duplicado durante la migración: ' + key);
    keys[key] = true;
  });
  validateDefinition_(definition);
  writeRaceDefinition_(definition);
  writeObjects_(ss.getSheetByName(SHEET_INSCRIPCIONES), INSCRIPCIONES_HEADERS, inscriptions);
  writeObjects_(resultSheet, RESULTADOS_HEADERS, resultados);
  updateConfigValue_(CONFIG_CARRERA_ID, carreraId);
  updateConfigValue_(CONFIG_CARRERA_ESTADO, config[CONFIG_CARRERA_ESTADO] || (resultados.length ? 'INICIADA' : 'CONFIGURACION'));
  updateConfigValue_(CONFIG_REVISION, config[CONFIG_REVISION] || '1');
  updateConfigValue_(CONFIG_CAMPEONATO_PUBLICADA, config[CONFIG_CAMPEONATO_PUBLICADA] || 'NO');
  if (config[CONFIG_CAMPEONATO_PUBLICACION_INICIADA] === undefined) updateConfigValue_(CONFIG_CAMPEONATO_PUBLICACION_INICIADA, '');
  refreshClasificacion_(inscriptions, resultados);
}

function buildLegacyRaceDefinition_() {
  const tramos = [{ tramoId: 'T1', nombre: 'Tramo 1', orden: 1, numIdas: 3, numVueltas: 3, pasadas: LEGACY_PASADAS.map(clonePass_) }];
  return { tramos: tramos, pasadas: LEGACY_PASADAS.map(clonePass_), totalPasadas: 6, numDescartes: 1 };
}

function clonePass_(pass) {
  return { pasadaId: pass.pasadaId, tramoId: pass.tramoId, label: pass.label, tipo: pass.tipo, numero: pass.numero, orden: pass.orden };
}

function validateDefinition_(definition) {
  if (!definition || !Array.isArray(definition.tramos) || !definition.tramos.length || definition.tramos.length > MAX_TRAMOS) {
    throw new Error('La definición de tramos no es válida');
  }
  if (!Array.isArray(definition.pasadas) || definition.pasadas.length < 1 || definition.pasadas.length > MAX_PASADAS) {
    throw new Error('La carrera debe tener entre 1 y ' + MAX_PASADAS + ' pasadas');
  }
  if (Number(definition.totalPasadas) !== definition.pasadas.length) {
    throw new Error('totalPasadas no coincide con la definición');
  }
  const expectedDiscards = definition.pasadas.length === 1 ? 0 : 1;
  if (Number(definition.numDescartes) !== expectedDiscards) throw new Error('El número de descartes no es válido');
  const stages = {};
  definition.tramos.forEach(function(stage, index) {
    if (!stage.tramoId || stages[stage.tramoId] || stage.orden !== index + 1 || stage.nombre !== 'Tramo ' + (index + 1)) {
      throw new Error('Tramo duplicado, desordenado o con nombre inválido');
    }
    if (!Number.isInteger(stage.numIdas) || stage.numIdas < 1 || !Number.isInteger(stage.numVueltas) ||
        (stage.numVueltas !== 0 && stage.numVueltas !== stage.numIdas)) {
      throw new Error('Configuración de idas/vueltas inválida');
    }
    stages[stage.tramoId] = stage;
  });
  const passIds = {};
  const passOrders = {};
  const stagePasses = {};
  definition.pasadas.forEach(function(pass) {
    if (!pass.pasadaId || passIds[pass.pasadaId] || !stages[pass.tramoId] ||
        (pass.tipo !== 'ida' && pass.tipo !== 'vuelta') || !Number.isInteger(pass.numero) || pass.numero < 1 ||
        !Number.isInteger(pass.orden) || pass.orden < 1 || passOrders[pass.orden]) {
      throw new Error('Definición de pasada inválida');
    }
    passIds[pass.pasadaId] = true;
    passOrders[pass.orden] = true;
    const key = pass.tramoId + '|' + pass.tipo + '|' + pass.numero;
    if (stagePasses[key]) throw new Error('Pasada lógica duplicada');
    stagePasses[key] = true;
  });
  definition.tramos.forEach(function(stage) {
    for (let number = 1; number <= stage.numIdas; number++) {
      if (!stagePasses[stage.tramoId + '|ida|' + number]) throw new Error('Falta una ida configurada');
      if (stage.numVueltas && !stagePasses[stage.tramoId + '|vuelta|' + number]) throw new Error('Falta una vuelta configurada');
    }
    const actual = definition.pasadas.filter(function(pass) { return pass.tramoId === stage.tramoId; }).length;
    if (actual !== stage.numIdas + stage.numVueltas) throw new Error('El número de pasadas del tramo no coincide');
  });
  return definition;
}

function getRaceDefinition_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tramos = readObjects_(ss.getSheetByName(SHEET_TRAMOS)).map(function(row) {
    return {
      tramoId: cellText_(row.tramoId), nombre: cellText_(row.nombre), orden: Number(row.orden),
      numIdas: Number(row.numIdas), numVueltas: Number(row.numVueltas), pasadas: []
    };
  }).sort(compareOrder_);
  const byId = {};
  tramos.forEach(function(stage) { byId[stage.tramoId] = stage; });
  const pasadas = readObjects_(ss.getSheetByName(SHEET_PASADAS)).map(function(row) {
    return {
      pasadaId: cellText_(row.pasadaId), tramoId: cellText_(row.tramoId), label: cellText_(row.label),
      tipo: cellText_(row.tipo), numero: Number(row.numero), orden: Number(row.orden)
    };
  }).sort(compareOrder_);
  pasadas.forEach(function(pass) { if (byId[pass.tramoId]) byId[pass.tramoId].pasadas.push(pass); });
  return validateDefinition_({
    tramos: tramos, pasadas: pasadas, totalPasadas: pasadas.length,
    numDescartes: pasadas.length === 1 ? 0 : 1
  });
}

function compareOrder_(a, b) {
  return Number(a.orden) - Number(b.orden);
}

function normalizeTramosConfig_(rawTramos) {
  if (!Array.isArray(rawTramos) || !rawTramos.length || rawTramos.length > MAX_TRAMOS) {
    throw new Error('La carrera debe contener entre 1 y ' + MAX_TRAMOS + ' tramos');
  }
  const ids = {};
  return rawTramos.map(function(raw, index) {
    const id = cellText_(raw && (raw.tramoId || raw.id)) || Utilities.getUuid();
    const idas = Number(raw && raw.numIdas);
    const vueltas = Number(raw && raw.numVueltas || 0);
    if (!/^[A-Za-z0-9][A-Za-z0-9_-]{0,99}$/.test(id) || ids[id]) throw new Error('Identificador de tramo inválido o duplicado');
    if (!Number.isInteger(idas) || idas < 1 || !Number.isInteger(vueltas) || (vueltas !== 0 && vueltas !== idas)) {
      throw new Error('Cada tramo requiere idas >= 1 y vueltas cero o iguales a las idas');
    }
    ids[id] = true;
    return { tramoId: id, nombre: 'Tramo ' + (index + 1), orden: index + 1, numIdas: idas, numVueltas: vueltas, pasadas: [] };
  });
}

function buildDefinitionFromTramos_(tramos, existingDefinition) {
  const old = {};
  (existingDefinition ? existingDefinition.pasadas : []).forEach(function(pass) {
    old[pass.tramoId + '|' + pass.tipo + '|' + pass.numero] = pass;
  });
  const passes = [];
  tramos.forEach(function(stage) {
    let localOrder = 1;
    for (let number = 1; number <= stage.numIdas; number++) {
      ['ida', 'vuelta'].forEach(function(type) {
        if (type === 'vuelta' && number > stage.numVueltas) return;
        const previous = old[stage.tramoId + '|' + type + '|' + number];
        const pass = {
          pasadaId: previous ? previous.pasadaId : Utilities.getUuid(), tramoId: stage.tramoId,
          label: (type === 'ida' ? 'Ida ' : 'Vuelta ') + number, tipo: type, numero: number,
          orden: stage.orden * 1000 + localOrder++
        };
        stage.pasadas.push(pass);
        passes.push(pass);
      });
    }
  });
  return validateDefinition_({
    tramos: tramos, pasadas: passes, totalPasadas: passes.length,
    numDescartes: passes.length === 1 ? 0 : 1
  });
}

function writeRaceDefinition_(definition) {
  validateDefinition_(definition);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  writeObjects_(ss.getSheetByName(SHEET_TRAMOS), TRAMOS_HEADERS, definition.tramos);
  writeObjects_(ss.getSheetByName(SHEET_PASADAS), PASADAS_HEADERS, definition.pasadas);
}

function recoverCarreraPublicationStatus_(config) {
  const status = cellText_(config[CONFIG_CAMPEONATO_PUBLICADA] || 'NO').toUpperCase();
  if (status !== 'PUBLICANDO') return status;
  const started = new Date(config[CONFIG_CAMPEONATO_PUBLICACION_INICIADA] || '').getTime();
  if (isFinite(started) && Date.now() - started <= CHAMPIONSHIP_PUBLICATION_LEASE_MS) return status;
  updateConfigValue_(CONFIG_CAMPEONATO_PUBLICADA, 'PUBLICACION_INCIERTA');
  updateConfigValue_(CONFIG_CAMPEONATO_PUBLICACION_INICIADA, '');
  config[CONFIG_CAMPEONATO_PUBLICADA] = 'PUBLICACION_INCIERTA';
  return 'PUBLICACION_INCIERTA';
}

function assertRaceMutable_() {
  const status = recoverCarreraPublicationStatus_(getConfigValues_());
  if (status === 'PUBLICANDO') throw new Error('La carrera se está publicando en el campeonato');
  if (status === 'PUBLICACION_INCIERTA') {
    throw new Error('El estado de publicación es incierto. Reintenta la publicación antes de modificar la carrera');
  }
  if (status === 'SI') throw new Error('La carrera ya está publicada y no admite cambios');
}

function assertCarreraAdmiteResultados_() {
  assertRaceMutable_();
  if (getConfigValue_(CONFIG_CARRERA_ESTADO).toUpperCase() !== 'INICIADA') {
    throw new Error('La carrera está en configuración. Iníciala antes de registrar resultados');
  }
}

function assertExpectedRace_(config, carreraId, revision) {
  if (!carreraId || cellText_(carreraId) !== cellText_(config[CONFIG_CARRERA_ID])) {
    throw new Error('La carrera ha cambiado. Actualiza los datos antes de continuar');
  }
  if (!Number.isInteger(Number(revision)) || Number(revision) !== Number(config[CONFIG_REVISION])) {
    throw new Error('La configuración ha cambiado. Actualiza los datos antes de continuar');
  }
}

function assertExpectedRaceIfProvided_(config, payload) {
  if (payload && payload.carreraId !== undefined) {
    if (cellText_(payload.carreraId) !== cellText_(config[CONFIG_CARRERA_ID])) {
      throw new Error('La carrera ha cambiado. Actualiza las inscripciones');
    }
  }
}

function bumpConfigRevision_() {
  const revision = Math.max(0, Number(getConfigValue_(CONFIG_REVISION)) || 0) + 1;
  updateConfigValue_(CONFIG_REVISION, String(revision));
  return revision;
}

function raceConfigResponse_(config) {
  const definition = getRaceDefinition_();
  return {
    schemaVersion: CURRENT_SCHEMA_VERSION, carreraId: cellText_(config[CONFIG_CARRERA_ID]),
    carrera: cellText_(config.CARRERA_ACTIVA), estado: cellText_(config[CONFIG_CARRERA_ESTADO]),
    revision: Number(config[CONFIG_REVISION]) || 1,
    publicacion: cellText_(config[CONFIG_CAMPEONATO_PUBLICADA]) || 'NO',
    tramos: definition.tramos, pasadas: definition.pasadas,
    totalPasadas: definition.totalPasadas, numDescartes: definition.numDescartes
  };
}

function getConfiguracionCarrera(pin, adminPin) {
  assertCredentials_(pin, adminPin);
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const config = getConfigValues_();
    recoverCarreraPublicationStatus_(config);
    return raceConfigResponse_(config);
  } finally { lock.releaseLock(); }
}

function actualizarConfiguracionCarrera(pin, adminPin, payload) {
  assertCredentials_(pin, adminPin);
  payload = payload || {};
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    assertRaceMutable_();
    const config = getConfigValues_();
    assertExpectedRace_(config, payload.carreraId, payload.revision);
    if (config[CONFIG_CARRERA_ESTADO] !== 'CONFIGURACION') throw new Error('Los tramos solo se configuran antes de iniciar');
    const definition = buildDefinitionFromTramos_(normalizeTramosConfig_(payload.tramos), getRaceDefinition_());
    writeRaceDefinition_(definition);
    config[CONFIG_REVISION] = String(bumpConfigRevision_());
    refreshClasificacion_();
    return raceConfigResponse_(config);
  } finally { lock.releaseLock(); }
}

function guardarConfiguracionCarrera(pin, adminPin, payload) {
  return actualizarConfiguracionCarrera(pin, adminPin, payload);
}

function iniciarCarrera(pin, adminPin, expectedCarreraId, expectedRevision) {
  assertCredentials_(pin, adminPin);
  const request = typeof expectedCarreraId === 'object'
    ? expectedCarreraId : { carreraId: expectedCarreraId, revision: expectedRevision };
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    assertRaceMutable_();
    const config = getConfigValues_();
    assertExpectedRace_(config, request.carreraId, request.revision);
    if (config[CONFIG_CARRERA_ESTADO] === 'INICIADA') throw new Error('La carrera ya está iniciada');
    validateDefinition_(getRaceDefinition_());
    updateConfigValue_(CONFIG_CARRERA_ESTADO, 'INICIADA');
    config[CONFIG_CARRERA_ESTADO] = 'INICIADA';
    config[CONFIG_REVISION] = String(bumpConfigRevision_());
    return raceConfigResponse_(config);
  } finally { lock.releaseLock(); }
}

function aumentarPasadasCarrera(pin, adminPin, payload) {
  assertCredentials_(pin, adminPin);
  payload = payload || {};
  const increment = Number(payload.incremento === undefined ? 1 : payload.incremento);
  if (!Number.isInteger(increment) || increment < 1) throw new Error('Incremento de pasadas inválido');
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    assertRaceMutable_();
    const config = getConfigValues_();
    assertExpectedRace_(config, payload.carreraId, payload.revision);
    if (config[CONFIG_CARRERA_ESTADO] !== 'INICIADA') throw new Error('La carrera debe estar iniciada');
    const current = getRaceDefinition_();
    const selected = current.tramos.find(function(stage) { return stage.tramoId === cellText_(payload.tramoId); });
    if (!selected) throw new Error('El tramo ya no existe. Actualiza la configuración');
    const paired = selected.numVueltas > 0;
    if ((paired && payload.modo !== 'par') || (!paired && payload.modo !== 'ida')) {
      throw new Error(paired ? 'El tramo requiere parejas ida/vuelta' : 'El tramo solo admite idas');
    }
    const raw = current.tramos.map(function(stage) {
      return {
        tramoId: stage.tramoId,
        numIdas: stage.numIdas + (stage.tramoId === selected.tramoId ? increment : 0),
        numVueltas: stage.numVueltas + (stage.tramoId === selected.tramoId && paired ? increment : 0)
      };
    });
    const definition = buildDefinitionFromTramos_(normalizeTramosConfig_(raw), current);
    writeRaceDefinition_(definition);
    config[CONFIG_REVISION] = String(bumpConfigRevision_());
    refreshClasificacion_();
    return raceConfigResponse_(config);
  } finally { lock.releaseLock(); }
}

function incrementarPasadasCarrera(pin, adminPin, payload) {
  return aumentarPasadasCarrera(pin, adminPin, payload);
}

function getPilotosDB_() {
  return readObjects_(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PILOTOS_DB)).map(function(row) {
    return {
      pilotoId: cellText_(row.pilotoId), nombre: cellText_(row.nombre), alias: cellText_(row.alias),
      activo: cellText_(row.activo) || 'SI', notas: cellText_(row.notas),
      createdAt: formatDateForClient_(row.createdAt), updatedAt: formatDateForClient_(row.updatedAt)
    };
  }).filter(function(row) { return row.pilotoId && row.nombre; });
}

function getCategoriasDB_() {
  return readObjects_(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CATEGORIAS_DB)).map(function(row) {
    return {
      categoriaId: cellText_(row.categoriaId), nombre: cellText_(row.nombre), activa: cellText_(row.activa) || 'SI',
      orden: Number(row.orden || 999), notas: cellText_(row.notas),
      createdAt: formatDateForClient_(row.createdAt), updatedAt: formatDateForClient_(row.updatedAt)
    };
  }).filter(function(row) { return row.categoriaId && row.nombre; }).sort(function(a, b) {
    return a.orden - b.orden || a.nombre.localeCompare(b.nombre);
  });
}

function getInscripciones_() {
  return readObjects_(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_INSCRIPCIONES)).map(function(row) {
    return {
      inscripcionId: cellText_(row.inscripcionId), pilotoId: cellText_(row.pilotoId),
      categoriaId: cellText_(row.categoriaId), piloto: cellText_(row.piloto), categoria: cellText_(row.categoria)
    };
  }).filter(validInscripcion_);
}

function validInscripcion_(row) {
  return Boolean(row.inscripcionId && row.pilotoId && row.categoriaId && row.piloto && row.categoria);
}

function getResultados_(carreraId) {
  const currentCarreraId = cellText_(carreraId === undefined ? getConfigValue_(CONFIG_CARRERA_ID) : carreraId);
  return readObjects_(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_RESULTADOS)).map(normalizeResultObject_)
    .filter(function(row) {
      return row.resultadoId && row.inscripcionId && row.pasadaId && row.carreraId === currentCarreraId;
    });
}

function normalizeResultObject_(row) {
  return {
    resultadoId: cellText_(row.resultadoId), timestamp: row.timestamp, carreraId: cellText_(row.carreraId),
    inscripcionId: cellText_(row.inscripcionId), tramoId: cellText_(row.tramoId), pasadaId: cellText_(row.pasadaId),
    piloto: cellText_(row.piloto), categoria: cellText_(row.categoria), pasadaLabel: cellText_(row.pasadaLabel),
    tiempo: numberOrZero_(row.tiempo), penalizacion: numberOrZero_(row.penalizacion),
    total: numberOrZero_(row.total), juez: cellText_(row.juez)
  };
}

function buildResultadosFromValues_(values) {
  if (!values || !values.length) return [];
  const headers = values[0].map(cellText_);
  return values.slice(1).map(function(row) {
    const object = {};
    headers.forEach(function(header, index) { object[header] = row[index]; });
    return normalizeResultObject_(object);
  }).filter(function(row) { return row.resultadoId; });
}

function getCategoriasActivas_() {
  const categories = getCategoriasDB_().filter(function(row) { return row.activa.toUpperCase() !== 'NO'; })
    .map(function(row) { return row.nombre; });
  getInscripciones_().forEach(function(row) { categories.push(row.categoria); });
  return Array.from(new Set(categories.filter(Boolean))).sort();
}

function getCategorias(pin) {
  if (!validatePin_(pin)) throw new Error('PIN incorrecto');
  return getCategoriasActivas_();
}

function getRegistroData(pin) {
  const config = getConfigValues_();
  if (cellText_(pin) !== cellText_(config.PIN)) throw new Error('PIN incorrecto');
  const definition = getRaceDefinition_();
  return {
    categorias: getCategoriasActivas_(), inscripciones: getInscripciones_(), tramos: definition.tramos,
    pasadas: definition.pasadas, totalPasadas: definition.totalPasadas,
    carreraId: config[CONFIG_CARRERA_ID], carreraEstado: config[CONFIG_CARRERA_ESTADO],
    configRevision: Number(config[CONFIG_REVISION]) || 1, fetchedAt: Date.now(),
    schemaVersion: CURRENT_SCHEMA_VERSION, environment: config[CONFIG_ENVIRONMENT] || 'PRODUCTION'
  };
}

function getInscripcionesByCategoria(pin, categoria) {
  if (!validatePin_(pin)) throw new Error('PIN incorrecto');
  const filter = cellText_(categoria);
  return getInscripciones_().filter(function(row) { return row.categoria === filter || row.categoriaId === filter; })
    .sort(function(a, b) { return a.piloto.localeCompare(b.piloto); });
}

function resolvePassFromPayload_(payload, definition) {
  const id = cellText_(payload && payload.pasadaId);
  if (id) {
    const pass = definition.pasadas.find(function(row) { return row.pasadaId === id; });
    if (!pass) throw new Error('La pasada ya no existe. Actualiza la carrera');
    return pass;
  }
  const number = Number(payload && payload.pasada);
  const legacy = definition.pasadas.length === 6 && definition.pasadas.every(function(row, index) {
    return row.pasadaId === 'P' + (index + 1);
  });
  if (!legacy || !Number.isInteger(number) || number < 1 || number > 6) {
    throw new Error('Esta interfaz está desactualizada. Recarga la aplicación');
  }
  return definition.pasadas[number - 1];
}

function normalizeResultPayload_(payload, definition) {
  if (!payload) throw new Error('No se han recibido datos');
  const minutes = Number(payload.minutos || 0);
  const seconds = Number(payload.segundos || 0);
  const tenths = Number(payload.decimas || 0);
  const penalty = parseNumber_(payload.penalizacion || 0);
  if (!cellText_(payload.inscripcionId)) throw new Error('Selecciona un piloto');
  if (!Number.isInteger(minutes) || minutes < 0 || !Number.isInteger(seconds) || seconds < 0 || seconds > 59 ||
      !Number.isInteger(tenths) || tenths < 0 || tenths > 9 || !isFinite(penalty) || penalty < 0) {
    throw new Error('Tiempo o penalización inválidos');
  }
  return {
    inscripcionId: cellText_(payload.inscripcionId), pass: resolvePassFromPayload_(payload, definition),
    tiempo: round3_(minutes * 60 + seconds + tenths / 10), penalizacion: round3_(penalty),
    juez: cellText_(payload.juez) || 'Sin identificar'
  };
}

function findResultadoRowIndex_(values, carreraId, inscripcionId, pasadaId) {
  if (!values.length) return -1;
  const map = getHeaderMapFromHeaders_(values[0]);
  for (let i = 1; i < values.length; i++) {
    if (cellText_(values[i][map.carreraId]) === cellText_(carreraId) &&
        cellText_(values[i][map.inscripcionId]) === cellText_(inscripcionId) &&
        cellText_(values[i][map.pasadaId]) === cellText_(pasadaId)) return i + 1;
  }
  return -1;
}

function resultForClient_(row) {
  const result = Object.assign({}, row);
  result.timestamp = formatDateForClient_(row.timestamp);
  return result;
}

function saveResultado(pin, payload) {
  if (!validatePin_(pin)) throw new Error('PIN incorrecto');
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    assertCarreraAdmiteResultados_();
    const config = getConfigValues_();
    assertExpectedRace_(config, payload && payload.carreraId, payload && payload.configRevision);
    const data = normalizeResultPayload_(payload, getRaceDefinition_());
    const inscription = getInscripciones_().find(function(row) { return row.inscripcionId === data.inscripcionId; });
    if (!inscription) throw new Error('No se ha encontrado la inscripción');
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_RESULTADOS);
    const values = sheet.getDataRange().getValues();
    if (findResultadoRowIndex_(values, config[CONFIG_CARRERA_ID], data.inscripcionId, data.pass.pasadaId) !== -1) {
      const existing = getResultados_(config[CONFIG_CARRERA_ID]).find(function(row) {
        return row.carreraId === config[CONFIG_CARRERA_ID] && row.inscripcionId === data.inscripcionId && row.pasadaId === data.pass.pasadaId;
      });
      return { status: 'DUPLICATE', message: 'Ya existe un resultado para esta inscripción y pasada.', existing: resultForClient_(existing) };
    }
    const row = {
      resultadoId: Utilities.getUuid(), timestamp: new Date(), carreraId: config[CONFIG_CARRERA_ID],
      inscripcionId: inscription.inscripcionId, tramoId: data.pass.tramoId, pasadaId: data.pass.pasadaId,
      piloto: inscription.piloto, categoria: inscription.categoria, pasadaLabel: data.pass.label,
      tiempo: data.tiempo, penalizacion: data.penalizacion, total: round3_(data.tiempo + data.penalizacion), juez: data.juez
    };
    appendObject_(sheet, row);
    refreshClasificacion_();
    return { status: 'CREATED', message: 'Resultado guardado', saved: resultForClient_(row), pasadasRegistradas: getRegisteredPassIds_(data.inscripcionId) };
  } finally { lock.releaseLock(); }
}

function getRegisteredPassIds_(inscripcionId) {
  const definition = getRaceDefinition_();
  const order = {};
  definition.pasadas.forEach(function(pass) { order[pass.pasadaId] = pass.orden; });
  return Array.from(new Set(getResultados_().filter(function(row) {
    return row.inscripcionId === inscripcionId && order[row.pasadaId] !== undefined;
  }).map(function(row) { return row.pasadaId; }))).sort(function(a, b) { return order[a] - order[b]; });
}

function getPasadasRegistradas(pin, inscripcionId) {
  if (!validatePin_(pin)) throw new Error('PIN incorrecto');
  const ids = getRegisteredPassIds_(cellText_(inscripcionId));
  const definition = getRaceDefinition_();
  const legacy = definition.pasadas.length === 6 && definition.pasadas.every(function(row, index) { return row.pasadaId === 'P' + (index + 1); });
  return legacy ? ids.map(function(id) { return Number(id.slice(1)); }) : ids;
}

function getResultadoParaCorreccion(pin, adminPin, inscripcionId, pasada) {
  assertCredentials_(pin, adminPin);
  const request = typeof pasada === 'object' ? pasada : { pasada: pasada };
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const config = getConfigValues_();
    assertExpectedRace_(config, request.carreraId, request.configRevision);
    const pass = resolvePassFromPayload_(request, getRaceDefinition_());
    const result = getResultados_(config[CONFIG_CARRERA_ID]).find(function(row) {
      return row.inscripcionId === cellText_(inscripcionId) && row.pasadaId === pass.pasadaId;
    });
    if (!result) return { status: 'NOT_FOUND', message: 'No existe ningún resultado para la pasada seleccionada.' };
    const client = resultForClient_(result);
    client.configRevision = Number(config[CONFIG_REVISION]);
    Object.assign(client, secondsToTimeParts_(result.tiempo));
    return { status: 'FOUND', resultado: client };
  } finally { lock.releaseLock(); }
}

function corregirResultado(pin, adminPin, payload) {
  assertCredentials_(pin, adminPin);
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    assertCarreraAdmiteResultados_();
    const config = getConfigValues_();
    assertExpectedRace_(config, payload && payload.carreraId, payload && payload.configRevision);
    const data = normalizeResultPayload_(payload, getRaceDefinition_());
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_RESULTADOS);
    const values = sheet.getDataRange().getValues();
    const index = findResultadoRowIndex_(values, config[CONFIG_CARRERA_ID], data.inscripcionId, data.pass.pasadaId);
    if (index === -1) throw new Error('No existe ningún resultado para corregir');
    const map = getHeaderMapFromHeaders_(values[0]);
    const row = values[index - 1].slice();
    row[map.timestamp] = new Date();
    row[map.tiempo] = data.tiempo;
    row[map.penalizacion] = data.penalizacion;
    row[map.total] = round3_(data.tiempo + data.penalizacion);
    row[map.juez] = data.juez + ' - Corrección';
    sheet.getRange(index, 1, 1, row.length).setValues([row]);
    refreshClasificacion_();
    const saved = resultForClient_(buildResultadosFromValues_([values[0], row])[0]);
    saved.configRevision = Number(config[CONFIG_REVISION]);
    return { status: 'OK', message: 'Resultado corregido correctamente', saved: saved };
  } finally { lock.releaseLock(); }
}

function buildClasificacion_(categoriaFilter, inscripcionesInput, resultadosInput) {
  const hasInjectedResults = Array.isArray(resultadosInput);
  const currentCarreraId = hasInjectedResults ? '' : getConfigValue_(CONFIG_CARRERA_ID);
  const definition = getRaceDefinition_();
  const passById = {};
  definition.pasadas.forEach(function(pass) { passById[pass.pasadaId] = pass; });
  const inscriptions = (inscripcionesInput || getInscripciones_()).filter(function(row) {
    return !categoriaFilter || row.categoria === categoriaFilter || row.categoriaId === categoriaFilter;
  });
  const grouped = {};
  (hasInjectedResults ? resultadosInput : getResultados_(currentCarreraId)).filter(function(result) {
    return !currentCarreraId || result.carreraId === currentCarreraId;
  }).forEach(function(result) {
    if (!passById[result.pasadaId]) return;
    if (!grouped[result.inscripcionId]) grouped[result.inscripcionId] = {};
    grouped[result.inscripcionId][result.pasadaId] = result;
  });
  const rows = inscriptions.map(function(inscription) {
    const byPass = grouped[inscription.inscripcionId] || {};
    const passes = definition.pasadas.map(function(pass) {
      const result = byPass[pass.pasadaId];
      return {
        pasadaId: pass.pasadaId, tramoId: pass.tramoId, label: pass.label, tipo: pass.tipo,
        numero: pass.numero, orden: pass.orden, tiempo: result ? result.tiempo : '',
        penalizacion: result ? result.penalizacion : '', total: result ? result.total : '', descartada: false
      };
    });
    const registered = passes.filter(function(pass) { return pass.total !== ''; });
    const complete = registered.length === definition.totalPasadas;
    let discarded = null;
    if (complete && definition.numDescartes === 1) {
      discarded = registered.slice().sort(function(a, b) { return Number(b.total) - Number(a.total) || b.orden - a.orden; })[0];
      discarded.descartada = true;
    }
    const total = registered.length ? round3_(registered.reduce(function(sum, pass) {
      return sum + (discarded && pass.pasadaId === discarded.pasadaId ? 0 : Number(pass.total));
    }, 0)) : '';
    const penalties = round3_(registered.reduce(function(sum, pass) { return sum + Number(pass.penalizacion || 0); }, 0));
    return {
      inscripcionId: inscription.inscripcionId, pilotoId: inscription.pilotoId, categoriaId: inscription.categoriaId,
      piloto: inscription.piloto, categoria: inscription.categoria, pasadas: passes,
      descartada: discarded ? discarded.label + ' - ' + discarded.total : '', penalizaciones: penalties,
      total: total, gap: '', completadas: registered.length, previstas: definition.totalPasadas,
      firmaCompletadas: registered.map(function(pass) { return pass.pasadaId; }).sort().join('|'),
      estado: !registered.length ? 'Sin resultados' : complete
        ? (penalties ? 'Completo con penalización' : 'Completo')
        : (penalties ? 'Pendiente con penalización' : 'Pendiente')
    };
  });
  rows.sort(function(a, b) {
    if (a.total === '' && b.total !== '') return 1;
    if (a.total !== '' && b.total === '') return -1;
    if (a.total !== b.total) return Number(a.total) - Number(b.total);
    return a.categoria.localeCompare(b.categoria) || a.piloto.localeCompare(b.piloto);
  });
  rows.forEach(function(row, index) {
    if (row.total === '') return;
    const previous = index ? rows[index - 1] : null;
    row.gap = previous && previous.total !== '' && previous.categoriaId === row.categoriaId &&
      previous.firmaCompletadas === row.firmaCompletadas
      ? round3_(Number(row.total) - Number(previous.total)) : '-';
  });
  return rows;
}

function refreshClasificacion_(inscripciones, resultados) {
  writeObjects_(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CLASIFICACION),
    CLASIFICACION_HEADERS, buildClasificacion_('', inscripciones, resultados));
}

function getClasificacion(pin, categoria) {
  if (!validatePin_(pin)) throw new Error('PIN incorrecto');
  return buildClasificacion_(categoria);
}

function calcularMetricasConstancia_(values, type) {
  const numbers = values.filter(function(row) { return !type || row.tipo === type; }).map(function(row) { return Number(row.valor); });
  if (!numbers.length) return { mejor: '', peor: '', media: '', diferencia: '' };
  const best = Math.min.apply(null, numbers);
  const worst = Math.max.apply(null, numbers);
  return {
    mejor: round3_(best), peor: round3_(worst),
    media: round3_(numbers.reduce(function(sum, value) { return sum + value; }, 0) / numbers.length),
    diferencia: numbers.length > 1 ? round3_(worst - best) : ''
  };
}

function buildDashboardConstancia_() {
  return buildClasificacion_('').map(function(row) {
    const values = row.pasadas.filter(function(pass) { return pass.total !== ''; }).map(function(pass) {
      return { pasadaId: pass.pasadaId, tramoId: pass.tramoId, label: pass.label, tipo: pass.tipo, valor: pass.total };
    });
    const all = calcularMetricasConstancia_(values, '');
    const idas = calcularMetricasConstancia_(values, 'ida');
    const vueltas = calcularMetricasConstancia_(values, 'vuelta');
    let accumulated = 0;
    return Object.assign({
      inscripcionId: row.inscripcionId, pilotoId: row.pilotoId, categoriaId: row.categoriaId,
      piloto: row.piloto, categoria: row.categoria, pasadas: row.pasadas,
      tiemposPorPasada: values,
      progresionAcumulada: row.pasadas.map(function(pass) {
        if (pass.total === '') return { pasadaId: pass.pasadaId, tramoId: pass.tramoId, label: pass.label, tipo: pass.tipo, valor: '' };
        accumulated = round3_(accumulated + Number(pass.total));
        return { pasadaId: pass.pasadaId, tramoId: pass.tramoId, label: pass.label, tipo: pass.tipo, valor: accumulated };
      }),
      completadas: row.completadas, previstas: row.previstas, progreso: row.completadas + '/' + row.previstas,
      estado: row.estado
    }, {
      mejor: all.mejor, peor: all.peor, media: all.media, diferencia: all.diferencia,
      mejorIdas: idas.mejor, peorIdas: idas.peor, mediaIdas: idas.media, diferenciaIdas: idas.diferencia,
      mejorVueltas: vueltas.mejor, peorVueltas: vueltas.peor, mediaVueltas: vueltas.media, diferenciaVueltas: vueltas.diferencia
    });
  });
}

function getDashboardData(pin) {
  if (!validatePin_(pin)) throw new Error('PIN incorrecto');
  const definition = getRaceDefinition_();
  const rows = buildDashboardConstancia_();
  return {
    carrera: getConfigValue_('CARRERA_ACTIVA') || 'Rally RC', updatedAt: nowString_(),
    tramos: definition.tramos, pasadas: definition.pasadas,
    categorias: Array.from(new Set(rows.map(function(row) { return row.categoria; }))).sort(),
    pilotos: Array.from(new Set(rows.map(function(row) { return row.piloto; }))).sort(), rows: rows
  };
}

function getDashboardPilotos(pin, categoria) {
  if (!validatePin_(pin)) throw new Error('PIN incorrecto');
  return Array.from(new Set(getInscripciones_().filter(function(row) {
    return !categoria || row.categoria === categoria || row.categoriaId === categoria;
  }).map(function(row) { return row.piloto; }))).sort();
}

function getAdminCatalogosData(pin, adminPin) {
  assertCredentials_(pin, adminPin);
  return { pilotos: getPilotosDB_(), categorias: getCategoriasDB_(), inscripciones: getInscripciones_() };
}

function getPilotoDBById_(id) {
  return getPilotosDB_().find(function(row) { return row.pilotoId === cellText_(id); });
}

function getCategoriaDBByIdOrName_(value) {
  const text = cellText_(value);
  return getCategoriasDB_().find(function(row) { return row.categoriaId === text || normalizeText_(row.nombre) === normalizeText_(text); });
}

function validateEnrollment_(payload, currentId, inscriptions) {
  payload = payload || {};
  const pilot = getPilotoDBById_(payload.pilotoId);
  const category = getCategoriaDBByIdOrName_(payload.categoriaId || payload.categoria);
  if (!pilot || pilot.activo.toUpperCase() === 'NO') throw new Error('Piloto inexistente o inactivo');
  if (!category || category.activa.toUpperCase() === 'NO') throw new Error('Categoría inexistente o inactiva');
  if ((inscriptions || getInscripciones_()).some(function(row) {
    return row.inscripcionId !== currentId && row.pilotoId === pilot.pilotoId && row.categoriaId === category.categoriaId;
  })) throw new Error('Este piloto ya está inscrito en esta categoría');
  return { piloto: pilot, categoria: category };
}

function getNextInscripcionId_() {
  return getNextId_('I', getInscripciones_().map(function(row) { return row.inscripcionId; }));
}

function getInscripcionRowIndex_(id) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_INSCRIPCIONES);
  const values = sheet.getDataRange().getValues();
  const map = getHeaderMapFromHeaders_(values[0]);
  for (let i = 1; i < values.length; i++) if (cellText_(values[i][map.inscripcionId]) === cellText_(id)) return i + 1;
  return -1;
}

function inscripcionTieneResultados_(id) {
  return getResultados_().some(function(row) { return row.inscripcionId === cellText_(id); });
}

function crearInscripcionAdmin(pin, adminPin, payload) {
  assertCredentials_(pin, adminPin);
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    assertRaceMutable_();
    const config = getConfigValues_();
    assertExpectedRaceIfProvided_(config, payload);
    const validated = validateEnrollment_(payload, '', getInscripciones_());
    const row = {
      inscripcionId: getNextInscripcionId_(), pilotoId: validated.piloto.pilotoId,
      categoriaId: validated.categoria.categoriaId, piloto: validated.piloto.nombre, categoria: validated.categoria.nombre
    };
    appendObject_(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_INSCRIPCIONES), row);
    refreshClasificacion_();
    return { status: 'OK', message: 'Inscripción creada correctamente', inscripcion: row };
  } finally { lock.releaseLock(); }
}

function editarInscripcionAdmin(pin, adminPin, payload) {
  assertCredentials_(pin, adminPin);
  payload = payload || {};
  const id = cellText_(payload.inscripcionId);
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    assertRaceMutable_();
    assertExpectedRaceIfProvided_(getConfigValues_(), payload);
    if (!id) throw new Error('inscripcionId es obligatorio');
    if (inscripcionTieneResultados_(id)) throw new Error('No se puede editar una inscripción con resultados');
    const validated = validateEnrollment_(payload, id, getInscripciones_());
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_INSCRIPCIONES);
    const index = getInscripcionRowIndex_(id);
    if (index === -1) throw new Error('No se ha encontrado la inscripción');
    const map = getHeaderMap_(sheet);
    const row = { inscripcionId: id, pilotoId: validated.piloto.pilotoId, categoriaId: validated.categoria.categoriaId,
      piloto: validated.piloto.nombre, categoria: validated.categoria.nombre };
    INSCRIPCIONES_HEADERS.forEach(function(header) { sheet.getRange(index, map[header] + 1).setValue(row[header]); });
    refreshClasificacion_();
    return { status: 'OK', message: 'Inscripción actualizada correctamente', inscripcion: row };
  } finally { lock.releaseLock(); }
}

function eliminarInscripcionAdmin(pin, adminPin, inscripcionId) {
  assertCredentials_(pin, adminPin);
  const request = typeof inscripcionId === 'object' ? inscripcionId : { inscripcionId: inscripcionId };
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    assertRaceMutable_();
    assertExpectedRaceIfProvided_(getConfigValues_(), request);
    const id = cellText_(request.inscripcionId);
    if (!id) throw new Error('inscripcionId es obligatorio');
    if (inscripcionTieneResultados_(id)) throw new Error('No se puede eliminar una inscripción con resultados');
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_INSCRIPCIONES);
    const index = getInscripcionRowIndex_(id);
    if (index === -1) throw new Error('No se ha encontrado la inscripción');
    sheet.deleteRow(index);
    refreshClasificacion_();
    return { status: 'OK', message: 'Inscripción eliminada correctamente' };
  } finally { lock.releaseLock(); }
}

function createPilotUnlocked_(payload) {
  payload = payload || {};
  const name = cellText_(payload.nombre);
  if (!name) throw new Error('El nombre del piloto es obligatorio');
  const pilots = getPilotosDB_();
  if (pilots.some(function(row) { return normalizeText_(row.nombre) === normalizeText_(name); })) throw new Error('Ya existe un piloto con ese nombre');
  const now = nowString_();
  const pilot = {
    pilotoId: getNextId_('P', pilots.map(function(row) { return row.pilotoId; })), nombre: name,
    alias: cellText_(payload.alias), activo: cellText_(payload.activo).toUpperCase() === 'NO' ? 'NO' : 'SI',
    notas: cellText_(payload.notas), createdAt: now, updatedAt: now
  };
  appendObject_(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PILOTOS_DB), pilot);
  return pilot;
}

function crearPilotoAdmin(pin, adminPin, payload) {
  assertCredentials_(pin, adminPin);
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    assertRaceMutable_();
    const pilot = createPilotUnlocked_(payload);
    return { status: 'OK', message: 'Piloto creado correctamente', piloto: pilot };
  } finally { lock.releaseLock(); }
}

function editarPilotoAdmin(pin, adminPin, payload) {
  assertCredentials_(pin, adminPin);
  payload = payload || {};
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    assertRaceMutable_();
    const id = cellText_(payload.pilotoId);
    const name = cellText_(payload.nombre);
    const pilots = getPilotosDB_();
    if (!id || !name) throw new Error('pilotoId y nombre son obligatorios');
    if (pilots.some(function(row) { return row.pilotoId !== id && normalizeText_(row.nombre) === normalizeText_(name); })) throw new Error('Ya existe otro piloto con ese nombre');
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PILOTOS_DB);
    const values = sheet.getDataRange().getValues();
    const map = getHeaderMapFromHeaders_(values[0]);
    let index = -1;
    for (let i = 1; i < values.length; i++) if (cellText_(values[i][map.pilotoId]) === id) index = i + 1;
    if (index === -1) throw new Error('No se ha encontrado el piloto');
    sheet.getRange(index, map.nombre + 1).setValue(name);
    sheet.getRange(index, map.alias + 1).setValue(cellText_(payload.alias));
    sheet.getRange(index, map.activo + 1).setValue(cellText_(payload.activo).toUpperCase() === 'NO' ? 'NO' : 'SI');
    sheet.getRange(index, map.notas + 1).setValue(cellText_(payload.notas));
    sheet.getRange(index, map.updatedAt + 1).setValue(nowString_());
    syncActiveDenormalizedName_('pilotoId', id, 'piloto', name);
    refreshClasificacion_();
    return { status: 'OK', message: 'Piloto actualizado correctamente' };
  } finally { lock.releaseLock(); }
}

function createCategoryUnlocked_(payload) {
  payload = payload || {};
  const name = cellText_(payload.nombre);
  if (!name) throw new Error('El nombre de la categoría es obligatorio');
  const categories = getCategoriasDB_();
  if (categories.some(function(row) { return normalizeText_(row.nombre) === normalizeText_(name); })) throw new Error('Ya existe una categoría con ese nombre');
  const now = nowString_();
  const category = {
    categoriaId: getNextId_('C', categories.map(function(row) { return row.categoriaId; })), nombre: name,
    activa: cellText_(payload.activa).toUpperCase() === 'NO' ? 'NO' : 'SI',
    orden: Number(payload.orden) > 0 ? Number(payload.orden) : categories.length + 1,
    notas: cellText_(payload.notas), createdAt: now, updatedAt: now
  };
  appendObject_(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CATEGORIAS_DB), category);
  return category;
}

function crearCategoriaAdmin(pin, adminPin, payload) {
  assertCredentials_(pin, adminPin);
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    assertRaceMutable_();
    const category = createCategoryUnlocked_(payload);
    return { status: 'OK', message: 'Categoría creada correctamente', categoria: category };
  } finally { lock.releaseLock(); }
}

function editarCategoriaAdmin(pin, adminPin, payload) {
  assertCredentials_(pin, adminPin);
  payload = payload || {};
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    assertRaceMutable_();
    const id = cellText_(payload.categoriaId);
    const name = cellText_(payload.nombre);
    const categories = getCategoriasDB_();
    if (!id || !name) throw new Error('categoriaId y nombre son obligatorios');
    if (categories.some(function(row) { return row.categoriaId !== id && normalizeText_(row.nombre) === normalizeText_(name); })) throw new Error('Ya existe otra categoría con ese nombre');
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CATEGORIAS_DB);
    const values = sheet.getDataRange().getValues();
    const map = getHeaderMapFromHeaders_(values[0]);
    let index = -1;
    for (let i = 1; i < values.length; i++) if (cellText_(values[i][map.categoriaId]) === id) index = i + 1;
    if (index === -1) throw new Error('No se ha encontrado la categoría');
    sheet.getRange(index, map.nombre + 1).setValue(name);
    sheet.getRange(index, map.activa + 1).setValue(cellText_(payload.activa).toUpperCase() === 'NO' ? 'NO' : 'SI');
    sheet.getRange(index, map.orden + 1).setValue(Number(payload.orden) > 0 ? Number(payload.orden) : 999);
    sheet.getRange(index, map.notas + 1).setValue(cellText_(payload.notas));
    sheet.getRange(index, map.updatedAt + 1).setValue(nowString_());
    syncActiveDenormalizedName_('categoriaId', id, 'categoria', name);
    refreshClasificacion_();
    return { status: 'OK', message: 'Categoría actualizada correctamente' };
  } finally { lock.releaseLock(); }
}

function syncActiveDenormalizedName_(idHeader, id, nameHeader, name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inscriptionSheet = ss.getSheetByName(SHEET_INSCRIPCIONES);
  const inscriptionIds = {};
  if (inscriptionSheet.getLastRow() >= 2) {
    const values = inscriptionSheet.getDataRange().getValues();
    const map = getHeaderMapFromHeaders_(values[0]);
    for (let i = 1; i < values.length; i++) {
      if (cellText_(values[i][map[idHeader]]) === id) {
        inscriptionIds[cellText_(values[i][map.inscripcionId])] = true;
        inscriptionSheet.getRange(i + 1, map[nameHeader] + 1).setValue(name);
      }
    }
  }
  const resultSheet = ss.getSheetByName(SHEET_RESULTADOS);
  if (!Object.keys(inscriptionIds).length || resultSheet.getLastRow() < 2) return;
  const resultValues = resultSheet.getDataRange().getValues();
  const resultMap = getHeaderMapFromHeaders_(resultValues[0]);
  const carreraId = getConfigValue_(CONFIG_CARRERA_ID);
  for (let i = 1; i < resultValues.length; i++) {
    if (cellText_(resultValues[i][resultMap.carreraId]) === carreraId &&
        inscriptionIds[cellText_(resultValues[i][resultMap.inscripcionId])]) {
      resultSheet.getRange(i + 1, resultMap[nameHeader] + 1).setValue(name);
    }
  }
}

function crearPilotoEInscribirAdmin(pin, adminPin, payload) {
  assertCredentials_(pin, adminPin);
  payload = payload || {};
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    assertRaceMutable_();
    const config = getConfigValues_();
    assertExpectedRaceIfProvided_(config, payload);
    const category = getCategoriaDBByIdOrName_(payload.categoriaId || payload.categoria);
    if (!category || category.activa.toUpperCase() === 'NO') throw new Error('Categoría inexistente o inactiva');
    const name = cellText_(payload.nombre);
    if (!name) throw new Error('El nombre del piloto es obligatorio');
    if (getPilotosDB_().some(function(row) { return normalizeText_(row.nombre) === normalizeText_(name); })) throw new Error('Ya existe un piloto con ese nombre');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const pilotSheet = ss.getSheetByName(SHEET_PILOTOS_DB);
    const inscriptionSheet = ss.getSheetByName(SHEET_INSCRIPCIONES);
    const pilotLastRow = pilotSheet.getLastRow();
    const inscriptionLastRow = inscriptionSheet.getLastRow();
    const pilot = createPilotUnlocked_(Object.assign({}, payload, { activo: 'SI' }));
    const row = {
      inscripcionId: getNextInscripcionId_(), pilotoId: pilot.pilotoId, categoriaId: category.categoriaId,
      piloto: pilot.nombre, categoria: category.nombre
    };
    try {
      appendObject_(inscriptionSheet, row);
      refreshClasificacion_();
    } catch (error) {
      if (inscriptionSheet.getLastRow() > inscriptionLastRow) {
        inscriptionSheet.deleteRow(inscriptionSheet.getLastRow());
      }
      if (pilotSheet.getLastRow() > pilotLastRow) {
        pilotSheet.deleteRow(pilotSheet.getLastRow());
      }
      try {
        refreshClasificacion_();
      } catch (refreshError) {
        // Preserve the original write error; active pilot/enrollment rows are already rolled back.
      }
      throw error;
    }
    return { status: 'OK', message: 'Piloto creado e inscrito correctamente', piloto: pilot, inscripcion: row };
  } finally { lock.releaseLock(); }
}

function normalizeRaceEnrollments_(rows) {
  if (!Array.isArray(rows)) throw new Error('Las inscripciones de la nueva carrera no son válidas');
  const pilots = {};
  getPilotosDB_().forEach(function(row) { pilots[row.pilotoId] = row; });
  const categories = {};
  getCategoriasDB_().forEach(function(row) { categories[row.categoriaId] = row; });
  const ids = {};
  const pairs = {};
  return rows.map(function(raw) {
    const pilot = pilots[cellText_(raw.pilotoId)];
    const category = categories[cellText_(raw.categoriaId)];
    if (!pilot || pilot.activo.toUpperCase() === 'NO' || !category || category.activa.toUpperCase() === 'NO') {
      throw new Error('La nueva carrera contiene una inscripción no canónica o inactiva');
    }
    const id = cellText_(raw.inscripcionId) || Utilities.getUuid();
    const pair = pilot.pilotoId + '|' + category.categoriaId;
    if (ids[id] || pairs[pair]) throw new Error('La nueva carrera contiene inscripciones duplicadas');
    ids[id] = true;
    pairs[pair] = true;
    return { inscripcionId: id, pilotoId: pilot.pilotoId, categoriaId: category.categoriaId, piloto: pilot.nombre, categoria: category.nombre };
  });
}

function backupActiveRaceSheets_(ss, label) {
  const suffix = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm-ss');
  [SHEET_CONFIG, SHEET_TRAMOS, SHEET_PASADAS, SHEET_INSCRIPCIONES, SHEET_RESULTADOS, SHEET_CLASIFICACION]
    .forEach(function(name) {
      const sheet = ss.getSheetByName(name);
      if (sheet) sheet.copyTo(ss).setName(uniqueBackupName_(ss, name + '_backup_' + label + '_' + suffix));
    });
}

function generarNuevaCarrera(pin, adminPin, nombreCarrera, tramosPayload) {
  assertCredentials_(pin, adminPin);
  const request = nombreCarrera && typeof nombreCarrera === 'object'
    ? nombreCarrera : { nombre: nombreCarrera, tramos: tramosPayload };
  const name = cellText_(request.nombre || request.nombreCarrera);
  if (!name || name.length > 200) throw new Error('Nombre de carrera inválido');
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const config = getConfigValues_();
    assertExpectedRace_(config, request.carreraId, request.revision);
    const publicationStatus = recoverCarreraPublicationStatus_(config);
    if (publicationStatus === 'PUBLICANDO') throw new Error('Espera a que termine la publicación');
    if (publicationStatus === 'PUBLICACION_INCIERTA') {
      throw new Error('Reintenta la publicación incierta antes de crear una carrera nueva');
    }
    const definition = buildDefinitionFromTramos_(normalizeTramosConfig_(request.tramos), null);
    const enrollments = normalizeRaceEnrollments_(Array.isArray(request.inscripciones) ? request.inscripciones : getInscripciones_());
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    backupActiveRaceSheets_(ss, 'nueva_carrera');
    writeRaceDefinition_(definition);
    writeObjects_(ss.getSheetByName(SHEET_INSCRIPCIONES), INSCRIPCIONES_HEADERS, enrollments);
    writeObjects_(ss.getSheetByName(SHEET_RESULTADOS), RESULTADOS_HEADERS, []);
    updateConfigValue_('CARRERA_ACTIVA', name);
    updateConfigValue_(CONFIG_CARRERA_ID, Utilities.getUuid());
    updateConfigValue_(CONFIG_CARRERA_ESTADO, 'CONFIGURACION');
    updateConfigValue_(CONFIG_REVISION, '1');
    updateConfigValue_(CONFIG_CAMPEONATO_PUBLICADA, 'NO');
    updateConfigValue_(CONFIG_CAMPEONATO_PUBLICACION_INICIADA, '');
    refreshClasificacion_();
    return { status: 'OK', message: 'Nueva carrera creada en estado CONFIGURACION', carrera: name, configuracion: raceConfigResponse_(getConfigValues_()) };
  } finally { lock.releaseLock(); }
}

function getEstadoPublicacionCampeonato(pin, adminPin) {
  assertCredentials_(pin, adminPin);
  const lock = LockService.getScriptLock();
  let config;
  try {
    lock.waitLock(10000);
    config = getConfigValues_();
    recoverCarreraPublicationStatus_(config);
  } finally { lock.releaseLock(); }
  const properties = PropertiesService.getScriptProperties();
  const classification = buildClasificacion_('');
  const expected = getRaceDefinition_().totalPasadas;
  const endpoint = Boolean(cellText_(properties.getProperty(CHAMPIONSHIP_ENDPOINT_PROPERTY)));
  const token = Boolean(String(properties.getProperty(CHAMPIONSHIP_TOKEN_PROPERTY) || ''));
  return {
    carreraId: config[CONFIG_CARRERA_ID], carrera: config.CARRERA_ACTIVA,
    carreraEstado: config[CONFIG_CARRERA_ESTADO], configRevision: Number(config[CONFIG_REVISION]) || 1,
    estado: cellText_(config[CONFIG_CAMPEONATO_PUBLICADA] || 'NO').toUpperCase(), participantes: classification.length,
    incompletos: classification.filter(function(row) { return row.completadas < expected; }).length,
    endpointConfigurado: endpoint, tokenConfigurado: token, configurado: endpoint && token
  };
}

function buildCampeonatoSnapshot_(config) {
  if (config[CONFIG_CARRERA_ESTADO] !== 'INICIADA') throw new Error('La carrera debe estar iniciada');
  const definition = getRaceDefinition_();
  const inscriptions = getInscripciones_();
  if (!inscriptions.length) throw new Error('No hay inscripciones para publicar');
  const rows = [];
  Array.from(new Set(inscriptions.map(function(row) { return row.categoriaId; }))).forEach(function(categoryId) {
    buildClasificacion_(categoryId).forEach(function(row, index) {
      rows.push({
        inscripcionId: row.inscripcionId, pilotoId: row.pilotoId, piloto: row.piloto,
        categoriaId: row.categoriaId, categoria: row.categoria, posicion: index + 1,
        pasadas: row.pasadas.filter(function(pass) { return pass.total !== ''; }).map(function(pass) {
          return { pasadaId: pass.pasadaId, tramoId: pass.tramoId, tiempo: pass.tiempo,
            penalizacion: pass.penalizacion, total: pass.total, descartada: pass.descartada };
        }),
        descartada: row.descartada, penalizaciones: row.penalizaciones, total: row.total,
        gap: row.gap, completadas: row.completadas, previstas: row.previstas, estado: row.estado
      });
    });
  });
  const snapshot = {
    schemaVersion: 2,
    carrera: {
      carreraId: config[CONFIG_CARRERA_ID], nombre: config.CARRERA_ACTIVA,
      definicion: {
        tramos: definition.tramos.map(function(stage) { return { tramoId: stage.tramoId, nombre: stage.nombre, orden: stage.orden }; }),
        pasadas: definition.pasadas, totalPasadas: definition.totalPasadas, numDescartes: definition.numDescartes
      }
    },
    fechaPublicacion: new Date().toISOString(),
    incompletos: rows.filter(function(row) { return row.completadas < row.previstas; }).length,
    resultados: rows
  };
  snapshot.payloadHash = sha256Hex_(JSON.stringify({
    schemaVersion: snapshot.schemaVersion,
    carrera: snapshot.carrera,
    incompletos: snapshot.incompletos,
    resultados: snapshot.resultados
  }));
  return snapshot;
}

function sha256Hex_(value) {
  return Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    value,
    Utilities.Charset.UTF_8
  ).map(function(byte) {
    return ('0' + ((byte + 256) % 256).toString(16)).slice(-2);
  }).join('');
}

function setCarreraPublicationStatus_(carreraId, status) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const config = getConfigValues_();
    if (config[CONFIG_CARRERA_ID] === carreraId) {
      updateConfigValue_(CONFIG_CAMPEONATO_PUBLICADA, status);
      updateConfigValue_(CONFIG_CAMPEONATO_PUBLICACION_INICIADA, '');
    }
  } finally { lock.releaseLock(); }
}

function publicarResultadosCampeonato(pin, adminPin, expectedCarreraId, expectedRevision) {
  assertCredentials_(pin, adminPin);
  const request = typeof expectedCarreraId === 'object'
    ? expectedCarreraId : { carreraId: expectedCarreraId, revision: expectedRevision };
  const properties = PropertiesService.getScriptProperties();
  const endpoint = cellText_(properties.getProperty(CHAMPIONSHIP_ENDPOINT_PROPERTY));
  const token = String(properties.getProperty(CHAMPIONSHIP_TOKEN_PROPERTY) || '');
  if (!endpoint || !token) throw new Error('Configura CHAMPIONSHIP_ENDPOINT y CHAMPIONSHIP_TOKEN');
  const lock = LockService.getScriptLock();
  let snapshot;
  try {
    lock.waitLock(10000);
    const config = getConfigValues_();
    assertExpectedRace_(config, request.carreraId, request.revision);
    const status = recoverCarreraPublicationStatus_(config);
    if (status === 'SI') throw new Error('La carrera ya está publicada');
    if (status === 'PUBLICANDO') throw new Error('Ya hay una publicación en curso');
    snapshot = buildCampeonatoSnapshot_(config);
    updateConfigValue_(CONFIG_CAMPEONATO_PUBLICADA, 'PUBLICANDO');
    updateConfigValue_(CONFIG_CAMPEONATO_PUBLICACION_INICIADA, new Date().toISOString());
  } finally { lock.releaseLock(); }
  try {
    const response = UrlFetchApp.fetch(endpoint, {
      method: 'post', contentType: 'application/json', payload: JSON.stringify({ token: token, payload: snapshot }),
      followRedirects: true, muteHttpExceptions: true
    });
    let data;
    try { data = JSON.parse(response.getContentText()); }
    catch (error) { throw new Error('El campeonato no devolvió JSON válido (HTTP ' + response.getResponseCode() + ')'); }
    if (data && data.ok === false) {
      const rejection = new Error(data.message || 'La publicación fue rechazada');
      rejection.publicationDefinitive = true;
      throw rejection;
    }
    if (!data || data.ok !== true) throw new Error('La respuesta del campeonato es ambigua');
    setCarreraPublicationStatus_(snapshot.carrera.carreraId, 'SI');
    return {
      status: data.status, message: 'Resultados publicados correctamente', carreraId: snapshot.carrera.carreraId,
      incompletos: snapshot.incompletos, payloadHash: snapshot.payloadHash
    };
  } catch (error) {
    setCarreraPublicationStatus_(
      snapshot.carrera.carreraId,
      error && error.publicationDefinitive ? 'NO' : 'PUBLICACION_INCIERTA'
    );
    throw error;
  }
}

function resetDemoData_() {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    assertRaceMutable_();
    writeObjects_(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_RESULTADOS), RESULTADOS_HEADERS, []);
    refreshClasificacion_();
  } finally { lock.releaseLock(); }
}
