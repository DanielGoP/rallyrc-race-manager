const SHEET_CONFIG = 'Config';
const SHEET_INSCRIPCIONES = 'Inscripciones';
const SHEET_RESULTADOS = 'Resultados';
const SHEET_CLASIFICACION = 'Clasificacion';
const SHEET_PILOTOS_DB = 'PilotosDB';
const SHEET_CATEGORIAS_DB = 'CategoriasDB';
const CONFIG_SCHEMA_VERSION = 'SCHEMA_VERSION';
const CONFIG_ENVIRONMENT = 'ENVIRONMENT';
const CONFIG_CARRERA_ID = 'CARRERA_ID';
const CONFIG_CAMPEONATO_PUBLICADA = 'CAMPEONATO_PUBLICADA';
const CONFIG_CAMPEONATO_PUBLICACION_INICIADA = 'CAMPEONATO_PUBLICACION_INICIADA';
// Son nombres de Script Properties; no sustituirlos por la URL ni por el token.
const CHAMPIONSHIP_ENDPOINT_PROPERTY = 'CHAMPIONSHIP_ENDPOINT';
const CHAMPIONSHIP_TOKEN_PROPERTY = 'CHAMPIONSHIP_TOKEN';
const CHAMPIONSHIP_PUBLICATION_LEASE_MS = 10 * 60 * 1000;
const CURRENT_SCHEMA_VERSION = '3';

const PASADAS = [
  { id: 'P1', num: 1, label: 'Ida 1', tipo: 'ida', orden: 1 },
  { id: 'P2', num: 2, label: 'Vuelta 1', tipo: 'vuelta', orden: 2 },
  { id: 'P3', num: 3, label: 'Ida 2', tipo: 'ida', orden: 3 },
  { id: 'P4', num: 4, label: 'Vuelta 2', tipo: 'vuelta', orden: 4 },
  { id: 'P5', num: 5, label: 'Ida 3', tipo: 'ida', orden: 5 },
  { id: 'P6', num: 6, label: 'Vuelta 3', tipo: 'vuelta', orden: 6 }
];

function doGet() {
  setupSheets();

  return HtmlService
    .createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Rally RC - Resultados')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  createSheetIfNotExists_(ss, SHEET_CONFIG, [
  ['key', 'value'],
  ['PIN', 'CAMBIAR_PIN_ACCESO'],
  ['ADMIN_PIN', 'CAMBIAR_PIN_ADMIN'],
  ['CARRERA_ACTIVA', 'Rally RC'],
  [CONFIG_CARRERA_ID, Utilities.getUuid()],
  [CONFIG_CAMPEONATO_PUBLICADA, 'NO'],
  [CONFIG_CAMPEONATO_PUBLICACION_INICIADA, ''],
  [CONFIG_ENVIRONMENT, 'PRODUCTION'],
  [CONFIG_SCHEMA_VERSION, '0']
]);
  createSheetIfNotExists_(ss, SHEET_PILOTOS_DB, [
  [
    'pilotoId',
    'nombre',
    'alias',
    'activo',
    'notas',
    'createdAt',
    'updatedAt'
  ]
]);

createSheetIfNotExists_(ss, SHEET_CATEGORIAS_DB, [
  [
    'categoriaId',
    'nombre',
    'activa',
    'orden',
    'notas',
    'createdAt',
    'updatedAt'
  ]
]);

  createSheetIfNotExists_(ss, SHEET_INSCRIPCIONES, [
    ['inscripcionId', 'pilotoId', 'dorsal', 'piloto', 'categoria'],
    ['I001', 'P001', '1', 'Piloto Demo 1', 'Rally 1/10'],
    ['I002', 'P002', '2', 'Piloto Demo 2', 'Rally 1/10'],
    ['I003', 'P003', '3', 'Piloto Demo 3', 'Rally 1/10'],
    ['I004', 'P001', '1', 'Piloto Demo 1', 'Clásicos']
  ]);

  createSheetIfNotExists_(ss, SHEET_RESULTADOS, [
    [
      'resultadoId',
      'timestamp',
      'inscripcionId',
      'dorsal',
      'piloto',
      'categoria',
      'pasada',
      'pasadaLabel',
      'tiempo',
      'penalizacion',
      'total',
      'juez'
    ]
  ]);

  createSheetIfNotExists_(ss, SHEET_CLASIFICACION, [
   [
    'categoria',
    'dorsal',
    'piloto',
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
    'estado'
  ]
  ]);

  const storedSchemaVersion = parseSchemaVersion_(getConfigValue_(CONFIG_SCHEMA_VERSION));
  const currentSchemaVersion = parseSchemaVersion_(CURRENT_SCHEMA_VERSION);

  if (storedSchemaVersion > currentSchemaVersion) {
    throw new Error('La hoja usa un esquema más reciente que esta versión de la aplicación');
  }

  if (storedSchemaVersion < currentSchemaVersion) {
    const lock = LockService.getScriptLock();

    try {
      lock.waitLock(10000);

      const lockedSchemaVersion = parseSchemaVersion_(getConfigValue_(CONFIG_SCHEMA_VERSION));

      if (lockedSchemaVersion > currentSchemaVersion) {
        throw new Error('La hoja usa un esquema más reciente que esta versión de la aplicación');
      }

      if (lockedSchemaVersion < 2) {
        migrarPilotosYCategoriasDesdeInscripciones_();
        refreshClasificacion_();
      }

      if (lockedSchemaVersion < 3) {
        ensureCarreraPublicationConfig_();
      }

      if (lockedSchemaVersion < currentSchemaVersion) {
        updateConfigValue_(CONFIG_SCHEMA_VERSION, CURRENT_SCHEMA_VERSION);
      }
    } finally {
      lock.releaseLock();
    }
  }
}

function parseSchemaVersion_(value) {
  const normalized = String(value || '0').trim();

  if (!/^\d+$/.test(normalized)) {
    throw new Error('SCHEMA_VERSION no es un entero válido');
  }

  return Number(normalized);
}

function createSheetIfNotExists_(ss, sheetName, initialRows) {
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.getRange(1, 1, initialRows.length, initialRows[0].length).setValues(initialRows);
    sheet.setFrozenRows(1);
    autoResize_(sheet);
  }
}

function autoResize_(sheet) {
  const lastColumn = sheet.getLastColumn();
  if (lastColumn > 0) {
    sheet.autoResizeColumns(1, lastColumn);
  }
}

function getConfigValue_(key) {
  return getConfigValues_()[key] || '';
}

function getConfigValues_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_CONFIG);
  const values = sheet.getDataRange().getValues();
  const config = {};

  for (let i = 1; i < values.length; i++) {
    const key = String(values[i][0] || '').trim();

    if (key) {
      config[key] = String(values[i][1] || '').trim();
    }
  }

  return config;
}

function validatePin_(pin) {
  const configuredPin = getConfigValue_('PIN');
  return String(pin || '').trim() === configuredPin;
}

function validatePin(pin) {
  return validatePin_(pin);
}

function validateAdminPin_(adminPin) {
  const configuredAdminPin = getConfigValue_('ADMIN_PIN');
  return String(adminPin || '').trim() === configuredAdminPin;
}

function validateAdminPin(adminPin) {
  return validateAdminPin_(adminPin);
}

function getCategorias(pin) {
  if (!validatePin_(pin)) {
    throw new Error('PIN incorrecto');
  }

  return getCategoriasActivas_();
}

function getRegistroData(pin) {
  const config = getConfigValues_();

  if (String(pin || '').trim() !== String(config.PIN || '').trim()) {
    throw new Error('PIN incorrecto');
  }

  return {
    categorias: getCategoriasActivas_(),
    inscripciones: getInscripciones_(),
    fetchedAt: Date.now(),
    schemaVersion: CURRENT_SCHEMA_VERSION,
    environment: config[CONFIG_ENVIRONMENT] || 'PRODUCTION'
  };
}

function getCategoriasActivas_() {
  const categoriasFromDb = getCategoriasDB_()
    .filter(categoria => String(categoria.activa || '').toUpperCase() !== 'NO')
    .map(categoria => categoria.nombre)
    .filter(Boolean);

  if (categoriasFromDb.length > 0) {
    return categoriasFromDb;
  }

  const inscripciones = getInscripciones_();

  return [...new Set(inscripciones.map(item => item.categoria))]
    .filter(Boolean)
    .sort();
}

function getInscripcionesByCategoria(pin, categoria) {
  if (!validatePin_(pin)) {
    throw new Error('PIN incorrecto');
  }

  const cleanCategoria = String(categoria || '').trim();

  return getInscripciones_()
    .filter(item => String(item.categoria || '').trim() === cleanCategoria)
    .sort((a, b) => {
      const dorsalA = Number(a.dorsal);
      const dorsalB = Number(b.dorsal);

      if (!isNaN(dorsalA) && !isNaN(dorsalB)) {
        return dorsalA - dorsalB;
      }

      return String(a.dorsal).localeCompare(String(b.dorsal));
    });
}

function getInscripciones_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_INSCRIPCIONES);

  if (!sheet || sheet.getLastRow() < 2) {
    return [];
  }

  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(function(header) {
    return String(header || '').trim();
  });

  const result = [];

  const hasPilotoId = headers.includes('pilotoId');

  if (hasPilotoId) {
    const map = getHeaderMapFromHeaders_(headers);

    for (let i = 1; i < values.length; i++) {
      const row = values[i];

      const inscripcionId = String(row[map.inscripcionId] || '').trim();
      const pilotoId = String(row[map.pilotoId] || '').trim();
      const dorsal = String(row[map.dorsal] || '').trim();
      const piloto = String(row[map.piloto] || '').trim();
      const categoria = String(row[map.categoria] || '').trim();

      if (!inscripcionId || !piloto || !categoria) {
        continue;
      }

      result.push({
        inscripcionId: inscripcionId,
        pilotoId: pilotoId,
        dorsal: dorsal,
        piloto: piloto,
        categoria: categoria
      });
    }

    return result;
  }

  for (let i = 1; i < values.length; i++) {
    const row = values[i];

    const inscripcionId = String(row[0] || '').trim();
    const dorsal = String(row[1] || '').trim();
    const piloto = String(row[2] || '').trim();
    const categoria = String(row[3] || '').trim();

    if (!inscripcionId || !piloto || !categoria) {
      continue;
    }

    result.push({
      inscripcionId: inscripcionId,
      pilotoId: '',
      dorsal: dorsal,
      piloto: piloto,
      categoria: categoria
    });
  }

  return result;
}

function saveResultado(pin, payload) {
  if (!validatePin_(pin)) {
    throw new Error('PIN incorrecto');
  }

  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(10000);
    assertCarreraAdmiteResultados_();

    const data = normalizePayload_(payload);
    const inscripciones = getInscripciones_();
    const inscripcion = inscripciones.find(function(item) {
      return item.inscripcionId === data.inscripcionId;
    });

    if (!inscripcion) {
      throw new Error('No se ha encontrado la inscripción seleccionada');
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_RESULTADOS);
    const values = sheet.getDataRange().getValues();

    const existingRowIndex = findResultadoRowIndex_(
      values,
      data.inscripcionId,
      data.pasada
    );

    const pasadaLabel = getPasadaLabel_(data.pasada);
    const total = round3_(data.tiempo + data.penalizacion);

    if (existingRowIndex !== -1) {
      const existing = values[existingRowIndex - 1];

      return {
        status: 'DUPLICATE',
        message: 'Ya existe un resultado para esta inscripción y pasada. No se ha guardado ningún cambio.',
        existing: {
          resultadoId: String(existing[0] || ''),
          timestamp: formatDateForClient_(existing[1]),
          inscripcionId: String(existing[2] || ''),
          dorsal: String(existing[3] || ''),
          piloto: String(existing[4] || ''),
          categoria: String(existing[5] || ''),
          pasada: String(existing[6] || ''),
          pasadaLabel: String(existing[7] || ''),
          tiempo: String(existing[8] || ''),
          penalizacion: String(existing[9] || ''),
          total: String(existing[10] || ''),
          juez: String(existing[11] || '')
        }
      };
    }

    const row = [
      existingRowIndex !== -1 ? values[existingRowIndex - 1][0] : Utilities.getUuid(),
      new Date(),
      inscripcion.inscripcionId,
      inscripcion.dorsal,
      inscripcion.piloto,
      inscripcion.categoria,
      data.pasada,
      pasadaLabel,
      data.tiempo,
      data.penalizacion,
      total,
      data.juez
    ];

    if (existingRowIndex !== -1) {
      sheet.getRange(existingRowIndex, 1, 1, row.length).setValues([row]);
    } else {
      sheet.appendRow(row);
    }

    const resultadosActualizados = buildResultadosFromValues_(values.concat([row]));
    refreshClasificacion_(inscripciones, resultadosActualizados);

    const pasadasRegistradas = values
      .slice(1)
      .filter(function(existingRow) {
        return String(existingRow[2] || '').trim() === data.inscripcionId;
      })
      .map(function(existingRow) {
        return Number(existingRow[6]);
      })
      .concat([data.pasada])
      .filter(function(pasada, index, all) {
        return PASADAS.some(function(item) {
          return item.num === pasada;
        }) && all.indexOf(pasada) === index;
      })
      .sort(function(a, b) {
        return a - b;
      });

    return {
      status: existingRowIndex !== -1 ? 'UPDATED' : 'CREATED',
      message: existingRowIndex !== -1 ? 'Resultado actualizado' : 'Resultado guardado',
      saved: {
        dorsal: inscripcion.dorsal,
        piloto: inscripcion.piloto,
        categoria: inscripcion.categoria,
        pasada: data.pasada,
        pasadaLabel,
        tiempo: data.tiempo,
        penalizacion: data.penalizacion,
        total
      },
      pasadasRegistradas: pasadasRegistradas
    };

  } finally {
    lock.releaseLock();
  }
}

function normalizePayload_(payload) {
  if (!payload) {
    throw new Error('No se han recibido datos');
  }

  const inscripcionId = String(payload.inscripcionId || '').trim();
  const pasada = Number(payload.pasada);
  const minutos = Number(payload.minutos || 0);
  const segundos = Number(payload.segundos || 0);
  const decimas = Number(payload.decimas || 0);
  const tiempo = minutos * 60 + segundos + (decimas / 10);
  const penalizacion = parseNumber_(payload.penalizacion || 0);
  const juez = String(payload.juez || '').trim() || 'Sin identificar';
  const overwrite = Boolean(payload.overwrite);

  if (!inscripcionId) {
    throw new Error('Debes seleccionar un piloto');
  }

  if (!PASADAS.some(p => p.num === pasada)) {
    throw new Error('La pasada seleccionada no es válida');
  }

  if (
    isNaN(minutos) || minutos < 0 ||
    isNaN(segundos) || segundos < 0 || segundos > 59 ||
    isNaN(decimas) || decimas < 0 || decimas > 9 ||
    isNaN(tiempo) || tiempo < 0
  ) {
    throw new Error('El tiempo no es válido');
  }

  if (isNaN(penalizacion) || penalizacion < 0) {
    throw new Error('La penalización no es válida');
  }

  return {
    inscripcionId,
    pasada,
    tiempo: round3_(tiempo),
    penalizacion: round3_(penalizacion),
    juez,
    overwrite
  };
}

function parseNumber_(value) {
  if (typeof value === 'number') {
    return value;
  }

  const normalized = String(value || '')
    .trim()
    .replace(',', '.');

  return Number(normalized);
}

function round3_(number) {
  return Math.round(Number(number) * 1000) / 1000;
}

function findInscripcion_(inscripcionId) {
  return getInscripciones_().find(item => item.inscripcionId === inscripcionId);
}

function findResultadoRowIndex_(values, inscripcionId, pasada) {
  for (let i = 1; i < values.length; i++) {
    const rowInscripcionId = String(values[i][2] || '').trim();
    const rowPasada = Number(values[i][6]);

    if (rowInscripcionId === inscripcionId && rowPasada === pasada) {
      return i + 1;
    }
  }

  return -1;
}

function getPasadaLabel_(pasada) {
  const found = PASADAS.find(p => p.num === Number(pasada));
  return found ? found.label : '';
}

function getClasificacion(pin, categoria) {
  if (!validatePin_(pin)) {
    throw new Error('PIN incorrecto');
  }

  return buildClasificacion_(categoria);
}

function refreshClasificacion_(inscripciones, resultados) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_CLASIFICACION);

  const headers = [
    'categoria',
    'dorsal',
    'piloto',
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
    'estado'
  ];

  const rows = buildClasificacion_('', inscripciones, resultados);

  const values = rows.map(function(row) {
    return [
      row.categoria,
      row.dorsal,
      row.piloto,
      row.ida1,
      row.vuelta1,
      row.ida2,
      row.vuelta2,
      row.ida3,
      row.vuelta3,
      row.descartada,
      row.penalizaciones,
      row.total,
      row.gap,
      row.estado
    ];
  });

  sheet.clearContents();

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  if (values.length > 0) {
    sheet.getRange(2, 1, values.length, headers.length).setValues(values);
  }

}

function buildClasificacion_(categoriaFilter, inscripcionesInput, resultadosInput) {
  const inscripciones = (inscripcionesInput || getInscripciones_())
    .filter(item => !categoriaFilter || item.categoria === categoriaFilter);

  const resultados = resultadosInput || getResultados_();
  const resultadosByInscripcion = {};

  resultados.forEach(function(resultado) {
    if (!resultadosByInscripcion[resultado.inscripcionId]) {
      resultadosByInscripcion[resultado.inscripcionId] = [];
    }

    resultadosByInscripcion[resultado.inscripcionId].push(resultado);
  });

  const rows = inscripciones.map(inscripcion => {
    const pilotoResultados = resultadosByInscripcion[inscripcion.inscripcionId] || [];

    const tiempos = {};

    PASADAS.forEach(pasada => {
      const found = pilotoResultados.find(r => Number(r.pasada) === pasada.num);
      tiempos[pasada.num] = found ? Number(found.total) : '';
    });

    const penalizaciones = pilotoResultados.reduce((sum, r) => {
      return sum + Number(r.penalizacion || 0);
    }, 0);

    const completedCount = PASADAS.filter(p => tiempos[p.num] !== '').length;

    let total = '';
    let estado = 'Sin resultados';
    let descartada = '';

    const valoresRegistrados = PASADAS
      .map(p => ({
        pasada: p.num,
        label: p.label,
        valor: tiempos[p.num] !== '' ? Number(tiempos[p.num]) : ''
      }))
      .filter(item => item.valor !== '');

    if (completedCount > 0 && completedCount < 6) {
      total = round3_(
        valoresRegistrados.reduce((sum, item) => sum + Number(item.valor), 0)
      );

      estado = penalizaciones > 0 ? 'Pendiente con penalización' : 'Pendiente';
    }

    if (completedCount === 6) {
      const peor = [...valoresRegistrados].sort((a, b) => b.valor - a.valor)[0];

      descartada = peor.label + ' - ' + peor.valor;

      const valoresValidos = valoresRegistrados.filter(item => item.pasada !== peor.pasada);

      total = round3_(
        valoresValidos.reduce((sum, item) => sum + Number(item.valor), 0)
      );

      estado = penalizaciones > 0 ? 'Completo con penalización' : 'Completo';
    }

    return {
      inscripcionId: inscripcion.inscripcionId,
      categoria: inscripcion.categoria,
      dorsal: inscripcion.dorsal,
      piloto: inscripcion.piloto,
      ida1: tiempos[1],
      vuelta1: tiempos[2],
      ida2: tiempos[3],
      vuelta2: tiempos[4],
      ida3: tiempos[5],
      vuelta3: tiempos[6],
      descartada,
      penalizaciones: round3_(penalizaciones),
      total,
      gap: '',
      completadas: completedCount,
      estado
    };
  });

  rows.sort((a, b) => {
    const aHasTotal = a.total !== '';
    const bHasTotal = b.total !== '';

    if (aHasTotal && !bHasTotal) {
      return -1;
    }

    if (!aHasTotal && bHasTotal) {
      return 1;
    }

    if (aHasTotal && bHasTotal) {
      const totalDiff = Number(a.total) - Number(b.total);

      if (totalDiff !== 0) {
        return totalDiff;
      }
    }

    if (a.categoria !== b.categoria) {
      return a.categoria.localeCompare(b.categoria);
    }

    const dorsalA = Number(a.dorsal);
    const dorsalB = Number(b.dorsal);

    if (!isNaN(dorsalA) && !isNaN(dorsalB)) {
      return dorsalA - dorsalB;
    }

    return String(a.dorsal).localeCompare(String(b.dorsal));
  });

  rows.forEach(function(row, index) {
    if (row.total === '') {
      row.gap = '';
      return;
    }

    if (index === 0) {
      row.gap = '-';
      return;
    }

    const previousRow = rows[index - 1];

    const sameCompletedCount =
      previousRow &&
      previousRow.total !== '' &&
      Number(previousRow.completadas) === Number(row.completadas);

    if (!sameCompletedCount) {
      row.gap = '-';
      return;
    }

    row.gap = round3_(Number(row.total) - Number(previousRow.total));
  });

  return rows;
}

function getResultados_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_RESULTADOS);
  const values = sheet.getDataRange().getValues();

  return buildResultadosFromValues_(values);
}

function buildResultadosFromValues_(values) {

  const result = [];

  for (let i = 1; i < values.length; i++) {
    const row = values[i];

    if (!row[0]) {
      continue;
    }

    result.push({
      resultadoId: row[0],
      timestamp: row[1],
      inscripcionId: String(row[2] || '').trim(),
      dorsal: String(row[3] || '').trim(),
      piloto: String(row[4] || '').trim(),
      categoria: String(row[5] || '').trim(),
      pasada: Number(row[6]),
      pasadaLabel: String(row[7] || '').trim(),
      tiempo: Number(row[8] || 0),
      penalizacion: Number(row[9] || 0),
      total: Number(row[10] || 0),
      juez: String(row[11] || '').trim()
    });
  }

  return result;
}

function resetDemoData_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const resultados = ss.getSheetByName(SHEET_RESULTADOS);
  resultados.clearContents();
  resultados.getRange(1, 1, 1, 12).setValues([[
    'resultadoId',
    'timestamp',
    'inscripcionId',
    'dorsal',
    'piloto',
    'categoria',
    'pasada',
    'pasadaLabel',
    'tiempo',
    'penalizacion',
    'total',
    'juez'
  ]]);

  refreshClasificacion_();
}
function updateConfigValue_(key, value) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_CONFIG);
  const values = sheet.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]).trim() === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }

  sheet.appendRow([key, value]);
}

function ensureCarreraPublicationConfig_() {
  const config = getConfigValues_();

  if (!String(config[CONFIG_CARRERA_ID] || '').trim()) {
    updateConfigValue_(CONFIG_CARRERA_ID, Utilities.getUuid());
  }

  if (!String(config[CONFIG_CAMPEONATO_PUBLICADA] || '').trim()) {
    updateConfigValue_(CONFIG_CAMPEONATO_PUBLICADA, 'NO');
  }

  if (config[CONFIG_CAMPEONATO_PUBLICACION_INICIADA] === undefined) {
    updateConfigValue_(CONFIG_CAMPEONATO_PUBLICACION_INICIADA, '');
  }
}

function assertCarreraAdmiteResultados_() {
  const status = recoverCarreraPublicationStatus_(getConfigValues_());

  if (status === 'PUBLICANDO') {
    throw new Error('La carrera se está publicando en el campeonato');
  }

  if (status === 'SI') {
    throw new Error('La carrera ya está publicada. Las correcciones deben hacerse en la hoja del campeonato');
  }
}

function recoverCarreraPublicationStatus_(config) {
  const status = String(config[CONFIG_CAMPEONATO_PUBLICADA] || 'NO').toUpperCase();

  if (status !== 'PUBLICANDO') {
    return status;
  }

  const startedAt = new Date(
    String(config[CONFIG_CAMPEONATO_PUBLICACION_INICIADA] || '')
  ).getTime();
  const stale = !isFinite(startedAt) || Date.now() - startedAt > CHAMPIONSHIP_PUBLICATION_LEASE_MS;

  if (!stale) {
    return status;
  }

  updateConfigValue_(CONFIG_CAMPEONATO_PUBLICADA, 'NO');
  updateConfigValue_(CONFIG_CAMPEONATO_PUBLICACION_INICIADA, '');
  config[CONFIG_CAMPEONATO_PUBLICADA] = 'NO';
  config[CONFIG_CAMPEONATO_PUBLICACION_INICIADA] = '';
  return 'NO';
}

function generarNuevaCarrera(pin, adminPin, nombreCarrera) {
  if (!validatePin_(pin)) {
    throw new Error('PIN de carrera incorrecto');
  }

  if (!validateAdminPin_(adminPin)) {
    throw new Error('Contraseña de administración incorrecta');
  }

  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(10000);

    if (recoverCarreraPublicationStatus_(getConfigValues_()) === 'PUBLICANDO') {
      throw new Error('Espera a que termine la publicación del campeonato');
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const timestamp = Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      'yyyy-MM-dd_HH-mm-ss'
    );

    const resultadosSheet = ss.getSheetByName(SHEET_RESULTADOS);
    const clasificacionSheet = ss.getSheetByName(SHEET_CLASIFICACION);

    // Crear backups antes de limpiar
    resultadosSheet.copyTo(ss).setName('Resultados_backup_' + timestamp);
    clasificacionSheet.copyTo(ss).setName('Clasificacion_backup_' + timestamp);

    // Limpiar resultados
    resultadosSheet.clearContents();
    resultadosSheet.getRange(1, 1, 1, 12).setValues([[
      'resultadoId',
      'timestamp',
      'inscripcionId',
      'dorsal',
      'piloto',
      'categoria',
      'pasada',
      'pasadaLabel',
      'tiempo',
      'penalizacion',
      'total',
      'juez'
    ]]);
    resultadosSheet.setFrozenRows(1);

    // Limpiar clasificación
    clasificacionSheet.clearContents();
    clasificacionSheet.getRange(1, 1, 1, 14).setValues([[
      'categoria',
      'dorsal',
      'piloto',
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
      'estado'
    ]]);
    clasificacionSheet.setFrozenRows(1);

    const carrera = String(nombreCarrera || '').trim() || 'Nueva carrera';

    updateConfigValue_('CARRERA_ACTIVA', carrera);
    updateConfigValue_(CONFIG_CARRERA_ID, Utilities.getUuid());
    updateConfigValue_(CONFIG_CAMPEONATO_PUBLICADA, 'NO');
    updateConfigValue_(CONFIG_CAMPEONATO_PUBLICACION_INICIADA, '');

    // Regenera clasificación vacía con las inscripciones actuales
    refreshClasificacion_();

    return {
      status: 'OK',
      message: 'Nueva carrera generada correctamente. Se han creado copias de seguridad de Resultados y Clasificación.',
      carrera: carrera
    };

  } finally {
    lock.releaseLock();
  }
}

function getEstadoPublicacionCampeonato(pin, adminPin) {
  if (!validatePin_(pin)) {
    throw new Error('PIN de carrera incorrecto');
  }

  if (!validateAdminPin_(adminPin)) {
    throw new Error('Contraseña de administración incorrecta');
  }

  const lock = LockService.getScriptLock();
  let config;

  try {
    lock.waitLock(10000);
    config = getConfigValues_();
    recoverCarreraPublicationStatus_(config);
  } finally {
    lock.releaseLock();
  }

  const properties = PropertiesService.getScriptProperties();
  const clasificacion = buildClasificacion_('');
  const endpointConfigurado = Boolean(
    String(properties.getProperty(CHAMPIONSHIP_ENDPOINT_PROPERTY) || '').trim()
  );
  const tokenConfigurado = Boolean(
    String(properties.getProperty(CHAMPIONSHIP_TOKEN_PROPERTY) || '')
  );

  return {
    carreraId: String(config[CONFIG_CARRERA_ID] || ''),
    carrera: String(config.CARRERA_ACTIVA || 'Rally RC'),
    estado: String(config[CONFIG_CAMPEONATO_PUBLICADA] || 'NO').toUpperCase(),
    participantes: clasificacion.length,
    incompletos: clasificacion.filter(function(row) {
      return Number(row.completadas) < 6;
    }).length,
    endpointConfigurado: endpointConfigurado,
    tokenConfigurado: tokenConfigurado,
    configurado: endpointConfigurado && tokenConfigurado
  };
}

function publicarResultadosCampeonato(pin, adminPin, expectedCarreraId) {
  if (!validatePin_(pin)) {
    throw new Error('PIN de carrera incorrecto');
  }

  if (!validateAdminPin_(adminPin)) {
    throw new Error('Contraseña de administración incorrecta');
  }

  const properties = PropertiesService.getScriptProperties();
  const endpoint = String(properties.getProperty(CHAMPIONSHIP_ENDPOINT_PROPERTY) || '').trim();
  const token = String(properties.getProperty(CHAMPIONSHIP_TOKEN_PROPERTY) || '');

  if (!endpoint || !token) {
    throw new Error('Configura CHAMPIONSHIP_ENDPOINT y CHAMPIONSHIP_TOKEN en las propiedades del script');
  }

  const lock = LockService.getScriptLock();
  let snapshot;

  try {
    lock.waitLock(10000);

    const config = getConfigValues_();
    const status = recoverCarreraPublicationStatus_(config);
    const carreraId = String(config[CONFIG_CARRERA_ID] || '');

    if (!expectedCarreraId || String(expectedCarreraId) !== carreraId) {
      throw new Error('La carrera ha cambiado desde la confirmación. Actualiza el estado antes de publicar');
    }

    if (status === 'SI') {
      throw new Error('Esta carrera ya está publicada en el campeonato');
    }

    if (status === 'PUBLICANDO') {
      throw new Error('Ya hay una publicación del campeonato en curso');
    }

    snapshot = buildCampeonatoSnapshot_(config);
    updateConfigValue_(CONFIG_CAMPEONATO_PUBLICADA, 'PUBLICANDO');
    updateConfigValue_(CONFIG_CAMPEONATO_PUBLICACION_INICIADA, new Date().toISOString());
  } finally {
    lock.releaseLock();
  }

  try {
    const response = UrlFetchApp.fetch(endpoint, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        token: token,
        payload: snapshot
      }),
      followRedirects: true,
      muteHttpExceptions: true
    });

    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    let responseData;

    try {
      responseData = JSON.parse(responseText);
    } catch (error) {
      throw new Error('La app de campeonato no devolvió una respuesta válida (HTTP ' + responseCode + ')');
    }

    if (!responseData || !responseData.ok) {
      throw new Error(
        responseData && responseData.message
          ? responseData.message
          : 'La app de campeonato rechazó la publicación'
      );
    }

    setCarreraPublicationStatus_(snapshot.carrera.carreraId, 'SI');

    return {
      status: responseData.status,
      message: responseData.status === 'ALREADY_EXISTS'
        ? 'La carrera ya existía en el campeonato y se ha marcado como publicada'
        : 'Resultados publicados correctamente en el campeonato',
      carreraId: snapshot.carrera.carreraId,
      incompletos: snapshot.incompletos
    };
  } catch (error) {
    setCarreraPublicationStatus_(snapshot.carrera.carreraId, 'NO');
    throw error;
  }
}

function buildCampeonatoSnapshot_(config) {
  const carreraId = String(config[CONFIG_CARRERA_ID] || '').trim();
  const carrera = String(config.CARRERA_ACTIVA || 'Rally RC').trim();

  if (!carreraId) {
    throw new Error('La carrera no tiene identificador. Ejecuta setupSheets() antes de publicar');
  }

  const inscripciones = getInscripciones_();
  const resultados = getResultados_();

  if (!inscripciones.length) {
    throw new Error('No hay inscripciones para publicar');
  }

  const inscripcionesById = {};
  inscripciones.forEach(function(inscripcion) {
    inscripcionesById[inscripcion.inscripcionId] = inscripcion;
  });

  const categoriasByName = {};
  getCategoriasDB_().forEach(function(categoria) {
    categoriasByName[normalizeText_(categoria.nombre)] = categoria;
  });

  const categorias = [...new Set(inscripciones.map(function(inscripcion) {
    return inscripcion.categoria;
  }))].filter(Boolean).sort();

  const rows = [];

  categorias.forEach(function(categoriaNombre) {
    const categoria = categoriasByName[normalizeText_(categoriaNombre)] || {};
    const clasificacion = buildClasificacion_(categoriaNombre, inscripciones, resultados);
    const completos = clasificacion.filter(function(row) {
      return Number(row.completadas) === 6 && /^Completo/i.test(String(row.estado || ''));
    });
    const incompletos = clasificacion.filter(function(row) {
      return completos.indexOf(row) === -1;
    });
    const clasificacionPublicable = completos.concat(incompletos);

    clasificacionPublicable.forEach(function(row, index) {
      const inscripcion = inscripcionesById[row.inscripcionId] || {};

      rows.push({
        inscripcionId: row.inscripcionId,
        pilotoId: String(inscripcion.pilotoId || ''),
        piloto: row.piloto,
        categoriaId: String(categoria.categoriaId || ''),
        categoria: row.categoria,
        dorsal: row.dorsal,
        posicion: index + 1,
        ida1: row.ida1,
        vuelta1: row.vuelta1,
        ida2: row.ida2,
        vuelta2: row.vuelta2,
        ida3: row.ida3,
        vuelta3: row.vuelta3,
        descartada: row.descartada,
        penalizaciones: row.penalizaciones,
        total: row.total,
        gap: row.gap,
        completadas: row.completadas,
        estado: row.estado
      });
    });
  });

  return {
    schemaVersion: 1,
    carrera: {
      carreraId: carreraId,
      nombre: carrera
    },
    fechaPublicacion: new Date().toISOString(),
    incompletos: rows.filter(function(row) {
      return Number(row.completadas) < 6;
    }).length,
    resultados: rows
  };
}

function setCarreraPublicationStatus_(carreraId, status) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(10000);
    const config = getConfigValues_();

    if (String(config[CONFIG_CARRERA_ID] || '') === String(carreraId || '')) {
      updateConfigValue_(CONFIG_CAMPEONATO_PUBLICADA, status);
      updateConfigValue_(CONFIG_CAMPEONATO_PUBLICACION_INICIADA, '');
    }
  } finally {
    lock.releaseLock();
  }
}

function formatDateForClient_(value) {
  if (!value) {
    return '';
  }

  if (Object.prototype.toString.call(value) === '[object Date]') {
    return Utilities.formatDate(
      value,
      Session.getScriptTimeZone(),
      'yyyy-MM-dd HH:mm:ss'
    );
  }

  return String(value);
}

function secondsToTimeParts_(secondsValue) {
  const totalTenths = Math.round(Number(secondsValue || 0) * 10);

  const minutes = Math.floor(totalTenths / 600);
  const remainingTenths = totalTenths % 600;
  const seconds = Math.floor(remainingTenths / 10);
  const decimas = remainingTenths % 10;

  return {
    minutos: minutes,
    segundos: seconds,
    decimas: decimas
  };
}

function getResultadoParaCorreccion(pin, adminPin, inscripcionId, pasada) {
  if (!validatePin_(pin)) {
    throw new Error('PIN de carrera incorrecto');
  }

  if (!validateAdminPin_(adminPin)) {
    throw new Error('Contraseña de administración incorrecta');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_RESULTADOS);
  const values = sheet.getDataRange().getValues();

  const rowIndex = findResultadoRowIndex_(values, inscripcionId, Number(pasada));

  if (rowIndex === -1) {
    return {
      status: 'NOT_FOUND',
      message: 'No existe ningún resultado para la categoría, piloto y pasada seleccionados.'
    };
  }

  const row = values[rowIndex - 1];
  const timeParts = secondsToTimeParts_(row[8]);

  return {
    status: 'FOUND',
    resultado: {
      resultadoId: String(row[0] || ''),
      timestamp: formatDateForClient_(row[1]),
      inscripcionId: String(row[2] || ''),
      dorsal: String(row[3] || ''),
      piloto: String(row[4] || ''),
      categoria: String(row[5] || ''),
      pasada: Number(row[6]),
      pasadaLabel: String(row[7] || ''),
      tiempo: Number(row[8] || 0),
      minutos: timeParts.minutos,
      segundos: timeParts.segundos,
      decimas: timeParts.decimas,
      penalizacion: Number(row[9] || 0),
      total: Number(row[10] || 0),
      juez: String(row[11] || '')
    }
  };
}

function corregirResultado(pin, adminPin, payload) {
  if (!validatePin_(pin)) {
    throw new Error('PIN de carrera incorrecto');
  }

  if (!validateAdminPin_(adminPin)) {
    throw new Error('Contraseña de administración incorrecta');
  }

  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(10000);
    assertCarreraAdmiteResultados_();

    if (!payload) {
      throw new Error('No se han recibido datos para corregir');
    }

    const inscripcionId = String(payload.inscripcionId || '').trim();
    const pasada = Number(payload.pasada);
    const minutos = Number(payload.minutos || 0);
    const segundos = Number(payload.segundos || 0);
    const decimas = Number(payload.decimas || 0);
    const penalizacion = parseNumber_(payload.penalizacion || 0);
    const juez = String(payload.juez || '').trim() || 'Administración';

    if (!inscripcionId) {
      throw new Error('Debes seleccionar un piloto');
    }

    if (!PASADAS.some(p => p.num === pasada)) {
      throw new Error('La pasada seleccionada no es válida');
    }

    if (
      isNaN(minutos) || minutos < 0 ||
      isNaN(segundos) || segundos < 0 || segundos > 59 ||
      isNaN(decimas) || decimas < 0 || decimas > 9
    ) {
      throw new Error('El tiempo no es válido. Revisa minutos, segundos y décimas.');
    }

    if (isNaN(penalizacion) || penalizacion < 0) {
      throw new Error('La penalización no es válida');
    }

    const tiempo = round3_(minutos * 60 + segundos + (decimas / 10));
    const total = round3_(tiempo + penalizacion);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_RESULTADOS);
    const values = sheet.getDataRange().getValues();

    const rowIndex = findResultadoRowIndex_(values, inscripcionId, pasada);

    if (rowIndex === -1) {
      throw new Error('No existe ningún resultado para corregir');
    }

    const existing = values[rowIndex - 1];

    const updatedRow = [
      existing[0],
      new Date(),
      existing[2],
      existing[3],
      existing[4],
      existing[5],
      existing[6],
      existing[7],
      tiempo,
      round3_(penalizacion),
      total,
      juez + ' - Corrección'
    ];

    sheet.getRange(rowIndex, 1, 1, updatedRow.length).setValues([updatedRow]);

    refreshClasificacion_();

    return {
      status: 'OK',
      message: 'Resultado corregido correctamente',
      saved: {
        dorsal: String(existing[3] || ''),
        piloto: String(existing[4] || ''),
        categoria: String(existing[5] || ''),
        pasadaLabel: String(existing[7] || ''),
        tiempo: tiempo,
        penalizacion: round3_(penalizacion),
        total: total
      }
    };

  } finally {
    lock.releaseLock();
  }
}
function getPasadasRegistradas(pin, inscripcionId) {
  if (!validatePin_(pin)) {
    throw new Error('PIN incorrecto');
  }

  const cleanInscripcionId = String(inscripcionId || '').trim();

  if (!cleanInscripcionId) {
    return [];
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_RESULTADOS);
  const values = sheet.getDataRange().getValues();

  const pasadas = [];

  for (let i = 1; i < values.length; i++) {
    const rowInscripcionId = String(values[i][2] || '').trim();
    const rowPasada = Number(values[i][6]);

    if (rowInscripcionId === cleanInscripcionId && rowPasada) {
      pasadas.push(rowPasada);
    }
  }

  return [...new Set(pasadas)].sort((a, b) => a - b);
}
function getDashboardPilotos(pin, categoria) {
  if (!validatePin_(pin)) {
    throw new Error('PIN incorrecto');
  }

  const cleanCategoria = String(categoria || '').trim();

  const inscripciones = getInscripciones_()
    .filter(item => !cleanCategoria || item.categoria === cleanCategoria);

  const pilotos = [...new Set(inscripciones.map(item => item.piloto))]
    .filter(Boolean)
    .sort();

  return pilotos;
}
function getDashboardData(pin) {
  if (!validatePin_(pin)) {
    throw new Error('PIN incorrecto');
  }

  const carreraActiva = getConfigValue_('CARRERA_ACTIVA') || 'Rally RC';
  const rows = buildDashboardConstancia_();

  const categorias = [...new Set(rows.map(row => row.categoria))]
    .filter(Boolean)
    .sort();

  const pilotos = [...new Set(rows.map(row => row.piloto))]
    .filter(Boolean)
    .sort();

  return {
    carrera: carreraActiva,
    updatedAt: Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      'yyyy-MM-dd HH:mm:ss'
    ),
    categorias: categorias,
    pilotos: pilotos,
    pasadas: PASADAS,
    rows: rows
  };
}
function buildDashboardConstancia_() {
  const inscripciones = getInscripciones_();
  const resultados = getResultados_();

  const rows = inscripciones.map(inscripcion => {
    const pilotoResultados = resultados.filter(r =>
      r.inscripcionId === inscripcion.inscripcionId
    );

    const tiempos = {};

    PASADAS.forEach(pasada => {
      const found = pilotoResultados.find(r => Number(r.pasada) === pasada.num);
      tiempos[pasada.num] = found ? Number(found.total) : '';
    });

    const valores = PASADAS
      .map(p => ({
        pasada: p.num,
        label: p.label,
        tipo: p.tipo,
        valor: tiempos[p.num] !== '' ? Number(tiempos[p.num]) : ''
      }))
      .filter(item => item.valor !== '');

    const completadas = valores.length;

    const metricasTodas = calcularMetricasConstancia_(valores, '');
    const metricasIdas = calcularMetricasConstancia_(valores, 'ida');
    const metricasVueltas = calcularMetricasConstancia_(valores, 'vuelta');

    return {
      categoria: inscripcion.categoria,
      dorsal: inscripcion.dorsal,
      piloto: inscripcion.piloto,
      inscripcionId: inscripcion.inscripcionId,

      ida1: tiempos[1],
      vuelta1: tiempos[2],
      ida2: tiempos[3],
      vuelta2: tiempos[4],
      ida3: tiempos[5],
      vuelta3: tiempos[6],

      completadas: completadas,
      progreso: completadas + '/6',

      mejor: metricasTodas.mejor,
      peor: metricasTodas.peor,
      media: metricasTodas.media,
      diferencia: metricasTodas.diferencia,

      mejorIdas: metricasIdas.mejor,
      peorIdas: metricasIdas.peor,
      mediaIdas: metricasIdas.media,
      diferenciaIdas: metricasIdas.diferencia,

      mejorVueltas: metricasVueltas.mejor,
      peorVueltas: metricasVueltas.peor,
      mediaVueltas: metricasVueltas.media,
      diferenciaVueltas: metricasVueltas.diferencia,

      tiemposPorPasada: buildTiemposPorPasada_(tiempos),
      progresionAcumulada: buildProgresionAcumulada_(tiempos),

      estado: completadas === 0
        ? 'Sin resultados'
        : completadas === 6
          ? 'Completo'
          : 'En curso'
    };
  });

  rows.sort((a, b) => {
    if (a.categoria !== b.categoria) {
      return a.categoria.localeCompare(b.categoria);
    }

    const dorsalA = Number(a.dorsal);
    const dorsalB = Number(b.dorsal);

    if (!isNaN(dorsalA) && !isNaN(dorsalB)) {
      return dorsalA - dorsalB;
    }

    return String(a.dorsal).localeCompare(String(b.dorsal));
  });

  return rows;
}
function calcularMetricasConstancia_(valores, tipoFilter) {
  const filtrados = valores.filter(item => {
    if (!tipoFilter) {
      return true;
    }

    return item.tipo === tipoFilter;
  });

  if (!filtrados.length) {
    return {
      mejor: '',
      peor: '',
      media: '',
      diferencia: ''
    };
  }

  const numeros = filtrados.map(item => Number(item.valor));

  const mejor = round3_(Math.min.apply(null, numeros));
  const peor = round3_(Math.max.apply(null, numeros));
  const media = round3_(
    numeros.reduce((sum, value) => sum + value, 0) / numeros.length
  );

  const diferencia = numeros.length >= 2
    ? round3_(peor - mejor)
    : '';

  return {
    mejor: mejor,
    peor: peor,
    media: media,
    diferencia: diferencia
  };
}

function buildTiemposPorPasada_(tiempos) {
  return PASADAS.map(pasada => {
    const valor = tiempos[pasada.num];

    return {
      pasada: pasada.num,
      label: pasada.label,
      tipo: pasada.tipo,
      valor: valor !== '' ? Number(valor) : ''
    };
  });
}

function buildProgresionAcumulada_(tiempos) {
  let acumulado = 0;

  return PASADAS.map(pasada => {
    const valor = tiempos[pasada.num];

    if (valor === '') {
      return {
        pasada: pasada.num,
        label: pasada.label,
        tipo: pasada.tipo,
        valor: ''
      };
    }

    acumulado = round3_(acumulado + Number(valor));

    return {
      pasada: pasada.num,
      label: pasada.label,
      tipo: pasada.tipo,
      valor: acumulado
    };
  });
}

function getHeaderMap_(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  return getHeaderMapFromHeaders_(headers);
}

function getHeaderMapFromHeaders_(headers) {
  const map = {};

  headers.forEach(function(header, index) {
    map[String(header || '').trim()] = index;
  });

  return map;
}

function normalizeText_(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');
}

function nowString_() {
  return Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    'yyyy-MM-dd HH:mm:ss'
  );
}

function getNextId_(prefix, existingIds) {
  let max = 0;

  existingIds.forEach(function(id) {
    const cleanId = String(id || '').trim();

    if (!cleanId.startsWith(prefix)) {
      return;
    }

    const numberPart = Number(cleanId.replace(prefix, ''));

    if (!isNaN(numberPart) && numberPart > max) {
      max = numberPart;
    }
  });

  return prefix + String(max + 1).padStart(3, '0');
}

function getPilotosDB_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_PILOTOS_DB);

  if (!sheet || sheet.getLastRow() < 2) {
    return [];
  }

  const values = sheet.getDataRange().getValues();
  const map = getHeaderMapFromHeaders_(values[0]);
  const result = [];

  for (let i = 1; i < values.length; i++) {
    const row = values[i];

    const pilotoId = String(row[map.pilotoId] || '').trim();
    const nombre = String(row[map.nombre] || '').trim();

    if (!pilotoId || !nombre) {
      continue;
    }

    result.push({
      pilotoId: pilotoId,
      nombre: nombre,
      alias: String(row[map.alias] || '').trim(),
      activo: String(row[map.activo] || 'SI').trim() || 'SI',
      notas: String(row[map.notas] || '').trim(),
      createdAt: String(row[map.createdAt] || '').trim(),
      updatedAt: String(row[map.updatedAt] || '').trim()
    });
  }

  return result;
}

function getCategoriasDB_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_CATEGORIAS_DB);

  if (!sheet || sheet.getLastRow() < 2) {
    return [];
  }

  const values = sheet.getDataRange().getValues();
  const map = getHeaderMapFromHeaders_(values[0]);
  const result = [];

  for (let i = 1; i < values.length; i++) {
    const row = values[i];

    const categoriaId = String(row[map.categoriaId] || '').trim();
    const nombre = String(row[map.nombre] || '').trim();

    if (!categoriaId || !nombre) {
      continue;
    }

    result.push({
      categoriaId: categoriaId,
      nombre: nombre,
      activa: String(row[map.activa] || 'SI').trim() || 'SI',
      orden: Number(row[map.orden] || 999),
      notas: String(row[map.notas] || '').trim(),
      createdAt: String(row[map.createdAt] || '').trim(),
      updatedAt: String(row[map.updatedAt] || '').trim()
    });
  }

  result.sort(function(a, b) {
    if (Number(a.orden) !== Number(b.orden)) {
      return Number(a.orden) - Number(b.orden);
    }

    return a.nombre.localeCompare(b.nombre);
  });

  return result;
}

function migrarPilotosYCategoriasDesdeInscripciones_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const inscripcionesSheet = ss.getSheetByName(SHEET_INSCRIPCIONES);
  const pilotosSheet = ss.getSheetByName(SHEET_PILOTOS_DB);
  const categoriasSheet = ss.getSheetByName(SHEET_CATEGORIAS_DB);

  if (!inscripcionesSheet || inscripcionesSheet.getLastRow() < 2) {
    return;
  }

  const inscripcionesValues = inscripcionesSheet.getDataRange().getValues();
  const inscripcionesHeaders = inscripcionesValues[0].map(function(header) {
    return String(header || '').trim();
  });

  const oldFormat = !inscripcionesHeaders.includes('pilotoId');

  let inscripciones = [];

  if (oldFormat) {
    for (let i = 1; i < inscripcionesValues.length; i++) {
      const row = inscripcionesValues[i];

      const inscripcionId = String(row[0] || '').trim();
      const dorsal = String(row[1] || '').trim();
      const piloto = String(row[2] || '').trim();
      const categoria = String(row[3] || '').trim();

      if (!inscripcionId || !piloto || !categoria) {
        continue;
      }

      inscripciones.push({
        inscripcionId: inscripcionId,
        pilotoId: '',
        dorsal: dorsal,
        piloto: piloto,
        categoria: categoria
      });
    }
  } else {
    const map = getHeaderMapFromHeaders_(inscripcionesHeaders);

    for (let i = 1; i < inscripcionesValues.length; i++) {
      const row = inscripcionesValues[i];

      const inscripcionId = String(row[map.inscripcionId] || '').trim();
      const pilotoId = String(row[map.pilotoId] || '').trim();
      const dorsal = String(row[map.dorsal] || '').trim();
      const piloto = String(row[map.piloto] || '').trim();
      const categoria = String(row[map.categoria] || '').trim();

      if (!inscripcionId || !piloto || !categoria) {
        continue;
      }

      inscripciones.push({
        inscripcionId: inscripcionId,
        pilotoId: pilotoId,
        dorsal: dorsal,
        piloto: piloto,
        categoria: categoria
      });
    }
  }

  const pilotosExistentes = getPilotosDB_();
  const categoriasExistentes = getCategoriasDB_();

  const pilotosByName = {};
  pilotosExistentes.forEach(function(piloto) {
    pilotosByName[normalizeText_(piloto.nombre)] = piloto;
  });

  const categoriasByName = {};
  categoriasExistentes.forEach(function(categoria) {
    categoriasByName[normalizeText_(categoria.nombre)] = categoria;
  });

  const existingPilotoIds = pilotosExistentes.map(function(p) {
    return p.pilotoId;
  });

  const existingCategoriaIds = categoriasExistentes.map(function(c) {
    return c.categoriaId;
  });

  const nuevosPilotos = [];
  const nuevasCategorias = [];

  inscripciones.forEach(function(inscripcion) {
    const pilotoKey = normalizeText_(inscripcion.piloto);

    if (!pilotosByName[pilotoKey]) {
      const pilotoId = getNextId_(
        'P',
        existingPilotoIds.concat(nuevosPilotos.map(function(p) {
          return p[0];
        }))
      );

      const now = nowString_();

      const pilotoRow = [
        pilotoId,
        inscripcion.piloto,
        '',
        'SI',
        '',
        now,
        now
      ];

      nuevosPilotos.push(pilotoRow);

      pilotosByName[pilotoKey] = {
        pilotoId: pilotoId,
        nombre: inscripcion.piloto
      };
    }

    const categoriaKey = normalizeText_(inscripcion.categoria);

    if (!categoriasByName[categoriaKey]) {
      const categoriaId = getNextId_(
        'C',
        existingCategoriaIds.concat(nuevasCategorias.map(function(c) {
          return c[0];
        }))
      );

      const now = nowString_();

      const categoriaRow = [
        categoriaId,
        inscripcion.categoria,
        'SI',
        categoriasExistentes.length + nuevasCategorias.length + 1,
        '',
        now,
        now
      ];

      nuevasCategorias.push(categoriaRow);

      categoriasByName[categoriaKey] = {
        categoriaId: categoriaId,
        nombre: inscripcion.categoria
      };
    }
  });

  if (nuevosPilotos.length > 0) {
    pilotosSheet
      .getRange(pilotosSheet.getLastRow() + 1, 1, nuevosPilotos.length, nuevosPilotos[0].length)
      .setValues(nuevosPilotos);
  }

  if (nuevasCategorias.length > 0) {
    categoriasSheet
      .getRange(categoriasSheet.getLastRow() + 1, 1, nuevasCategorias.length, nuevasCategorias[0].length)
      .setValues(nuevasCategorias);
  }

  migrarFormatoInscripciones_(inscripciones, pilotosByName);
}

function migrarFormatoInscripciones_(inscripciones, pilotosByName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_INSCRIPCIONES);

  const headers = [
    'inscripcionId',
    'pilotoId',
    'dorsal',
    'piloto',
    'categoria'
  ];

  const rows = inscripciones.map(function(inscripcion) {
    const pilotoKey = normalizeText_(inscripcion.piloto);
    const piloto = pilotosByName[pilotoKey];

    return [
      inscripcion.inscripcionId,
      piloto ? piloto.pilotoId : inscripcion.pilotoId,
      inscripcion.dorsal,
      inscripcion.piloto,
      inscripcion.categoria
    ];
  });

  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }

  sheet.setFrozenRows(1);
  autoResize_(sheet);
}

function getAdminCatalogosData(pin, adminPin) {
  if (!validatePin_(pin)) {
    throw new Error('PIN incorrecto');
  }

  if (!validateAdminPin_(adminPin)) {
    throw new Error('PIN de administración incorrecto');
  }

  return {
    pilotos: getPilotosDB_(),
    categorias: getCategoriasDB_(),
    inscripciones: getInscripciones_()
  };
}

function crearPilotoAdmin(pin, adminPin, payload) {
  if (!validatePin_(pin)) {
    throw new Error('PIN incorrecto');
  }

  if (!validateAdminPin_(adminPin)) {
    throw new Error('PIN de administración incorrecto');
  }

  assertCarreraAdmiteResultados_();

  const nombre = String(payload.nombre || '').trim();
  const alias = String(payload.alias || '').trim();
  const notas = String(payload.notas || '').trim();
  const activo = String(payload.activo || 'SI').trim().toUpperCase() === 'NO'
    ? 'NO'
    : 'SI';

  if (!nombre) {
    throw new Error('El nombre del piloto es obligatorio');
  }

  const pilotos = getPilotosDB_();
  const nombreNormalizado = normalizeText_(nombre);

  const existe = pilotos.some(function(piloto) {
    return normalizeText_(piloto.nombre) === nombreNormalizado;
  });

  if (existe) {
    throw new Error('Ya existe un piloto con ese nombre');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_PILOTOS_DB);

  const pilotoId = getNextId_(
    'P',
    pilotos.map(function(piloto) {
      return piloto.pilotoId;
    })
  );

  const now = nowString_();

  sheet.appendRow([
    pilotoId,
    nombre,
    alias,
    activo,
    notas,
    now,
    now
  ]);

  autoResize_(sheet);

  return {
    status: 'OK',
    message: 'Piloto creado correctamente',
    piloto: {
      pilotoId: pilotoId,
      nombre: nombre,
      alias: alias,
      activo: activo,
      notas: notas,
      createdAt: now,
      updatedAt: now
    }
  };
}

function editarPilotoAdmin(pin, adminPin, payload) {
  if (!validatePin_(pin)) {
    throw new Error('PIN incorrecto');
  }

  if (!validateAdminPin_(adminPin)) {
    throw new Error('PIN de administración incorrecto');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    assertCarreraAdmiteResultados_();

  const pilotoId = String(payload.pilotoId || '').trim();
  const nombre = String(payload.nombre || '').trim();
  const alias = String(payload.alias || '').trim();
  const notas = String(payload.notas || '').trim();
  const activo = String(payload.activo || 'SI').trim().toUpperCase() === 'NO'
    ? 'NO'
    : 'SI';

  if (!pilotoId) {
    throw new Error('El pilotoId es obligatorio');
  }

  if (!nombre) {
    throw new Error('El nombre del piloto es obligatorio');
  }

  const pilotos = getPilotosDB_();
  const nombreNormalizado = normalizeText_(nombre);

  const duplicado = pilotos.some(function(piloto) {
    return piloto.pilotoId !== pilotoId &&
      normalizeText_(piloto.nombre) === nombreNormalizado;
  });

  if (duplicado) {
    throw new Error('Ya existe otro piloto con ese nombre');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_PILOTOS_DB);
  const values = sheet.getDataRange().getValues();
  const map = getHeaderMap_(sheet);

  let rowIndex = -1;

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][map.pilotoId] || '').trim() === pilotoId) {
      rowIndex = i + 1;
      break;
    }
  }

  if (rowIndex === -1) {
    throw new Error('No se ha encontrado el piloto');
  }

  const now = nowString_();

  sheet.getRange(rowIndex, map.nombre + 1).setValue(nombre);
  sheet.getRange(rowIndex, map.alias + 1).setValue(alias);
  sheet.getRange(rowIndex, map.activo + 1).setValue(activo);
  sheet.getRange(rowIndex, map.notas + 1).setValue(notas);
  sheet.getRange(rowIndex, map.updatedAt + 1).setValue(now);

  sincronizarNombrePilotoEnInscripciones_(pilotoId, nombre);

  autoResize_(sheet);

    return {
      status: 'OK',
      message: 'Piloto actualizado correctamente'
    };
  } finally {
    lock.releaseLock();
  }
}

function sincronizarNombrePilotoEnInscripciones_(pilotoId, nombre) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_INSCRIPCIONES);

  if (!sheet || sheet.getLastRow() < 2) {
    return;
  }

  const values = sheet.getDataRange().getValues();
  const map = getHeaderMap_(sheet);

  if (map.pilotoId === undefined || map.piloto === undefined) {
    return;
  }

  for (let i = 1; i < values.length; i++) {
    const rowPilotoId = String(values[i][map.pilotoId] || '').trim();

    if (rowPilotoId === pilotoId) {
      sheet.getRange(i + 1, map.piloto + 1).setValue(nombre);
    }
  }
}

function crearCategoriaAdmin(pin, adminPin, payload) {
  if (!validatePin_(pin)) {
    throw new Error('PIN incorrecto');
  }

  if (!validateAdminPin_(adminPin)) {
    throw new Error('PIN de administración incorrecto');
  }

  assertCarreraAdmiteResultados_();

  const nombre = String(payload.nombre || '').trim();
  const notas = String(payload.notas || '').trim();
  const activa = String(payload.activa || 'SI').trim().toUpperCase() === 'NO'
    ? 'NO'
    : 'SI';

  let orden = Number(payload.orden || 0);

  if (!nombre) {
    throw new Error('El nombre de la categoría es obligatorio');
  }

  const categorias = getCategoriasDB_();
  const nombreNormalizado = normalizeText_(nombre);

  const existe = categorias.some(function(categoria) {
    return normalizeText_(categoria.nombre) === nombreNormalizado;
  });

  if (existe) {
    throw new Error('Ya existe una categoría con ese nombre');
  }

  if (isNaN(orden) || orden <= 0) {
    orden = categorias.length + 1;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_CATEGORIAS_DB);

  const categoriaId = getNextId_(
    'C',
    categorias.map(function(categoria) {
      return categoria.categoriaId;
    })
  );

  const now = nowString_();

  sheet.appendRow([
    categoriaId,
    nombre,
    activa,
    orden,
    notas,
    now,
    now
  ]);

  autoResize_(sheet);

  return {
    status: 'OK',
    message: 'Categoría creada correctamente',
    categoria: {
      categoriaId: categoriaId,
      nombre: nombre,
      activa: activa,
      orden: orden,
      notas: notas,
      createdAt: now,
      updatedAt: now
    }
  };
}

function editarCategoriaAdmin(pin, adminPin, payload) {
  if (!validatePin_(pin)) {
    throw new Error('PIN incorrecto');
  }

  if (!validateAdminPin_(adminPin)) {
    throw new Error('PIN de administración incorrecto');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    assertCarreraAdmiteResultados_();

  const categoriaId = String(payload.categoriaId || '').trim();
  const nombre = String(payload.nombre || '').trim();
  const notas = String(payload.notas || '').trim();
  const activa = String(payload.activa || 'SI').trim().toUpperCase() === 'NO'
    ? 'NO'
    : 'SI';

  let orden = Number(payload.orden || 0);

  if (!categoriaId) {
    throw new Error('El categoriaId es obligatorio');
  }

  if (!nombre) {
    throw new Error('El nombre de la categoría es obligatorio');
  }

  if (isNaN(orden) || orden <= 0) {
    orden = 999;
  }

  const categorias = getCategoriasDB_();
  const categoriaActual = categorias.find(function(categoria) {
    return categoria.categoriaId === categoriaId;
  });

  if (!categoriaActual) {
    throw new Error('No se ha encontrado la categoría');
  }

  const nombreNormalizado = normalizeText_(nombre);

  const duplicada = categorias.some(function(categoria) {
    return categoria.categoriaId !== categoriaId &&
      normalizeText_(categoria.nombre) === nombreNormalizado;
  });

  if (duplicada) {
    throw new Error('Ya existe otra categoría con ese nombre');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_CATEGORIAS_DB);
  const values = sheet.getDataRange().getValues();
  const map = getHeaderMap_(sheet);

  let rowIndex = -1;

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][map.categoriaId] || '').trim() === categoriaId) {
      rowIndex = i + 1;
      break;
    }
  }

  if (rowIndex === -1) {
    throw new Error('No se ha encontrado la categoría');
  }

  const oldName = categoriaActual.nombre;
  const now = nowString_();

  sheet.getRange(rowIndex, map.nombre + 1).setValue(nombre);
  sheet.getRange(rowIndex, map.activa + 1).setValue(activa);
  sheet.getRange(rowIndex, map.orden + 1).setValue(orden);
  sheet.getRange(rowIndex, map.notas + 1).setValue(notas);
  sheet.getRange(rowIndex, map.updatedAt + 1).setValue(now);

  sincronizarNombreCategoriaEnInscripciones_(oldName, nombre);

  autoResize_(sheet);

    return {
      status: 'OK',
      message: 'Categoría actualizada correctamente'
    };
  } finally {
    lock.releaseLock();
  }
}

function sincronizarNombreCategoriaEnInscripciones_(oldName, newName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_INSCRIPCIONES);

  if (!sheet || sheet.getLastRow() < 2) {
    return;
  }

  const values = sheet.getDataRange().getValues();
  const map = getHeaderMap_(sheet);

  if (map.categoria === undefined) {
    return;
  }

  const oldNameNormalized = normalizeText_(oldName);

  for (let i = 1; i < values.length; i++) {
    const rowCategoria = String(values[i][map.categoria] || '').trim();

    if (normalizeText_(rowCategoria) === oldNameNormalized) {
      sheet.getRange(i + 1, map.categoria + 1).setValue(newName);
    }
  }
}
function getNextInscripcionId_() {
  const inscripciones = getInscripciones_();

  return getNextId_(
    'I',
    inscripciones.map(function(item) {
      return item.inscripcionId;
    })
  );
}

function getPilotoDBById_(pilotoId) {
  return getPilotosDB_().find(function(piloto) {
    return piloto.pilotoId === pilotoId;
  });
}

function getCategoriaDBByNombre_(nombre) {
  const nombreNormalizado = normalizeText_(nombre);

  return getCategoriasDB_().find(function(categoria) {
    return normalizeText_(categoria.nombre) === nombreNormalizado;
  });
}

function getInscripcionRowIndex_(inscripcionId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_INSCRIPCIONES);

  if (!sheet || sheet.getLastRow() < 2) {
    return -1;
  }

  const values = sheet.getDataRange().getValues();
  const map = getHeaderMap_(sheet);

  for (let i = 1; i < values.length; i++) {
    const rowInscripcionId = String(values[i][map.inscripcionId] || '').trim();

    if (rowInscripcionId === inscripcionId) {
      return i + 1;
    }
  }

  return -1;
}

function inscripcionTieneResultados_(inscripcionId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_RESULTADOS);

  if (!sheet || sheet.getLastRow() < 2) {
    return false;
  }

  const values = sheet.getDataRange().getValues();
  const map = getHeaderMap_(sheet);

  const inscripcionIndex = map.inscripcionId !== undefined
    ? map.inscripcionId
    : 2;

  for (let i = 1; i < values.length; i++) {
    const rowInscripcionId = String(values[i][inscripcionIndex] || '').trim();

    if (rowInscripcionId === inscripcionId) {
      return true;
    }
  }

  return false;
}

function validarInscripcionAdmin_(payload, currentInscripcionId) {
  const pilotoId = String(payload.pilotoId || '').trim();
  const dorsal = String(payload.dorsal || '').trim();
  const categoria = String(payload.categoria || '').trim();

  if (!pilotoId) {
    throw new Error('Selecciona un piloto');
  }

  if (!categoria) {
    throw new Error('Selecciona una categoría');
  }

  if (!dorsal) {
    throw new Error('Introduce un dorsal');
  }

  const piloto = getPilotoDBById_(pilotoId);

  if (!piloto) {
    throw new Error('No se ha encontrado el piloto en la base de datos');
  }

  if (String(piloto.activo || '').toUpperCase() === 'NO') {
    throw new Error('El piloto seleccionado está inactivo');
  }

  const categoriaDB = getCategoriaDBByNombre_(categoria);

  if (!categoriaDB) {
    throw new Error('No se ha encontrado la categoría en la base de datos');
  }

  if (String(categoriaDB.activa || '').toUpperCase() === 'NO') {
    throw new Error('La categoría seleccionada está inactiva');
  }

  const inscripciones = getInscripciones_();
  const categoriaNormalizada = normalizeText_(categoria);

  const duplicadoPilotoCategoria = inscripciones.some(function(item) {
    return item.inscripcionId !== currentInscripcionId &&
      String(item.pilotoId || '').trim() === pilotoId &&
      normalizeText_(item.categoria) === categoriaNormalizada;
  });

  if (duplicadoPilotoCategoria) {
    throw new Error('Este piloto ya está inscrito en esta categoría');
  }

  const duplicadoDorsalCategoria = inscripciones.some(function(item) {
    return item.inscripcionId !== currentInscripcionId &&
      String(item.dorsal || '').trim() === dorsal &&
      normalizeText_(item.categoria) === categoriaNormalizada;
  });

  if (duplicadoDorsalCategoria) {
    throw new Error('Ya existe una inscripción con ese dorsal en esta categoría');
  }

  return {
    piloto: piloto,
    categoria: categoriaDB,
    dorsal: dorsal
  };
}
function crearInscripcionAdmin(pin, adminPin, payload) {
  if (!validatePin_(pin)) {
    throw new Error('PIN incorrecto');
  }

  if (!validateAdminPin_(adminPin)) {
    throw new Error('PIN de administración incorrecto');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    assertCarreraAdmiteResultados_();
    const validated = validarInscripcionAdmin_(payload, '');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_INSCRIPCIONES);

    const inscripcionId = getNextInscripcionId_();

    sheet.appendRow([
      inscripcionId,
      validated.piloto.pilotoId,
      validated.dorsal,
      validated.piloto.nombre,
      validated.categoria.nombre
    ]);

    autoResize_(sheet);
    refreshClasificacion_();

    return {
      status: 'OK',
      message: 'Inscripción creada correctamente',
      inscripcion: {
        inscripcionId: inscripcionId,
        pilotoId: validated.piloto.pilotoId,
        dorsal: validated.dorsal,
        piloto: validated.piloto.nombre,
        categoria: validated.categoria.nombre
      }
    };
  } finally {
    lock.releaseLock();
  }
}
function editarInscripcionAdmin(pin, adminPin, payload) {
  if (!validatePin_(pin)) {
    throw new Error('PIN incorrecto');
  }

  if (!validateAdminPin_(adminPin)) {
    throw new Error('PIN de administración incorrecto');
  }

  const inscripcionId = String(payload.inscripcionId || '').trim();

  if (!inscripcionId) {
    throw new Error('El inscripcionId es obligatorio');
  }

  if (inscripcionTieneResultados_(inscripcionId)) {
    throw new Error('No se puede editar una inscripción que ya tiene resultados registrados');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    assertCarreraAdmiteResultados_();

    if (inscripcionTieneResultados_(inscripcionId)) {
      throw new Error('No se puede editar una inscripción que ya tiene resultados registrados');
    }

    const validated = validarInscripcionAdmin_(payload, inscripcionId);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_INSCRIPCIONES);
    const rowIndex = getInscripcionRowIndex_(inscripcionId);

    if (rowIndex === -1) {
      throw new Error('No se ha encontrado la inscripción');
    }

    sheet.getRange(rowIndex, 1, 1, 5).setValues([[
      inscripcionId,
      validated.piloto.pilotoId,
      validated.dorsal,
      validated.piloto.nombre,
      validated.categoria.nombre
    ]]);

    autoResize_(sheet);
    refreshClasificacion_();

    return {
      status: 'OK',
      message: 'Inscripción actualizada correctamente'
    };
  } finally {
    lock.releaseLock();
  }
}
function eliminarInscripcionAdmin(pin, adminPin, inscripcionId) {
  if (!validatePin_(pin)) {
    throw new Error('PIN incorrecto');
  }

  if (!validateAdminPin_(adminPin)) {
    throw new Error('PIN de administración incorrecto');
  }

  const cleanInscripcionId = String(inscripcionId || '').trim();

  if (!cleanInscripcionId) {
    throw new Error('El inscripcionId es obligatorio');
  }

  if (inscripcionTieneResultados_(cleanInscripcionId)) {
    throw new Error('No se puede eliminar una inscripción que ya tiene resultados registrados');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    assertCarreraAdmiteResultados_();

    if (inscripcionTieneResultados_(cleanInscripcionId)) {
      throw new Error('No se puede eliminar una inscripción que ya tiene resultados registrados');
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_INSCRIPCIONES);
    const rowIndex = getInscripcionRowIndex_(cleanInscripcionId);

    if (rowIndex === -1) {
      throw new Error('No se ha encontrado la inscripción');
    }

    sheet.deleteRow(rowIndex);

    autoResize_(sheet);
    refreshClasificacion_();

    return {
      status: 'OK',
      message: 'Inscripción eliminada correctamente'
    };
  } finally {
    lock.releaseLock();
  }
}
function crearPilotoEInscribirAdmin(pin, adminPin, payload) {
  if (!validatePin_(pin)) {
    throw new Error('PIN incorrecto');
  }

  if (!validateAdminPin_(adminPin)) {
    throw new Error('PIN de administración incorrecto');
  }

  assertCarreraAdmiteResultados_();

  const nombre = String(payload.nombre || '').trim();
  const alias = String(payload.alias || '').trim();
  const notas = String(payload.notas || '').trim();
  const categoria = String(payload.categoria || '').trim();
  const dorsal = String(payload.dorsal || '').trim();

  if (!nombre) {
    throw new Error('El nombre del piloto es obligatorio');
  }

  if (!categoria) {
    throw new Error('Selecciona una categoría');
  }

  if (!dorsal) {
    throw new Error('Introduce un dorsal');
  }

  const pilotoResult = crearPilotoAdmin(pin, adminPin, {
    nombre: nombre,
    alias: alias,
    activo: 'SI',
    notas: notas
  });

  const inscripcionResult = crearInscripcionAdmin(pin, adminPin, {
    pilotoId: pilotoResult.piloto.pilotoId,
    categoria: categoria,
    dorsal: dorsal
  });

  return {
    status: 'OK',
    message: 'Piloto creado e inscrito correctamente',
    piloto: pilotoResult.piloto,
    inscripcion: inscripcionResult.inscripcion
  };
}
/*
function testAdminCatalogos() {
  const pin = '1234';
  const adminPin = '9999';

  const data = getAdminCatalogosData(pin, adminPin);

  Logger.log('Pilotos:');
  Logger.log(JSON.stringify(data.pilotos, null, 2));

  Logger.log('Categorías:');
  Logger.log(JSON.stringify(data.categorias, null, 2));

  Logger.log('Inscripciones:');
  Logger.log(JSON.stringify(data.inscripciones, null, 2));
}

function testEditarCategoriaAdmin() {
  const pin = '1234';
  const adminPin = '9999';

  const result = editarCategoriaAdmin(pin, adminPin, {
    categoriaId: 'C004',
    nombre: 'Categoria Test Editada',
    activa: 'SI',
    orden: 99,
    notas: 'Editada desde test'
  });

  Logger.log(JSON.stringify(result, null, 2));
}

function testCrearPilotoAdmin() {
  const pin = '1234';
  const adminPin = '9999';

  const result = crearPilotoAdmin(pin, adminPin, {
    nombre: 'Piloto Test',
    alias: '',
    activo: 'SI',
    notas: 'Creado desde test'
  });

  Logger.log(JSON.stringify(result, null, 2));
}
function testCrearCategoriaAdmin() {
  const pin = '1234';
  const adminPin = '9999';

  const result = crearCategoriaAdmin(pin, adminPin, {
    nombre: 'Categoria Test',
    activa: 'SI',
    orden: 99,
    notas: 'Creada desde test'
  });

  Logger.log(JSON.stringify(result, null, 2));
}

function testEditarPilotoAdmin() {
  const pin = '1234';
  const adminPin = '9999';

  const result = editarPilotoAdmin(pin, adminPin, {
    pilotoId: 'P012',
    nombre: 'Piloto Test Editado 3',
    alias: '',
    activo: 'SI',
    notas: 'Editado desde test 3'
  });

  Logger.log(JSON.stringify(result, null, 2));
}

function testDashboardData() {
  const data = getDashboardData('1234');

  Logger.log(JSON.stringify(data, null, 2));
}

function testMigracionPilotosCategorias() {
  setupSheets();

  Logger.log('PilotosDB');
  Logger.log(JSON.stringify(getPilotosDB_(), null, 2));

  Logger.log('CategoriasDB');
  Logger.log(JSON.stringify(getCategoriasDB_(), null, 2));

  Logger.log('Inscripciones');
  Logger.log(JSON.stringify(getInscripciones_(), null, 2));
}
function testRegistroInscripciones() {
  const categorias = getCategorias('1234');
  Logger.log('Categorías:');
  Logger.log(JSON.stringify(categorias, null, 2));

  if (categorias.length > 0) {
    const inscripciones = getInscripcionesByCategoria('1234', categorias[0]);
    Logger.log('Inscripciones de ' + categorias[0] + ':');
    Logger.log(JSON.stringify(inscripciones, null, 2));
  }
}*/
