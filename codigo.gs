const SHEET_CONFIG = 'Config';
const SHEET_INSCRIPCIONES = 'Inscripciones';
const SHEET_RESULTADOS = 'Resultados';
const SHEET_CLASIFICACION = 'Clasificacion';

const PASADAS = [
  { num: 1, label: 'Ida 1' },
  { num: 2, label: 'Vuelta 1' },
  { num: 3, label: 'Ida 2' },
  { num: 4, label: 'Vuelta 2' },
  { num: 5, label: 'Ida 3' },
  { num: 6, label: 'Vuelta 3' }
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
  ['PIN', '1234'],
  ['ADMIN_PIN', '9999'],
  ['CARRERA_ACTIVA', 'Rally RC']
]);

  createSheetIfNotExists_(ss, SHEET_INSCRIPCIONES, [
    ['inscripcionId', 'dorsal', 'piloto', 'categoria'],
    ['I001', '1', 'Piloto Demo 1', 'Rally 1/10'],
    ['I002', '2', 'Piloto Demo 2', 'Rally 1/10'],
    ['I003', '3', 'Piloto Demo 3', 'Rally 1/10'],
    ['I004', '1', 'Piloto Demo 1', 'Clásicos']
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
    'estado'
  ]
  ]);
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
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_CONFIG);
  const values = sheet.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]).trim() === key) {
      return String(values[i][1]).trim();
    }
  }

  return '';
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

  const inscripciones = getInscripciones_();
  const categorias = [...new Set(inscripciones.map(item => item.categoria))]
    .filter(Boolean)
    .sort();

  return categorias;
}

function getInscripcionesByCategoria(pin, categoria) {
  if (!validatePin_(pin)) {
    throw new Error('PIN incorrecto');
  }

  return getInscripciones_()
    .filter(item => item.categoria === categoria)
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
  const values = sheet.getDataRange().getValues();

  const result = [];

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
      inscripcionId,
      dorsal,
      piloto,
      categoria
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

    const data = normalizePayload_(payload);
    const inscripcion = findInscripcion_(data.inscripcionId);

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

    refreshClasificacion_();

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
      }
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

function refreshClasificacion_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_CLASIFICACION);

  const allRows = buildClasificacion_();

  sheet.clearContents();

  const headers = [
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
      'estado'
    ]
  ];

  const rows = allRows.map(item => [
    item.categoria,
    item.dorsal,
    item.piloto,
    item.ida1,
    item.vuelta1,
    item.ida2,
    item.vuelta2,
    item.ida3,
    item.vuelta3,
    item.descartada,
    item.penalizaciones,
    item.total,
    item.estado
  ]);

  sheet.getRange(1, 1, headers.length, headers[0].length).setValues(headers);

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers[0].length).setValues(rows);
  }

  sheet.setFrozenRows(1);
  autoResize_(sheet);
}

function buildClasificacion_(categoriaFilter) {
  const inscripciones = getInscripciones_()
    .filter(item => !categoriaFilter || item.categoria === categoriaFilter);

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
      estado
    };
  });

    rows.sort((a, b) => {
      const aHasTotal = a.total !== '';
      const bHasTotal = b.total !== '';

      // Los pilotos con algún tiempo registrado van arriba
      if (aHasTotal && !bHasTotal) {
        return -1;
      }

      if (!aHasTotal && bHasTotal) {
        return 1;
      }

      // Si ambos tienen total, ordenar siempre por mejor tiempo acumulado
      if (aHasTotal && bHasTotal) {
        const totalDiff = Number(a.total) - Number(b.total);

        if (totalDiff !== 0) {
          return totalDiff;
        }
      }

      // En empate o sin tiempos, ordenar por categoría
      if (a.categoria !== b.categoria) {
        return a.categoria.localeCompare(b.categoria);
      }

      // Después por dorsal
      const dorsalA = Number(a.dorsal);
      const dorsalB = Number(b.dorsal);

      if (!isNaN(dorsalA) && !isNaN(dorsalB)) {
        return dorsalA - dorsalB;
      }

      return String(a.dorsal).localeCompare(String(b.dorsal));
    });

  return rows;
}

function getResultados_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_RESULTADOS);
  const values = sheet.getDataRange().getValues();

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

function resetDemoData() {
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
    clasificacionSheet.getRange(1, 1, 1, 13).setValues([[
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
      'estado'
    ]]);
    clasificacionSheet.setFrozenRows(1);

    const carrera = String(nombreCarrera || '').trim() || 'Nueva carrera';

    updateConfigValue_('CARRERA_ACTIVA', carrera);

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