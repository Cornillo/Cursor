/*
Estructura de llamadas:

1. depuradorActive()
   ├── obtenerKeysToDepurate()
   │   └── obtenerFechaLimite()
   │
   ├── procesarRegistros(keysToDepurate)
   │   ├── Itera sobre ordenProcesamiento[]
   │   ├── Obtiene datos de cada hoja
   │   ├── Identifica registros a depurar usando relacionesHojas{}
   │   ├── flushBuffer() - copia a hojas 'Total' y 'Logs'
   │   └── clearContent() en registros originales
   │
   └── logTotales()

2. verificarCasos() [función independiente]
   └── Verifica registros del ProjectID específico
      ├── Busca DWOs relacionados en hoja principal
      └── Verifica cada hoja siguiendo ordenProcesamiento[]

Constantes y Configuraciones Principales:
├── relacionesHojas{} - Define columnas de ID y DWO para cada hoja
├── ordenProcesamiento[] - Define secuencia de procesamiento preservando integridad referencial
└── CONFIG{
    ├── BUFFER_SIZE - Tamaño del lote para procesamiento
    ├── DIAS_ANTIGUEDAD - Límite para considerar registros antiguos
    └── SHEET_NAMES - Nombres de hojas especiales (Total, Logs)
    }
*/

// Variables globales para las spreadsheets
let ssActive, ssTotal, ssLogs;

// Configuración global
const CONFIG = {
  CHUNK_SIZE: 5000,    // Tamaño de chunks para lectura
  BATCH_SIZE: 100,     // Tamaño de lotes para procesamiento
  BUFFER_SIZE: 1000,   // Tamaño del buffer de escritura
  CACHE_TTL: 21600,    // Tiempo de vida de caché (6 horas)
  VERBOSE: true        // Control de logging
};

// Definición de relaciones entre hojas y sus columnas clave
const relacionesHojas = {
  'DWO': {
    columna: 'B',  // Columna de ID
    dwoCol: 'B'    // Columna de DWO ID (misma que ID en este caso)
  },
  'DWO-Production': {
    columna: 'A',
    dwoCol: 'B'
  },
  'DWO-Event': {
    columna: 'A',
    dwoCol: 'BX'
  },
  'DWO_Character': {
    columna: 'A',
    dwoCol: 'B'
  },
  'DWO_CharacterProduction': {
    columna: 'A',
    dwoCol: 'X'
  },
  'DWO_Files': {
    columna: 'A',
    dwoCol: 'P'
  },
  'DWO-MixAndEdit': {
    columna: 'A',
    dwoCol: 'Q'
  },
  'DWO-Observation': {
    columna: 'A',
    dwoCol: 'Z'
  },
  'DWO_FilesLines': {
    columna: 'A',
    dwoCol: 'Q'
  },
  'DWO_FilesCharacter': {
    columna: 'A',
    dwoCol: 'N'
  },
  'DWO_Song': {
    columna: 'A',
    dwoCol: 'P'
  },
  'DWO_SongDetail': {
    columna: 'A',
    dwoCol: 'G'
  },
  'DWO-SynopsisProject': {
    columna: 'A',
    dwoCol: 'A'
  },
  'DWO-SynopsisProduction': {
    columna: 'A',
    dwoCol: 'T'
  }
};

// Definir el orden de procesamiento para preservar integridad referencial
const ordenProcesamiento = [
  // 1. Hojas de nivel más bajo (más alejadas de DWO)
  'DWO_FilesCharacter',
  'DWO_FilesLines',
  'DWO_SongDetail',
  
  // 2. Hojas de nivel medio
  'DWO_Files',
  'DWO_Song',
  'DWO_CharacterProduction',
  'DWO-SynopsisProduction',
  
  // 3. Hojas directamente relacionadas con DWO
  'DWO-Production',
  'DWO-Event',
  'DWO_Character',
  'DWO-MixAndEdit',
  'DWO-Observation',
  'DWO-SynopsisProject',
  
  // 4. Hoja principal (última en procesarse)
  'DWO'
];

function initializeSpreadsheets() {
  const allIDs = databaseID.getID();
  ssActive = SpreadsheetApp.openById(allIDs['activeID']);
  ssTotal = SpreadsheetApp.openById(allIDs['totalID']);
  ssLogs = SpreadsheetApp.openById(allIDs['logsID']);
}

function depuradorActive() {
  try {
    console.time('Tiempo total de migración');
    initializeSpreadsheets();
    
    // Implementar caché para mejorar rendimiento
    const cache = CacheService.getScriptCache();
    const cacheKey = `migration_${new Date().toISOString().split('T')[0]}`;
    
    if (cache.get(cacheKey)) {
      console.log('Ya se realizó la migración hoy');
      return;
    }

    // Obtener y procesar registros
    const keysToDepurate = obtenerRegistrosADepurar(ssActive);
    console.log(`Se encontraron ${keysToDepurate.size} registros para depurar`);
    
    if (keysToDepurate.size > 0) {
      const totalesPorHoja = procesarRegistros(keysToDepurate);
      
      // Guardar resultados en caché
      cache.put(cacheKey, 'completed');
      
      // Mostrar resumen
      mostrarResumen(totalesPorHoja);
      console.timeEnd('Tiempo total de migración');

      // Ejecutar conteo auxiliar después de la migración
      console.log('\n=== CONTEO DE REGISTROS POST-MIGRACIÓN ===');
      conteoAux();
    } else {
      console.log('No se encontraron registros para depurar');
    }
    
  } catch (error) {
    console.error(`Error en depuradorActive: ${error.message}`);
    throw error;
  }
}

function obtenerRegistrosADepurar(ss) {
  const dwSheet = ss.getSheetByName('DWO');
  const data = dwSheet.getDataRange().getValues();
  const today = new Date();
  const limitDate = new Date(today.getFullYear(), today.getMonth() - 1, today.getDate() - 15);
  
  const keysToDepurate = new Set();
  
  // Empezar desde 1 para saltar headers
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const key = row[1];                  // Columna B
    const status = row[58];              // Columna BG
    const dateStr = row[68];             // Columna BQ
    
    const date = new Date(dateStr);
    
    if (status !== "(01) On track: DWO" && date < limitDate) {
      keysToDepurate.add(key.toString().trim());
    }
  }
  
  return keysToDepurate;
}

function columnToNumber(column) {
  let result = 0;
  for (let i = 0; i < column.length; i++) {
    result *= 26;
    result += column.charCodeAt(i) - 'A'.charCodeAt(0) + 1;
  }
  return result;
}

function conteoAux() {
  const allIDs = databaseID.getID();
  const ssActive = SpreadsheetApp.openById(allIDs['activeID']);

  Logger.log('\nRegistros en blanco en Active:');
  
  // Hojas a revisar
  const sheetNames = Object.keys(relacionesHojas);

  // Recorrer cada hoja
  for (const sheetName of sheetNames) {
    const sheet = ssActive.getSheetByName(sheetName);
    if (!sheet) {
      Logger.log(`  ${sheetName}: Hoja no encontrada`);
      continue;
    }

    // Obtener todos los datos sin importar filtros
    const data = sheet.getDataRange().getValues();
    let blankCount = 0;

    // Contar filas con columna A en blanco (excluyendo encabezado)
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0] || data[i][0].toString().trim() === '') {
        blankCount++;
      }
    }

    Logger.log(`  ${sheetName}: ${blankCount} filas con columna A en blanco`);
  }
}

function acumularResultados(totalesPorHoja, resultadosLote) {
  for (const [hoja, cantidad] of Object.entries(resultadosLote)) {
    totalesPorHoja[hoja] = (totalesPorHoja[hoja] || 0) + cantidad;
  }
}

function obtenerOCrearHoja(spreadsheet, nombreHoja, hojaOrigen) {
  let hoja = spreadsheet.getSheetByName(nombreHoja);
  if (!hoja) {
    hoja = spreadsheet.insertSheet(nombreHoja);
    // Copiar encabezados de la hoja origen
    const headers = hojaOrigen.getRange(1, 1, 1, hojaOrigen.getLastColumn()).getValues();
    hoja.getRange(1, 1, 1, headers[0].length).setValues(headers);
  }
  return hoja;
}

function procesarRegistros(keysToDepurate) {
  const totalesPorHoja = {};

  // Procesar hojas en el orden definido
  for (const sheetName of ordenProcesamiento) {
    const config = relacionesHojas[sheetName];
    console.log(`\nProcesando hoja: ${sheetName}`);
    
    const sourceSheet = ssActive.getSheetByName(sheetName);
    if (!sourceSheet) {
      console.log(`  Hoja no encontrada: ${sheetName}`);
      continue;
    }

    const data = sourceSheet.getDataRange().getValues();
    const keyColIndex = columnToNumber(config.columna) - 1;
    const dwoColIndex = config.dwoCol ? columnToNumber(config.dwoCol) - 1 : null;
    
    // Procesar registros en chunks
    let registrosProcesados = 0;
    let buffer = [];
    
    // Empezar desde 1 para saltar headers
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const dwoKey = (dwoColIndex !== null ? row[dwoColIndex] : row[keyColIndex])?.toString().trim();
      
      if (keysToDepurate.has(dwoKey)) {
        // Copiar a Total y Logs
        buffer.push(row);
        registrosProcesados++;
        
        // Limpiar fila original usando clearContent()
        sourceSheet.getRange(i + 1, 1, 1, row.length).clearContent();
      }
      
      // Flush buffer si está lleno
      if (buffer.length >= CONFIG.BUFFER_SIZE) {
        flushBuffer(buffer, sheetName);
        buffer = [];
      }
    }
    
    // Flush buffer restante
    if (buffer.length > 0) {
      flushBuffer(buffer, sheetName);
    }
    
    totalesPorHoja[sheetName] = registrosProcesados;
    console.log(`  Registros procesados: ${registrosProcesados}`);
  }
  
  return totalesPorHoja;
}

function flushBuffer(buffer, sheetName) {
  if (buffer.length === 0) return;
  
  // Copiar a Total
  const totalSheet = obtenerOCrearHoja(ssTotal, sheetName, ssActive.getSheetByName(sheetName));
  const totalLastRow = totalSheet.getLastRow();
  totalSheet.getRange(totalLastRow + 1, 1, buffer.length, buffer[0].length).setValues(buffer);
  
  // Copiar a Logs
  const logsSheet = obtenerOCrearHoja(ssLogs, sheetName, ssActive.getSheetByName(sheetName));
  const logsLastRow = logsSheet.getLastRow();
  logsSheet.getRange(logsLastRow + 1, 1, buffer.length, buffer[0].length).setValues(buffer);
}

function mostrarResumen(totales) {
  console.log('\n=== RESUMEN DE MIGRACIÓN ===');
  Object.entries(totales).forEach(([hoja, cantidad]) => {
    console.log(`${hoja}: ${cantidad} registros migrados`);
  });
  console.log('===========================\n');
}

function verificarCasos() {
  const PROJECT_ID = 'A3654F4F-35EB-4CBB-9BB5-04805E2DD708';
  const SOURCE_ID = '1_KBAXZfhQ502UXrK6Kf-kXj1ytDx6zKgLBF6gFBcEHo';
  const TARGET_ID = '1d258eRpy3NovzQ1wKX_WVRCfXhxc2u2pwpYUuTDyfIw';
  
  console.log('=== INICIANDO VERIFICACIÓN Y RESTAURACIÓN DE CASOS ===');
  
  const sourceSpreadsheet = SpreadsheetApp.openById(SOURCE_ID);
  const targetSpreadsheet = SpreadsheetApp.openById(TARGET_ID);
  
  // Primero buscar en todas las hojas para obtener los DWOs relacionados
  const dwosEncontrados = new Set();
  
  // Buscar en cada hoja según el orden de procesamiento
  for (const sheetName of ordenProcesamiento) {
    const sourceSheet = sourceSpreadsheet.getSheetByName(sheetName);
    if (!sourceSheet) continue;
    
    const sourceData = sourceSheet.getDataRange().getValues();
    const config = relacionesHojas[sheetName];
    const dwoColIndex = columnToNumber(config.dwoCol) - 1;
    
    // Buscar registros relacionados
    for (let i = 1; i < sourceData.length; i++) {
      const dwoId = sourceData[i][dwoColIndex]?.toString().trim();
      
      if (dwoId === PROJECT_ID) {
        dwosEncontrados.add(dwoId);
      }
    }
  }
  
  if (dwosEncontrados.size === 0) {
    console.log('No se encontraron DWOs relacionados al ProjectID');
    return;
  }
  
  let seEncontraronFaltantes = false;
  
  // Verificar y restaurar registros en el target
  for (const sheetName of ordenProcesamiento) {
    const sourceSheet = sourceSpreadsheet.getSheetByName(sheetName);
    const targetSheet = targetSpreadsheet.getSheetByName(sheetName);
    
    if (!sourceSheet || !targetSheet) continue;
    
    const sourceData = sourceSheet.getDataRange().getValues();
    const targetData = targetSheet.getDataRange().getValues();
    const config = relacionesHojas[sheetName];
    const dwoColIndex = columnToNumber(config.dwoCol) - 1;
    
    let registrosParaRestaurar = [];
    
    // Verificar registros
    for (let i = 1; i < sourceData.length; i++) {
      const dwoId = sourceData[i][dwoColIndex]?.toString().trim();
      
      if (dwosEncontrados.has(dwoId)) {
        const existeEnTarget = targetData.some(row => 
          row[dwoColIndex]?.toString().trim() === dwoId
        );
        
        if (!existeEnTarget) {
          registrosParaRestaurar.push(sourceData[i]);
        }
      }
    }
    
    // Restaurar registros faltantes
    if (registrosParaRestaurar.length > 0) {
      if (!seEncontraronFaltantes) {
        console.log('\nRestaurando registros faltantes:');
        seEncontraronFaltantes = true;
      }
      
      console.log(`\nHoja: ${sheetName}`);
      console.log(`  Restaurando ${registrosParaRestaurar.length} registros`);
      
      // Copiar registros al target
      const targetLastRow = targetSheet.getLastRow();
      targetSheet.getRange(
        targetLastRow + 1, 
        1, 
        registrosParaRestaurar.length, 
        registrosParaRestaurar[0].length
      ).setValues(registrosParaRestaurar);
      
      registrosParaRestaurar.forEach(row => {
        console.log(`  Restaurado - DWO ID: ${row[dwoColIndex]}`);
      });
    }
  }
  
  if (!seEncontraronFaltantes) {
    console.log('\nTodos los registros están correctamente migrados');
  } else {
    console.log('\nRestauración completada');
  }
  
  console.log('\n=== PROCESO COMPLETADO ===');
}