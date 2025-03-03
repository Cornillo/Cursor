/**
 * IMPORTANTE: Las siguientes variables globales deben estar declaradas 
 * en cualquier función que use LazyLoad:
 * 
 * @var {GoogleAppsScript.Spreadsheet.Sheet} containerSheet - Hoja actual
 * @var {Array<Array>} containerValues - Valores de la hoja
 * @var {Array<string>} containerNDX - Índice principal
 * @var {Array<string>} containerNDX2 - Índice secundario (opcional)
 * @var {Object} sheetCache - Cache de hojas {initialized: boolean, sheets: Map}
 * 
 * LazyLoad se encargará de inicializar:
 * - Los Spreadsheets (ssActive, ssTotal, ssLogs, ssNoTrack)
 * - Los datos de Labels y Users (labelValues, userValues, labelNDX, userNDX)
 * - El cache de hojas (sheetCache)
 */

// Variables globales
/*let containerSheet, containerValues, containerNDX, containerNDX2;
let ssActive, ssTotal, ssLogs, ssNoTrack;
let labelValues, userValues, labelNDX, userNDX;

const sheetCache = {
  initialized: false,
  sheets: new Map()
};*/

// Configuración completa de índices por hoja
const SHEET_CONFIG = {
  // DubAppActive y DubAppTotal
  'DWO': { ndxCol: 1 },
  'DWO-Channel': { ndxCol: 0 },
  'DWO-ChannelEventType': { ndxCol: 1 },
  'DWO-Production': { ndxCol: 0, ndx2Col: 1 },
  'DWO-Event': { ndxCol: 0, ndx2Col: 1 },
  'DWO-MixAndEdit': { ndxCol: 0, ndx2Col: 1 },
  'DWO-Observation': { ndxCol: 0, ndx2Col: 1 },
  'DWO-Series': { ndxCol: 1 },
  'DWO-Settlement': { ndxCol: 0 },
  'DWO-SettlementResource': { ndxCol: 0 },
  'DWO-SettlementResourceDetail': { ndxCol: 0 },
  'DWO-SynopsisProduction': { ndxCol: 0 },
  'DWO-SynopsisProject': { ndxCol: 0 },
  'DWO-SynopsisSeries': { ndxCol: 1 },

  // DubAppLogs
  'DWO_FilesCharacterLog': { ndxCol: 0, ndx2Col: 1 },
  'DWO_CharacterLog': { ndxCol: 0, ndx2Col: 1 },
  'DWO_CharacterProductionLog': { ndxCol: 0, ndx2Col: 1 },
  'DWO_SongLog': { ndxCol: 0, ndx2Col: 1 },
  'DWO_SongDetailLog': { ndxCol: 0, ndx2Col: 1 },
  'DWO_FilesLinesLog': { ndxCol: 0, ndx2Col: 1 },
  'DWO_FilesLog': { ndxCol: 0, ndx2Col: 1 },
  
  // Hojas con NDX en A y NDX2 en B
  'DWO_Song': { ndxCol: 0, ndx2Col: 1 },
  'DWO_SongDetail': { ndxCol: 0, ndx2Col: 1 },
  'DWO_Character': { ndxCol: 0, ndx2Col: 1 },
  
  // Caso especial - NDX en A y NDX2 en D
  'DWO_Files': { ndxCol: 0, ndx2Col: 3 },
  
  // Caso especial - NDX en A y NDX2 en C
  'DWO_CharacterProduction': { ndxCol: 0, ndx2Col: 2 },
  'DWO_FilesCharacter': { ndxCol: 0, ndx2Col: 2 }, // Corregido: NDX en A (0) y NDX2 en C (2)
  
  // Resto de configuraciones sin cambios...

  // DubAppNoTrack
  'App-User': { ndxCol: 0 },

  // Caso especial - DWO_FilesLines
  'DWO_FilesLines': { ndxCol: 0, ndx2Col: 1 }
};

function LazyLoad(ssAux, sheetNameAux) {
  try {
    // Inicialización básica (una sola vez)
    if (!sheetCache.initialized) {
      initializeBasicData();
      sheetCache.initialized = true;
    }

    // Generar clave única para el cache
    const cacheKey = `${ssAux}_${sheetNameAux}`;
    
    // Retornar datos del cache si existen
    if (sheetCache.sheets.has(cacheKey)) {
      assignToGlobalVariables(sheetCache.sheets.get(cacheKey));
      return;
    }

    // Si no existe en cache, cargar los datos
    loadSheetData(ssAux, sheetNameAux);
  } catch (err) {
    console.error(`Error in LazyLoad for ${sheetNameAux}: ${err.toString()}`);
    throw err;
  }
}

function initializeBasicData() {
  try {
    // Cargar spreadsheets principales
    var allIDs = databaseID.getID();
    ssActive = SpreadsheetApp.openById(allIDs['activeID']);
    ssTotal = SpreadsheetApp.openById(allIDs['totalID']);
    ssLogs = SpreadsheetApp.openById(allIDs['logsID']);
    ssNoTrack = SpreadsheetApp.openById(allIDs['noTrackID']);

    // Cargar datos básicos (labels y users) con mejor manejo de errores
    const labelSheet = ssActive.getSheetByName("DWO-LogLabels");
    const userSheet = ssNoTrack.getSheetByName("App-User");

    if (!labelSheet) {
      console.error('No se encontró la hoja DWO-LogLabels');
      labelValues = [];
      labelNDX = [];
    } else {
      labelValues = getSheetData(labelSheet);
      labelNDX = labelValues.length ? labelValues.map(r => r[0].toString()) : [];
    }

    if (!userSheet) {
      console.error('No se encontró la hoja App-User');
      userValues = [];
      userNDX = [];
    } else {
      userValues = getSheetData(userSheet);
      userNDX = userValues.length ? userValues.map(r => r[0].toString()) : [];
    }

    // Validar que al menos tengamos arrays vacíos
    labelValues = labelValues || [];
    labelNDX = labelNDX || [];
    userValues = userValues || [];
    userNDX = userNDX || [];

  } catch (err) {
    console.error('Error en initializeBasicData: ' + err.toString());
    console.error('Stack: ' + err.stack);
    // Inicializar con arrays vacíos en caso de error
    labelValues = [];
    labelNDX = [];
    userValues = [];
    userNDX = [];
  }
}

function loadSheetData(ssAux, sheetNameAux) {
  try {
    const ss = getSpreadsheet(ssAux);
    const sheet = ss.getSheetByName(sheetNameAux);
    
    if (!sheet) {
      console.error(`Hoja no encontrada: ${sheetNameAux} en ${ssAux}`);
      return;
    }

    // Caso especial para hojas voluminosas
    if (sheetNameAux === 'DWO-Event' || sheetNameAux === 'DWO_FilesLines') {
      const lastRow = sheet.getLastRow();
      const lastCol = sheet.getLastColumn();
      
      // Aumentamos el tamaño del chunk y eliminamos las pausas
      const CHUNK_SIZE = 10000;
      const chunks = Math.ceil((lastRow - 1) / CHUNK_SIZE);
      let values = [];
      
      // Cargamos todos los datos de una vez si son menos de 20000 filas
      if (lastRow <= 20000) {
        values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
      } else {
        // Procesamiento por chunks sin pausas
        for (let i = 0; i < chunks; i++) {
          const startRow = 2 + (i * CHUNK_SIZE);
          const rows = Math.min(CHUNK_SIZE, lastRow - startRow + 1);
          if (rows <= 0) break;
          
          const chunkValues = sheet.getRange(startRow, 1, rows, lastCol).getValues();
          values = values.concat(chunkValues);
          
          if (verboseFlag === true) {
            console.log(`Chunk ${i+1}/${chunks} procesado. Filas: ${values.length}`);
          }
        }
      }
      
      if (!values.length) {
        if(verboseFlag === true) {
          console.log(`No data found in ${sheetNameAux}`);
        }
        return;
      }

      const config = SHEET_CONFIG[sheetNameAux];
      containerSheet = sheet;
      containerValues = values;
      containerNDX = values.map(r => r[config.ndxCol].toString());
      containerNDX2 = config.ndx2Col !== undefined 
        ? values.map(r => r[config.ndx2Col].toString())
        : null;
    } else {
      containerSheet = sheet;
      containerValues = getSheetData(sheet);
      
      if (!containerValues || containerValues.length === 0) {
        if(verboseFlag === true) {
          console.log(`No data found in ${sheetNameAux}`);
        }
        return;
      }

      const config = SHEET_CONFIG[sheetNameAux] || { ndxCol: 0 };
      containerNDX = containerValues.map(r => r[config.ndxCol].toString());
      containerNDX2 = config.ndx2Col !== undefined 
        ? containerValues.map(r => r[config.ndx2Col].toString())
        : null;
    }

    // Cache
    sheetCache.sheets.set(`${ssAux}_${sheetNameAux}`, {
      sheet: containerSheet,
      values: containerValues,
      ndx: containerNDX,
      ndx2: containerNDX2
    });
  } catch (err) {
    console.error(`Error loading sheet ${sheetNameAux}: ${err.toString()}`);
    console.error('Stack: ' + err.stack);
    throw err;
  }
}

function assignToGlobalVariables(cachedData) {
  try {
    if (!cachedData || !cachedData.values) {
      throw new Error('Datos de cache inválidos');
    }
    containerSheet = cachedData.sheet;
    containerValues = cachedData.values;
    containerNDX = cachedData.ndx;
    containerNDX2 = cachedData.ndx2;
  } catch (err) {
    console.error('Error assigning cached data: ' + err.toString());
    throw err;
  }
}

function getSheetData(sheet) {
  try {
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    
    if (lastRow <= 1) {
      if(verboseFlag === true) {
        console.log(`Sheet ${sheet.getName()} is empty or has only headers`);
      }
      return [];
    }
    
    return sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  } catch (e) {
    console.error(`Error getting data from sheet ${sheet.getName()}: ${e.toString()}`);
    return [];
  }
}

function getSpreadsheet(ssAux) {
  try {
    const ssName = ssAux.slice(0, -2);
    
    switch(ssName) {
      case 'DubAppActive': return ssActive;
      case 'DubAppTotal': return ssTotal;
      case 'DubAppLogs': return ssLogs;
      case 'DubAppNoTrack': return ssNoTrack;
      default: throw new Error(`Spreadsheet desconocido: ${ssAux}`);
    }
  } catch (err) {
    console.error(`Error getting spreadsheet ${ssAux}: ${err.toString()}`);
    throw err;
  }
}