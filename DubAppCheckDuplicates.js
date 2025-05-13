/**
 * @fileoverview Script para detectar y gestionar registros duplicados en hojas de cálculo.
 * 
 * Este script busca registros duplicados basándose en una columna clave definida en la estructura
 * de la base de datos. Puede operar en dos modos:
 * - Modo visualización: Solo muestra los duplicados encontrados
 * - Modo limpieza: Limpia el contenido de los registros duplicados más antiguos
 * 
 */

// Variables de estado globales
let lastProcessedSheet = 0;
let lastProcessedRow = 0;
const MAX_EXECUTION_TIME = 350000; // 5.8 minutos (dejando un pequeño margen)
const TRIGGER_FUNCTION_NAME = 'checkDuplicates';
const CHUNK_SIZE = 2000; // Aumentar el tamaño del chunk para procesar más filas
const startTime = Date.now();

// Variables globales de la librería databaseID
const allIDs = databaseID.getID();

// Constante para el ID por defecto
const DEFAULT_WORKSHEET_ID = allIDs.activeID;

/**
 * Guarda el estado actual de procesamiento
 */
function saveState(sheetIndex, rowIndex, worksheetId, eraseFlag) {
  PropertiesService.getScriptProperties().setProperties({
    'lastProcessedSheet': sheetIndex.toString(),
    'lastProcessedRow': rowIndex.toString(),
    'worksheetId': worksheetId,
    'eraseFlag': eraseFlag.toString()
  });
}

/**
 * Limpia el estado guardado
 */
function clearState() {
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.deleteAllProperties();
  Logger.log('Estado limpiado');
}

/**
 * Verifica si se está cerca del timeout
 */
function isNearTimeout() {
  const elapsedTime = Date.now() - startTime;
  return elapsedTime > MAX_EXECUTION_TIME;
}

/**
 * Crea un trigger para continuar el proceso
 */
function createTrigger() {
  // Borrar triggers existentes primero
  deleteTriggers();
  
  // Crear nuevo trigger para ejecutar en 1 minuto
  ScriptApp.newTrigger(TRIGGER_FUNCTION_NAME)
    .timeBased()
    .after(60 * 1000) // 1 minuto
    .create();
  
  Logger.log('Nuevo trigger creado para continuar el proceso en 1 minuto');
}

/**
 * Borra todos los triggers existentes para esta función
 */
function deleteTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === TRIGGER_FUNCTION_NAME) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  Logger.log('Triggers anteriores eliminados');
}

/**
 * Verifica si tenemos acceso a la worksheet con reintentos
 */
function verifyAccess(worksheetId) {
  const maxRetries = 3;
  let retryCount = 0;

  while (retryCount < maxRetries) {
    try {
      const spreadsheet = SpreadsheetApp.openById(worksheetId);
      spreadsheet.getName(); // Verificar que podemos acceder
      return true;
    } catch (error) {
      retryCount++;
      if (retryCount === maxRetries) {
        Logger.log(`Error de acceso a worksheet ${worksheetId}: ${error.message}`);
        Logger.log('Por favor, verifica que:');
        Logger.log('1. El ID de la worksheet es correcto');
        Logger.log('2. Tienes permisos de acceso al documento');
        Logger.log('3. El documento existe');
        return false;
      }
      Logger.log(`Intento ${retryCount} fallido, reintentando verificación de acceso...`);
      Utilities.sleep(2000);
    }
  }
  return false;
}

/**
 * Función principal que inicia el proceso de verificación de duplicados.
 * @param {string} worksheetId - ID de la hoja de cálculo a procesar (opcional)
 * @param {boolean} eraseFlag - Si es true, limpia los registros duplicados más antiguos (opcional)
 */
function call(worksheetId = null, eraseFlag = false) {
  const lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(10000)) {
      Logger.log('Script ya en ejecución. Saliendo...');
      return;
    }

    // Si no se proporciona worksheetId, usar el ID por defecto
    const targetWorksheetId = worksheetId || DEFAULT_WORKSHEET_ID;

    // Verificar acceso antes de continuar
    if (!verifyAccess(targetWorksheetId)) {
      Logger.log('No se pudo acceder a la worksheet. Abortando proceso.');
      return;
    }

    // Verificar si hay un estado guardado
    const scriptProps = PropertiesService.getScriptProperties();
    const lastSheet = scriptProps.getProperty('lastProcessedSheet');
    const lastRow = scriptProps.getProperty('lastProcessedRow');

    if (!lastSheet && !lastRow) {
      Logger.log('=== INICIANDO NUEVO PROCESO DE DUPLICADOS ===');
    } else {
      Logger.log('=== CONTINUANDO PROCESO EXISTENTE ===');
    }

    Logger.log(`Worksheet ID: ${targetWorksheetId}`);
    Logger.log(`Modo borrado: ${eraseFlag ? 'Activado' : 'Solo lectura'}`);
    Logger.log('----------------------------------------');
    
    // Iniciar el proceso
    checkDuplicates(targetWorksheetId, eraseFlag);

  } catch(err) {
    Logger.log(`Error en call: ${err.toString()}`);
    Logger.log(`Stack trace: ${err.stack}`);
    console.error(err);
  } finally {
    if (lock.hasLock()) {
      lock.releaseLock();
    }
  }
}

/**
 * Verifica y gestiona registros duplicados en las hojas de cálculo especificadas.
 * 
 * @param {string} worksheetId - ID de la hoja de cálculo a revisar
 * @param {boolean} eraseFlag - Si es true, limpia los registros duplicados más antiguos
 */
function checkDuplicates(worksheetId, eraseFlag = false) {
  const startTime = Date.now();
  const lock = LockService.getScriptLock();
  
  try {
    // Intentar obtener el lock
    if (!lock.tryLock(30000)) {
      Logger.log('No se pudo obtener el lock. Otra instancia está ejecutándose.');
      return;
    }

    // Recuperar estado anterior
    const scriptProps = PropertiesService.getScriptProperties();
    lastProcessedSheet = parseInt(scriptProps.getProperty('lastProcessedSheet')) || 0;
    lastProcessedRow = parseInt(scriptProps.getProperty('lastProcessedRow')) || 0;
    worksheetId = scriptProps.getProperty('worksheetId') || worksheetId;
    eraseFlag = scriptProps.getProperty('eraseFlag') === 'true' || eraseFlag;

    // Solo abrir la worksheet si es la primera ejecución o si no tenemos la referencia
    if (lastProcessedSheet === 0 && lastProcessedRow === 0) {
      let spreadsheet = null;
      let retryCount = 0;
      const maxRetries = 3;

      while (!spreadsheet && retryCount < maxRetries) {
        try {
          spreadsheet = SpreadsheetApp.openById(worksheetId);
          if (!spreadsheet) throw new Error('No se pudo abrir la worksheet');
        } catch (error) {
          retryCount++;
          if (retryCount === maxRetries) {
            Logger.log('Error crítico al abrir la worksheet. Deteniendo el proceso.');
            Logger.log('Por favor, verifica el acceso al documento y reinicia manualmente.');
            clearState();
            deleteTriggers();
            return;
          }
          Logger.log(`Intento ${retryCount} fallido, reintentando en 2 segundos...`);
          Utilities.sleep(2000);
        }
      }
    }

    // Obtener la referencia a la worksheet
    const spreadsheet = SpreadsheetApp.openById(worksheetId);
    const sheets = spreadsheet.getSheets();

    // Obtener todas las hojas
    if (!sheets || sheets.length === 0) {
      throw new Error('La worksheet no contiene hojas');
    }

    // Modificar el mensaje del log para mostrar el estado real
    if (lastProcessedSheet > 0 || lastProcessedRow > 0) {
      Logger.log(`Retomando proceso desde Hoja ${lastProcessedSheet}, Fila ${lastProcessedRow}`);
    } else {
      Logger.log('Iniciando nuevo proceso desde el principio');
    }
    Logger.log(`Worksheet ID: ${worksheetId}`);

    let allErrorMessages = [];  // Array para acumular mensajes de error

    // Para cada hoja en el spreadsheet
    for (let i = lastProcessedSheet; i < sheets.length; i++) {
      const sheet = sheets[i];
      const sheetName = sheet.getName();
      
      if (sheetName === 'Index' || sheetName === 'DWO-LogLabels') continue;

      Logger.log(`Procesando hoja: ${sheetName}`);
      const data = sheet.getDataRange().getValues();
      Logger.log(`Número total de filas: ${data.length}`);
      
      if (data.length <= 1) {
        Logger.log(`Hoja ${sheetName} está vacía o solo tiene encabezados. Continuando...`);
        continue;
      }
      
      // Modificar la selección de la columna clave
      const keyColumn = ['DWO', 'DWO-ChannelEventType', 'DWO-Series'].includes(sheetName) ? 1 : 0;
      Logger.log(`Usando columna ${keyColumn + 1} (${String.fromCharCode(65 + keyColumn)}) como clave para ${sheetName}`);
      
      const duplicatesMap = new Map();
      const rowsToDelete = new Set();
      
      // Primera pasada: identificar duplicados
      Logger.log(`Buscando duplicados en la columna ${keyColumn}`);
      let duplicateCount = 0;

      for (let row = (i === lastProcessedSheet ? lastProcessedRow : 1); row < data.length; row++) {
        // Verificar timeout con menos frecuencia
        if (row % CHUNK_SIZE === 0 && isNearTimeout()) {
          saveState(i, row, worksheetId, eraseFlag);
          Logger.log(`Timeout cercano - Guardando estado: Hoja ${i}, Fila ${row}`);
          createTrigger();
          return;
        }

        const key = data[row][keyColumn]?.toString().trim();
        if (!key) continue;
        
        if (duplicatesMap.has(key)) {
          duplicatesMap.get(key).push(row);
          duplicateCount++;
        } else {
          duplicatesMap.set(key, [row]);
        }
      }

      Logger.log(`Se encontraron ${duplicateCount} registros duplicados en ${sheetName}`);
      
      // Si hay duplicados, mostrarlos y acumular mensaje de error
      if (duplicateCount > 0) {
        let errorMessage = `=== ALERTA: DUPLICADOS ENCONTRADOS ===\n`;
        errorMessage += `Worksheet ID: ${worksheetId}\n`;
        errorMessage += `Hoja: ${sheetName}\n`;
        errorMessage += `Total duplicados: ${duplicateCount}\n`;
        
        for (const [key, rows] of duplicatesMap) {
          if (rows.length > 1) {
            const filasStr = rows.map(r => r + 1).join(', ');
            errorMessage += `Valor duplicado "${key}" encontrado en las filas: ${filasStr}\n`;
            Logger.log(`Valor duplicado "${key}" encontrado en las filas: ${filasStr}`);
          }
        }
        
        if (!eraseFlag) { // Acumular error en modo visualización
          allErrorMessages.push(errorMessage);
        }
      }
      
      // Segunda pasada: limpiar duplicados
      if (eraseFlag && duplicatesMap.size > 0) {
        // Verificar timeout solo si hay muchas filas para limpiar
        if (isNearTimeout()) {
          saveState(i, 0, worksheetId, eraseFlag);
          Logger.log(`Timeout durante limpieza - Guardando estado: Hoja ${i}`);
          createTrigger();
          return;
        }

        let clearedCount = 0;
        const rowsToClear = [];

        // Identificar filas a limpiar
        for (const [key, rows] of duplicatesMap) {
          if (rows.length > 1) {
            rows.sort((a, b) => a - b); // Cambio aquí: ordenar de menor a mayor para mantener el más antiguo
            rowsToClear.push(...rows.slice(1)); // Mantiene el primero (más antiguo) y marca el resto para limpiar
            clearedCount += rows.length - 1;
          }
        }

        // Limpiar en bloques para evitar timeouts
        if (rowsToClear.length > 0) {
          rowsToClear.sort((a, b) => b - a); // Ordenar de mayor a menor
          
          for (let j = 0; j < rowsToClear.length; j++) {
            if (j % 100 === 0 && isNearTimeout()) {
              saveState(i, rowsToClear[j], worksheetId, eraseFlag);
              Logger.log(`Timeout durante limpieza - Guardando estado: Hoja ${i}, Fila ${rowsToClear[j]}`);
              createTrigger();
              return;
            }
            
            // Limpiar la fila en lugar de eliminarla
            const numColumns = sheet.getLastColumn();
            sheet.getRange(rowsToClear[j] + 1, 1, 1, numColumns).clearContent();
          }
          
          Logger.log(`${sheetName}: ${clearedCount} filas duplicadas limpiadas`);
        }
      }

      // Actualizar estado al terminar cada hoja
      saveState(i + 1, 0, worksheetId, eraseFlag);
    }
    
    // Al final del proceso, si hay errores acumulados, lanzarlos todos juntos
    if (!eraseFlag && allErrorMessages.length > 0) {
      throw new Error(allErrorMessages.join('\n----------------------------------------\n'));
    }
    
    // Proceso completado
    clearState();
    deleteTriggers();
    Logger.log('\n=== PROCESO COMPLETADO ===');
    
  } catch (error) {
    Logger.log(`Error en checkDuplicates: ${error.message}`);
    Logger.log(error.stack);
    saveState(lastProcessedSheet, lastProcessedRow, worksheetId, eraseFlag);
    createTrigger();
    return;
  } finally {
    if (lock.hasLock()) {
      lock.releaseLock();
    }
  }
}

/**
 * Función inicial que inicia el proceso
 * @param {string} worksheetId - ID de la hoja de cálculo a procesar
 * @param {boolean} eraseFlag - Si es true, limpia los registros duplicados más antiguos
 */
function startDuplicateCheck(worksheetId, eraseFlag = false) {
  if (!worksheetId) {
    throw new Error('Se requiere un ID de worksheet válido');
  }

  // Limpiar estado y triggers anteriores
  clearState();
  deleteTriggers();
  
  // Iniciar el proceso
  checkDuplicates(worksheetId, eraseFlag);
}

/**
 * Función para resetear manualmente el estado del script y los triggers.
 * Usar esta función cuando el proceso se haya quedado en un estado inconsistente.
 */
function resetScript() {
  // Limpiar todas las propiedades del script
  const props = PropertiesService.getScriptProperties();
  props.deleteAllProperties();
  
  // Eliminar todos los triggers existentes
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'call') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  Logger.log('=== RESET COMPLETO ===');
  Logger.log('- Variables de estado limpiadas');
  Logger.log('- Triggers eliminados');
  Logger.log('El script está listo para una nueva ejecución');
}

/**
 * Compara dos worksheets para verificar que todos los casos del primero existan en el segundo.
 * @param {string} sourceWorksheetId - ID del worksheet fuente
 * @param {string} targetWorksheetId - ID del worksheet objetivo para comparar
 */
function compareWorksheets(sourceWorksheetId, targetWorksheetId) {
  try {
    const sourceSpreadsheet = SpreadsheetApp.openById(sourceWorksheetId);
    const targetSpreadsheet = SpreadsheetApp.openById(targetWorksheetId);
    
    Logger.log('=== INICIANDO COMPARACIÓN DE WORKSHEETS ===');
    Logger.log(`Worksheet Fuente: ${sourceSpreadsheet.getName()}`);
    Logger.log(`Worksheet Objetivo: ${targetSpreadsheet.getName()}`);
    Logger.log('----------------------------------------');

    const sourceSheets = sourceSpreadsheet.getSheets();
    let totalCasosFaltantes = 0;
    
    // Procesar cada hoja en chunks para evitar timeouts
    for (const sourceSheet of sourceSheets) {
      if (isNearTimeout()) {
        Logger.log('Tiempo límite alcanzado durante la comparación');
        createTrigger();
        return;
      }

      const sheetName = sourceSheet.getName();
      
      // Ignorar hojas específicas
      if (sheetName === 'Index' || sheetName === 'DWO-LogLabels') continue;
      
      const targetSheet = targetSpreadsheet.getSheetByName(sheetName);
      if (!targetSheet) {
        Logger.log(`ADVERTENCIA: La hoja "${sheetName}" no existe en el worksheet objetivo`);
        continue;
      }

      // Determinar la columna clave
      const keyColumn = sheetName === 'DWO' ? 1 : 0;
      
      // Obtener datos en chunks
      const sourceData = sourceSheet.getDataRange().getValues();
      const targetData = targetSheet.getDataRange().getValues();
      
      const targetKeys = new Set(targetData.slice(1).map(row => row[keyColumn]?.toString().trim()));
      
      let casosFaltantes = 0;
      Logger.log(`\nAnalizando hoja: ${sheetName}`);
      
      // Procesar en chunks
      for (let i = 1; i < sourceData.length; i += CHUNK_SIZE) {
        if (i % CHUNK_SIZE === 0 && isNearTimeout()) {
          Logger.log(`Timeout durante comparación - Guardando estado en hoja ${sheetName}`);
          createTrigger();
          return;
        }

        const endIndex = Math.min(i + CHUNK_SIZE, sourceData.length);
        for (let j = i; j < endIndex; j++) {
          const key = sourceData[j][keyColumn]?.toString().trim();
          if (!key || key === '' || key === 'null' || key === 'undefined') continue;
          
          if (!targetKeys.has(key)) {
            casosFaltantes++;
            totalCasosFaltantes++;
            Logger.log(`Caso faltante en ${sheetName}: ${key} (fila ${j + 1})`);
          }
        }
      }
      
      if (casosFaltantes > 0) {
        Logger.log(`Total casos faltantes en ${sheetName}: ${casosFaltantes}`);
      } else {
        Logger.log(`${sheetName}: Todos los casos están presentes en el objetivo`);
      }
    }
    
    Logger.log('\n=== RESUMEN DE COMPARACIÓN ===');
    Logger.log(`Total de casos faltantes: ${totalCasosFaltantes}`);
    
  } catch(err) {
    Logger.log(`Error en compareWorksheets: ${err.message}`);
    Logger.log(err.stack);
    createTrigger();
  }
}

/**
 * Ejemplo de uso de la función compareWorksheets
 * 
 * @description
 * Esta función compara dos worksheets para verificar que todos los casos del primero existan en el segundo.
 * Es útil para validar que no se hayan perdido datos durante una copia o migración.
 * 
 * @example
 * // Comparar worksheet fuente con worksheet objetivo
 * const sourceId = '1rynSNh2wO6Izuty-DO1Nw5tVu0T-9saDGwleKYETIDM'; // ID de la worksheet original
 * const targetId = DEFAULT_WORKSHEET_ID; // ID de la worksheet copia/destino
 * compareWorksheets(sourceId, targetId);
 * 
 * @param {string} sourceId - ID de la worksheet fuente (original)
 * @param {string} targetId - ID de la worksheet objetivo (copia) para comparar
 * 
 * @returns {void} No retorna valor, pero genera un log detallado con:
 * - Casos faltantes por hoja
 * - ID de los casos faltantes y su ubicación (número de fila)
 * - Resumen total de casos faltantes
 */
function testCompareWorksheets() {
  // ID de la worksheet fuente (original)
  const sourceId = '1y8MTd8lUblM3bUWqe_1_LstkcvqlenRjAg92gyp_eIw';
  
  // ID de la worksheet objetivo (copia) para comparar
  const targetId = DEFAULT_WORKSHEET_ID;
  
  // Ejecutar la comparación
  compareWorksheets(sourceId, targetId);
}

// Ejemplos de uso:

// 1. Llamada básica (usa ID por defecto, solo detecta duplicados)
function testCallBasic() {
  clearState();
  deleteTriggers();
  call();
}

// 2. Llamada con ID específico (solo detecta duplicados)
function testCallWithId() {
  clearState();
  deleteTriggers();
  call(allIDs.totalID);
}

// 3. Llamada con borrado de duplicados
function testCallWithErase() {
  clearState();
  deleteTriggers();
  call(DEFAULT_WORKSHEET_ID, true);
}

// Ejemplo de uso directo
function startDuplicateCheck(worksheetId, eraseFlag = false) {
  clearState();
  deleteTriggers();
  if (!worksheetId) {
    throw new Error('Se requiere un ID de worksheet válido');
  }
  call(worksheetId, eraseFlag);
}