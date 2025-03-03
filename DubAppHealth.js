/*
### Estructura General
DubAppHealth.js
├── checkActive()
│   ├── getProgressSheet()
│   ├── handleTriggeredExecution()
│   └── processSheetWithLock()
│       └── checkSheet() [llamada paralela para diferentes hojas]
└── Funciones Utilitarias
    ├── checkDuplicates()
    ├── checkVersions()
    ├── dateHandle()
    ├── isValidDate() 
    ├── String.prototype.replaceAll()
    └── daysBetweenToday()

### Explicación de Funciones

Funciones Principales:
1. checkActive()
   - Función principal que inicia y coordina el proceso de verificación
   - Gestiona el procesamiento paralelo de múltiples hojas
   - Monitorea el progreso y maneja errores globales
   - Utiliza triggers para paralelización

2. getProgressSheet()
   - Obtiene o crea la hoja de seguimiento de progreso
   - Mantiene registro del estado de procesamiento de cada hoja

3. handleTriggeredExecution()
   - Maneja la ejecución de triggers individuales
   - Coordina el procesamiento de cada hoja
   - Limpia recursos después de la ejecución

4. processSheetWithLock()
   - Procesa una hoja individual con protección de Lock Service
   - Maneja reintentos y registro de progreso
   - Previene conflictos de escritura simultánea

5. checkSheet(sourceSS, sheet2check)
   - Compara una hoja específica entre dos spreadsheets
   - Verifica duplicados y diferencias en timestamps
   - Registra casos que necesitan ser actualizados

Funciones Utilitarias:
6. dateHandle(d, timezone, timestamp_format)
   - Formatea fechas según zona horaria y formato especificado
   
7. isValidDate(d)
   - Valida si un objeto es una fecha válida

8. String.prototype.replaceAll(search, replacement)
   - Extensión del prototipo String para reemplazar todas las ocurrencias

9. daysBetweenToday(dateParam)
   - Calcula días entre una fecha dada y hoy

Variables y Constantes Importantes:
- PROGRESS_STATUS: Estados de progreso (PENDING, IN_PROGRESS, COMPLETED, FAILED)
- LOCK_TIMEOUT_SECONDS: Tiempo máximo de bloqueo (300 segundos)
- PARALLEL_BATCH_SIZE: Número de hojas a procesar en paralelo (3)
- Configuración de timezone y formatos de timestamp
- Arrays y objetos para tracking de diferentes hojas

El sistema está diseñado para mantener la sincronización entre diferentes hojas de 
cálculo en un sistema de gestión de dubbing, con énfasis en:
- Procesamiento paralelo eficiente
- Control de versiones
- Detección de inconsistencias
- Registro detallado de progreso
- Recuperación ante fallos
*/

//Global declaration
const allIDs = databaseID.getID();
const timezone = allIDs.timezone;
const timestamp_format = allIDs.timestamp_format;
const timestamp_format2 = "yyyy/MM/dd HH:mm:ss"; // Timestamp alt Format.

const sheetCache = {
  initialized: false,
  sheets: new Map()
};

//Flexible loading
let ssActive, ssTotal, ssLogs, ssNoTrack;
let labelValues, userValues, labelNDX, userNDX;
var sourceSheet = null;
var sourceValues = null;
var sourceNDX = null;
var sourceRow = null;
var witnessSheet = null;  
var witnessValues = null;
var witnessNDX = null;
var witnessRow = null;
var logSheet = null;
var logValues = null;
var logNDX = null;
var containerSheet = null;
var containerValues = null;
var containerNDX = null;
var containerNDX2 = null;
var conTask = null;
var countDWO2Total = 0;

//Previous loading witness
var ssWitness = null;
var sheetNameWitness = null;

//Count check
var counter = {
  dwo: 0,
  production: 0,
  synopsisproject: 0, 
  event: 0,
  synopsisproduction: 0,
  mixandedit: 0,
  observation: 0};

var counterRows = {
  dwo: [],
  production: [],
  synopsisproject: [], 
  event: [],
  synopsisproduction: [],
  mixandedit: [],
  observation: []};

//General variables
var resendCases = 0; var duplicateTotal = 0; var duplicateDetail = "";

// Agregar al inicio del archivo después de las variables globales
const PROGRESS_SHEET_NAME = "Health_Check_Progress";
const PROGRESS_STATUS = {
  PENDING: "PENDING",
  IN_PROGRESS: "IN_PROGRESS",
  COMPLETED: "COMPLETED",
  FAILED: "FAILED"
};

// Ajustar constantes al inicio del archivo
const LOCK_TIMEOUT_SECONDS = 600; // 10 minutos para locks normales
const SCRIPT_LOCK_KEY = 'DUB_APP_HEALTH_CHECK';
const PARALLEL_BATCH_SIZE = 2; // Reducir a 2 hojas simultáneas para evitar sobrecarga

// Constantes para el manejo de hojas grandes
const LARGE_SHEETS = {
  'DWO_FilesLines': {
    timeout: 1200,  // 20 minutos
    batchSize: 5000 // procesar en lotes de 5000 filas
  }
};

function getProgressSheet() {
  const ss = SpreadsheetApp.openById(allIDs['controlID']);
  let progressSheet = ss.getSheetByName(PROGRESS_SHEET_NAME);
  
  if (!progressSheet) {
    progressSheet = ss.insertSheet(PROGRESS_SHEET_NAME);
    progressSheet.appendRow([
      "Fecha", "Hoja", "Estado", "Filas_Procesadas", 
      "Total_Filas", "Errores", "Resend_Cases", "Duplicate_Total"
    ]);
  }
  return progressSheet;
}

function processSheetWithLock(hoja, fecha, progressSheet) {
  const lock = LockService.getScriptLock();
  try {
    // Intenta obtener un lock para esta hoja específica
    if (lock.tryLock(LOCK_TIMEOUT_SECONDS * 1000)) {
      console.log(`Lock obtenido para hoja: ${hoja}`);
      
      // Registrar inicio de procesamiento
      progressSheet.appendRow([
        fecha, hoja, PROGRESS_STATUS.IN_PROGRESS, 
        0, 0, "", resendCases, duplicateTotal
      ]);

      let retryCount = 0;
      const maxRetries = 3;
      let result = false;

      while (retryCount < maxRetries) {
        try {
          result = checkSheet("DubAppActive01", hoja);
          break;
        } catch (error) {
          retryCount++;
          if (retryCount === maxRetries) {
            throw error;
          }
          console.log(`Intento ${retryCount} fallido para ${hoja}, reintentando en 5 segundos...`);
          Utilities.sleep(5000);
        }
      }
      
      // Actualizar progreso
      progressSheet.appendRow([
        fecha, 
        hoja, 
        result ? PROGRESS_STATUS.COMPLETED : PROGRESS_STATUS.SKIPPED,
        sourceValues?.length || 0, 
        witnessValues?.length || 0,
        result ? "" : "Hoja saltada",
        resendCases,
        duplicateTotal
      ]);

      return result;
    } else {
      console.warn(`No se pudo obtener lock para ${hoja} - ya está siendo procesada`);
      return false;
    }
  } catch (error) {
    console.error(`Error procesando ${hoja}: ${error.message}`);
    progressSheet.appendRow([
      fecha, hoja, PROGRESS_STATUS.FAILED,
      sourceValues?.length || 0, witnessValues?.length || 0,
      error.message, resendCases, duplicateTotal
    ]);
    return false;
  } finally {
    if (lock.hasLock()) {
      lock.releaseLock();
      console.log(`Lock liberado para hoja: ${hoja}`);
    }
  }
}

function checkActive() {
  const startTime = new Date();
  console.log(`[${startTime.toISOString()}] Iniciando DubApp Health Check`);
  
  try {
    const progressSheet = getProgressSheet();
    const fecha = new Date();
    
    // Obtener la estructura desde databaseID
    console.log('Obteniendo estructura de hojas...');
    const estructura = databaseID.getStructure();
    const hojas = Object.keys(estructura);
    
    // Identificar hojas grandes
    const hojasGrandes = ['DWO-Event', 'DWO_FilesLines'];
    const hojasNormales = hojas.filter(hoja => !hojasGrandes.includes(hoja));
    
    console.log('Iniciando procesamiento paralelo...');
    
    // Iniciar procesamiento de hojas grandes inmediatamente
    console.log('Procesando hojas grandes...');
    hojasGrandes.forEach(hoja => {
      const triggerId = ScriptApp.newTrigger('handleTriggeredExecution')
        .timeBased()
        .after(1000)
        .create()
        .getUniqueId();
      
      PropertiesService.getScriptProperties().setProperty(
        triggerId,
        JSON.stringify({
          hoja: hoja,
          fecha: fecha.getTime(),
          started: new Date().getTime(),
          isLargeSheet: true,
          batchSize: 5000 // tamaño de lote para hojas grandes
        })
      );
      
      console.log(`Trigger creado para hoja grande: ${hoja}`);
    });
    
    // Procesar hojas normales en paralelo
    console.log('Procesando hojas normales...');
    procesarLotes(hojasNormales, fecha, progressSheet);
    
    // Esperar a que terminen las hojas grandes
    console.log('Esperando finalización de hojas grandes...');
    const largeSheetTimeout = 1200; // 20 minutos
    let timeoutCounter = 0;
    
    while (true) {
      const pendingTriggers = PropertiesService.getScriptProperties()
        .getKeys()
        .filter(key => {
          const prop = PropertiesService.getScriptProperties().getProperty(key);
          return prop && JSON.parse(prop).isLargeSheet;
        });
      
      if (pendingTriggers.length === 0) break;
      
      if (timeoutCounter % 30 === 0) {
        console.log(`Hojas grandes aún en proceso (${timeoutCounter} segundos):`);
        pendingTriggers.forEach(key => {
          const data = JSON.parse(PropertiesService.getScriptProperties().getProperty(key));
          console.log(`- ${data.hoja}: ${(new Date().getTime() - data.started) / 1000}s`);
        });
      }
      
      Utilities.sleep(5000);
      timeoutCounter += 5;
      
      if (timeoutCounter > largeSheetTimeout) {
        throw new Error('Timeout esperando hojas grandes');
      }
    }

    // Verificar resultados finales
    const endTime = new Date();
    const executionTime = (endTime - startTime) / 1000;
    
    console.log('\n=== Resumen Final ===');
    console.log(`Tiempo total de ejecución: ${executionTime} segundos`);
    
  } catch (error) {
    console.error(`[ERROR] en checkActive: ${error.message}`);
    console.error('Stack trace:', error.stack);
    throw error;
  }
}

function handleTriggeredExecution(e) {
  const triggerId = e.triggerUid;
  const props = PropertiesService.getScriptProperties();
  const triggerData = JSON.parse(props.getProperty(triggerId));
  
  if (!triggerData) return; // El trigger ya fue procesado
  
  try {
    const progressSheet = getProgressSheet();
    processSheetWithLock(
      triggerData.hoja,
      new Date(triggerData.fecha),
      progressSheet
    );
  } finally {
    // Limpiar el trigger y sus datos
    ScriptApp.getProjectTriggers()
      .filter(t => t.getUniqueId() === triggerId)
      .forEach(t => ScriptApp.deleteTrigger(t));
    props.deleteProperty(triggerId);
  }
}

function checkSheet(sourceSS, sheet2check) {
  try {
    // Inicializar CON-TaskCurrent
    const controlSS = SpreadsheetApp.openById(allIDs['controlID']);
    const taskSheet = controlSS.getSheetByName('CON-TaskCurrent');
    if (!taskSheet) {
      throw new Error('No se encontró la hoja CON-TaskCurrent');
    }

    // Formatear fecha actual para el comentario
    const currentTimestamp = dateHandle(new Date(), timezone, timestamp_format);
    const checkComment = "checkSheet " + currentTimestamp;

    // Verificar que no sea una hoja especial
    if (sheet2check === 'Index' || sheet2check === 'DWO-LogLabels') {
      console.log(`Saltando hoja especial: ${sheet2check}`);
      return false;
    }

    //Obtain confront SS
    witnessSS = "DubAppTotal01";

    console.log(`Intentando cargar hoja ${sheet2check} desde ${sourceSS}...`);
    
    //Load source
    try {
      LazyLoad(sourceSS, sheet2check);
      if (!containerSheet || !containerValues) {
        throw new Error(`LazyLoad no pudo cargar ${sheet2check} desde ${sourceSS}`);
      }
    } catch (error) {
      console.error(`Error cargando source: ${error.message}`);
      return false;
    }
    
    sourceSheet = containerSheet;
    sourceValues = containerValues;
    sourceNDX = containerNDX;
    
    console.log(`Intentando cargar hoja ${sheet2check} desde ${witnessSS}...`);
    
    //Load witness
    try {
      LazyLoad(witnessSS, sheet2check);
      if (!containerSheet || !containerValues) {
        throw new Error(`LazyLoad no pudo cargar ${sheet2check} desde ${witnessSS}`);
      }
    } catch (error) {
      console.error(`Error cargando witness: ${error.message}`);
      return false;
    }
    
    witnessSheet = containerSheet;
    witnessValues = containerValues;
    witnessNDX = containerNDX;

    // Solo procedemos con la comparación si ambas cargas fueron exitosas
    if (!sourceValues || !witnessValues) {
      console.log(`Saltando comparación de ${sheet2check} - datos incompletos`);
      return false;
    }

    console.log(`Procesando ${sheet2check}: ${sourceValues?.length || 0} filas en source, ${witnessValues?.length || 0} filas en witness`);

    //Clear filters
    let aux = sourceSheet.getFilter(); 
    if (aux != null) {
      try {
        aux.remove();
      } catch (error) {
        console.warn(`No se pudo eliminar el filtro en source: ${error.message}`);
      }
    }
    
    aux = witnessSheet.getFilter(); 
    if (aux != null) {
      try {
        aux.remove();
      } catch (error) {
        console.warn(`No se pudo eliminar el filtro en witness: ${error.message}`);
      }
    }
    
    // Obtain label data
    if (!labelNDX || !labelValues) {
      console.warn('Datos de etiquetas no inicializados correctamente');
      return false;
    }

    labelRow = labelNDX.indexOf(sheet2check, 0);
    if (labelRow === -1) {
      console.warn(`No se encontraron etiquetas para ${sheet2check} - saltando verificación`);
      return false;
    }

    taskLabelNames = labelValues[labelRow];
    taskLabelActions = labelValues[labelRow + 1];
    
    if (!taskLabelNames || !taskLabelActions) {
      console.warn(`Configuración de etiquetas incompleta para ${sheet2check} - saltando verificación`);
      return false;
    }

    taskColKey = taskLabelActions.indexOf("K") - 1;
    taskColUser = taskLabelNames.indexOf("Last user", 0) - 1; 
    taskColChange = taskLabelNames.indexOf("Last change", 0) - 1;

    if (taskColKey === -2 || taskColUser === -2 || taskColChange === -2) {
      throw new Error(`Columnas requeridas no encontradas en ${sheet2check}`);
    }

    // Array para acumular las filas a insertar en CON-TaskCurrent
    const tasksToAppend = [];

    if (LARGE_SHEETS[sheet2check]) {
      const config = LARGE_SHEETS[sheet2check];
      const totalRows = sourceValues.length;
      let processedRows = 0;
      
      // Procesar en lotes más pequeños
      for (let i = 0; i < sourceValues.length; i += config.batchSize) {
        const endIndex = Math.min(i + config.batchSize, sourceValues.length);
        const batchValues = sourceValues.slice(i, endIndex);
        
        // Procesar el lote
        for (const sourceRow of batchValues) {
          const sourceKey = sourceRow[taskColKey];
          
          if (!sourceKey) continue; // Saltar filas sin clave

          // Buscar la fila correspondiente en witness
          const witnessRowIndex = witnessValues.findIndex(row => row[taskColKey] === sourceKey);
          
          if (witnessRowIndex === -1) {
            // Caso NO_ENCONTRADO - registrar para inserción
            tasksToAppend.push([
              sheet2check,                                    // Table
              sourceKey,                                      // Key
              sourceRow[taskColChange],                       // Timestamp
              "DubAppActive" + allIDs.instalation,           // Origin
              "INSERT_ROW",                                   // Action type
              sourceRow[taskColUser],                        // User
              "01 Pending",                                  // Intake Status
              null,                                          // Columna H (null)
              checkComment                                   // Columna I (checkSheet comment)
            ]);
            resendCases++;
            continue;
          }

          const witnessRow = witnessValues[witnessRowIndex];
          
          // Verificar timestamps
          const sourceDate = new Date(sourceRow[taskColChange]);
          const witnessDate = new Date(witnessRow[taskColChange]);
          
          if (isValidDate(sourceDate) && isValidDate(witnessDate)) {
            if (sourceDate > witnessDate) {
              // Caso DESACTUALIZADO - registrar para modificación
              tasksToAppend.push([
                sheet2check,                                // Table
                sourceKey,                                  // Key
                sourceRow[taskColChange],                   // Timestamp
                "DubAppActive" + allIDs.instalation,       // Origin
                "EDIT",                                     // Action type
                sourceRow[taskColUser],                    // User
                "01 Pending",                              // Intake Status
                null,                                      // Columna H (null)
                checkComment                               // Columna I (checkSheet comment)
              ]);
              resendCases++;
            }
          }
        }

        // Agregar el lote actual a CON-TaskCurrent
        if (tasksToAppend.length > 0) {
          taskSheet.getRange(taskSheet.getLastRow() + 1, 1, tasksToAppend.length, tasksToAppend[0].length)
            .setValues(tasksToAppend);
          tasksToAppend.length = 0; // Limpiar el array después de escribir
        }
        
        processedRows = endIndex;
        const progress = Math.round((processedRows / totalRows) * 100);
        
        // Actualizar progreso en las propiedades
        const triggerId = PropertiesService.getScriptProperties().getKeys()
          .find(key => {
            const prop = PropertiesService.getScriptProperties().getProperty(key);
            return prop && JSON.parse(prop).hoja === sheet2check;
          });
        
        if (triggerId) {
          const prop = JSON.parse(PropertiesService.getScriptProperties().getProperty(triggerId));
          prop.progress = `${progress}%`;
          PropertiesService.getScriptProperties().setProperty(triggerId, JSON.stringify(prop));
        }
        
        console.log(`[${new Date().toISOString()}] ${sheet2check}: Procesado ${progress}% (${processedRows}/${totalRows} filas)`);
      }
    } else {
      // Procesamiento normal para hojas pequeñas
      for (let i = 0; i < sourceValues.length; i++) {
        const sourceRow = sourceValues[i];
        const sourceKey = sourceRow[taskColKey];
        
        if (!sourceKey) continue; // Saltar filas sin clave

        // Buscar la fila correspondiente en witness
        const witnessRowIndex = witnessValues.findIndex(row => row[taskColKey] === sourceKey);
        
        if (witnessRowIndex === -1) {
          // Caso NO_ENCONTRADO - registrar para inserción
          tasksToAppend.push([
            sheet2check,                                    // Table
            sourceKey,                                      // Key
            sourceRow[taskColChange],                       // Timestamp
            "DubAppActive" + allIDs.instalation,           // Origin
            "INSERT_ROW",                                   // Action type
            sourceRow[taskColUser],                        // User
            "01 Pending",                                  // Intake Status
            null,                                          // Columna H (null)
            checkComment                                   // Columna I (checkSheet comment)
          ]);
          resendCases++;
          continue;
        }

        const witnessRow = witnessValues[witnessRowIndex];
        
        // Verificar timestamps
        const sourceDate = new Date(sourceRow[taskColChange]);
        const witnessDate = new Date(witnessRow[taskColChange]);
        
        if (isValidDate(sourceDate) && isValidDate(witnessDate)) {
          if (sourceDate > witnessDate) {
            // Caso DESACTUALIZADO - registrar para modificación
            tasksToAppend.push([
              sheet2check,                                // Table
              sourceKey,                                  // Key
              sourceRow[taskColChange],                   // Timestamp
              "DubAppActive" + allIDs.instalation,       // Origin
              "EDIT",                                     // Action type
              sourceRow[taskColUser],                    // User
              "01 Pending",                              // Intake Status
              null,                                      // Columna H (null)
              checkComment                               // Columna I (checkSheet comment)
            ]);
            resendCases++;
          }
        }
      }

      // Agregar las filas acumuladas a CON-TaskCurrent
      if (tasksToAppend.length > 0) {
        taskSheet.getRange(taskSheet.getLastRow() + 1, 1, tasksToAppend.length, tasksToAppend[0].length)
          .setValues(tasksToAppend);
      }
    }
    
    return true;
  } catch (error) {
    console.error(`Error en checkSheet para ${sheet2check}: ${error.message}`);
    return false;
  }
}

/*GENERAL UTILITIES*/

function dateHandle(d,timezone, timestamp_format) {
  if ( isValidDate(d) )
  {
    return Utilities.formatDate(d, timezone, timestamp_format);
  }
  else
  {
    return d;
  }
}

function isValidDate(d) {
  if ( Object.prototype.toString.call(d) !== "[object Date]" )
    return false;
  return !isNaN(d.getTime());
}

String.prototype.replaceAll = function(search, replacement) {
  var target = this;
  if(search == null || replacement == null) {
    return target;
  } else {
    return target.replace(new RegExp(search, 'g'), replacement);
  }
};

function daysBetweenToday(dateParam) {
  let auxNow = new Date();
  let auxDays = Math.floor((auxNow.getTime() - dateParam.getTime()) / (1000 * 60 * 60 * 24));
  return auxDays; 
}

// Función auxiliar para procesar lotes
function procesarLotes(hojas, fecha, progressSheet) {
  let processingStats = {
    currentBatch: 0,
    completed: 0,
    failed: 0,
    skipped: 0,
    total: hojas.length
  };

  console.log(`Iniciando procesamiento de ${hojas.length} hojas normales`);

  for (let i = 0; i < hojas.length; i += PARALLEL_BATCH_SIZE) {
    processingStats.currentBatch++;
    const batch = hojas.slice(i, i + PARALLEL_BATCH_SIZE);
    
    console.log(`\n[Lote ${processingStats.currentBatch}] Iniciando procesamiento de ${batch.length} hojas:`);
    console.log(`Hojas en este lote: ${batch.join(', ')}`);
    
    // Crear triggers para procesamiento paralelo
    const triggers = batch.map(hoja => {
      const triggerId = ScriptApp.newTrigger('handleTriggeredExecution')
        .timeBased()
        .after(1000)
        .create()
        .getUniqueId();
      
      PropertiesService.getScriptProperties().setProperty(
        triggerId,
        JSON.stringify({
          hoja: hoja,
          fecha: fecha.getTime(),
          started: new Date().getTime()
        })
      );
      
      console.log(`[${new Date().toISOString()}] Trigger creado para hoja: ${hoja} (ID: ${triggerId})`);
      return triggerId;
    });
    
    // Esperar a que termine el lote actual
    console.log(`[${new Date().toISOString()}] Esperando finalización del lote ${processingStats.currentBatch}...`);
    let timeoutCounter = 0;
    
    while (triggers.some(id => PropertiesService.getScriptProperties().getProperty(id))) {
      Utilities.sleep(5000);
      timeoutCounter += 5;
      
      if (timeoutCounter % 30 === 0) {
        console.log(`... aún procesando lote ${processingStats.currentBatch} (${timeoutCounter} segundos)`);
        triggers.forEach(id => {
          const prop = PropertiesService.getScriptProperties().getProperty(id);
          if (prop) {
            const data = JSON.parse(prop);
            console.log(`  - ${data.hoja}: ${(new Date().getTime() - data.started) / 1000}s`);
          }
        });
      }
      
      if (timeoutCounter > LOCK_TIMEOUT_SECONDS) {
        console.error(`[${new Date().toISOString()}] Timeout en lote ${processingStats.currentBatch}`);
        console.error('Estado final de triggers:');
        triggers.forEach(id => {
          const prop = PropertiesService.getScriptProperties().getProperty(id);
          if (prop) {
            const data = JSON.parse(prop);
            console.error(`  - ${data.hoja}: ${(new Date().getTime() - data.started) / 1000}s`);
          }
        });
        throw new Error(`Timeout después de ${LOCK_TIMEOUT_SECONDS} segundos esperando procesamiento del lote ${processingStats.currentBatch}`);
      }
    }
    
    // Actualizar estadísticas del lote - CORREGIDO
    const lastRow = progressSheet.getLastRow();
    if (lastRow > 1) {
      const allResults = progressSheet.getRange(2, 1, lastRow-1, 8).getValues();
      const currentResults = allResults.filter(row => 
        new Date(row[0]).getTime() === fecha.getTime() && 
        hojas.includes(row[1])
      );

      // Reiniciar contadores antes de recalcular
      processingStats.completed = 0;
      processingStats.failed = 0;
      processingStats.skipped = 0;

      // Contar todos los resultados hasta ahora
      currentResults.forEach(row => {
        switch(row[2]) {
          case PROGRESS_STATUS.COMPLETED:
            processingStats.completed++;
            break;
          case PROGRESS_STATUS.FAILED:
            processingStats.failed++;
            break;
          case PROGRESS_STATUS.SKIPPED:
            processingStats.skipped++;
            break;
        }
      });
    }

    console.log(`[Lote ${processingStats.currentBatch}] Completado`);
    console.log('Estadísticas actuales:', {
      'Completadas': processingStats.completed,
      'Fallidas': processingStats.failed,
      'Saltadas': processingStats.skipped,
      'Pendientes': processingStats.total - (processingStats.completed + processingStats.failed + processingStats.skipped)
    });
  }

  return processingStats;
}

function unlockAllSheets() {
  try {
    console.log('Iniciando liberación de todos los locks...');
    
    // Obtener todos los locks activos
    const scriptLock = LockService.getScriptLock();
    if (scriptLock.hasLock()) {
      scriptLock.releaseLock();
      console.log('Script lock principal liberado');
    }

    // Obtener la estructura desde databaseID
    const estructura = databaseID.getStructure();
    const hojas = Object.keys(estructura);
    
    // Intentar liberar el lock de cada hoja
    hojas.forEach(hoja => {
      try {
        const lock = LockService.getScriptLock();
        if (lock.hasLock()) {
          lock.releaseLock();
          console.log(`Lock liberado para hoja: ${hoja}`);
        }
      } catch (error) {
        console.warn(`No se pudo liberar lock para ${hoja}: ${error.message}`);
      }
    });

    // Limpiar todas las propiedades de script relacionadas con triggers
    const scriptProperties = PropertiesService.getScriptProperties();
    const allKeys = scriptProperties.getKeys();
    const triggerKeys = allKeys.filter(key => {
      const prop = scriptProperties.getProperty(key);
      try {
        return prop && JSON.parse(prop).hoja;
      } catch {
        return false;
      }
    });

    triggerKeys.forEach(key => {
      scriptProperties.deleteProperty(key);
      console.log(`Propiedad de trigger eliminada: ${key}`);
    });

    // Eliminar todos los triggers existentes
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
      try {
        ScriptApp.deleteTrigger(trigger);
        console.log('Trigger eliminado:', trigger.getUniqueId());
      } catch (error) {
        console.warn(`No se pudo eliminar trigger: ${error.message}`);
      }
    });

    console.log('Proceso de liberación de locks completado exitosamente');
    return true;
  } catch (error) {
    console.error('Error al liberar locks:', error.message);
    return false;
  }
}