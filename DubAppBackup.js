/*
=== ÁRBOLES DE LLAMADA ===

1. BackupControlQueue() - Trigger programado
   └── BackupActive()
       ├── LazyLoad()
       ├── MakeBackup()
       │   └── CleanFilters()
       └── BackupTotal()
           ├── MakeBackup()
           └── BackupLogs()
               ├── MakeBackup()
               └── BackupNoTrack()
                   ├── MakeBackup()
                   └── FinalCleanup()
                       └── CleanOldBackups()

2. CleanAllSheets() - Trigger programado
   └── SheetClean()
       └── LazyLoad()

3. cleanDubAppControl() - Trigger programado
   └── databaseID.getID()

=== DESCRIPCIÓN DE FUNCIONES ===

BackupControlQueue: Inicia la secuencia de backups
BackupActive: Gestiona el backup de DubAppActive01
BackupTotal: Gestiona el backup de DubAppTotal01
BackupLogs: Gestiona el backup de DubAppLogs01
BackupNoTrack: Gestiona el backup de DubAppNoTrack01
FinalCleanup: Ejecuta la limpieza final de backups antiguos
CleanAllSheets: Limpia filas vacías en todas las hojas
CleanOldBackups: Gestiona la retención de backups según antigüedad
SheetClean: Limpia filas vacías en una hoja específica
MakeBackup: Crea una copia de respaldo de una hoja
LazyLoad: Carga datos de una hoja (función externa)
CleanFilters: Elimina filtros de todas las hojas
setCache: Almacena datos en caché
cleanDubAppControl: Gestiona la limpieza y backup de registros en CON-Task
  - Cambia el estado de procesamiento en CON-Control A2 (false durante proceso)
  - Elimina filas con estado "05 Discarded" o "06 Unchanged"
  - Mantiene filas con estado "08 Jumped off", "04 Retry" o "01 Pending"
  - Mueve las demás filas a CON-TaskBackup
  - Implementa bloqueo para prevenir modificaciones concurrentes
  - Restaura el estado de procesamiento en CON-Control A2 (true al finalizar)

Tabla de llamadas y descripción de funciones:

cleanDubApp()
   └─ cleanConTask() - Limpia y respalda registros de CON-Task
        └─ moveToBackup() - Mueve registros a hoja de respaldo y elimina descartados

Descripción de funciones:
- cleanDubApp: Función principal que inicia el proceso de limpieza
- cleanConTask: Gestiona la limpieza de la hoja CON-Task
- moveToBackup: Mueve registros a hoja de respaldo y elimina registros descartados/sin cambios
*/

//Global declaration
const allIDs = databaseID.getID();
const timezone = allIDs["timezone"];
const timestamp_format = allIDs["timestamp_format"];
var verboseFlag = null;
var descreetFlag = null;

//Count check
var counter = {
  dwo: 0,
  production: 0,
  synopsisproject: 0, 
  event: 0,
  synopsisproduction: 0,
  mixandedit: 0,
  observation: 0
};

function BackupControlQueue() {
  // Verificar fin de semana
  const today = new Date();
  const dayOfWeek = today.getDay();
  if (dayOfWeek === 0 || dayOfWeek === 6) {
    console.log('No se ejecutan backups en fin de semana');
    return;
  }

  // Configurar flags
  const ss = SpreadsheetApp.openById(allIDs['controlID']);
  const ConControl = ss.getSheetByName("CON-Control");
  const controlArray = ConControl.getRange('A2:M2').getValues();
  verboseFlag = controlArray[0][11];
  descreetFlag = controlArray[0][12];

  if(!controlArray[0][0]) {
    if(verboseFlag) console.log('system not operational.');
    return;
  }

  // Limpiar todos los triggers existentes del proceso de backup
  const triggersToClean = ['BackupActive', 'BackupTotal', 'BackupLogs', 'BackupNoTrack', 'FinalCleanup'];
  deleteTriggersByFunctionNames(triggersToClean);

  // Programar backups secuenciales
  ScriptApp.newTrigger('BackupActive')
    .timeBased()
    .after(1 * 60 * 1000) // 1 minuto
    .create();
}

function BackupActive() {
  try {
    const ssActive = SpreadsheetApp.openById(allIDs['activeID']);
    MakeBackup(ssActive, allIDs['activeID']);
    
    // Pequeña pausa para asegurar que el backup se completó
    Utilities.sleep(5000); // 5 segundos de espera
    
    // Intentar crear el trigger con reintentos
    let retryCount = 0;
    const maxRetries = 3;
    let trigger = null;
    
    while (retryCount < maxRetries) {
      try {
        trigger = ScriptApp.newTrigger('BackupTotal')
          .timeBased()
          .after(2 * 60 * 1000)
          .create();
        break; // Si tiene éxito, salir del bucle
      } catch (triggerError) {
        retryCount++;
        if (retryCount === maxRetries) {
          throw new Error(`No se pudo crear el trigger después de ${maxRetries} intentos: ${triggerError.message}`);
        }
        console.log(`Intento ${retryCount} fallido, reintentando en 5 segundos...`);
        Utilities.sleep(5000);
      }
    }
    
    console.log('Trigger para BackupTotal creado exitosamente');
    
  } catch (error) {
    console.error(`Error en BackupActive: ${error.message}`);
    // Intentar enviar email de error directamente
    try {
      MailApp.sendEmail({
        to: Session.getEffectiveUser().getEmail(),
        subject: "Error en proceso de Backup - BackupActive",
        body: `Error en BackupActive:\n${error.message}\n${error.stack}`
      });
    } catch (emailError) {
      console.error(`No se pudo enviar email de error: ${emailError.message}`);
    }
    throw error;
  }
}

function BackupTotal() {
  try {
    const ssTotal = SpreadsheetApp.openById(allIDs['totalID']);
    MakeBackup(ssTotal, allIDs['totalID']);
    
    Utilities.sleep(5000);
    
    let retryCount = 0;
    const maxRetries = 3;
    let trigger = null;
    
    while (retryCount < maxRetries) {
      try {
        trigger = ScriptApp.newTrigger('BackupLogs')
          .timeBased()
          .after(2 * 60 * 1000)
          .create();
        break;
      } catch (triggerError) {
        retryCount++;
        if (retryCount === maxRetries) {
          throw new Error(`No se pudo crear el trigger después de ${maxRetries} intentos: ${triggerError.message}`);
        }
        console.log(`Intento ${retryCount} fallido, reintentando en 5 segundos...`);
        Utilities.sleep(5000);
      }
    }
    
    console.log('Trigger para BackupLogs creado exitosamente');
    
  } catch (error) {
    console.error(`Error en BackupTotal: ${error.message}`);
    try {
      MailApp.sendEmail({
        to: Session.getEffectiveUser().getEmail(),
        subject: "Error en proceso de Backup - BackupTotal",
        body: `Error en BackupTotal:\n${error.message}\n${error.stack}`
      });
    } catch (emailError) {
      console.error(`No se pudo enviar email de error: ${emailError.message}`);
    }
    throw error;
  }
}

function BackupLogs() {
  try {
    const ssLogs = SpreadsheetApp.openById(allIDs['logsID']);
    MakeBackup(ssLogs, allIDs['logsID']);
    
    Utilities.sleep(5000);
    
    let retryCount = 0;
    const maxRetries = 3;
    let trigger = null;
    
    while (retryCount < maxRetries) {
      try {
        trigger = ScriptApp.newTrigger('BackupNoTrack')
          .timeBased()
          .after(2 * 60 * 1000)
          .create();
        break;
      } catch (triggerError) {
        retryCount++;
        if (retryCount === maxRetries) {
          throw new Error(`No se pudo crear el trigger después de ${maxRetries} intentos: ${triggerError.message}`);
        }
        console.log(`Intento ${retryCount} fallido, reintentando en 5 segundos...`);
        Utilities.sleep(5000);
      }
    }
    
    console.log('Trigger para BackupNoTrack creado exitosamente');
    
  } catch (error) {
    console.error(`Error en BackupLogs: ${error.message}`);
    try {
      MailApp.sendEmail({
        to: Session.getEffectiveUser().getEmail(),
        subject: "Error en proceso de Backup - BackupLogs",
        body: `Error en BackupLogs:\n${error.message}\n${error.stack}`
      });
    } catch (emailError) {
      console.error(`No se pudo enviar email de error: ${emailError.message}`);
    }
    throw error;
  }
}

function BackupNoTrack() {
  try {
    const ssNoTrack = SpreadsheetApp.openById(allIDs['noTrackID']);
    MakeBackup(ssNoTrack, allIDs['noTrackID']);
    
    Utilities.sleep(5000);
    
    let retryCount = 0;
    const maxRetries = 3;
    let trigger = null;
    
    while (retryCount < maxRetries) {
      try {
        trigger = ScriptApp.newTrigger('FinalCleanup')
          .timeBased()
          .after(2 * 60 * 1000)
          .create();
        break;
      } catch (triggerError) {
        retryCount++;
        if (retryCount === maxRetries) {
          throw new Error(`No se pudo crear el trigger después de ${maxRetries} intentos: ${triggerError.message}`);
        }
        console.log(`Intento ${retryCount} fallido, reintentando en 5 segundos...`);
        Utilities.sleep(5000);
      }
    }
    
    console.log('Trigger para FinalCleanup creado exitosamente');
    
  } catch (error) {
    console.error(`Error en BackupNoTrack: ${error.message}`);
    try {
      MailApp.sendEmail({
        to: Session.getEffectiveUser().getEmail(),
        subject: "Error en proceso de Backup - BackupNoTrack",
        body: `Error en BackupNoTrack:\n${error.message}\n${error.stack}`
      });
    } catch (emailError) {
      console.error(`No se pudo enviar email de error: ${emailError.message}`);
    }
    throw error;
  }
}

function FinalCleanup() {
  try {
    CleanOldBackups();
  } catch (error) {
    console.error(`Error en FinalCleanup: ${error.message}`);
    throw error;
  }
}

function CleanOldBackups() {
  const backupFolder = DriveApp.getFolderById(allIDs["backupID"]);
  const files = backupFolder.getFiles();
  const now = new Date();
  const twoYearsAgo = new Date(now.getFullYear() - 2, now.getMonth(), 1);
  const sixMonthsAgo = new Date(now.getFullYear(), now.getMonth() - 6, 1);
  const twoMonthsAgo = new Date(now.getFullYear(), now.getMonth() - 2, 1);
  
  // Contadores para el resumen
  let totalFiles = 0;
  let deletedFiles = 0;
  let keptFiles = 0;
  
  // Agrupar archivos por fecha y tipo
  const filesByDate = new Map(); // fecha -> Map(tipo -> [archivos])
  
  while (files.hasNext()) {
    const file = files.next();
    const fileName = file.getName();
    totalFiles++;
    
    // Extraer fecha y tipo del nombre del archivo
    const match = fileName.match(/Backup (\d{4}-\d{2}-\d{2}) \d{2}:\d{2}:\d{2} (DubApp\w+)/);
    if (!match) continue;
    
    const dateStr = match[1];
    const fileType = match[2];
    const fileDate = new Date(dateStr);
    
    if (!filesByDate.has(dateStr)) {
      filesByDate.set(dateStr, new Map());
    }
    const typesMap = filesByDate.get(dateStr);
    if (!typesMap.has(fileType)) {
      typesMap.set(fileType, []);
    }
    typesMap.get(fileType).push(file);
  }

  // Procesar cada fecha
  for (const [dateStr, typesMap] of filesByDate) {
    const fileDate = new Date(dateStr);
    
    // Mantener solo un backup por día según las reglas
    let keepBackup = false;
    
    if (fileDate < twoYearsAgo) {
      // Para > 2 años: eliminar todo
      keepBackup = false;
    } else if (fileDate < sixMonthsAgo) {
      // Para > 6 meses hasta 2 años: mantener solo el primer día del mes
      keepBackup = fileDate.getDate() === 1;
    } else if (fileDate < twoMonthsAgo) {
      // Para > 2 meses hasta 6 meses: mantener solo los viernes
      keepBackup = fileDate.getDay() === 5;
    } else {
      // Para archivos más recientes que 2 meses, mantener todos
      keepBackup = true;
    }

    if (!keepBackup) {
      // Eliminar todos los backups del día
      for (const [fileType, files] of typesMap) {
        files.forEach(file => {
          if(verboseFlag === true) {
            console.log(`Eliminando backup: ${file.getName()}`);
          }
          file.setTrashed(true);
          deletedFiles++;
        });
      }
    } else if (typesMap.size > 0) {
      // Si hay que mantener el backup, conservar solo el primero del día para cada tipo
      for (const [fileType, files] of typesMap) {
        // Ordenar por hora (el nombre incluye la hora)
        files.sort((a, b) => a.getName().localeCompare(b.getName()));
        keptFiles++; // Contamos el archivo que mantenemos
        // Eliminar todos excepto el primero
        for (let i = 1; i < files.length; i++) {
          if(verboseFlag === true) {
            console.log(`Eliminando backup duplicado del día: ${files[i].getName()}`);
          }
          files[i].setTrashed(true);
          deletedFiles++;
        }
      }
    }
  }

  // Log del resumen
  console.log("\n=== RESUMEN DE LIMPIEZA DE BACKUPS ===");
  console.log(`Total de archivos procesados: ${totalFiles}`);
  console.log(`Archivos eliminados: ${deletedFiles}`);
  console.log(`Archivos mantenidos: ${keptFiles}`);
  console.log("=====================================\n");
}

function CleanAllSheets() {
  let ConControl;
  
  try {
    const ss = SpreadsheetApp.openById(allIDs['controlID']);
    
    ConControl = ss.getSheetByName("CON-Control");
    if (!ConControl) {
      throw new Error('No se pudo encontrar la hoja CON-Control');
    }

    const verboseFlag = ConControl.getRange('L2').getValue();
    const descreetFlag = ConControl.getRange('M2').getValue();

    //Block operation
    ConControl.getRange(2,1).setValue(false);

    // Lista de spreadsheets a procesar
    const spreadsheets = [
      { name: 'DubAppActive01', id: 'activeID' },
      { name: 'DubAppTotal01', id: 'totalID' },
      { name: 'DubAppLogs01', id: 'logsID' },
      { name: 'DubAppNoTrack01', id: 'noTrackID' }
    ];

    const sheetsToClean = [
      "DWO-MixAndEdit",
      "DWO-Observation",
      "DWO-Event",
      "DWO-SynopsisProduction",
      "DWO-Production",
      "DWO-SynopsisProject",
      "DWO_Files",
      "DWO_FilesLines",
      "DWO_FilesCharacter",
      "DWO_Character",
      "DWO_CharacterProduction",
      "DWO"
    ];

    // Procesar cada spreadsheet
    for (const ssInfo of spreadsheets) {
      try {
        const currentSpreadsheet = SpreadsheetApp.openById(allIDs[ssInfo.id]);
        if(verboseFlag) {
          console.log(`\n=== Procesando spreadsheet: ${ssInfo.name} ===`);
        }

        // Procesar cada hoja en el spreadsheet actual
        for (const sheet of sheetsToClean) {
          try {
            // Obtener la hoja actual
            const currentSheet = currentSpreadsheet.getSheetByName(sheet);
            
            if (!currentSheet) {
              if(verboseFlag) {
                console.warn(`Hoja ${sheet} no encontrada, continuando...`);
              }
              continue;
            }

            const beforeRows = currentSheet.getLastRow();
            
            // Verificación adicional para hojas grandes
            if (beforeRows > 5000) {
              console.log(`Precaución: Hoja grande detectada (${sheet}: ${beforeRows} filas)`);
            }
            
            // Intentar limpiar la hoja
            try {
              SheetClean(currentSpreadsheet, sheet);
              SpreadsheetApp.flush();
              
            } catch (cleanError) {
              console.error(`Error limpiando ${sheet}: ${cleanError.message}`);
              if (cleanError.message.includes("Timeout")) {
                continue;
              }
              throw cleanError;
            }
            
            // Verificar resultado después de la limpieza
            try {
              const afterRows = currentSheet.getLastRow();
              if(verboseFlag) {
                console.log(`Estado final de ${sheet}:`);
                console.log(`- Filas antes: ${beforeRows}, después: ${afterRows}`);
                console.log(`- Filas eliminadas: ${beforeRows - afterRows}`);
              }
            } catch (verifyError) {
              console.error(`Error verificando ${sheet}: ${verifyError.message}`);
            }
            
            // Pausa entre hojas grandes
            if (beforeRows > 5000) {
              Utilities.sleep(1000);
            }
            
          } catch (sheetError) {
            console.error(`Error en ${sheet}: ${sheetError.message}`);
            continue;
          }
        }
      } catch (ssError) {
        console.error(`Error en ${ssInfo.name}: ${ssError.message}`);
        continue;
      }
    }

  } catch (error) {
    console.error(`Error en CleanAllSheets: ${error.message}`);
    throw error;
  } finally {
    //Unblock operation
    if (ConControl) {
      console.log("CleanAllSheets: Proceso desbloqueado");
      ConControl.getRange(2,1).setValue(true);
    }
  }
}

function SheetClean(ss, wSheet) {
  try {
    const sheet = ss.getSheetByName(wSheet);
    if (!sheet) {
      throw new Error(`Hoja ${wSheet} no encontrada`);
    }

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    
    if (lastRow < 2) return;

    // Optimización: Primero verificar rápidamente si hay filas vacías
    const checkRange = sheet.getRange(2, 1, lastRow - 1, lastCol);
    const checkFormula = `=COUNTA(A2:${sheet.getRange(1, lastCol).getA1Notation()})`;
    const nonEmptyCheck = checkRange.getFormula();

    // Si no hay filas vacías, terminar
    if (nonEmptyCheck === checkFormula) {
      return;
    }

    // Obtener datos solo si es necesario
    const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    const rowsToKeep = data.filter((row, index) => 
      index === 0 || row.some(cell => cell !== '' && cell !== null)
    );

    // Si se encontraron filas vacías
    if (rowsToKeep.length < data.length) {
      sheet.clearContents();
      sheet.getRange(1, 1, rowsToKeep.length, lastCol).setValues(rowsToKeep);
    }

    SpreadsheetApp.flush();

  } catch (error) {
    console.error(`Error en SheetClean para ${wSheet}: ${error.message}`);
    throw error;
  }
}

function setCache(auxKey, valueToCache) {

 var cache = CacheService.getScriptCache(); // Obtener la caché del script
 var valueType = typeof valueToCache; // Obtener el tipo de la variable

 // Crear un objeto JSON que contenga el valor y su tipo
 var cacheObject = {
   value: valueToCache,
   type: valueType
 };

 // Convertir el objeto JSON a cadena y guardar en la caché
 cache.put(auxKey, JSON.stringify(cacheObject), 21600); // 6 horas
}


function MakeBackup(ss, fileID) {
 CleanFilters(ss);

 const today = new Date();
 const formattedDate = Utilities.formatDate(today, timezone, 'yyyy-MM-dd HH:mm:ss');
 const nameBackup = "Backup " + formattedDate + " " + ss.getName();
 const target = DriveApp.getFolderById(allIDs["backupID"]);
 
 // Verificar si ya existe un backup del día
 const todayStr = Utilities.formatDate(today, timezone, 'yyyy-MM-dd');
 const files = target.getFiles();
 let backupExists = false;
 
 while (files.hasNext()) {
   const file = files.next();
   const fileName = file.getName();
   // Buscar si hay un backup del mismo archivo para hoy
   if (fileName.includes(todayStr) && fileName.includes(ss.getName())) {
     backupExists = true;
     if(verboseFlag === true) {
       console.log(`Ya existe un backup de ${ss.getName()} para hoy (${fileName})`);
     }
     break;
   }
 }

 if (!backupExists) {
   const currentFile = DriveApp.getFileById(fileID);
   const newFileObj = currentFile.makeCopy(target);
   newFileObj.setName(nameBackup);
   console.log(nameBackup + " created");
 }
}

function CleanFilters(ss) {

 var currentSheets = ss.getSheets();

 for (var i = 0; i < currentSheets.length; i++) {
   var filterAux = currentSheets[i].getFilter();
   if (filterAux !== null) {
     filterAux.remove();
   }
 }
}

/**
 * Script para gestionar backups y limpieza de DubApp
 * 
 * Árbol de llamadas:
 * ├── cleanDubAppControl()
 * │   └── databaseID.getID()
 * 
 * La función cleanDubAppControl:
 * - Cambia el estado de procesamiento en CON-Control A2 (false durante proceso)
 * - Elimina filas con estado "05 Discarded" o "06 Unchanged"
 * - Mantiene filas con estado "08 Jumped off", "04 Retry" o "01 Pending"
 * - Mueve las demás filas a CON-TaskBackup
 * - Implementa bloqueo para prevenir modificaciones concurrentes
 * - Restaura el estado de procesamiento en CON-Control A2 (true al finalizar)
 */
function cleanDubAppControl() {
  // Configuración de lotes para evitar timeouts
  const BATCH_SIZE = 1000; // Tamaño de lote para procesar
  const scriptProperties = PropertiesService.getScriptProperties();
  let startRow = parseInt(scriptProperties.getProperty('CLEANUP_START_ROW') || '0');
  let isFirstRun = startRow === 0;
  
  try {
    // Obtener la hoja de control
    const ss = SpreadsheetApp.openById(allIDs['controlID']);
    const currentSheet = ss.getSheetByName("CON-TaskCurrent");
    const backupSheet = ss.getSheetByName("CON-TaskBackup");
    
    // Índice de la columna G (status)
    const STATUS_COLUMN_INDEX = 6; // 0-based index para columna G
    
    // Valores para las diferentes acciones
    const MOVE_TO_BACKUP = ["02 Incorporated", "07 Source missed key"];
    const DELETE_ONLY = ["05 Discarded", "06 Unchanged"];
    const KEEP = ["08 Jumped off", "01 Pending", "04 Retry"];
    
    // Si es la primera ejecución, registrar el estado inicial y ordenar la hoja
    if (isFirstRun) {
      console.log("=== ESTADO INICIAL ===");
      console.log(`CON-TaskCurrent: ${currentSheet.getLastRow() - 1} registros`);
      console.log(`CON-TaskBackup: ${backupSheet.getLastRow() - 1} registros`);
      console.log("\n=== PROCESAMIENTO ===");
      
      // Ordenar la hoja por la columna G para optimizar el procesamiento
      const dataRange = currentSheet.getRange(2, 1, currentSheet.getLastRow() - 1, currentSheet.getLastColumn());
      dataRange.sort({column: STATUS_COLUMN_INDEX + 1, ascending: true}); // +1 porque sort usa 1-based index
      
      // Guardar el total para seguimiento
      scriptProperties.setProperty('TOTAL_ROWS_TO_PROCESS', currentSheet.getLastRow() - 1);
    }
    
    // Obtener datos de la hoja actual a partir de la fila de inicio
    const totalRows = currentSheet.getLastRow() - 1;
    if (totalRows <= 0 || startRow >= totalRows) {
      // Proceso completado, limpiar propiedades y finalizar
      scriptProperties.deleteProperty('CLEANUP_START_ROW');
      scriptProperties.deleteProperty('TOTAL_ROWS_TO_PROCESS');
      console.log("=== LIMPIEZA COMPLETADA ===");
      return;
    }
    
    // Calcular cuántas filas procesar en este lote
    const rowsToProcess = Math.min(BATCH_SIZE, totalRows - startRow);
    const endRow = startRow + rowsToProcess;
    
    // Obtener datos del lote actual
    const dataRange = currentSheet.getRange(startRow + 2, 1, rowsToProcess, currentSheet.getLastColumn());
    const data = dataRange.getValues();
    
    // Arrays para almacenar registros según su acción
    const toBackup = [];
    const rowsToDelete = [];
    
    // Contadores para estadísticas
    let moveToBackupCount = 0;
    let deleteOnlyCount = 0;
    let keepCount = 0;
    
    // Procesar cada fila según el valor de la columna G
    for (let i = 0; i < data.length; i++) {
      const status = data[i][STATUS_COLUMN_INDEX];
      const actualRowIndex = startRow + i + 2;
      
      if (MOVE_TO_BACKUP.includes(status)) {
        // Mover a backup y luego eliminar
        toBackup.push(data[i]);
        rowsToDelete.push(actualRowIndex);
        moveToBackupCount++;
      } else if (DELETE_ONLY.includes(status)) {
        // Solo eliminar
        rowsToDelete.push(actualRowIndex);
        deleteOnlyCount++;
      } else {
        // No hacer nada (mantener en la hoja actual)
        keepCount++;
      }
    }
    
    // Registrar estadísticas del lote actual
    console.log(`Lote actual (filas ${startRow+1}-${endRow}):`);
    console.log(`- Registros a mover a backup: ${moveToBackupCount}`);
    console.log(`- Registros solo a eliminar: ${deleteOnlyCount}`);
    console.log(`- Registros a mantener: ${keepCount}`);
    
    // Mover registros al backup si hay alguno
    if (toBackup.length > 0) {
      const backupLastRow = backupSheet.getLastRow();
      backupSheet.getRange(backupLastRow + 1, 1, toBackup.length, toBackup[0].length).setValues(toBackup);
      console.log(`${toBackup.length} registros copiados a CON-TaskBackup`);
    }
    
    // Eliminar filas en orden inverso para evitar problemas con los índices
    if (rowsToDelete.length > 0) {
      rowsToDelete.sort((a, b) => b - a); // Ordenar en orden descendente
      
      // Optimización: Agrupar filas consecutivas para eliminar rangos en lugar de filas individuales
      const rangesToDelete = [];
      let rangeStart = rowsToDelete[0];
      let rangeEnd = rowsToDelete[0];
      
      for (let i = 1; i < rowsToDelete.length; i++) {
        if (rowsToDelete[i] === rangeEnd - 1) {
          // Fila consecutiva, extender el rango
          rangeEnd = rowsToDelete[i];
        } else {
          // No consecutiva, guardar el rango actual y comenzar uno nuevo
          rangesToDelete.push({start: rangeEnd, end: rangeStart});
          rangeStart = rowsToDelete[i];
          rangeEnd = rowsToDelete[i];
        }
      }
      // Añadir el último rango
      rangesToDelete.push({start: rangeEnd, end: rangeStart});
      
      // Eliminar los rangos
      for (const range of rangesToDelete) {
        const rowCount = range.end - range.start + 1;
        currentSheet.deleteRows(range.start, rowCount);
        console.log(`Eliminado rango de filas ${range.start}-${range.end} (${rowCount} filas)`);
      }
    }
    
    // Actualizar el punto de inicio para la próxima ejecución
    // Ajustar por las filas eliminadas
    const deletedCount = rowsToDelete.length;
    startRow = endRow - deletedCount;
    scriptProperties.setProperty('CLEANUP_START_ROW', startRow.toString());
    
    // Programar la siguiente ejecución si aún hay datos por procesar
    if (startRow < totalRows - deletedCount) {
      console.log(`Procesado hasta la fila ${startRow} de ${totalRows - deletedCount}. Programando siguiente lote...`);
      // Crear un trigger para continuar el proceso en 1 minuto
      deleteTriggersByFunctionNames(['cleanDubAppControl']);
      ScriptApp.newTrigger('cleanDubAppControl')
        .timeBased()
        .after(60 * 1000) // 1 minuto
        .create();
    } else {
      // Proceso completado
      scriptProperties.deleteProperty('CLEANUP_START_ROW');
      scriptProperties.deleteProperty('TOTAL_ROWS_TO_PROCESS');
      console.log("=== LIMPIEZA COMPLETADA ===");
    }
  } catch (e) {
    console.error(`Error en cleanDubAppControl: ${e.toString()}`);
    // Guardar el estado para poder continuar después
    scriptProperties.setProperty('CLEANUP_START_ROW', startRow.toString());
    // Enviar email con el error
    sendErrorEmail({
      triggerUid: 'cleanDubAppControl',
      error: e
    });
    throw e;
  }
}

// Función auxiliar para eliminar triggers por nombres de función
function deleteTriggersByFunctionNames(functionNames) {
  const triggers = ScriptApp.getProjectTriggers();
  for (let trigger of triggers) {
    if (functionNames.includes(trigger.getHandlerFunction())) {
      ScriptApp.deleteTrigger(trigger);
      if(verboseFlag === true) {
        console.log(`Trigger eliminado para la función: ${trigger.getHandlerFunction()}`);
      }
    }
  }
}

// Función para manejar errores de triggers
function sendErrorEmail(e) {
  MailApp.sendEmail({
    to: Session.getEffectiveUser().getEmail(),
    subject: `Error en proceso de Backup - ${e.triggerUid}`,
    body: `Se produjo un error en la ejecución del trigger:\n${e.error.toString()}`
  });
}