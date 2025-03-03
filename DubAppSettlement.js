/**
 * Proceso de liquidación para DubApp
 * 
 * Este script maneja el proceso de liquidación a partir de varias tablas que cambian valores.
 * Se pasa a la función el parámetro Settlement ID.
 * 
 * @param {string} settlementID - ID de liquidación a procesar
 * @return {string} - Cadena vacía si todo está bien, mensaje de error si hay problemas
 */

// Obtener IDs de las bases de datos
const allIDs = databaseID.getID();
const timezone = allIDs["timezone"];
const timestamp_format = allIDs["timestamp_format"];

// Constantes para optimización
const BATCH_SIZE = 50; // Tamaño de lote para operaciones en bloque
const MAX_EXECUTION_TIME = 350000; // Tiempo máximo de ejecución en ms (5.8 minutos)

/**
 * Función principal para procesar liquidaciones
 * @param {string} settlementID - ID de liquidación a procesar
 * @return {string} - Cadena vacía si todo está bien, mensaje de error si hay problemas
 */

function call(){
    processSettlement("8d535699");
  }

function processSettlement(settlementID) {
  const startTime = Date.now();
  
  try {
    console.log("Iniciando proceso de liquidación para Settlement ID: " + settlementID);
    
    // Abrir las hojas de cálculo necesarias
    const ssActive = SpreadsheetApp.openById(allIDs["activeID"]);
    const ssControl = SpreadsheetApp.openById(allIDs["controlID"]);
    
    // Obtener las hojas necesarias
    const sheetSettlement = ssActive.getSheetByName("DWO-Settlement");
    const sheetTaskCurrent = ssControl.getSheetByName("CON-TaskCurrent");
    
    if (!sheetSettlement || !sheetTaskCurrent) {
      throw new Error("No se pudieron encontrar las hojas necesarias");
    }
    
    // Buscar el Settlement ID en la columna A
    const settlementData = sheetSettlement.getDataRange().getValues();
    let settlementRow = -1;
    let userID = "";
    let currentTimestamp = "";
    let settlementStatus = "";
    
    for (let i = 0; i < settlementData.length; i++) {
      if (settlementData[i][0] === settlementID) {
        settlementRow = i;
        userID = settlementData[i][10]; // Columna K (índice 10)
        // Tomar el timestamp de la columna L (índice 11) de DWO-Settlement
        currentTimestamp = settlementData[i][11] || Utilities.formatDate(new Date(), timezone, timestamp_format);
        settlementStatus = settlementData[i][9]; // Columna J (índice 9)
        break;
      }
    }
    
    if (settlementRow === -1) {
      throw new Error("No se encontró el Settlement ID: " + settlementID);
    }
    
    console.log("Settlement encontrado en fila " + (settlementRow + 1) + ", estado: " + settlementStatus);
    console.log("Usando timestamp: " + currentTimestamp + " y userID: " + userID);
    
    // Determinar qué caso procesar
    let resultado = "";
    if (settlementStatus === "(03) Released: DWOSettlementBatch") {
      resultado = procesarCasoLiquidar(ssActive, sheetTaskCurrent, settlementID, userID, currentTimestamp, startTime);
    } else if (settlementStatus === "(01) In preparation: DWOSettlementBatch") {
      resultado = procesarCasoDesliquidar(ssActive, sheetTaskCurrent, settlementID, userID, currentTimestamp, startTime);
    } else {
      return "Estado no procesable: " + settlementStatus;
    }
    
    if (resultado !== "") {
      return resultado; // Si hay un error en alguno de los casos, devolver el mensaje
    }
    
    console.log("Proceso de liquidación completado con éxito");
    return ""; // Retorno vacío indica éxito
    
  } catch (error) {
    console.error("Error en processSettlement: " + error.message);
    return "Error en processSettlement: " + error.message;
  }
}

/**
 * Procesa el caso de liquidación
 * @param {SpreadsheetApp.Spreadsheet} ssActive - Hoja de cálculo activa
 * @param {SpreadsheetApp.Sheet} sheetTaskCurrent - Hoja de tareas actuales
 * @param {string} settlementID - ID de liquidación
 * @param {string} userID - ID de usuario
 * @param {string} timestamp - Marca de tiempo actual
 * @param {number} startTime - Tiempo de inicio de la ejecución
 * @return {string} - Cadena vacía si todo está bien, mensaje de error si hay problemas
 */
function procesarCasoLiquidar(ssActive, sheetTaskCurrent, settlementID, userID, timestamp, startTime) {
  try {
    console.log("Iniciando CASO LIQUIDAR");
    
    // 1. Obtener DWO-SettlementResource donde columna B es el Settlement ID
    const sheetSettlementResource = ssActive.getSheetByName("DWO-SettlementResource");
    
    // Cargar datos de manera más eficiente
    const resourceData = sheetSettlementResource.getDataRange().getValues();
    
    // Crear índices para búsquedas más eficientes
    // Filtrar directamente los recursos que coinciden con el settlementID
    const resourceRows = resourceData
      .map((row, index) => ({ row, index }))
      .filter((item, index) => index > 0 && item.row[1] === settlementID) // Columna B (índice 1)
      .map(item => item.index);
    
    console.log("Encontrados " + resourceRows.length + " registros en DWO-SettlementResource");
    
    if (resourceRows.length === 0) {
      return "No se encontraron registros en DWO-SettlementResource para el Settlement ID: " + settlementID;
    }
    
    // Cargar todas las hojas necesarias una sola vez
    const sheetResourceDetail = ssActive.getSheetByName("DWO-SettlementResourceDetail");
    const detailData = sheetResourceDetail.getDataRange().getValues();
    const sheetEvent = ssActive.getSheetByName("DWO-Event");
    const eventData = sheetEvent.getDataRange().getValues();
    
    // Crear mapas de índices para búsquedas más eficientes
    // Mapa de resourceID -> filas de detalle
    const detailRowsMap = new Map();
    // Mapa de eventID -> fila de evento
    const eventRowMap = new Map();
    
    // Construir mapa de eventos (solo una vez)
    for (let i = 1; i < eventData.length; i++) {
      const eventID = eventData[i][0]; // Columna A (índice 0)
      if (eventID) {
        eventRowMap.set(eventID, i);
      }
    }
    
    // Construir mapa de detalles (solo una vez)
    for (let i = 1; i < detailData.length; i++) {
      const resourceID = detailData[i][1]; // Columna B (índice 1)
      if (resourceID) {
        if (!detailRowsMap.has(resourceID)) {
          detailRowsMap.set(resourceID, []);
        }
        detailRowsMap.get(resourceID).push(i);
      }
    }
    
    // Preparar arrays para actualizaciones en lote
    const resourceUpdates = [];
    const taskCurrentRows = [];
    const eventUpdates = [];
    
    // Procesar cada registro de DWO-SettlementResource
    for (const resourceRowIndex of resourceRows) {
      // Verificar tiempo de ejecución
      if (Date.now() - startTime > MAX_EXECUTION_TIME) {
        return "Tiempo de ejecución excedido. El proceso se completó parcialmente.";
      }
      
      const resourceID = resourceData[resourceRowIndex][0]; // Columna A (índice 0)
      
      // Obtener detalles relacionados usando el mapa
      const detailRows = detailRowsMap.get(resourceID) || [];
      
      // Calcular sumas de columnas D e I
      let sumColumnD = 0;
      let sumColumnI = 0;
      
      // Procesar detalles habilitados
      for (const detailRowIndex of detailRows) {
        if (detailData[detailRowIndex][13] === "(01) Enabled: DWOSettlementResourceDetail") { // Columna N (índice 13)
          sumColumnD += isNaN(detailData[detailRowIndex][3]) ? 0 : Number(detailData[detailRowIndex][3]); // Columna D (índice 3)
          sumColumnI += isNaN(detailData[detailRowIndex][8]) ? 0 : Number(detailData[detailRowIndex][8]); // Columna I (índice 8)
        }
      }
      
      // Redondear las sumas a dos decimales
      sumColumnD = Math.round(sumColumnD * 100) / 100;
      sumColumnI = Math.round(sumColumnI * 100) / 100;
      const totalSum = Math.round((sumColumnD + sumColumnI) * 100) / 100;
      
      console.log("Encontrados " + detailRows.length + " registros en DWO-SettlementResourceDetail para resourceID " + resourceID);
      console.log("Suma de columna D: " + sumColumnD.toFixed(2) + ", Suma de columna I: " + sumColumnI.toFixed(2));
      
      // 1. Preparar actualización para DWO-SettlementResource
      resourceUpdates.push({
        row: resourceRowIndex + 1,
        values: [
          ["(02) Settled: DWOSettlementUser", userID, timestamp, totalSum],
          [17, 18, 19, 10] // Columnas Q, R, S, J
        ]
      });
      
      // 2. Preparar fila para CON-TaskCurrent
      taskCurrentRows.push([
        "DWO-SettlementResource",
        resourceID,
        timestamp,
        "DubAppActive01",
        "EDIT",
        userID,
        "01 Pending"
      ]);
      
      // 3. Procesar cada detalle y su evento asociado
      for (const detailRowIndex of detailRows) {
        // Verificar tiempo de ejecución
        if (Date.now() - startTime > MAX_EXECUTION_TIME) {
          return "Tiempo de ejecución excedido. El proceso se completó parcialmente.";
        }
        
        const eventID = detailData[detailRowIndex][9]; // Columna J (índice 9)
        const detailStatus = detailData[detailRowIndex][13]; // Columna N (índice 13)
        
        // Buscar el evento asociado usando el mapa
        const eventRowIndex = eventRowMap.get(eventID);
        
        if (eventRowIndex !== undefined) {
          // Determinar el valor para la columna BN según el estado
          let bnValue = "";
          if (detailStatus === "(03) Disabled: DWOSettlementResourceDetail") {
            bnValue = "(02) Dismissed: DWOSettlement";
          } else {
            bnValue = "(05) Settled: DWOSettlement";
          }
          
          // Preparar actualización para DWO-Event
          eventUpdates.push({
            row: eventRowIndex + 1,
            values: [
              [bnValue, settlementID, userID, timestamp],
              [66, 67, 60, 61] // Columnas BN, BO, BH, BI
            ]
          });
          
          // Preparar fila para CON-TaskCurrent
          taskCurrentRows.push([
            "DWO-Event",
            eventID,
            timestamp,
            "DubAppActive01",
            "EDIT",
            userID,
            "01 Pending"
          ]);
        } else {
          console.log("No se encontró el evento con ID " + eventID);
        }
      }
    }
    
    // Aplicar actualizaciones en lote
    
    // 1. Actualizar DWO-SettlementResource en lotes
    const resourceBatchSize = 20;
    for (let i = 0; i < resourceUpdates.length; i += resourceBatchSize) {
      const batch = resourceUpdates.slice(i, i + resourceBatchSize);
      for (const update of batch) {
        for (let j = 0; j < update.values[0].length; j++) {
          sheetSettlementResource.getRange(update.row, update.values[1][j]).setValue(update.values[0][j]);
        }
      }
      // Pequeña pausa para evitar sobrecargar la API
      if (i + resourceBatchSize < resourceUpdates.length) {
        Utilities.sleep(50);
      }
    }
    
    // 2. Actualizar DWO-Event en lotes
    const eventBatchSize = 20;
    for (let i = 0; i < eventUpdates.length; i += eventBatchSize) {
      const batch = eventUpdates.slice(i, i + eventBatchSize);
      for (const update of batch) {
        for (let j = 0; j < update.values[0].length; j++) {
          sheetEvent.getRange(update.row, update.values[1][j]).setValue(update.values[0][j]);
        }
      }
      // Pequeña pausa para evitar sobrecargar la API
      if (i + eventBatchSize < eventUpdates.length) {
        Utilities.sleep(50);
      }
    }
    
    // 3. Agregar filas a CON-TaskCurrent en lotes
    for (let i = 0; i < taskCurrentRows.length; i += BATCH_SIZE) {
      const batch = taskCurrentRows.slice(i, i + BATCH_SIZE);
      if (batch.length > 0) {
        sheetTaskCurrent.getRange(sheetTaskCurrent.getLastRow() + 1, 1, batch.length, batch[0].length)
          .setValues(batch);
      }
      // Pequeña pausa para evitar sobrecargar la API
      if (i + BATCH_SIZE < taskCurrentRows.length) {
        Utilities.sleep(50);
      }
    }
    
    return ""; // Retorno vacío indica éxito
    
  } catch (error) {
    console.error("Error en procesarCasoLiquidar: " + error.message);
    return "Error en procesarCasoLiquidar: " + error.message;
  }
}

/**
 * Procesa el caso de desliquidación
 * @param {SpreadsheetApp.Spreadsheet} ssActive - Hoja de cálculo activa
 * @param {SpreadsheetApp.Sheet} sheetTaskCurrent - Hoja de tareas actuales
 * @param {string} settlementID - ID de liquidación
 * @param {string} userID - ID de usuario
 * @param {string} timestamp - Marca de tiempo actual
 * @param {number} startTime - Tiempo de inicio de la ejecución
 * @return {string} - Cadena vacía si todo está bien, mensaje de error si hay problemas
 */
function procesarCasoDesliquidar(ssActive, sheetTaskCurrent, settlementID, userID, timestamp, startTime) {
  try {
    console.log("Iniciando CASO DESLIQUIDAR");
    
    // 1. Obtener DWO-SettlementResource donde columna B es el Settlement ID
    const sheetSettlementResource = ssActive.getSheetByName("DWO-SettlementResource");
    
    // Cargar datos de manera más eficiente
    const resourceData = sheetSettlementResource.getDataRange().getValues();
    
    // Filtrar directamente los recursos que coinciden con el settlementID
    const resourceRows = resourceData
      .map((row, index) => ({ row, index }))
      .filter((item, index) => index > 0 && item.row[1] === settlementID) // Columna B (índice 1)
      .map(item => item.index);
    
    console.log("Encontrados " + resourceRows.length + " registros en DWO-SettlementResource");
    
    if (resourceRows.length === 0) {
      return "No se encontraron registros en DWO-SettlementResource para el Settlement ID: " + settlementID;
    }
    
    // Cargar todas las hojas necesarias una sola vez
    const sheetResourceDetail = ssActive.getSheetByName("DWO-SettlementResourceDetail");
    const detailData = sheetResourceDetail.getDataRange().getValues();
    const sheetEvent = ssActive.getSheetByName("DWO-Event");
    const eventData = sheetEvent.getDataRange().getValues();
    
    // Crear mapas de índices para búsquedas más eficientes
    // Mapa de resourceID -> filas de detalle
    const detailRowsMap = new Map();
    // Mapa de eventID -> fila de evento
    const eventRowMap = new Map();
    
    // Construir mapa de eventos (solo una vez)
    for (let i = 1; i < eventData.length; i++) {
      const eventID = eventData[i][0]; // Columna A (índice 0)
      if (eventID) {
        eventRowMap.set(eventID, i);
      }
    }
    
    // Construir mapa de detalles (solo una vez)
    for (let i = 1; i < detailData.length; i++) {
      const resourceID = detailData[i][1]; // Columna B (índice 1)
      if (resourceID) {
        if (!detailRowsMap.has(resourceID)) {
          detailRowsMap.set(resourceID, []);
        }
        detailRowsMap.get(resourceID).push(i);
      }
    }
    
    // Preparar arrays para actualizaciones en lote
    const resourceUpdates = [];
    const taskCurrentRows = [];
    const eventUpdates = [];
    
    // Procesar cada registro de DWO-SettlementResource
    for (const resourceRowIndex of resourceRows) {
      // Verificar tiempo de ejecución
      if (Date.now() - startTime > MAX_EXECUTION_TIME) {
        return "Tiempo de ejecución excedido. El proceso se completó parcialmente.";
      }
      
      const resourceID = resourceData[resourceRowIndex][0]; // Columna A (índice 0)
      
      // 1. Preparar actualización para DWO-SettlementResource
      resourceUpdates.push({
        row: resourceRowIndex + 1,
        values: [
          ["(01) Pending publication: DWOSettlementUser", userID, timestamp, 0.00],
          [17, 18, 19, 10] // Columnas Q, R, S, J
        ]
      });
      
      // 2. Preparar fila para CON-TaskCurrent
      taskCurrentRows.push([
        "DWO-SettlementResource",
        resourceID,
        timestamp,
        "DubAppActive01",
        "EDIT",
        userID,
        "01 Pending"
      ]);
      
      // Obtener detalles relacionados usando el mapa
      const detailRows = detailRowsMap.get(resourceID) || [];
      
      console.log("Encontrados " + detailRows.length + " registros en DWO-SettlementResourceDetail para resourceID " + resourceID);
      
      // 3. Procesar cada detalle y su evento asociado
      for (const detailRowIndex of detailRows) {
        // Verificar tiempo de ejecución
        if (Date.now() - startTime > MAX_EXECUTION_TIME) {
          return "Tiempo de ejecución excedido. El proceso se completó parcialmente.";
        }
        
        const eventID = detailData[detailRowIndex][9]; // Columna J (índice 9)
        
        // Buscar el evento asociado usando el mapa
        const eventRowIndex = eventRowMap.get(eventID);
        
        if (eventRowIndex !== undefined) {
          // Preparar actualización para DWO-Event
          eventUpdates.push({
            row: eventRowIndex + 1,
            values: [
              ["(01) Settlement pending: DWOSettlement", "", userID, timestamp],
              [66, 67, 60, 61] // Columnas BN, BO, BH, BI
            ]
          });
          
          // Preparar fila para CON-TaskCurrent
          taskCurrentRows.push([
            "DWO-Event",
            eventID,
            timestamp,
            "DubAppActive01",
            "EDIT",
            userID,
            "01 Pending"
          ]);
        } else {
          console.log("No se encontró el evento con ID " + eventID);
        }
      }
    }
    
    // Aplicar actualizaciones en lote
    
    // 1. Actualizar DWO-SettlementResource en lotes
    const resourceBatchSize = 20;
    for (let i = 0; i < resourceUpdates.length; i += resourceBatchSize) {
      const batch = resourceUpdates.slice(i, i + resourceBatchSize);
      for (const update of batch) {
        for (let j = 0; j < update.values[0].length; j++) {
          sheetSettlementResource.getRange(update.row, update.values[1][j]).setValue(update.values[0][j]);
        }
      }
      // Pequeña pausa para evitar sobrecargar la API
      if (i + resourceBatchSize < resourceUpdates.length) {
        Utilities.sleep(50);
      }
    }
    
    // 2. Actualizar DWO-Event en lotes
    const eventBatchSize = 20;
    for (let i = 0; i < eventUpdates.length; i += eventBatchSize) {
      const batch = eventUpdates.slice(i, i + eventBatchSize);
      for (const update of batch) {
        for (let j = 0; j < update.values[0].length; j++) {
          sheetEvent.getRange(update.row, update.values[1][j]).setValue(update.values[0][j]);
        }
      }
      // Pequeña pausa para evitar sobrecargar la API
      if (i + eventBatchSize < eventUpdates.length) {
        Utilities.sleep(50);
      }
    }
    
    // 3. Agregar filas a CON-TaskCurrent en lotes
    for (let i = 0; i < taskCurrentRows.length; i += BATCH_SIZE) {
      const batch = taskCurrentRows.slice(i, i + BATCH_SIZE);
      if (batch.length > 0) {
        sheetTaskCurrent.getRange(sheetTaskCurrent.getLastRow() + 1, 1, batch.length, batch[0].length)
          .setValues(batch);
      }
      // Pequeña pausa para evitar sobrecargar la API
      if (i + BATCH_SIZE < taskCurrentRows.length) {
        Utilities.sleep(50);
      }
    }
    
    return ""; // Retorno vacío indica éxito
    
  } catch (error) {
    console.error("Error en procesarCasoDesliquidar: " + error.message);
    return "Error en procesarCasoDesliquidar: " + error.message;
  }
}

/**
 * Función para realizar pruebas del proceso de liquidación
 * Permite ejecutar el proceso con diferentes IDs de liquidación
 * 
 * @param {string} settlementID - ID de liquidación a procesar (opcional)
 * @return {object} - Objeto con el resultado de la prueba
 */
function testSettlement(settlementID) {
  try {
    console.log("Iniciando prueba de liquidación");
    
    // Si no se proporciona un ID, usar un ID de prueba
    if (!settlementID) {
      // Buscar un ID válido en la hoja DWO-Settlement
      const ssActive = SpreadsheetApp.openById(allIDs["activeID"]);
      const sheetSettlement = ssActive.getSheetByName("DWO-Settlement");
      
      if (!sheetSettlement) {
        return { 
          success: false, 
          message: "No se pudo encontrar la hoja DWO-Settlement" 
        };
      }
      
      const settlementData = sheetSettlement.getDataRange().getValues();
      // Buscar el primer ID disponible (omitir la fila de encabezado)
      for (let i = 1; i < settlementData.length; i++) {
        if (settlementData[i][0]) {
          settlementID = settlementData[i][0];
          console.log("Usando ID de prueba: " + settlementID);
          break;
        }
      }
      
      if (!settlementID) {
        return { 
          success: false, 
          message: "No se encontró ningún ID de liquidación válido para pruebas" 
        };
      }
    }
    
    // Ejecutar el proceso de liquidación
    const startTime = new Date();
    const resultado = processSettlement(settlementID);
    const endTime = new Date();
    const executionTime = (endTime - startTime) / 1000; // Tiempo en segundos
    
    // Preparar el resultado
    if (resultado === "") {
      return {
        success: true,
        message: "Proceso completado con éxito para Settlement ID: " + settlementID,
        executionTime: executionTime,
        settlementID: settlementID
      };
    } else {
      return {
        success: false,
        message: resultado,
        executionTime: executionTime,
        settlementID: settlementID
      };
    }
    
  } catch (error) {
    console.error("Error en testSettlement: " + error.message);
    return {
      success: false,
      message: "Error en testSettlement: " + error.message,
      settlementID: settlementID
    };
  }
}

/**
 * Función para ejecutar pruebas en lote con múltiples IDs de liquidación
 * 
 * @param {string[]} settlementIDs - Array de IDs de liquidación a procesar
 * @return {object[]} - Array de resultados de las pruebas
 */
function testBatchSettlement(settlementIDs) {
  const results = [];
  
  if (!Array.isArray(settlementIDs) || settlementIDs.length === 0) {
    console.log("No se proporcionaron IDs para procesar en lote");
    return [{ 
      success: false, 
      message: "No se proporcionaron IDs para procesar en lote" 
    }];
  }
  
  console.log("Iniciando procesamiento en lote de " + settlementIDs.length + " liquidaciones");
  
  for (const id of settlementIDs) {
    console.log("Procesando Settlement ID: " + id);
    const result = testSettlement(id);
    results.push(result);
    
    // Pequeña pausa entre ejecuciones para evitar sobrecargar el sistema
    Utilities.sleep(1000);
  }
  
  // Resumen de resultados
  const successCount = results.filter(r => r.success).length;
  console.log("Procesamiento en lote completado. Éxitos: " + successCount + "/" + settlementIDs.length);
  
  return results;
}

/**
 * Función para ejecutar el proceso desde la interfaz de usuario
 * Esta función puede ser llamada desde un botón o menú en la hoja de cálculo
 */
function runSettlementFromUI() {
  const ui = SpreadsheetApp.getUi();
  
  // Solicitar el ID de liquidación
  const response = ui.prompt(
    'Procesar Liquidación',
    'Ingrese el ID de liquidación:',
    ui.ButtonSet.OK_CANCEL
  );
  
  // Verificar si el usuario hizo clic en "Cancelar"
  if (response.getSelectedButton() === ui.Button.CANCEL) {
    ui.alert('Operación cancelada por el usuario.');
    return;
  }
  
  // Obtener el ID ingresado
  const settlementID = response.getResponseText().trim();
  
  if (!settlementID) {
    ui.alert('Debe ingresar un ID de liquidación válido.');
    return;
  }
  
  // Ejecutar el proceso
  const result = testSettlement(settlementID);
  
  // Mostrar el resultado
  if (result.success) {
    ui.alert(
      'Proceso completado',
      'El proceso de liquidación se completó con éxito.\n\nID: ' + result.settlementID + '\nTiempo de ejecución: ' + result.executionTime + ' segundos',
      ui.ButtonSet.OK
    );
  } else {
    ui.alert(
      'Error en el proceso',
      'Ocurrió un error durante el proceso de liquidación:\n\n' + result.message,
      ui.ButtonSet.OK
    );
  }
}
