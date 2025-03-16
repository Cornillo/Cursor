/**
 * Proceso de liquidación para DubApp
 * 
 * Este script maneja el proceso de liquidación a partir de varias tablas que cambian valores.
 * Se pasa a la función el parámetro Settlement ID.
 * 
 * ESTRUCTURA Y ORDEN DE LLAMADAS:
 * ------------------------------
 * 1. Punto de entrada principal: processSettlement(settlementID)
 *    - Determina el estado del settlement y llama a la función correspondiente:
 *      - Si estado = "(03) Released: DWOSettlementBatch" → procesarCasoLiquidar()
 *      - Si estado = "(01) In preparation: DWOSettlementBatch" → procesarCasoDesliquidar()
 * 
 * 2. Funciones principales:
 *    - procesarCasoLiquidar(): Procesa la liquidación de recursos y eventos
 *    - procesarCasoDesliquidar(): Revierte la liquidación de recursos y eventos
 * 
 * 3. Funciones auxiliares:
 *    - call(): Función de prueba que llama a processSettlement con un ID específico
 *    - testSettlement(): Función para probar el proceso con un ID específico
 *    - testBatchSettlement(): Función para probar el proceso con múltiples IDs
 *    - runSettlementFromUI(): Función para ejecutar el proceso desde la interfaz de usuario
 * 
 * FLUJO DE PROCESAMIENTO:
 * ----------------------
 * 1. Obtener datos del settlement
 * 2. Obtener recursos relacionados (DWO-SettlementResource)
 * 3. Para cada recurso:
 *    - Obtener detalles relacionados (DWO-SettlementResourceDetail)
 *    - Calcular sumas y preparar actualizaciones
 *    - Para cada detalle, procesar el evento asociado (DWO-Event)
 * 4. Aplicar todas las actualizaciones en lotes
 * 
 * NOTAS IMPORTANTES:
 * -----------------
 * - El script evita procesar eventos duplicados verificando si ya han sido procesados
 * - Se implementan contadores para monitorear el número de eventos procesados
 * - Se utilizan mapas para búsquedas eficientes y procesamiento en lotes para optimizar rendimiento
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

// Variables para el envío de correos electrónicos
const DEBUG = true; // Cambiar a TRUE para enviar todos los correos a appsheet@mediaaccesscompany.com
const EMAIL_TEMPLATE_ID = "1A32AoqCjHrfSaW35ZapKtolw9JqMR0jnDhCd3c8FOnY";
const EMAIL_SENDER = "florencia.cardacci@mediaaccesscompany.com";
const EMAIL_SUBJECT = "DubApp notification / Settlement released";
const ccAux = ""; // Agregar direcciones CC si es necesario
const bccAux = ""; // Agregar direcciones BCC si es necesario

// Agregar después de las constantes globales existentes
const TEMP_PROPERTY_KEY = "pendingSettlementEmails";
const TRIGGER_FUNCTION_NAME = "procesarEnvioCorreosPendientes";
const TRIGGER_MINUTES = 1; // Ejecutar el trigger cada 1 minuto

/**
 * Función principal para procesar liquidaciones
 * @param {string} settlementID - ID de liquidación a procesar
 * @return {string} - Cadena vacía si todo está bien, mensaje de error si hay problemas
 */

/**
 * Función de prueba para ejecutar el proceso con un ID específico
 * Esta función se utiliza principalmente para pruebas y depuración
 * Llama directamente a processSettlement con el ID "8d535699"
 */
function call(){
    processSettlement("8d535699");
}

/**
 * Función principal que coordina todo el proceso de liquidación
 * 
 * Pasos:
 * 1. Busca el settlement por ID
 * 2. Determina su estado actual
 * 3. Llama a la función correspondiente según el estado:
 *    - procesarCasoLiquidar: Para settlements en estado "(03) Released: DWOSettlementBatch"
 *    - procesarCasoDesliquidar: Para settlements en estado "(01) In preparation: DWOSettlementBatch"
 * 
 * @param {string} settlementID - ID de liquidación a procesar
 * @return {string} - Cadena vacía si todo está bien, mensaje de error si hay problemas
 */
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
 * 
 * Esta función realiza el proceso de liquidación cuando el estado del settlement es "(03) Released: DWOSettlementBatch".
 * 
 * Flujo detallado:
 * 1. Obtiene todos los recursos (DWO-SettlementResource) asociados al settlement
 * 2. Para cada recurso:
 *    a. Obtiene todos los detalles (DWO-SettlementResourceDetail) asociados
 *    b. Calcula sumas de columnas D e I para detalles habilitados
 *    c. Actualiza el estado del recurso a "(02) Settled: DWOSettlementUser"
 *    d. Para cada detalle, procesa el evento asociado (DWO-Event):
 *       - Actualiza el estado del evento a "(05) Settled: DWOSettlement" o "(02) Dismissed: DWOSettlement"
 *       - Evita procesar eventos duplicados verificando si ya han sido procesados
 * 3. Aplica todas las actualizaciones en lotes para optimizar rendimiento
 * 
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
    
    // Contador para eventos procesados
    let totalEventosEncontrados = 0;
    let totalEventosProcesados = 0;
    
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
      
      // Contador para eventos procesados por recurso
      let eventosEncontradosRecurso = 0;
      let eventosProcesadosRecurso = 0;
      
      // Contador para detalles con "(01) Authorized: DWOApproval"
      let detallesAutorizados = 0;
      
      // Procesar detalles habilitados
      for (const detailRowIndex of detailRows) {
        if (detailData[detailRowIndex][13] === "(01) Enabled: DWOSettlementResourceDetail") { // Columna N (índice 13)
          // Sumar columna D solo si columna G (índice 6) es "(01) Authorized: DWOApproval"
          if (detailData[detailRowIndex][6] === "(01) Authorized: DWOApproval") {
            sumColumnD += isNaN(detailData[detailRowIndex][3]) ? 0 : Number(detailData[detailRowIndex][3]); // Columna D (índice 3)
            detallesAutorizados++;
          }
          sumColumnI += isNaN(detailData[detailRowIndex][8]) ? 0 : Number(detailData[detailRowIndex][8]); // Columna I (índice 8)
        }
      }
      
      // Redondear las sumas a dos decimales
      sumColumnD = Math.round(sumColumnD * 100) / 100;
      sumColumnI = Math.round(sumColumnI * 100) / 100;
      const totalSum = Math.round((sumColumnD + sumColumnI) * 100) / 100;
      
      console.log("Encontrados " + detailRows.length + " registros en DWO-SettlementResourceDetail para resourceID " + resourceID);
      console.log("Detalles con '(01) Authorized: DWOApproval': " + detallesAutorizados);
      console.log("Suma de columna D (solo autorizados): " + sumColumnD.toFixed(2) + ", Suma de columna I: " + sumColumnI.toFixed(2));
      
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
        
        // Incrementar contador de eventos encontrados
        totalEventosEncontrados++;
        eventosEncontradosRecurso++;
        
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
          
          // Verificar si este evento ya ha sido procesado para evitar duplicados
          const existingEventIndex = eventUpdates.findIndex(update => update.row === eventRowIndex + 1);
          if (existingEventIndex === -1) {
            // Incrementar contador de eventos procesados
            totalEventosProcesados++;
            eventosProcesadosRecurso++;
            
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
            console.log("Evento con ID " + eventID + " ya procesado, evitando duplicado");
          }
        } else {
          console.log("No se encontró el evento con ID " + eventID);
        }
      }
      
      console.log("Para resourceID " + resourceID + ": eventos encontrados: " + eventosEncontradosRecurso + ", eventos procesados: " + eventosProcesadosRecurso);
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
    
    console.log("Total de eventos encontrados: " + totalEventosEncontrados + ", eventos procesados: " + totalEventosProcesados);
    
    // En lugar de enviar correos directamente, programar el envío asincrónico
    try {
      programarEnvioCorreos(ssActive.getId(), settlementID, resourceRows, resourceData, detailRowsMap, detailData, timestamp);
      console.log(`Envío de correos programado para el settlement ID: ${settlementID}`);
    } catch (emailError) {
      console.error("Error al programar envío de correos: " + emailError.message);
      // No interrumpimos el proceso principal si falla la programación
    }
    
    return ""; // Retorno vacío indica éxito
    
  } catch (error) {
    console.error("Error en procesarCasoLiquidar: " + error.message);
    return "Error en procesarCasoLiquidar: " + error.message;
  }
}

/**
 * Procesa el caso de desliquidación
 * 
 * Esta función revierte el proceso de liquidación cuando el estado del settlement es "(01) In preparation: DWOSettlementBatch".
 * 
 * Flujo detallado:
 * 1. Obtiene todos los recursos (DWO-SettlementResource) asociados al settlement
 * 2. Para cada recurso:
 *    a. Actualiza el estado del recurso a "(01) Pending publication: DWOSettlementUser"
 *    b. Obtiene todos los detalles (DWO-SettlementResourceDetail) asociados
 *    c. Para cada detalle, procesa el evento asociado (DWO-Event):
 *       - Actualiza el estado del evento a "(01) Settlement pending: DWOSettlement"
 *       - Limpia el campo de settlementID
 *       - Evita procesar eventos duplicados verificando si ya han sido procesados
 * 3. Aplica todas las actualizaciones en lotes para optimizar rendimiento
 * 
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
    const detailUpdates = []; // Nuevo array para actualizaciones de detalles
    
    // Contador para eventos procesados
    let totalEventosEncontrados = 0;
    let totalEventosProcesados = 0;
    
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
          ["(01) Pending publication: DWOSettlementUser", userID, timestamp, null],
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
      
      // Contador para eventos procesados por recurso
      let eventosEncontradosRecurso = 0;
      let eventosProcesadosRecurso = 0;
      
      // 3. Procesar cada detalle y su evento asociado
      for (const detailRowIndex of detailRows) {
        // Verificar tiempo de ejecución
        if (Date.now() - startTime > MAX_EXECUTION_TIME) {
          return "Tiempo de ejecución excedido. El proceso se completó parcialmente.";
        }
        
        // Preparar actualización para DWO-SettlementResourceDetail (limpiar columnas D a I)
        detailUpdates.push({
          row: detailRowIndex + 1,
          values: [
            [null, null, null, null, null, null], // Valores nulos para columnas D a I
            [4, 5, 6, 7, 8, 9] // Columnas D, E, F, G, H, I (índices + 1)
          ]
        });
        
        const eventID = detailData[detailRowIndex][9]; // Columna J (índice 9)
        
        // Incrementar contador de eventos encontrados
        totalEventosEncontrados++;
        eventosEncontradosRecurso++;
        
        // Buscar el evento asociado usando el mapa
        const eventRowIndex = eventRowMap.get(eventID);
        
        if (eventRowIndex !== undefined) {
          // Verificar si este evento ya ha sido procesado para evitar duplicados
          const existingEventIndex = eventUpdates.findIndex(update => update.row === eventRowIndex + 1);
          if (existingEventIndex === -1) {
            // Incrementar contador de eventos procesados
            totalEventosProcesados++;
            eventosProcesadosRecurso++;
            
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
            console.log("Evento con ID " + eventID + " ya procesado, evitando duplicado");
          }
        } else {
          console.log("No se encontró el evento con ID " + eventID);
        }
      }
      
      console.log("Para resourceID " + resourceID + ": eventos encontrados: " + eventosEncontradosRecurso + ", eventos procesados: " + eventosProcesadosRecurso);
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
    
    // 3. Actualizar DWO-SettlementResourceDetail en lotes (nuevo)
    const detailBatchSize = 20;
    for (let i = 0; i < detailUpdates.length; i += detailBatchSize) {
      const batch = detailUpdates.slice(i, i + detailBatchSize);
      for (const update of batch) {
        for (let j = 0; j < update.values[0].length; j++) {
          sheetResourceDetail.getRange(update.row, update.values[1][j]).setValue(update.values[0][j]);
        }
      }
      // Pequeña pausa para evitar sobrecargar la API
      if (i + detailBatchSize < detailUpdates.length) {
        Utilities.sleep(50);
      }
    }
    
    // 4. Agregar filas a CON-TaskCurrent en lotes
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
    
    console.log("Total de eventos encontrados: " + totalEventosEncontrados + ", eventos procesados: " + totalEventosProcesados);
    return ""; // Retorno vacío indica éxito
    
  } catch (error) {
    console.error("Error en procesarCasoDesliquidar: " + error.message);
    return "Error en procesarCasoDesliquidar: " + error.message;
  }
}

/**
 * Función para realizar pruebas del proceso de liquidación
 * 
 * Esta función permite ejecutar el proceso con diferentes IDs de liquidación.
 * Si no se proporciona un ID, busca automáticamente un ID válido en la hoja DWO-Settlement.
 * 
 * Flujo:
 * 1. Verifica si se proporcionó un ID, si no, busca uno en la hoja
 * 2. Ejecuta processSettlement con el ID seleccionado
 * 3. Mide el tiempo de ejecución y devuelve un objeto con los resultados
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
 * Esta función permite procesar varios settlements en secuencia.
 * Útil para pruebas masivas o procesamiento por lotes.
 * 
 * Flujo:
 * 1. Verifica que se haya proporcionado un array de IDs válido
 * 2. Procesa cada ID secuencialmente llamando a testSettlement
 * 3. Recopila los resultados de cada procesamiento
 * 4. Devuelve un resumen con todos los resultados
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
 * 
 * Esta función proporciona una interfaz amigable para ejecutar el proceso de liquidación.
 * Muestra un cuadro de diálogo para que el usuario ingrese el ID de liquidación.
 * 
 * Flujo:
 * 1. Muestra un cuadro de diálogo solicitando el ID
 * 2. Valida la entrada del usuario
 * 3. Ejecuta el proceso con el ID proporcionado
 * 4. Muestra un mensaje con el resultado del proceso
 * 
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

/**
 * Envía correos electrónicos de notificación a los recursos que tienen la marca "Send alert when published: Settlement attributes"
 * 
 * @param {SpreadsheetApp.Spreadsheet} ssActive - Hoja de cálculo activa
 * @param {string} settlementID - ID de liquidación
 * @param {Array} resourceRows - Índices de las filas de recursos
 * @param {Array} resourceData - Datos de los recursos
 * @param {Map} detailRowsMap - Mapa de resourceID -> filas de detalle
 * @param {Array} detailData - Datos de los detalles
 * @param {string} currentTimestamp - Marca de tiempo actual
 * @return {number} - Número de correos enviados
 */
function enviarCorreosLiquidacion(ssActive, settlementID, resourceRows, resourceData, detailRowsMap, detailData, currentTimestamp) {
  try {
    console.log("Iniciando envío de correos de notificación");
    
    // Obtener datos del settlement
    const sheetSettlement = ssActive.getSheetByName("DWO-Settlement");
    const settlementData = sheetSettlement.getDataRange().getValues();
    
    // Buscar el settlement por ID
    let settlementRow = -1;
    let settlementMonth = "";
    let settlementYear = "";
    
    for (let i = 0; i < settlementData.length; i++) {
      if (settlementData[i][0] === settlementID) {
        settlementRow = i;
        settlementMonth = settlementData[i][4]; // Columna E (índice 4)
        settlementYear = settlementData[i][3];  // Columna D (índice 3)
        break;
      }
    }
    
    if (settlementRow === -1) {
      console.error("No se encontró el Settlement ID para enviar correos: " + settlementID);
      return 0;
    }
    
    // Abrir la hoja App-User para buscar destinatarios
    let ssNoTrack;
    try {
      ssNoTrack = SpreadsheetApp.openById(allIDs["noTrackID"]);
    } catch (e) {
      console.error("Error al abrir la hoja noTrackID: " + e.message);
      return 0;
    }
    
    const sheetAppUser = ssNoTrack.getSheetByName("App-User");
    if (!sheetAppUser) {
      console.error("No se encontró la hoja App-User");
      return 0;
    }
    
    const appUserData = sheetAppUser.getDataRange().getValues();
    console.log(`Cargados ${appUserData.length} registros de la hoja App-User`);
    
    // Crear mapa de userID -> [nombre, email]
    const userMap = new Map();
    for (let i = 1; i < appUserData.length; i++) {
      const userId = appUserData[i][0]; // Columna A (índice 0)
      if (userId) {
        // Normalizar el ID de usuario (trim y lowercase)
        const normalizedUserId = String(userId).trim().toLowerCase();
        userMap.set(normalizedUserId, [
          appUserData[i][1], // Columna B (índice 1) - Nombre
          appUserData[i][2]  // Columna C (índice 2) - Email
        ]);
        
        // Verificar si es el usuario problemático
        if (normalizedUserId.includes("transcreator05")) {
          console.log(`Usuario transcreator05 encontrado en fila ${i+1}: ID=${userId}, Nombre=${appUserData[i][1]}, Email=${appUserData[i][2]}`);
        }
      }
    }
    
    // Contador de correos enviados
    let emailsEnviados = 0;
    
    // Procesar cada recurso que tenga "Send alert when published: Settlement attributes" en columna O (índice 14)
    for (const resourceRowIndex of resourceRows) {
      const resourceID = resourceData[resourceRowIndex][0]; // Columna A (índice 0)
      const userID = resourceData[resourceRowIndex][2];     // Columna C (índice 2)
      const alertConfig = resourceData[resourceRowIndex][14]; // Columna O (índice 14)
      const isAuthorized = resourceData[resourceRowIndex][6] === "(01) Authorized: DWOApproval"; // Columna G (índice 6)
      const additionalAmount = isAuthorized ? resourceData[resourceRowIndex][3] : 0; // Columna D (índice 3)
      
      // Verificar si debe enviar alerta
      if (alertConfig && alertConfig.includes("Send alert when published: Settlement attributes")) {
        // Normalizar el ID de usuario para la búsqueda
        const normalizedUserID = String(userID).trim().toLowerCase();
        
        // Verificar si es el usuario problemático
        if (normalizedUserID.includes("transcreator05")) {
          console.log(`Intentando buscar usuario transcreator05 con ID=${userID}, normalizado=${normalizedUserID}`);
        }
        
        // Buscar datos del usuario
        if (!userMap.has(normalizedUserID)) {
          console.log(`No se encontró el usuario con ID: ${userID} (normalizado: ${normalizedUserID})`);
          
          // Intentar buscar de manera alternativa si es el usuario problemático
          if (normalizedUserID.includes("transcreator05")) {
            // Buscar por coincidencia parcial
            let encontrado = false;
            for (const [mapKey, mapValue] of userMap.entries()) {
              if (mapKey.includes("transcreator05") || mapValue[1].includes("transcreator05")) {
                console.log(`Coincidencia alternativa encontrada para transcreator05: ${mapKey} -> ${mapValue[0]}, ${mapValue[1]}`);
                encontrado = true;
                // No usamos esta coincidencia automáticamente, solo la reportamos
              }
            }
            if (!encontrado) {
              console.log("No se encontró ninguna coincidencia alternativa para transcreator05");
            }
          }
          
          continue;
        }
        
        const [nombreUsuario, emailUsuario] = userMap.get(normalizedUserID);
        
        // Determinar el destinatario según modo DEBUG
        const destinatario = DEBUG ? "appsheet@mediaaccesscompany.com" : emailUsuario;
        
        // Obtener detalles relacionados usando el mapa
        const detailRows = detailRowsMap.get(resourceID) || [];
        
        // Contar detalles deshabilitados para el log
        const detallesDeshabilitados = detailRows.filter(idx => detailData[idx][13] === "(03) Disabled: DWOSettlementResourceDetail").length;
        
        // Ordenar detalles por columnas E y F, excluyendo los deshabilitados
        const detallesOrdenados = detailRows
          .filter(idx => detailData[idx][13] !== "(03) Disabled: DWOSettlementResourceDetail") // Excluir deshabilitados en columna N (índice 13)
          .map(idx => ({
            idx,
            project: detailData[idx][4], // Columna E (índice 4)
            production: detailData[idx][5], // Columna F (índice 5)
            amount: detailData[idx][8], // Columna I (índice 8)
            detail: detailData[idx][6] // Columna G (índice 6) - Cambiado de D a G
          }))
          .sort((a, b) => {
            // Ordenar primero por project, luego por production
            if (a.project !== b.project) return a.project.localeCompare(b.project);
            return a.production.localeCompare(b.production);
          });
        
        // Calcular el total de los montos
        let totalAmount = 0;
        detallesOrdenados.forEach(detalle => {
          totalAmount += (typeof detalle.amount === 'number') ? detalle.amount : 0;
        });
        
        // Redondear el total a 2 decimales
        totalAmount = Math.round(totalAmount * 100) / 100;
        
        // Si el total es 0, no enviar correo
        if (totalAmount === 0) {
          console.log(`No se envía correo a ${nombreUsuario} (${destinatario}) porque el total es 0`);
          continue;
        }
        
        // Crear tabla HTML con los detalles
        let tablaHTML = '<table border="1" cellpadding="10" cellspacing="0" style="border-collapse: collapse; width: 100%;">';
        tablaHTML += '<tr style="background-color: #f2f2f2;">';
        tablaHTML += '<th style="padding: 10px; width: 45%;"><b>Project</b></th>';
        tablaHTML += '<th style="padding: 10px; width: 20%;"><b>Production</b></th>';
        tablaHTML += '<th style="padding: 10px; text-align: right; width: 10%;"><b>Amount</b></th>';
        tablaHTML += '<th style="padding: 10px; text-align: right; width: 25%;"><b>Detail</b></th>';
        tablaHTML += '</tr>';
        
        // Agregar filas a la tabla
        for (const detalle of detallesOrdenados) {
          tablaHTML += '<tr>';
          tablaHTML += `<td style="padding: 10px; width: 45%;">${detalle.project || ''}</td>`;
          tablaHTML += `<td style="padding: 10px; width: 20%;">${detalle.production || ''}</td>`;
          tablaHTML += `<td style="padding: 10px; text-align: right; width: 10%;">${typeof detalle.amount === 'number' ? detalle.amount.toFixed(2) : ''}</td>`;
          
          // Quitar la cadena ": Trascreation: RateTeam" del detalle
          let detailText = detalle.detail || '';
          detailText = detailText.replace(": Trascreation: RateTeam", "");
          
          tablaHTML += `<td style="padding: 10px; text-align: right; width: 25%;">${detailText}</td>`;
          tablaHTML += '</tr>';
        }
        
        // Agregar fila de total con el mismo color de fondo que la cabecera
        tablaHTML += `<tr style="background-color: #f2f2f2;">`;
        tablaHTML += `<td style="padding: 10px; width: 45%;"></td>`;
        tablaHTML += `<td style="padding: 10px; width: 20%;"><b>Total</b></td>`;
        tablaHTML += `<td style="padding: 10px; text-align: right; width: 10%;"><b>${totalAmount.toFixed(2)}</b></td>`;
        tablaHTML += `<td style="padding: 10px; text-align: right; width: 25%;"></td>`;
        tablaHTML += '</tr>';
        
        tablaHTML += '</table>';
        
        // Log de detalles incluidos y excluidos
        console.log(`Para resourceID ${resourceID}: ${detallesOrdenados.length} detalles incluidos en el correo, ${detallesDeshabilitados} detalles excluidos por estar deshabilitados`);
        
        // Formatear la fecha en formato dd/mm/yyyy
        const fechaObj = new Date(currentTimestamp);
        const dia = String(fechaObj.getDate()).padStart(2, '0');
        const mes = String(fechaObj.getMonth() + 1).padStart(2, '0');
        const anio = fechaObj.getFullYear();
        const fechaFormateada = `${dia}/${mes}/${anio}`;
        
        // Construir parámetros para el correo
        let parametros = `Title::<span style="font-family: Arial; font-size: 20pt; color: #55c3c6;">DubApp: Settlement released</span>||Header::A new monthly settlement was released.<br><br>Period: ${settlementMonth}-${settlementYear}<br><br>Settled date: ${fechaFormateada}<br><br>`;
        
        // Agregar monto adicional si está autorizado
        if (isAuthorized) {
          parametros += `Additional amount: u$s ${additionalAmount.toFixed(2)}`;
        }
        
        parametros += `||Detail::Tasks settled for the period for <b>${nombreUsuario}</b>||Footer::${tablaHTML}`;
        
        // Enviar correo
        try {
          SendEmail.AppSendEmailX(
            destinatario,
            EMAIL_SENDER,
            nombreUsuario,
            EMAIL_TEMPLATE_ID,
            null, // Sin PDF adjunto
            EMAIL_SUBJECT,
            parametros,
            ccAux,
            bccAux
          );
          
          emailsEnviados++;
          console.log(`Correo enviado a ${nombreUsuario} (${destinatario})`);
          
          // Actualizar la columna O para quitar "Send alert when published: Settlement attributes"
          const currentValue = resourceData[resourceRowIndex][14] || "";
          const newValue = currentValue
            .split(",")
            .map(item => item.trim())
            .filter(item => item !== "Send alert when published: Settlement attributes")
            .join(", ");
          
          // Actualizar en la hoja
          const sheetSettlementResource = ssActive.getSheetByName("DWO-SettlementResource");
          sheetSettlementResource.getRange(resourceRowIndex + 1, 15).setValue(newValue); // Columna O (índice 14 + 1)
          
        } catch (e) {
          console.error(`Error al enviar correo a ${nombreUsuario} (${destinatario}): ${e.message}`);
        }
      }
    }
    
    console.log(`Total de correos enviados: ${emailsEnviados}`);
    return emailsEnviados;
  } catch (error) {
    console.error("Error en enviarCorreosLiquidacion: " + error.message);
    return 0;
  }
}

/**
 * Almacena los datos necesarios para el envío de correos y configura un trigger
 * Versión optimizada para reducir el tamaño de los datos almacenados
 * 
 * @param {string} ssActiveId - ID de la hoja de cálculo activa
 * @param {string} settlementID - ID de liquidación
 * @param {Array} resourceRows - Índices de las filas de recursos
 * @param {Array} resourceData - Datos de los recursos
 * @param {Map} detailRowsMap - Mapa de resourceID -> filas de detalle
 * @param {Array} detailData - Datos de los detalles
 * @param {string} timestamp - Marca de tiempo actual
 */
function programarEnvioCorreos(ssActiveId, settlementID, resourceRows, resourceData, detailRowsMap, detailData, timestamp) {
  try {
    // Limpiar propiedades antiguas antes de agregar nuevas
    limpiarPropiedadesAntiguas();
    
    // Optimizar los datos para reducir el tamaño
    // Solo almacenamos los IDs y referencias mínimas necesarias
    const resourcesInfo = [];
    
    for (const rowIndex of resourceRows) {
      const resourceID = resourceData[rowIndex][0]; // Columna A (índice 0)
      const userID = resourceData[rowIndex][2];     // Columna C (índice 2)
      const alertConfig = resourceData[rowIndex][14]; // Columna O (índice 14)
      const isAuthorized = resourceData[rowIndex][6] === "(01) Authorized: DWOApproval"; // Columna G (índice 6)
      const additionalAmount = isAuthorized ? resourceData[rowIndex][3] : 0; // Columna D (índice 3)
      
      // Solo incluimos recursos que necesitan envío de correo
      if (alertConfig && alertConfig.includes("Send alert when published: Settlement attributes")) {
        // Obtener detalles relacionados
        const detailRows = detailRowsMap.get(resourceID) || [];
        
        // Filtrar solo los detalles habilitados y extraer información mínima
        const detailsInfo = [];
        for (const detailRowIndex of detailRows) {
          if (detailData[detailRowIndex][13] !== "(03) Disabled: DWOSettlementResourceDetail") {
            detailsInfo.push({
              project: detailData[detailRowIndex][4],    // Columna E (índice 4)
              production: detailData[detailRowIndex][5], // Columna F (índice 5)
              amount: detailData[detailRowIndex][8],     // Columna I (índice 8)
              detail: detailData[detailRowIndex][6]      // Columna G (índice 6)
            });
          }
        }
        
        // Solo agregar recursos con detalles válidos
        if (detailsInfo.length > 0) {
          resourcesInfo.push({
            resourceID: resourceID,
            userID: userID,
            isAuthorized: isAuthorized,
            additionalAmount: additionalAmount,
            rowIndex: rowIndex,
            details: detailsInfo
          });
        }
      }
    }
    
    // Si no hay recursos para enviar correos, terminamos
    if (resourcesInfo.length === 0) {
      console.log("No hay recursos que requieran envío de correo");
      return;
    }
    
    // Crear objeto con los datos mínimos necesarios
    const emailData = {
      ssActiveId: ssActiveId,
      settlementID: settlementID,
      timestamp: timestamp,
      createdAt: new Date().toISOString(),
      resources: resourcesInfo
    };
    
    // Obtener datos existentes o inicializar array vacío
    const scriptProperties = PropertiesService.getScriptProperties();
    let pendingEmails = [];
    const existingData = scriptProperties.getProperty(TEMP_PROPERTY_KEY);
    
    if (existingData) {
      try {
        pendingEmails = JSON.parse(existingData);
      } catch (e) {
        console.error("Error al parsear datos existentes: " + e.message);
        pendingEmails = [];
      }
    }
    
    // Verificar si ya existe un settlement con el mismo ID
    const existingIndex = pendingEmails.findIndex(item => item.settlementID === settlementID);
    if (existingIndex >= 0) {
      // Reemplazar los datos existentes con los nuevos
      console.log(`Actualizando datos existentes para settlement ID: ${settlementID}`);
      pendingEmails[existingIndex] = emailData;
    } else {
      // Agregar nuevos datos
      pendingEmails.push(emailData);
    }
    
    // Guardar datos actualizados
    try {
      scriptProperties.setProperty(TEMP_PROPERTY_KEY, JSON.stringify(pendingEmails));
    } catch (e) {
      console.error("Error al guardar datos en PropertiesService: " + e.message);
      // Si falla por exceder la cuota, intentamos limpiar y guardar solo el nuevo
      if (e.message.includes("exceeded the property storage quota")) {
        limpiarTodasLasPropiedades();
        try {
          // Intentar guardar solo el nuevo dato
          scriptProperties.setProperty(TEMP_PROPERTY_KEY, JSON.stringify([emailData]));
          console.log("Se limpiaron todas las propiedades y se guardó solo el nuevo dato");
        } catch (e2) {
          console.error("Error al guardar datos incluso después de limpiar: " + e2.message);
          throw new Error("No se pudo guardar los datos para el envío asincrónico de correos");
        }
      } else {
        throw e; // Re-lanzar otros errores
      }
    }
    
    // Verificar si ya existe un trigger y crear uno solo si no existe
    crearTriggerSiNoExiste();
    
    console.log(`Programado envío de ${resourcesInfo.length} correos para settlement ID: ${settlementID}`);
  } catch (error) {
    console.error("Error en programarEnvioCorreos: " + error.message);
    throw error; // Re-lanzar el error para que se maneje en la función llamante
  }
}

/**
 * Crea un trigger solo si no existe uno para la función de procesamiento
 * Evita la creación de múltiples triggers que ejecuten la misma función
 */
function crearTriggerSiNoExiste() {
  // Verificar si ya existe un trigger
  const triggers = ScriptApp.getProjectTriggers();
  let triggerExists = false;
  
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === TRIGGER_FUNCTION_NAME) {
      triggerExists = true;
      console.log("Ya existe un trigger para envío de correos");
      break;
    }
  }
  
  // Crear trigger si no existe
  if (!triggerExists) {
    console.log("Creando nuevo trigger para envío de correos");
    ScriptApp.newTrigger(TRIGGER_FUNCTION_NAME)
      .timeBased()
      .everyMinutes(TRIGGER_MINUTES)
      .create();
  }
}

/**
 * Función que se ejecuta por el trigger para procesar correos pendientes
 * Versión optimizada para trabajar con el nuevo formato de datos
 * Incluye protección contra ejecuciones simultáneas
 */
function procesarEnvioCorreosPendientes() {
  // Verificar si ya hay una ejecución en curso
  const scriptProperties = PropertiesService.getScriptProperties();
  const lockKey = "PROCESSING_LOCK";
  const currentLock = scriptProperties.getProperty(lockKey);
  
  if (currentLock) {
    // Verificar si el bloqueo es antiguo (más de 10 minutos)
    const lockTime = new Date(currentLock);
    const now = new Date();
    const minutesElapsed = (now - lockTime) / (1000 * 60);
    
    if (minutesElapsed < 10) {
      console.log(`Ya hay una ejecución en curso iniciada hace ${minutesElapsed.toFixed(2)} minutos. Saliendo.`);
      return;
    } else {
      console.log(`Encontrado bloqueo antiguo de hace ${minutesElapsed.toFixed(2)} minutos. Continuando con la ejecución.`);
    }
  }
  
  // Establecer bloqueo
  scriptProperties.setProperty(lockKey, new Date().toISOString());
  
  try {
    const startTime = Date.now();
    const MAX_EXECUTION_TIME_MS = 4 * 60 * 1000; // 4 minutos (el límite de Apps Script es 6 minutos)
    const pendingEmailsStr = scriptProperties.getProperty(TEMP_PROPERTY_KEY);
    
    if (!pendingEmailsStr) {
      console.log("No hay correos pendientes para enviar");
      // Eliminar el trigger si no hay más correos pendientes
      eliminarTriggerSiExiste();
      return;
    }
    
    let pendingEmails = JSON.parse(pendingEmailsStr);
    
    if (pendingEmails.length === 0) {
      console.log("Lista de correos pendientes vacía");
      scriptProperties.deleteProperty(TEMP_PROPERTY_KEY);
      eliminarTriggerSiExiste();
      return;
    }
    
    console.log(`Procesando ${pendingEmails.length} settlements pendientes`);
    let settlementsProcessed = 0;
    let totalEmailsSent = 0;
    let settlementsWithErrors = 0;
    
    // Procesar settlements mientras haya tiempo disponible
    while (pendingEmails.length > 0 && (Date.now() - startTime) < MAX_EXECUTION_TIME_MS) {
      // Tomar el primer elemento de la lista
      const emailData = pendingEmails.shift();
      
      // Verificar si los datos son demasiado antiguos (más de 24 horas)
      const createdAt = new Date(emailData.createdAt);
      const now = new Date();
      const hoursElapsed = (now - createdAt) / (1000 * 60 * 60);
      
      if (hoursElapsed > 24) {
        console.log(`Datos de correo demasiado antiguos (${hoursElapsed.toFixed(2)} horas), omitiendo settlement ID: ${emailData.settlementID}`);
        continue; // Pasar al siguiente settlement
      }
      
      try {
        // Abrir la hoja de cálculo activa
        const ssActive = SpreadsheetApp.openById(emailData.ssActiveId);
        
        // Procesar cada recurso y enviar correos
        let emailsEnviados = 0;
        
        // Obtener datos necesarios para el envío de correos
        const sheetSettlement = ssActive.getSheetByName("DWO-Settlement");
        const settlementData = sheetSettlement.getDataRange().getValues();
        
        // Buscar el settlement por ID
        let settlementRow = -1;
        let settlementMonth = "";
        let settlementYear = "";
        
        for (let i = 0; i < settlementData.length; i++) {
          if (settlementData[i][0] === emailData.settlementID) {
            settlementRow = i;
            settlementMonth = settlementData[i][4]; // Columna E (índice 4)
            settlementYear = settlementData[i][3];  // Columna D (índice 3)
            break;
          }
        }
        
        if (settlementRow === -1) {
          console.error("No se encontró el Settlement ID para enviar correos: " + emailData.settlementID);
          continue; // Pasar al siguiente settlement
        }
        
        // Abrir la hoja App-User para buscar destinatarios
        let ssNoTrack;
        try {
          ssNoTrack = SpreadsheetApp.openById(allIDs["noTrackID"]);
        } catch (e) {
          console.error("Error al abrir la hoja noTrackID: " + e.message);
          settlementsWithErrors++;
          continue; // Pasar al siguiente settlement
        }
        
        const sheetAppUser = ssNoTrack.getSheetByName("App-User");
        if (!sheetAppUser) {
          console.error("No se encontró la hoja App-User");
          settlementsWithErrors++;
          continue; // Pasar al siguiente settlement
        }
        
        const appUserData = sheetAppUser.getDataRange().getValues();
        
        // Crear mapa de userID -> [nombre, email]
        const userMap = new Map();
        for (let i = 1; i < appUserData.length; i++) {
          const userId = appUserData[i][0]; // Columna A (índice 0)
          if (userId) {
            // Normalizar el ID de usuario (trim y lowercase)
            const normalizedUserId = String(userId).trim().toLowerCase();
            userMap.set(normalizedUserId, [
              appUserData[i][1], // Columna B (índice 1) - Nombre
              appUserData[i][2]  // Columna C (índice 2) - Email
            ]);
          }
        }
        
        // Procesar cada recurso
        for (const resourceInfo of emailData.resources) {
          // Normalizar el ID de usuario para la búsqueda
          const normalizedUserID = String(resourceInfo.userID).trim().toLowerCase();
          
          // Buscar datos del usuario
          if (!userMap.has(normalizedUserID)) {
            console.log(`No se encontró el usuario con ID: ${resourceInfo.userID} (normalizado: ${normalizedUserID})`);
            continue; // Pasar al siguiente recurso
          }
          
          const [nombreUsuario, emailUsuario] = userMap.get(normalizedUserID);
          
          // Determinar el destinatario según modo DEBUG
          const destinatario = DEBUG ? "appsheet@mediaaccesscompany.com" : emailUsuario;
          
          // Calcular el total de los montos
          let totalAmount = 0;
          resourceInfo.details.forEach(detalle => {
            totalAmount += (typeof detalle.amount === 'number') ? detalle.amount : 0;
          });
          
          // Redondear el total a 2 decimales
          totalAmount = Math.round(totalAmount * 100) / 100;
          
          // Si el total es 0, no enviar correo
          if (totalAmount === 0) {
            console.log(`No se envía correo a ${nombreUsuario} (${destinatario}) porque el total es 0`);
            continue; // Pasar al siguiente recurso
          }
          
          // Crear tabla HTML con los detalles
          let tablaHTML = '<table border="1" cellpadding="10" cellspacing="0" style="border-collapse: collapse; width: 100%;">';
          tablaHTML += '<tr style="background-color: #f2f2f2;">';
          tablaHTML += '<th style="padding: 10px; width: 45%;"><b>Project</b></th>';
          tablaHTML += '<th style="padding: 10px; width: 20%;"><b>Production</b></th>';
          tablaHTML += '<th style="padding: 10px; text-align: right; width: 10%;"><b>Amount</b></th>';
          tablaHTML += '<th style="padding: 10px; text-align: right; width: 25%;"><b>Detail</b></th>';
          tablaHTML += '</tr>';
          
          // Ordenar detalles por project y production
          const detallesOrdenados = [...resourceInfo.details].sort((a, b) => {
            if (a.project !== b.project) return a.project.localeCompare(b.project);
            return a.production.localeCompare(b.production);
          });
          
          // Agregar filas a la tabla
          for (const detalle of detallesOrdenados) {
            tablaHTML += '<tr>';
            tablaHTML += `<td style="padding: 10px; width: 45%;">${detalle.project || ''}</td>`;
            tablaHTML += `<td style="padding: 10px; width: 20%;">${detalle.production || ''}</td>`;
            tablaHTML += `<td style="padding: 10px; text-align: right; width: 10%;">${typeof detalle.amount === 'number' ? detalle.amount.toFixed(2) : ''}</td>`;
            
            // Quitar la cadena ": Trascreation: RateTeam" del detalle
            let detailText = detalle.detail || '';
            detailText = detailText.replace(": Trascreation: RateTeam", "");
            
            tablaHTML += `<td style="padding: 10px; text-align: right; width: 25%;">${detailText}</td>`;
            tablaHTML += '</tr>';
          }
          
          // Agregar fila de total con el mismo color de fondo que la cabecera
          tablaHTML += `<tr style="background-color: #f2f2f2;">`;
          tablaHTML += `<td style="padding: 10px; width: 45%;"></td>`;
          tablaHTML += `<td style="padding: 10px; width: 20%;"><b>Total</b></td>`;
          tablaHTML += `<td style="padding: 10px; text-align: right; width: 10%;"><b>${totalAmount.toFixed(2)}</b></td>`;
          tablaHTML += `<td style="padding: 10px; text-align: right; width: 25%;"></td>`;
          tablaHTML += '</tr>';
          
          tablaHTML += '</table>';
          
          // Formatear la fecha en formato dd/mm/yyyy
          const fechaObj = new Date(emailData.timestamp);
          const dia = String(fechaObj.getDate()).padStart(2, '0');
          const mes = String(fechaObj.getMonth() + 1).padStart(2, '0');
          const anio = fechaObj.getFullYear();
          const fechaFormateada = `${dia}/${mes}/${anio}`;
          
          // Construir parámetros para el correo
          let parametros = `Title::<span style="font-family: Arial; font-size: 20pt; color: #55c3c6;">DubApp: Settlement released</span>||Header::A new monthly settlement was released.<br><br>Period: ${settlementMonth}-${settlementYear}<br><br>Settled date: ${fechaFormateada}<br><br>`;
          
          // Agregar monto adicional si está autorizado
          if (resourceInfo.isAuthorized) {
            parametros += `Additional amount: u$s ${resourceInfo.additionalAmount.toFixed(2)}`;
          }
          
          parametros += `||Detail::Works for the settled period of <b>${nombreUsuario}</b>||Footer::${tablaHTML}`;
          
          // Verificar si ya se ha enviado este correo (comprobando si la marca ya fue eliminada)
          const sheetSettlementResource = ssActive.getSheetByName("DWO-SettlementResource");
          const currentValue = sheetSettlementResource.getRange(resourceInfo.rowIndex + 1, 15).getValue() || "";
          
          if (!currentValue.includes("Send alert when published: Settlement attributes")) {
            console.log(`Correo ya enviado previamente a ${nombreUsuario} (${destinatario}), omitiendo`);
            continue; // Pasar al siguiente recurso
          }
          
          // Enviar correo
          try {
            SendEmail.AppSendEmailX(
              destinatario,
              EMAIL_SENDER,
              nombreUsuario,
              EMAIL_TEMPLATE_ID,
              null, // Sin PDF adjunto
              EMAIL_SUBJECT,
              parametros,
              ccAux,
              bccAux
            );
            
            emailsEnviados++;
            console.log(`Correo enviado a ${nombreUsuario} (${destinatario})`);
            
            // Actualizar la columna O para quitar "Send alert when published: Settlement attributes"
            const newValue = currentValue
              .split(",")
              .map(item => item.trim())
              .filter(item => item !== "Send alert when published: Settlement attributes")
              .join(", ");
            
            sheetSettlementResource.getRange(resourceInfo.rowIndex + 1, 15).setValue(newValue);
            
          } catch (e) {
            console.error(`Error al enviar correo a ${nombreUsuario} (${destinatario}): ${e.message}`);
          }
        }
        
        console.log(`Settlement ID ${emailData.settlementID}: Se enviaron ${emailsEnviados} correos.`);
        totalEmailsSent += emailsEnviados;
        settlementsProcessed++;
        
        // Verificar tiempo restante
        if ((Date.now() - startTime) > MAX_EXECUTION_TIME_MS) {
          console.log("Tiempo máximo de ejecución alcanzado, pausando procesamiento");
          break;
        }
      } catch (error) {
        console.error(`Error procesando settlement ID ${emailData.settlementID}: ${error.message}`);
        // Si hay un error con este settlement, lo volvemos a agregar al final de la cola
        // pero solo si no es demasiado antiguo
        if (hoursElapsed <= 24) {
          pendingEmails.push(emailData);
          console.log(`Settlement ID ${emailData.settlementID} vuelto a encolar para reintento posterior`);
          settlementsWithErrors++;
        }
      }
    }
    
    // Actualizar la lista de pendientes
    if (pendingEmails.length > 0) {
      try {
        scriptProperties.setProperty(TEMP_PROPERTY_KEY, JSON.stringify(pendingEmails));
        console.log(`Quedan ${pendingEmails.length} settlements pendientes para procesar en la próxima ejecución`);
      } catch (e) {
        console.error("Error al actualizar lista de pendientes: " + e.message);
        // Si falla por exceder la cuota, intentamos limpiar y guardar solo los pendientes
        if (e.message.includes("exceeded the property storage quota")) {
          limpiarPropiedadesAntiguas();
          try {
            scriptProperties.setProperty(TEMP_PROPERTY_KEY, JSON.stringify(pendingEmails));
          } catch (e2) {
            console.error("Error al guardar pendientes incluso después de limpiar: " + e2.message);
          }
        }
      }
    } else {
      scriptProperties.deleteProperty(TEMP_PROPERTY_KEY);
      eliminarTriggerSiExiste();
      console.log("Todos los settlements han sido procesados, trigger eliminado");
    }
    
    console.log(`Resumen: Procesados ${settlementsProcessed} settlements, enviados ${totalEmailsSent} correos, con ${settlementsWithErrors} errores en ${((Date.now() - startTime)/1000).toFixed(2)} segundos`);
    
  } catch (error) {
    console.error("Error en procesarEnvioCorreosPendientes: " + error.message);
  } finally {
    // Liberar el bloqueo al finalizar
    scriptProperties.deleteProperty(lockKey);
  }
}

/**
 * Limpia propiedades antiguas (más de 24 horas)
 */
function limpiarPropiedadesAntiguas() {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    const pendingEmailsStr = scriptProperties.getProperty(TEMP_PROPERTY_KEY);
    
    if (!pendingEmailsStr) {
      return; // No hay datos para limpiar
    }
    
    const pendingEmails = JSON.parse(pendingEmailsStr);
    const now = new Date();
    
    // Filtrar solo los datos que tienen menos de 24 horas
    const filteredEmails = pendingEmails.filter(emailData => {
      if (!emailData.createdAt) return false;
      
      const createdAt = new Date(emailData.createdAt);
      const hoursElapsed = (now - createdAt) / (1000 * 60 * 60);
      return hoursElapsed <= 24;
    });
    
    // Si se eliminaron elementos, actualizar la propiedad
    if (filteredEmails.length < pendingEmails.length) {
      console.log(`Se eliminaron ${pendingEmails.length - filteredEmails.length} propiedades antiguas`);
      
      if (filteredEmails.length > 0) {
        scriptProperties.setProperty(TEMP_PROPERTY_KEY, JSON.stringify(filteredEmails));
      } else {
        scriptProperties.deleteProperty(TEMP_PROPERTY_KEY);
      }
    }
  } catch (error) {
    console.error("Error al limpiar propiedades antiguas: " + error.message);
    // No lanzamos el error para que no interrumpa el flujo principal
  }
}

/**
 * Limpia todas las propiedades relacionadas con el envío de correos
 */
function limpiarTodasLasPropiedades() {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.deleteProperty(TEMP_PROPERTY_KEY);
    console.log("Se eliminaron todas las propiedades de correos pendientes");
  } catch (error) {
    console.error("Error al limpiar todas las propiedades: " + error.message);
  }
}

/**
 * Elimina el trigger si existe
 */
function eliminarTriggerSiExiste() {
  const triggers = ScriptApp.getProjectTriggers();
  
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === TRIGGER_FUNCTION_NAME) {
      ScriptApp.deleteTrigger(trigger);
      console.log("Trigger de envío de correos eliminado");
      break;
    }
  }
}

/**
 * Función para forzar el procesamiento inmediato de correos pendientes
 * Útil para pruebas o para procesar correos sin esperar al trigger
 */
function forzarEnvioCorreosPendientes() {
  procesarEnvioCorreosPendientes();
}