/**
 * Servicio web DubApp para acciones con permisos elevados
 * Se ejecuta como: appsheet@mediaaccesscompany.com
 */

// Variables globales para LazyLoad
let containerSheet, containerValues, containerNDX, containerNDX2;
let ssActive, ssTotal, ssLogs, ssNoTrack;
let labelValues, userValues, labelNDX, userNDX;
let verboseFlag = false;

const sheetCache = {
  initialized: false,
  sheets: new Map()
};

/**
 * Maneja solicitudes GET al servicio web
 */
function doGet(e) {
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      message: "Servicio web DubApp activo. Este servicio solo acepta solicitudes POST.",
      version: "1.0"
    })).setMimeType(ContentService.MimeType.JSON);
  }
  
  /**
   * Maneja solicitudes POST al servicio web
   */
  function doPost(e) {
    try {
      // Registrar el inicio de la solicitud
      console.log('Iniciando doPost');
      
      // Verificar que existan datos de postData
      if (!e || !e.postData || !e.postData.contents) {
        console.error('Error: No hay datos en la solicitud POST');
        return ContentService.createTextOutput(JSON.stringify({
          success: false,
          message: 'No se recibieron datos en la solicitud'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      
      // Parsear la solicitud con manejo de errores
      let request;
      try {
        request = JSON.parse(e.postData.contents);
      } catch (parseError) {
        console.error('Error al parsear JSON: ' + parseError.toString());
        return ContentService.createTextOutput(JSON.stringify({
          success: false,
          message: 'Error al parsear la solicitud JSON: ' + parseError.message
        })).setMimeType(ContentService.MimeType.JSON);
      }
      
      // Verificar campos obligatorios
      const action = request.action;
      const userEmail = request.userEmail;
      
      if (!action || !userEmail) {
        console.error('Error: Faltan campos obligatorios');
        return ContentService.createTextOutput(JSON.stringify({
          success: false,
          message: 'Faltan campos obligatorios: action y userEmail son requeridos'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      
      // Registrar información para diagnóstico
      console.log(`Request recibido: action=${action}, userEmail=${userEmail}`);
      
      // Verificar permisos básicos
      if (!userEmail.endsWith('@mediaaccesscompany.com')) {
        console.log('Intento de acceso no autorizado: ' + userEmail);
        return ContentService.createTextOutput(JSON.stringify({
          success: false,
          message: 'Usuario no autorizado'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      
      // Ejecutar la acción solicitada
      let result;
      try {
        switch (action) {
          case 'procesarProjectID':
            if (!request.projectID) {
              throw new Error('El campo projectID es requerido');
            }
            result = apiProcesarProjectID(request.projectID, userEmail);
            break;
          case 'cargaMisProyectos':
            result = apiCargaMisProyectos(userEmail);
            break;
          case 'cargaTodosLosProyectos':
            result = apiCargaTodosLosProyectos(userEmail);
            break;
          case 'actualizarDWOEvent':
            if (!request.cambios) {
              throw new Error('El campo cambios es requerido');
            }
            result = apiActualizarDWOEvent(request.cambios, userEmail);
            break;
          case 'obtenerPlanillaUsuario':
            result = apiObtenerPlanillaUsuario(userEmail);
            break;
          default:
            throw new Error('Acción no reconocida: ' + action);
        }
      } catch (actionError) {
        console.error('Error al ejecutar acción ' + action + ': ' + actionError.toString());
        console.error('Stack: ' + actionError.stack);
        return ContentService.createTextOutput(JSON.stringify({
          success: false,
          message: 'Error al ejecutar la acción solicitada: ' + actionError.message,
          details: actionError.toString()
        })).setMimeType(ContentService.MimeType.JSON);
      }
      
      // Registrar la operación con manejo de errores
      try {
        registrarOperacion(userEmail, action, request.projectID || '');
      } catch (logError) {
        console.error('Error al registrar operación: ' + logError.toString());
        // Continuar aunque falle el registro
      }
      
      // Asegurar que result sea un objeto
      if (result === undefined || result === null) {
        result = {};
      }
      
      console.log('doPost completado con éxito');
      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        data: result
      })).setMimeType(ContentService.MimeType.JSON);
      
    } catch (error) {
      console.error('Error general en doPost: ' + error.toString());
      console.error('Stack: ' + error.stack);
      
      // Asegurar que siempre devolvemos un JSON válido
      return ContentService.createTextOutput(JSON.stringify({
        success: false,
        message: 'Error general en el servidor: ' + error.message,
        details: error.toString()
      })).setMimeType(ContentService.MimeType.JSON);
    }
  }
  
  /**
   * Registra operaciones para auditoría
   */
  function registrarOperacion(userEmail, action, projectID) {
    try {
      const allIDs = databaseID.getID();
      const timestamp = Utilities.formatDate(
        new Date(), 
        allIDs["timezone"] || "GMT-3", 
        allIDs["timestamp_format"] || "yyyy-MM-dd HH:mm:ss"
      );
      
      // Registrar en una hoja de control
      const ssControl = SpreadsheetApp.openById(allIDs["controlID"]);
      const sheetLog = ssControl.getSheetByName('CON-APILog') || 
                      ssControl.insertSheet('CON-APILog');
      
      sheetLog.appendRow([
        timestamp,
        userEmail,
        action,
        projectID
      ]);
      
    } catch (error) {
      Logger.log('Error al registrar operación: ' + error.toString());
      // Continuar aunque falle el registro
    }
  }
  
  /**
   * API para procesar Project ID
   */
  function apiProcesarProjectID(projectID, userEmail) {
    // Obtener IDs de las bases de datos
    const allIDs = databaseID.getID();
    const appName = 'DubAppActive' + allIDs["instalation"];
    
    // Para DWO
    OpenSht('DWO', 2, 2, projectID, appName);
    const dwoValues = auxFilteredValues;
    
    // Para DWO-Production
    OpenSht('DWO-Production', 1, 2, projectID, appName);
    const productionValues = auxFilteredValues;
    
    // Para DWO-Event
    OpenSht('DWO-Event', 1, 76, projectID, appName);
    const eventValues = auxFilteredValues;
    
    // Para DWO-ChannelEventType
    OpenSht('DWO-ChannelEventType', 1, 0, "", appName);
    const channelEventValues = containerValues;
    const channelEventNDX = containerNDX;
    
    // Cargar App-Status
    OpenSht('App-Status', 1, 1, 'DWOEvent', 'DubAppNoTrack'+allIDs["instalation"]);
    const appStatusValues = auxFilteredValues;
    
    // Matrices para almacenar datos
    const datosEdicion = [];
    const datosDatos = [];
    const datosDatos2 = [];
    const datosActuales = [];
    
    // Procesamiento de datos
    // Aquí implementa la lógica para procesar los datos y llenar las matrices
    
    // Ejemplo simplificado de procesamiento
    if (dwoValues && dwoValues.length > 0) {
      // Procesar información de DWO, Production y Event para llenar las matrices
      // Este código dependerá de la estructura específica de tus datos
      
      // Encabezados para las hojas
      datosEdicion.push(['Header1', 'Header2', 'Header3']);
      datosDatos.push(['Header1', 'Header2', 'Header3']);
      datosDatos2.push(['Header1', 'Header2', 'Header3']);
      datosActuales.push(['Header1', 'Header2', 'Header3']);
      
      // Datos procesados
      for (let i = 0; i < eventValues.length; i++) {
        if (eventValues[i] && eventValues[i].length > 0) {
          // Procesar cada fila según tu lógica específica
          datosEdicion.push([/* datos procesados */]);
          datosDatos.push([/* datos procesados */]);
          datosDatos2.push([/* datos procesados */]);
          datosActuales.push([/* datos procesados */]);
        }
      }
    }
    
    // Retornar las matrices procesadas
    return {
      datosEdicion: datosEdicion,
      datosDatos: datosDatos,
      datosDatos2: datosDatos2,
      datosActuales: datosActuales
    };
  }
  
  /**
   * API para cargar mis proyectos
   */
  function apiCargaMisProyectos(userEmail) {
    const allIDs = databaseID.getID();
    const appName = 'DubAppActive' + allIDs["instalation"];
    
    // Cargar datos de DWO
    OpenSht('DWO', 1, 0, "", appName);
    const dwoValues = containerValues;
    
    // Filtrar proyectos
    const proyectosFiltrados = dwoValues.filter(row => 
      row[58] === "(01) On track: DWO" &&  
      row[35] === userEmail
    );
    
    // Ordenar por ProjectID (columna A)
    const proyectosOrdenados = proyectosFiltrados.sort((a, b) => a[0].localeCompare(b[0]));
    
    // Preparar datos
    const datosConfig = proyectosOrdenados.map(row => {
      const proyectoID = row[0];
      const nombreCompleto = row[6] ? 
        `${proyectoID}: ${row[6]}` : 
        `${row[7]} / ${row[8]}`.trim();
      
      return [nombreCompleto, row[1]];
    });
    
    // Ordenar por nombre
    datosConfig.sort((a, b) => a[0].localeCompare(b[0], 'es', {sensitivity: 'base'}));
    
    return {
      proyectos: datosConfig
    };
  }
  
  /**
   * API para cargar todos los proyectos - Versión simplificada con mejor manejo de errores
   */
  function apiCargaTodosLosProyectos(userEmail) {
    try {
      console.log('Iniciando apiCargaTodosLosProyectos para ' + userEmail);
      
      // Obtener IDs desde una función de utilidad
      try {
        var allIDs = databaseID.getID();
        console.log('IDs obtenidos correctamente');
      } catch (idError) {
        console.error('Error al obtener IDs: ' + idError.toString());
        console.error('Stack: ' + idError.stack);
        return { 
          proyectos: [],
          error: 'Error al obtener IDs: ' + idError.message
        };
      }
      
      // Validación y log de IDs críticos
      if (!allIDs || !allIDs["instalation"]) {
        console.error('Error: allIDs no tiene la propiedad instalation');
        return { 
          proyectos: [],
          error: 'Configuración incompleta: falta instalation'
        };
      }
      
      // Continuar con el procesamiento
      const appName = 'DubAppActive' + allIDs["instalation"];
      console.log('Nombre de la aplicación: ' + appName);
      
      // Cargar datos de DWO con manejo de errores detallado
      try {
        console.log('Intentando cargar DWO desde ' + appName);
        OpenSht('DWO', 1, 0, "", appName);
        console.log('Datos cargados correctamente. Filas: ' + (containerValues ? containerValues.length : 0));
      } catch (loadError) {
        console.error('Error al cargar datos de DWO: ' + loadError.toString());
        console.error('Stack: ' + loadError.stack);
        return { 
          proyectos: [],
          error: 'Error al cargar datos: ' + loadError.message
        };
      }
      
      // Verificar que containerValues exista
      if (!containerValues || !Array.isArray(containerValues)) {
        console.log('containerValues no es un array válido');
        return { 
          proyectos: [], 
          error: 'Datos no disponibles: containerValues no es válido'
        };
      }
      
      if (containerValues.length === 0) {
        console.log('No se encontraron datos en DWO');
        return { proyectos: [] };
      }
      
      // Utilizar try-catch para cada operación crítica
      try {
        // Filtrar proyectos con validación adicional
        console.log('Filtrando proyectos con estado "(01) On track: DWO"');
        
        // Asegurarnos que todos los elementos son válidos antes de filtrar
        const proyectosFiltrados = containerValues.filter(row => {
          if (!row || !Array.isArray(row)) return false;
          if (row.length <= 58) return false;
          return row[58] === "(01) On track: DWO";
        });
        
        console.log('Proyectos filtrados: ' + proyectosFiltrados.length);
        
        // Datos simplificados y validados
        const datosConfig = proyectosFiltrados.map(row => {
          // Validación exhaustiva para evitar errores
          if (!row || !Array.isArray(row)) return ['Proyecto inválido', ''];
          
          const proyectoID = (row[0] !== undefined && row[0] !== null) ? String(row[0]) : '';
          let nombre = 'Sin nombre';
          
          // Validar cada campo antes de usarlo
          if (row[6] !== undefined && row[6] !== null && row[6] !== '') {
            nombre = String(row[6]);
          } else if (row[7] !== undefined && row[7] !== null && row[8] !== undefined && row[8] !== null) {
            nombre = `${String(row[7])} / ${String(row[8])}`;
          }
          
          const idProyecto = (row[1] !== undefined && row[1] !== null) ? String(row[1]) : '';
          return [`${proyectoID}: ${nombre}`.trim(), idProyecto];
        });
        
        console.log('Datos procesados correctamente. Total: ' + datosConfig.length);
        
        return {
          proyectos: datosConfig
        };
      } catch (processError) {
        console.error('Error al procesar datos: ' + processError.toString());
        console.error('Stack: ' + processError.stack);
        return { 
          proyectos: [],
          error: 'Error al procesar datos: ' + processError.message
        };
      }
    } catch (error) {
      console.error('Error general en apiCargaTodosLosProyectos: ' + error.toString());
      console.error('Stack: ' + error.stack);
      return { 
        proyectos: [],
        error: 'Error general: ' + error.message
      };
    }
  }
  
  /**
   * API para actualizar eventos
   */
  function apiActualizarDWOEvent(cambios, userEmail) {
    if (!cambios || cambios.length === 0) {
      return { celdasActualizadas: [] };
    }
    
    const allIDs = databaseID.getID();
    const appName = 'DubAppActive' + allIDs["instalation"];
    const timezone = allIDs["timezone"];
    const timeformat = allIDs["timestamp_format"];
    
    // Obtener fecha/hora
    const fechaHora = Utilities.formatDate(
      new Date(), 
      timezone, 
      timeformat
    );
    
    // Obtener acceso a DWO-Event
    OpenSht('DWO-Event', 1, 0, "", appName);
    
    // Obtener acceso a CON-TaskCurrent
    const ssControl = SpreadsheetApp.openById(allIDs["controlID"]);
    const sheetTaskCurrent = ssControl.getSheetByName('CON-TaskCurrent');
    
    // Procesar cada cambio
    const celdasActualizadas = [];
    
    cambios.forEach(cambio => {
      // Buscar eventID en auxNDX
      const filaEvent = auxNDX.indexOf(cambio.eventID);
      
      if (filaEvent !== -1) {
        // Actualizar DWO-Event
        const rangoEvento = containerSheet.getRange(filaEvent + 2, 1, 1, 61);
        rangoEvento.getCell(1, 5).setValue(cambio.valorNuevo); // Columna E
        rangoEvento.getCell(1, 60).setValue(userEmail);       // Columna BH
        rangoEvento.getCell(1, 61).setValue(fechaHora);       // Columna BI
        
        // Registrar en CON-TaskCurrent
        sheetTaskCurrent.appendRow([
          "DWO-Event",
          cambio.eventID,
          fechaHora,
          appName,
          "EDIT",
          userEmail,
          "01 Pending"
        ]);
        
        // Guardar celda actualizada para devolver
        celdasActualizadas.push({
          fila: cambio.fila,
          columna: cambio.columna,
          valor: cambio.valorNuevo
        });
      }
    });
    
    return {
      celdasActualizadas: celdasActualizadas
    };
  }
  
  /**
   * API para obtener o crear una planilla para un usuario específico
   */
  function apiObtenerPlanillaUsuario(userEmail) {
    try {
      // Obtener acceso a DubAppNoTrack
      const allIDs = databaseID.getID();
      OpenSht('App-UserVariable', 1, 0, "", 'DubAppNoTrack01');
      
      // Buscar la fila que coincida con el usuario y EditEventSheet
      const filaEncontrada = containerValues.find(row => 
        row[0] === userEmail &&  // Columna A = email
        row[1] === 'EditEventSheet'   // Columna B = EditEventSheet
      );
      
      if (filaEncontrada && filaEncontrada[4]) { // Columna E = SheetID
        // Verificar si la planilla aún existe
        try {
          SpreadsheetApp.openById(filaEncontrada[4]);
          Logger.log('Planilla existente encontrada para usuario: ' + userEmail);
          return {
            sheetId: filaEncontrada[4],
            isNew: false
          };
        } catch (e) {
          // Si la planilla no existe, crearemos una nueva
          Logger.log('Planilla no encontrada, creando nueva');
        }
      }
      
      // Si no se encontró planilla o no existe, crear una nueva
      const templateID = SpreadsheetApp.getActiveSpreadsheet().getId();
      const template = DriveApp.getFileById(templateID);
      
      // Extraer el nombre del usuario del email
      const userName = userEmail.split('@')[0];
      const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
      const nombrePlanilla = `EditEvent ${userName}_${timestamp}`;
      
      // Obtener el folder destino desde userTempID
      const userTempFolder = DriveApp.getFolderById(allIDs["userTempID"]);
      
      // Crear copia en el folder específico
      const nuevaPlanilla = template.makeCopy(nombrePlanilla, userTempFolder);
      
      // Dar permisos al usuario
      nuevaPlanilla.setSharing(
        DriveApp.Access.ANYONE_WITH_LINK,
        DriveApp.Permission.EDIT
      );
      nuevaPlanilla.addEditor(userEmail);
      
      // Abrir la planilla para verificar acceso
      const nuevoSpreadsheet = SpreadsheetApp.openById(nuevaPlanilla.getId());
      nuevoSpreadsheet.addEditor(userEmail);
      
      // Guardar el ID en App-UserVariable
      if (filaEncontrada) {
        // Actualizar fila existente
        const filaIndex = containerValues.findIndex(row => 
          row[0] === userEmail && row[1] === 'EditEventSheet'
        );
        containerSheet.getRange(filaIndex + 1, 5).setValue(nuevaPlanilla.getId());
      } else {
        // Agregar nueva fila
        containerSheet.appendRow([userEmail, 'EditEventSheet', '', '', nuevaPlanilla.getId()]);
      }
      
      Logger.log('Nueva planilla creada para usuario: ' + userEmail);
      return {
        sheetId: nuevaPlanilla.getId(),
        isNew: true
      };
      
    } catch (error) {
      Logger.log('Error en apiObtenerPlanillaUsuario: ' + error.toString());
      throw error;
    }
  }
  
  /**
   * Función auxiliar para cargar hojas y valores
   */
  let auxSheet, auxValues, auxNDX, auxFilteredValues;
  
  function OpenSht(sheetNameAux, ndxColValues, keyCol, keyValue, ssName) {
    try {
      console.log(`OpenSht: Cargando ${sheetNameAux} desde ${ssName}`);
      
      // Cargar hoja y valores con manejo de excepciones
      try {
        LazyLoad(ssName, sheetNameAux);
        console.log('LazyLoad ejecutado correctamente');
      } catch (lazyError) {
        console.error(`Error en LazyLoad: ${lazyError.toString()}`);
        console.error(`Stack: ${lazyError.stack}`);
        throw new Error(`Error al cargar hoja ${sheetNameAux} desde ${ssName}: ${lazyError.message}`);
      }
      
      // Verificar que containerSheet se haya cargado correctamente
      if (!containerSheet) {
        console.error(`Error: containerSheet es null o undefined después de LazyLoad`);
        throw new Error(`No se pudo cargar la hoja ${sheetNameAux}`);
      }
      
      auxSheet = containerSheet;
      var lastRow = 0;
      
      try {
        lastRow = auxSheet.getLastRow();
        console.log(`Última fila de ${sheetNameAux}: ${lastRow}`);
      } catch (rowError) {
        console.error(`Error al obtener última fila: ${rowError.toString()}`);
        throw rowError;
      }
      
      if (lastRow <= 1) {
        console.log(`La hoja ${sheetNameAux} está vacía o solo tiene encabezado`);
        auxValues = [];
        auxNDX = [];
        auxFilteredValues = [];
        return;
      }
      
      // Verificar que containerValues se haya cargado correctamente
      if (!containerValues || !Array.isArray(containerValues)) {
        console.error(`Error: containerValues no es un array válido después de LazyLoad`);
        throw new Error(`No se pudieron cargar los valores de ${sheetNameAux}`);
      }
      
      auxValues = containerValues;
      console.log(`Valores cargados: ${auxValues.length} filas`);
      
      // Crear índice si es necesario
      if (ndxColValues > 0) {
        try {
          auxNDX = auxValues.map(r => {
            if (!r || r.length <= ndxColValues - 1) return '';
            return (r[ndxColValues - 1] !== null && r[ndxColValues - 1] !== undefined) ? 
                   r[ndxColValues - 1].toString() : '';
          });
          console.log(`Índice creado con ${auxNDX.length} elementos`);
        } catch (ndxError) {
          console.error(`Error al crear índice: ${ndxError.toString()}`);
          auxNDX = [];
        }
      } else {
        auxNDX = [];
      }
      
      // Filtrar valores si se especifica
      if (keyCol !== 0 && keyValue !== "") {
        try {
          auxFilteredValues = auxValues.filter(row => {
            if (!row || row.length <= keyCol - 1) return false;
            return (row[keyCol - 1] !== null && row[keyCol - 1] !== undefined) ? 
                   row[keyCol - 1].toString() === keyValue : false;
          });
          console.log(`Valores filtrados: ${auxFilteredValues.length} filas`);
        } catch (filterError) {
          console.error(`Error al filtrar valores: ${filterError.toString()}`);
          auxFilteredValues = [];
        }
      } else {
        auxFilteredValues = [];
      }
      
    } catch (error) {
      console.error(`Error general en OpenSht: ${error.toString()}`);
      console.error(`Stack: ${error.stack}`);
      // Inicializar variables para evitar errores posteriores
      auxSheet = null;
      auxValues = [];
      auxNDX = [];
      auxFilteredValues = [];
      throw error;
    }
  }
  
  /**
   * Función para verificar el acceso a las hojas de cálculo
   */
  function verificarAccesoSpreadsheets() {
    try {
      const allIDs = databaseID.getID();
      const resultados = {};
      
      // Probar cada spreadsheet
      for (const [key, id] of Object.entries(allIDs)) {
        if (key.endsWith('ID') && !key.includes('instalation')) {
          try {
            const ss = SpreadsheetApp.openById(id);
            resultados[key] = {
              acceso: true,
              nombre: ss.getName(),
              hojas: ss.getSheets().map(s => s.getName())
            };
          } catch (e) {
            resultados[key] = {
              acceso: false,
              error: e.toString()
            };
          }
        }
      }
      
      return resultados;
    } catch (e) {
      return {
        error: e.toString(),
        stack: e.stack
      };
    }
  }
  
