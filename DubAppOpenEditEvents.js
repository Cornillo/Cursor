/**
 * Árbol de llamadas a funciones:
 * 
 * llamarProcesarProjectID()
 *   └─> procesarProjectID(projectID)
 *         ├─> limpiarHojas()
 *         └─> OpenSht() [llamada múltiple para diferentes hojas]
 *               └─> LazyLoad()
 * 
 * Descripción:
 * Este script procesa datos de un proyecto específico identificado por projectID.
 * - Limpia las hojas de trabajo antes de procesar
 * - Obtiene datos de múltiples hojas (DWO, DWO-Production, DWO-Event, etc.)
 * - Procesa eventos y actualiza las hojas: Edicion, Datos, Datos2 y Actuales
 * - Maneja la lógica de eventos y sus tipos correspondientes
 */

// Variables globales
let containerSheet, containerValues, containerNDX, containerNDX2;
let ssActive;
let labelValues, userValues, labelNDX, userNDX;
let verboseFlag = true;

const sheetCache = {
  initialized: false,
  sheets: new Map()
};

function procesarProjectID(projectID) {
  // Usar variable global en lugar de leer propiedad directamente
  
  // Obtener IDs de las bases de datos
  const allIDs = databaseID.getID();
  const sheetID = allIDs["activeID"];
  const appName = 'DubAppActive' + allIDs["instalation"];
  // Open sheet
  ssActive = SpreadsheetApp.openById(sheetID);
  
  // Obtener las hojas que vamos a usar
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheetEdicion = spreadsheet.getSheetByName('Edicion');
  const sheetDatos = spreadsheet.getSheetByName('Datos');
  const sheetDatos2 = spreadsheet.getSheetByName('Datos2');
  const sheetActuales = spreadsheet.getSheetByName('Actuales');
  
  // Verificar que las hojas existan
  if (!sheetEdicion || !sheetDatos || !sheetDatos2 || !sheetActuales) {
    throw new Error('No se encontraron todas las hojas necesarias');
  }
  
  // Para DWO
  OpenSht('DWO', 2, 2, projectID, appName);
  const dwoValues = auxFilteredValues;
  
  // Para DWO-Production
  OpenSht('DWO-Production', 1, 2, projectID, appName);
  const productionValues = auxFilteredValues;
  
  // Para DWO-Event
  OpenSht('DWO-Event', 1, 76, projectID, appName);
  const eventValues = auxFilteredValues;
  
  // Para DWO-ChannelEventType (este no se filtra)
  OpenSht('DWO-ChannelEventType', 1, 0, "", appName);
  const channelEventValues = containerValues;
  const channelEventNDX = containerNDX;
  
  // Cargar App-Status
  const ssNoTrack = SpreadsheetApp.openById(allIDs['noTrackID']);
  OpenSht('App-Status', 1, 1, 'DWOEvent', 'DubAppNoTrack'+allIDs["instalation"]);
  const appStatusValues = auxFilteredValues;
  
  // Preparar matrices con dimensiones más manejables
  const crearMatrizVacia = (filas, columnas) => {
    return Array(filas).fill().map(() => Array(columnas).fill(''));
  };

  // Crear matrices con dimensiones suficientes
  const datosEdicion = crearMatrizVacia(100, 30);  // Ajustar según necesidad
  const datosDatos = crearMatrizVacia(100, 30);
  const datosDatos2 = crearMatrizVacia(100, 30);
  const datosActuales = crearMatrizVacia(100, 30);

  // Paso 1 y 2: Obtener valor de DWO
  if (dwoValues.length === 0) throw new Error('ProjectID no encontrado en DWO');
  const rowValues = dwoValues[0];
  
  // Paso 3: Obtener valor de columna G o H + I y guardar en B2
  /*let valorEdicion = rowValues[6]; // Columna G
  if (!valorEdicion) {
    valorEdicion = `${rowValues[7]} ${rowValues[8]}`; // Columna H + I
  }
  datosEdicion[1][1] = valorEdicion; // B2
  */
  // Paso 4 y 5: Procesar columna Z de DWO y buscar en ChannelEventType
  const columnaZ = rowValues[25]; // Columna Z
  const matrizEventos = columnaZ.split(',').map(item => item.trim());
  
  // Paso 6: Procesar eventos y guardar en las hojas
  let columnaActual = 2; // Empezamos en C (índice 2)
  matrizEventos.forEach(evento => {
    const eventTypeIndex = channelEventNDX.indexOf(evento);
    if (eventTypeIndex !== -1) {
      const eventTypeRow = channelEventValues[eventTypeIndex];
      const eventGroup = eventTypeRow[2]; // Columna C
      
      if (eventGroup === "1 - Asset reception: EventTypeGroup" || 
          eventGroup === "2 - Service phase: EventTypeGroup") {
        // Guardar en las matrices
        datosEdicion[3][columnaActual] = eventTypeRow[13]; // Columna N en fila 4
        datosDatos[3][columnaActual] = eventTypeRow[1];    // Columna B en fila 4
        datosDatos2[3][columnaActual] = eventTypeRow[2];   // Columna C en fila 4
        columnaActual++;
      }
    }
  });
  
  // Paso 7 y 8: Procesar DWO-Production
  let filaActual = 4; // Empezamos en fila 5 (índice 4)
  if (productionValues.length > 0) {
    productionValues.forEach(row => {
      datosEdicion[filaActual][1] = row[3]; // Columna D en columna B
      datosDatos[filaActual][1] = row[0];   // Columna A en columna B
      filaActual++;
    });
  }
  
  // Paso 9 y 10: Procesar DWO-Event
  eventValues.forEach(row => {
    const eventType = row[3];  // Columna D
    const productionID = row[1]; // Columna B
    const eventID = row[0];    // Columna A
    const eventValue = row[4];  // Columna E
    const statusCode = row[58]; // Columna BG (índice 58)
    
    // Buscar columna en fila 4 de Datos
    let columnaIndex = -1;
    for (let c = 2; c < datosDatos[3].length; c++) {
      if (datosDatos[3][c] === eventType) {
        columnaIndex = c;
        break;
      }
    }
    
    // Buscar fila en columna B de Datos
    let filaIndex = -1;
    for (let r = 4; r < datosDatos.length; r++) {
      if (datosDatos[r][1] === productionID) {
        filaIndex = r;
        break;
      }
    }
    
    if (columnaIndex !== -1 && filaIndex !== -1) {
      datosDatos[filaIndex][columnaIndex] = eventID;
      datosEdicion[filaIndex][columnaIndex] = eventValue;
      datosActuales[filaIndex][columnaIndex] = eventValue;
      
      // Procesar estado para Datos2
      if (statusCode) {
        // Quitar ": DWOEvent" del código de estado
        const cleanStatusCode = statusCode.replace(': DWOEvent', '');
        
        // Buscar en App-Status
        const appStatusMatch = appStatusValues.find(statusRow => 
          statusRow[1] === cleanStatusCode // Columna B
        );
        
        if (appStatusMatch && Number(appStatusMatch[7]) === 999) { // Convertir a número para comparar
          datosDatos2[filaIndex][columnaIndex] = 'Finalizado';
        }
      }
    }
  });

  // Encontrar dimensiones reales usadas
  const ultimaFila = Math.max(
    filaActual + 1,
    datosEdicion.findLastIndex(row => row.some(cell => cell !== '')) + 1,
    5 // Mínimo hasta fila 5
  );

  const ultimaColumna = Math.max(
    columnaActual + 1,
    datosEdicion[3].findLastIndex(cell => cell !== '') + 1,
    3 // Mínimo hasta columna C
  );

  // Escribir todas las matrices de una vez
  if (ultimaFila > 0 && ultimaColumna > 0) {
    // Escribir solo B2 en Edicion
    //sheetEdicion.getRange('B2').setValue(datosEdicion[1][1]);
    
    // Escribir desde fila 4 en Edicion
    sheetEdicion.getRange(4, 1, ultimaFila - 3, ultimaColumna).setValues(
      datosEdicion.slice(3, ultimaFila).map(row => row.slice(0, ultimaColumna))
    );
    
    // Similar para las otras hojas
    sheetDatos.getRange(4, 1, ultimaFila - 3, ultimaColumna).setValues(
      datosDatos.slice(3, ultimaFila).map(row => row.slice(0, ultimaColumna))
    );
    sheetDatos2.getRange(4, 1, ultimaFila - 3, ultimaColumna).setValues(
      datosDatos2.slice(3, ultimaFila).map(row => row.slice(0, ultimaColumna))
    );
    
    // Actuales se mantiene igual
    if (ultimaFila > 4) {
      const actualesData = datosActuales
        .slice(4, ultimaFila)
        .map(row => row.slice(2, ultimaColumna));
      sheetActuales.getRange(5, 3, ultimaFila - 4, ultimaColumna - 2).setValues(actualesData);
    }
  }

  // Al final de la función, antes del cierre
  const sheetConfig = spreadsheet.getSheetByName('Config');
  
  // Guardar selección en Config!G2
  const valorC2 = sheetEdicion.getRange('C2').getValue();
  sheetConfig.getRange('G2').setValue(valorC2);
}

function llamarProcesarProjectID() {
  // Obtener el projectID desde algún lugar, por ejemplo, un campo de entrada o una variable
  var projectID = "C8CC5541-A241-4202-9F17-32FA8FDE51AC";
  
  // Llamar a procesarProjectID con el parámetro obtenido
  procesarProjectID(projectID);
}

function OpenSht(sheetNameAux, ndxColValues, keyCol, keyValue, ssName) {
  // Cargar hoja y valores
  LazyLoad(ssName, sheetNameAux);
  auxSheet = containerSheet;
  var lastRow = auxSheet.getLastRow();
  
  if (lastRow === 1) {
    auxValues = [];
    auxNDX = [];
    auxFilteredValues = [];
    return;
  }
  auxValues = containerValues;
  
  // Crear índice si es necesario
  if (ndxColValues > 0) {
    auxNDX = auxValues.map(r => r[ndxColValues - 1].toString());
  }
  
  // Filtrar valores si se especifica
  if (keyCol !== 0 && keyValue != "") {
    auxFilteredValues = auxValues
      .filter(row => row[keyCol - 1].toString() === keyValue);
  }
}
  
function limpiarHojas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetEdicion = ss.getSheetByName('Edicion');
  
  // Obtener última fila con datos en columna B de Edición
  const ultimaFilaB = sheetEdicion.getRange('B:B')
                                 .getValues()
                                 .filter(String)
                                 .length;
  
  // Obtener última columna con datos en fila 4 de Edición
  const valoresFila4 = sheetEdicion.getRange(4, 1, 1, sheetEdicion.getMaxColumns()).getValues()[0];
  const ultimaColumna = valoresFila4.reduce((max, cell, index) => 
    cell !== '' ? index + 1 : max, 0);

  // Limpiar hojas
  if(ultimaColumna>0 && ultimaFilaB>0) {
    const hojas = ['Edicion', 'Actuales', 'Datos', 'Datos2'];
    hojas.forEach(nombreHoja => {
      const hoja = ss.getSheetByName(nombreHoja);
      if (hoja) {
        const rango = hoja.getRange(4, 2, ultimaFilaB+1, ultimaColumna);
        rango.clearContent();
      }
    });
  }
}

/**
 * Actualiza DWO-Event cuando hay cambios en Edición vs Actuales
 * - Compara valores entre Edición y Actuales
 * - Verifica que no esté Finalizado en Datos2
 * - Obtiene el ID de DWO-Event desde Datos
 * - Actualiza DWO-Event con los nuevos valores
 * - Actualiza solo las celdas modificadas en Actuales
 */
function actualizarDWOEvent() {
  const allIDs = databaseID.getID();
  const appName = 'DubAppActive' + allIDs["instalation"];

  // Obtener las hojas necesarias
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetEdicion = ss.getSheetByName('Edicion');
  const sheetActuales = ss.getSheetByName('Actuales');
  const sheetDatos = ss.getSheetByName('Datos');
  const sheetDatos2 = ss.getSheetByName('Datos2');

  if (!sheetEdicion || !sheetActuales || !sheetDatos || !sheetDatos2) {
    throw new Error('No se encontraron todas las hojas necesarias');
  }

  // Obtener los rangos de datos
  const ultimaFila = sheetEdicion.getLastRow();
  const ultimaColumna = sheetEdicion.getLastColumn();
  
  // Obtener el ancho real basado en la fila 4
  const valoresFila4 = sheetDatos.getRange(4, 3, 1, ultimaColumna - 2).getValues()[0];
  const anchoReal = valoresFila4.reduce((count, cell) => cell !== '' ? count + 1 : count, 0);

  // Obtener todos los valores de una vez
  const valoresEdicion = sheetEdicion.getRange(5, 3, ultimaFila - 4, anchoReal).getValues();
  const valoresActuales = sheetActuales.getRange(5, 3, ultimaFila - 4, anchoReal).getValues();
  const valoresDatos = sheetDatos.getRange(5, 3, ultimaFila - 4, anchoReal).getValues();
  const valoresDatos2 = sheetDatos2.getRange(5, 3, ultimaFila - 4, anchoReal).getValues();

  // Obtener acceso a DWO-Event
  const timezone = allIDs["timezone"];
  const timeformat = allIDs["timestamp_format"];
  const usuario = Session.getActiveUser().getEmail();
  const ssControl = SpreadsheetApp.openById(allIDs["controlID"]);
  const sheetTaskCurrent = ssControl.getSheetByName('CON-TaskCurrent');

  // Obtener fecha/hora una sola vez
  const fechaHora = Utilities.formatDate(
    new Date(), 
    timezone, 
    timeformat
  );

  // Obtener acceso a DWO-Event
  OpenSht('DWO-Event', 1, 0, "", appName);

  // Array para almacenar las actualizaciones
  const actualizaciones = [];
  const celdasActuales = [];

  // Comparar valores y preparar actualizaciones
  for (let i = 0; i < valoresEdicion.length; i++) {
    for (let j = 0; j < anchoReal; j++) {
      const valorEdicion = valoresEdicion[i][j];
      const valorActuales = valoresActuales[i][j];
      const valorDatos = valoresDatos[i][j];      // Valor específico de esta iteración
      const valorDatos2 = valoresDatos2[i][j];    // Valor específico de esta iteración

      // Convertir las fechas a timestamp o string para comparación
      const fechaEdicion = valorEdicion instanceof Date ? valorEdicion.getTime() : valorEdicion;
      const fechaActual = valorActuales instanceof Date ? valorActuales.getTime() : valorActuales;

      if (fechaEdicion !== fechaActual && 
          valorDatos && 
          valorDatos2 !== 'Finalizado') {
        
        // Usar auxNDX para encontrar la fila
        const eventID = valorDatos;
        const filaEvent = auxNDX.indexOf(eventID);
        
        if (filaEvent !== -1) {
          actualizaciones.push({
            fila: filaEvent + 2,
            valor: valorEdicion,
            eventID: valorDatos    // Agregar el eventID a la actualización
          });
          
          // Guardar la celda de Actuales que necesita actualizarse
          celdasActuales.push({
            fila: i + 5, // +5 porque empezamos desde fila 5
            columna: j + 3, // +3 porque empezamos desde columna C
            valor: valorEdicion
          });
        }
      }
    }
  }

  // Aplicar las actualizaciones a DWO-Event y registrar en CON-TaskCurrent
  actualizaciones.forEach(update => {
    // Actualizar DWO-Event
    const rangoEvento = containerSheet.getRange(update.fila, 1, 1, 61);
    rangoEvento.getCell(1, 5).setValue(update.valor);
    rangoEvento.getCell(1, 60).setValue(usuario);
    rangoEvento.getCell(1, 61).setValue(fechaHora);
    
    // Registrar en CON-TaskCurrent
    sheetTaskCurrent.appendRow([
      "DWO-Event",
      update.eventID,
      fechaHora,
      appName,              // Usar appName aquí
      "EDIT",
      usuario,
      "01 Pending"
    ]);
  });

  // Actualizar solo las celdas modificadas en Actuales
  celdasActuales.forEach(celda => {
    sheetActuales.getRange(celda.fila, celda.columna).setValue(celda.valor);
  });
}

/**
 * Función para actualizar la planilla con el ProjectID
 * Se llama de forma asincrónica desde AppSheet después de abrir la URL
 */
function actualizarPlanillaConProjectID(projectID) {
  try {
    Logger.log('Iniciando actualización con ProjectID: ' + projectID);
    
    // Procesar el ProjectID
    procesarProjectID(projectID);
    
    Logger.log('Actualización completada exitosamente');
    return {
      success: true,
      message: 'Planilla actualizada correctamente'
    };
    
  } catch (error) {
    Logger.log('Error en actualizarPlanillaConProjectID: ' + error.toString());
    return {
      success: false,
      message: 'Error: ' + error.message
    };
  }
}

/**
 * Obtiene o crea una planilla para un usuario específico
 * @param {string} userEmail - Email del usuario
 * @returns {Object} Objeto con el ID de la planilla y si es nueva
 */
function obtenerPlanillaUsuario(userEmail) {
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
    Logger.log('Error en obtenerPlanillaUsuario: ' + error.toString());
    throw error;
  }
}

// Trigger instalable para limpiar al abrir
function onOpenInstalable() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const hojaEdicion = spreadsheet.getSheetByName('Edicion');
  hojaEdicion.getRange('C2:C2').clearContent();

  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Proyectos DubApp')
    .addItem('Cargar mis proyectos', 'cargaMisProyectos')
    .addItem('Cargar todos los proyectos', 'cargaTodosLosProyectos')
    .addSeparator()
    .addItem('Grabar cambios', 'actualizarDWOEvent')
    .addToUi();

  limpiarHojas();
}

// Mantener el trigger simple para el menú
function onOpen() {
  const allIDs = databaseID.getID();
  const appName = 'DubAppActive' + allIDs["instalation"];
  
  onOpenInstalable();
  OpenSht('DWO-Event', 1, 0, "", appName);
}

// Nueva función cargaMisProyectos
function cargaMisProyectos() {
  const usuario = Session.getActiveUser().getEmail();
  const allIDs = databaseID.getID();
  const appName = 'DubAppActive' + allIDs["instalation"];
  // Cargar datos de DWO
  OpenSht('DWO', 1, 0, "", appName);
  const dwoValues = containerValues;
  
  // Filtrar proyectos
  const proyectosFiltrados = dwoValues.filter(row => 
    row[58] === "(01) On track: DWO" &&  // Columna BJ (índice 61)
    row[35] === usuario                   // Columna AJ (índice 35)
  );
  
  // Ordenar por ProjectID (columna A)
  const proyectosOrdenados = proyectosFiltrados.sort((a, b) => a[0].localeCompare(b[0]));
  actualizarConfig(proyectosOrdenados);
}

// Nueva función cargaTodosLosProyectos
function cargaTodosLosProyectos() {
  const allIDs = databaseID.getID();
  const appName = 'DubAppActive' + allIDs["instalation"];
  // Cargar datos de DWO
  OpenSht('DWO', 1, 0, "", appName);
  const dwoValues = containerValues;
  
  // Filtrar proyectos
  const proyectosFiltrados = dwoValues.filter(row => 
    row[58] === "(01) On track: DWO"  // Columna BJ (índice 61)
  );
  
  // Ordenar por ProjectID (columna A)
  const proyectosOrdenados = proyectosFiltrados.sort((a, b) => a[0].localeCompare(b[0]));
  actualizarConfig(proyectosOrdenados);
}

// Función común para actualizar la hoja Config
function actualizarConfig(proyectos) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetConfig = ss.getSheetByName('Config');
  
  // Limpiar datos anteriores
  sheetConfig.getRange('B3:C').clearContent();
  
  // Preparar datos
  const datosConfig = proyectos.map(row => {
    const proyectoID = row[0];  // Columna A
    const nombreCompleto = row[6] ? 
      `${proyectoID}: ${row[6]}` :  // Si G tiene contenido: A + G
      `${row[7]} / ${row[8]}`.trim(); // Si G vacío: H + I
    
    return [nombreCompleto, row[1]]; // [B, C]
  });

  // Ordenar por nombre (columna B) - Modificado aquí
  datosConfig.sort((a, b) => a[0].localeCompare(b[0], 'es', {sensitivity: 'base'}));
  
  // Escribir datos
  if (datosConfig.length > 0) {
    sheetConfig.getRange(3, 2, datosConfig.length, 2).setValues(datosConfig);
  }
}

/**
 * Trigger que detecta cambios en celda C2 de Edicion
 */
function onEditInstalable() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const editSheet = ss.getSheetByName('Edicion'); 
  const nuevoValor = editSheet.getRange('C2').getValue();
  
  // Buscar en Config
  const configSheet = ss.getSheetByName('Config');
  const configData = configSheet.getRange('B3:C' + configSheet.getLastRow()).getValues();
  const actualValor = configSheet.getRange('G2').getValue();

  if(nuevoValor===actualValor) {return};
  
  // Buscar coincidencia en columna B (índice 0) y obtener ProjectID de columna C (índice 1)
  const proyectoEncontrado = configData.find(row => row[0] === nuevoValor);
  
  if (proyectoEncontrado) {
    const projectID = proyectoEncontrado[1];
    limpiarHojas();
    procesarProjectID(projectID);
  } else {
    SpreadsheetApp.getUi().alert('Proyecto no encontrado en Config');
  }

}

// Asegurar que el trigger onEdit tenga permisos
/*function installableOnEdit(e) {
  onEdit(e);
}
*/