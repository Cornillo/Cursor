/**
 * Script para probar las funciones del servicio web sin la intermediación del servicio POST
 */

// Variables globales para LazyLoad - Deben ser las mismas que en DubAppDatesWebService.js
let containerSheet, containerValues, containerNDX, containerNDX2;
let ssActive, ssTotal, ssLogs, ssNoTrack;
let labelValues, userValues, labelNDX, userNDX;
let verboseFlag = false;

const sheetCache = {
  initialized: false,
  sheets: new Map()
};

// Para acceder a auxSheet, auxValues, etc.
let auxSheet, auxValues, auxNDX, auxFilteredValues;

/**
 * Función principal para ejecutar pruebas
 */
function ejecutarPruebas() {
  console.log("=== INICIANDO PRUEBAS ===");
  
  try {
    // 1. Probar databaseID
    console.log("\n== Prueba 1: databaseID.getID() ==");
    const allIDs = databaseID.getID();
    console.log("IDs obtenidos correctamente:", JSON.stringify(allIDs));
    
    // 2. Probar LazyLoad
    console.log("\n== Prueba 2: LazyLoad ==");
    try {
      LazyLoad("DubAppActive" + allIDs["instalation"], "DWO");
      console.log("LazyLoad ejecutado correctamente");
      console.log("containerSheet existe:", !!containerSheet);
      console.log("containerValues.length:", containerValues ? containerValues.length : "undefined");
    } catch (error) {
      console.error("Error en LazyLoad:", error.toString());
      console.error("Stack:", error.stack);
    }
    
    // 3. Probar OpenSht
    console.log("\n== Prueba 3: OpenSht ==");
    try {
      OpenSht("DWO", 1, 0, "", "DubAppActive" + allIDs["instalation"]);
      console.log("OpenSht ejecutado correctamente");
      console.log("auxSheet existe:", !!auxSheet);
      console.log("auxValues.length:", auxValues ? auxValues.length : "undefined");
      console.log("auxNDX.length:", auxNDX ? auxNDX.length : "undefined");
    } catch (error) {
      console.error("Error en OpenSht:", error.toString());
      console.error("Stack:", error.stack);
    }
    
    // 4. Probar apiCargaTodosLosProyectos
    console.log("\n== Prueba 4: apiCargaTodosLosProyectos ==");
    const emailPrueba = "usuario@mediaaccesscompany.com";
    try {
      const resultado = apiCargaTodosLosProyectos(emailPrueba);
      console.log("apiCargaTodosLosProyectos ejecutado correctamente");
      console.log("Número de proyectos:", resultado.proyectos ? resultado.proyectos.length : 0);
      console.log("Error:", resultado.error || "Ninguno");
    } catch (error) {
      console.error("Error en apiCargaTodosLosProyectos:", error.toString());
      console.error("Stack:", error.stack);
    }
    
  } catch (error) {
    console.error("Error general en las pruebas:", error.toString());
    console.error("Stack:", error.stack);
  }
  
  console.log("\n=== PRUEBAS FINALIZADAS ===");
}

/**
 * Función para probar cada función individualmente
 */
function probarDatabaseID() {
  console.log("Probando databaseID.getID()");
  const allIDs = databaseID.getID();
  console.log(JSON.stringify(allIDs, null, 2));
}

function probarLazyLoad() {
  console.log("Probando LazyLoad");
  const allIDs = databaseID.getID();
  LazyLoad("DubAppActive" + allIDs["instalation"], "DWO");
  console.log("LazyLoad ejecutado correctamente");
  console.log("containerValues.length:", containerValues.length);
}

function probarOpenSht() {
  console.log("Probando OpenSht");
  const allIDs = databaseID.getID();
  OpenSht("DWO", 1, 0, "", "DubAppActive" + allIDs["instalation"]);
  console.log("OpenSht ejecutado correctamente");
  console.log("auxValues.length:", auxValues.length);
}

function probarApiCargaTodosLosProyectos() {
  console.log("Probando apiCargaTodosLosProyectos");
  const emailPrueba = "usuario@mediaaccesscompany.com";
  const resultado = apiCargaTodosLosProyectos(emailPrueba);
  console.log(JSON.stringify(resultado, null, 2));
}

function probarApiCargaMisProyectos() {
  console.log("Probando apiCargaMisProyectos");
  const emailPrueba = "usuario@mediaaccesscompany.com";
  try {
    const resultado = apiCargaMisProyectos(emailPrueba);
    console.log(JSON.stringify(resultado, null, 2));
  } catch (error) {
    console.error("Error:", error.toString());
  }
}

function probarApiProcesarProjectID() {
  console.log("Probando apiProcesarProjectID");
  const projectIDPrueba = "PRJ12345"; // Reemplaza con un ID real
  const emailPrueba = "usuario@mediaaccesscompany.com";
  try {
    const resultado = apiProcesarProjectID(projectIDPrueba, emailPrueba);
    console.log(JSON.stringify(resultado, null, 2));
  } catch (error) {
    console.error("Error:", error.toString());
  }
}

function probarApiActualizarDWOEvent() {
  console.log("Probando apiActualizarDWOEvent");
  const emailPrueba = "usuario@mediaaccesscompany.com";
  // Ejemplo de cambios - adaptar según la estructura real
  const cambiosPrueba = [
    {
      eventID: "EVT12345", // Reemplazar con un ID real
      fila: 5,
      columna: 3,
      valorNuevo: "Valor actualizado"
    }
  ];
  try {
    const resultado = apiActualizarDWOEvent(cambiosPrueba, emailPrueba);
    console.log(JSON.stringify(resultado, null, 2));
  } catch (error) {
    console.error("Error:", error.toString());
  }
}

/**
 * Copia de la función apiCargaTodosLosProyectos del archivo DubAppDatesWebService.js
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
 * Copia de la función OpenSht del archivo DubAppDatesWebService.js
 * Debe mantenerse actualizada con la versión en ese archivo
 */
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
 * Prueba el servicio web realizando una solicitud HTTP
 */
function probarServicioWeb() {
  // URL del servicio web (reemplaza con la URL real de tu servicio web publicado)
  const url = PropertiesService.getScriptProperties().getProperty('WEB_SERVICE_URL') || 
              'https://script.google.com/macros/s/TU_ID_SERVICIO_WEB/exec';
  
  console.log('Probando servicio web en URL: ' + url);
  
  // Datos a enviar
  const payload = {
    action: 'cargaTodosLosProyectos',
    userEmail: 'usuario@mediaaccesscompany.com'
  };
  
  // Opciones de la solicitud
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  // Realizar la solicitud
  try {
    console.log('Enviando solicitud POST con payload:', JSON.stringify(payload));
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const contentText = response.getContentText();
    
    console.log('Código de respuesta HTTP:', responseCode);
    console.log('Headers:', JSON.stringify(response.getAllHeaders()));
    
    if (contentText.length > 1000) {
      console.log('Contenido de respuesta (primeros 500 caracteres):', contentText.substring(0, 500) + '...');
    } else {
      console.log('Contenido de respuesta completo:', contentText);
    }
    
    // Intentar parsear la respuesta como JSON
    try {
      const jsonResponse = JSON.parse(contentText);
      console.log('Respuesta JSON parseada:', JSON.stringify(jsonResponse, null, 2));
      
      // Análisis adicional si se pudo parsear como JSON
      if (jsonResponse.success === false) {
        console.log('El servicio web reportó un error:', jsonResponse.message);
      } else if (jsonResponse.success === true) {
        console.log('El servicio web respondió correctamente');
        if (jsonResponse.data && jsonResponse.data.proyectos) {
          console.log('Número de proyectos recibidos:', jsonResponse.data.proyectos.length);
        }
      }
    } catch (parseError) {
      console.error('Error al parsear respuesta como JSON:', parseError.toString());
      
      // Analizar el contenido si no es JSON (probablemente HTML)
      if (contentText.includes('<!DOCTYPE html>')) {
        console.error('La respuesta es HTML, posiblemente una página de error');
        
        // Buscar mensajes de error comunes en el HTML
        const errorMatch = contentText.match(/<div.*?class="error".*?>(.*?)<\/div>/i) || 
                          contentText.match(/<h1.*?>(.*?)<\/h1>/i) ||
                          contentText.match(/<title>(.*?)<\/title>/i);
        
        if (errorMatch) {
          console.error('Posible mensaje de error encontrado en el HTML:', errorMatch[1]);
        }
      }
    }
  } catch (fetchError) {
    console.error('Error al realizar la solicitud HTTP:', fetchError.toString());
    console.error('Stack:', fetchError.stack);
  }
}

/**
 * Configurar la URL del servicio web a través de las propiedades del script
 * @param {string} url - URL del servicio web (opcional)
 */
function configurarURLServicioWeb(url) {
  // Si no se proporciona la URL, usar un valor predeterminado
  if (!url) {
    url = 'https://script.google.com/macros/s/TU_ID_SERVICIO_WEB/exec';
    console.log('No se proporcionó URL. Debes editar manualmente el código para configurar la URL correcta.');
  }
  
  PropertiesService.getScriptProperties().setProperty('WEB_SERVICE_URL', url);
  console.log('URL del servicio web configurada:', url);
  
  return url;
}

/**
 * Verifica la configuración del servicio web en el proyecto original
 * Debe ejecutarse desde el script donde está definido el servicio web
 */
function verificarImplementacionServicioWeb() {
  console.log("Verificando implementación del servicio web...");
  
  try {
    // 1. Verificar si el script está publicado como aplicación web
    const deploymentInfo = ScriptApp.getService();
    console.log("Estado de implementación:");
    
    if (deploymentInfo && deploymentInfo.isEnabled()) {
      console.log("✅ El servicio web está publicado y habilitado");
      console.log("URL:", deploymentInfo.getUrl());
    } else {
      console.log("❌ El servicio web no está publicado o está deshabilitado");
    }
  } catch (e) {
    console.log("❌ Error al verificar implementación:", e.toString());
  }
  
  try {  
    // 2. Verificar los permisos del script
    const scriptProperties = PropertiesService.getScriptProperties().getProperties();
    console.log("\nPropiedades configuradas:");
    for (const key in scriptProperties) {
      console.log(`- ${key}: ${scriptProperties[key]}`);
    }
  } catch (e) {
    console.log("❌ Error al verificar propiedades:", e.toString());
  }
  
  try {
    // 3. Verificar la existencia de funciones críticas
    const projectFiles = DriveApp.getFileById(ScriptApp.getScriptId()).getBlob().getDataAsString();
    console.log("\nVerificando funciones críticas:");
    
    const funciones = [
      "doGet", "doPost", "apiCargaTodosLosProyectos", "OpenSht", "LazyLoad", "databaseID.getID"
    ];
    
    for (const funcion of funciones) {
      if (projectFiles.includes(`function ${funcion}`)) {
        console.log(`✅ Función '${funcion}' encontrada`);
      } else {
        console.log(`❓ Función '${funcion}' no encontrada directamente`);
      }
    }
  } catch (e) {
    console.log("❌ Error al verificar funciones:", e.toString());
  }
  
  console.log("\nVerificación completada.");
}

/**
 * Analiza la URL del servicio web y proporciona información de diagnóstico
 */
function analizarURLServicioWeb() {
  const url = PropertiesService.getScriptProperties().getProperty('WEB_SERVICE_URL') || '';
  
  if (!url) {
    console.log("❌ No hay URL configurada. Por favor, configura la URL primero.");
    return;
  }
  
  console.log("Analizando URL del servicio web:", url);
  
  // Verificar formato básico
  if (url.includes('script.google.com/macros/s/') && url.includes('/exec')) {
    console.log("✅ El formato de la URL parece correcto");
    
    // Extraer ID de despliegue
    const match = url.match(/script\.google\.com\/macros\/s\/([^\/]+)\/exec/);
    if (match && match[1]) {
      console.log("ID de despliegue:", match[1]);
    }
  } else {
    console.log("❌ El formato de la URL no coincide con el esperado para un despliegue de Apps Script");
  }
  
  // Verificar respuesta HTTP básica
  try {
    console.log("\nVerificando acceso básico a la URL...");
    const response = UrlFetchApp.fetch(url, {
      method: 'get',
      muteHttpExceptions: true
    });
    
    const responseCode = response.getResponseCode();
    console.log("Código de respuesta:", responseCode);
    
    if (responseCode >= 200 && responseCode < 300) {
      console.log("✅ La URL responde correctamente");
    } else {
      console.log("❌ La URL responde con error:", responseCode);
    }
    
    // Verificar el tipo de contenido
    const contentType = response.getHeaders()['Content-Type'] || '';
    console.log("Tipo de contenido:", contentType);
    
    if (contentType.includes('application/json')) {
      console.log("✅ El servicio responde con JSON");
    } else if (contentType.includes('text/html')) {
      console.log("❌ El servicio responde con HTML en lugar de JSON");
    }
  } catch (e) {
    console.log("❌ Error al verificar acceso a la URL:", e.toString());
  }
}

/**
 * Probar el servicio web con una URL específica
 * @param {string} url - URL del servicio web a probar
 */
function probarServicioWebConURL(url) {
  if (!url) {
    console.log('Error: Se requiere una URL para probar el servicio web');
    return;
  }
  
  console.log('Configurando URL:', url);
  configurarURLServicioWeb(url);
  
  console.log('Probando servicio web...');
  probarServicioWeb();
} 