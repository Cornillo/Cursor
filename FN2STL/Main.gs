/**
 * Forced Narratives to STL (FN2STL)
 * Herramienta para convertir subtítulos de Google Sheets al formato EBU STL
 * según la especificación EBU Tech 3264-1991
 * 
 * @author FN2STL Team
 * @version 1.0
 * @OnlyCurrentDoc
 */

/**
 * Configuración global del sistema
 */
const CONFIG = {
  COUNTRY_CODES: {
    ARG: "ARG", // Argentina
    BRA: "BRA", // Brasil
    MEX: "MEX"  // México
  },
  LANGUAGE_CODES: {
    SPANISH: "0A",    // Español
    PORTUGUESE: "21"  // Portugués
  },
  GSI_BLOCK_SIZE: 1024,  // Tamaño del bloque GSI en bytes
  TTI_BLOCK_SIZE: 128,   // Tamaño del bloque TTI en bytes
  FIRST_DATA_ROW: 2,    // Primera fila con datos de subtítulos (después del encabezado)
  DEFAULT_DISPLAY_STANDARD: "0", // 0 = Open subtitling
  DEFAULT_MAX_ROWS: 2,  // Número máximo de filas por subtítulo
  DEFAULT_MAX_CHARS_PER_ROW: 40 // Número máximo de caracteres por línea
};

// Variable global para controlar el modo verbose
let verboseFlag = false;

/**
 * Establece el modo verbose para depuración
 * @param {boolean} value - Valor para el modo verbose
 */
function setVerbose(value) {
  verboseFlag = !!value;
  Logger.log("Modo verbose: " + (verboseFlag ? "ACTIVADO" : "DESACTIVADO"));
}

/**
 * Registra un mensaje en el log si el modo verbose está activado
 * @param {string} message - Mensaje a registrar
 */
function logMessage(message) {
  if (verboseFlag) {
    Logger.log(message);
  }
}

/**
 * Convierte los subtítulos de una hoja de Google Sheets a formato STL
 * @param {string} sheetId - ID de la hoja de Google Sheets
 * @param {string} country - Código de país (3 caracteres)
 * @param {string} languageCode - Código de idioma (2 caracteres)
 * @param {string} folderId - ID de la carpeta de destino en Google Drive
 * @param {boolean} verboseFlag - Activar logs detallados
 * @return {object} Resultado de la conversión con URL del archivo generado
 */
function convertSheetToSTL(sheetId, country, languageCode, folderId, verboseFlag) {
  if (verboseFlag) Logger.log(`Iniciando conversión de hoja ${sheetId} a STL`);
  
  try {
    // 1. Abrir la hoja de cálculo
    const spreadsheet = SpreadsheetApp.openById(sheetId);
    if (!spreadsheet) {
      throw new Error(`No se pudo abrir la hoja con ID: ${sheetId}`);
    }
    
    const sheet = spreadsheet.getActiveSheet();
    const sheetName = sheet.getName();
    
    if (verboseFlag) Logger.log(`Hoja abierta: "${sheetName}"`);
    
    // 2. Analizar la estructura de la hoja
    const structure = analyzeSheetStructure(sheet, verboseFlag);
    
    if (!structure.valid) {
      throw new Error(`Estructura de hoja inválida: ${structure.error}`);
    }
    
    if (verboseFlag) {
      Logger.log(`Estructura de hoja detectada:`);
      Logger.log(`- Columna de número: ${structure.numberCol}`);
      Logger.log(`- Columna de tiempo inicial: ${structure.startTimeCol}`);
      Logger.log(`- Columna de tiempo final: ${structure.endTimeCol}`);
      Logger.log(`- Columna de texto: ${structure.textCol}`);
      Logger.log(`- Fila de inicio de datos: ${structure.dataStartRow}`);
    }
    
    // 3. Extraer los datos de subtítulos
    const subtitles = extractSubtitleData(sheet, structure, verboseFlag);
    
    if (subtitles.length === 0) {
      throw new Error("No se encontraron subtítulos válidos en la hoja");
    }
    
    if (verboseFlag) Logger.log(`Se extrajeron ${subtitles.length} subtítulos`);
    
    // 4. Crear el bloque GSI
    const gsiBlock = createGSIBlock(sheetName, country, languageCode, subtitles.length, verboseFlag);
    
    // 5. Crear los bloques TTI para cada subtítulo
    const ttiBlocks = [];
    
    for (let i = 0; i < subtitles.length; i++) {
      const subtitle = subtitles[i];
      
      if (verboseFlag && i < 5) {
        Logger.log(`Procesando subtítulo #${i+1}: ${subtitle.startTime} - ${subtitle.endTime} "${subtitle.text}"`);
      }
      
      const ttiBlock = createTTIBlock(
        subtitle.number,
        subtitle.startTime,
        subtitle.endTime,
        subtitle.text,
        verboseFlag
      );
      
      ttiBlocks.push(ttiBlock);
    }
    
    // 6. Combinar los bloques en un archivo STL
    const stlFile = combineBlocks(gsiBlock, ttiBlocks, verboseFlag);
    
    // 7. Guardar el archivo en Google Drive
    const fileName = `${sheetName.replace(/[^\w\s]/gi, '')}_${new Date().toISOString().slice(0,10)}`;
    
    // Crear el archivo en Drive
    let folder;
    if (folderId) {
      try {
        folder = DriveApp.getFolderById(folderId);
      } catch (e) {
        if (verboseFlag) Logger.log(`Error al obtener carpeta: ${e.message}. Usando carpeta raíz.`);
        folder = DriveApp.getRootFolder();
      }
    } else {
      folder = DriveApp.getRootFolder();
    }
    
    // Guardar el archivo usando la función saveSTLFile
    const fileInfo = saveSTLFile(stlFile, fileName, folder.getId(), verboseFlag);
    
    return {
      success: true,
      fileId: fileInfo.id,
      fileName: fileInfo.name,
      fileUrl: fileInfo.url,
      downloadUrl: fileInfo.downloadUrl,
      subtitleCount: subtitles.length
    };
    
  } catch (e) {
    Logger.log(`Error en la conversión: ${e.message}`);
    return {
      success: false,
      error: e.message
    };
  }
}

/**
 * Valida los parámetros de entrada para la conversión
 * 
 * @param {string} sheetId - ID de la hoja de cálculo
 * @param {string} country - Código del país
 * @param {string} languageCode - Código del idioma
 * @param {string} tempFolderId - ID de la carpeta temporal
 * @throws {Error} Si algún parámetro no es válido
 */
function validateParameters(sheetId, country, languageCode, tempFolderId) {
  logMessage("Validando parámetros de entrada...");
  
  // Validar ID de hoja de cálculo
  if (!sheetId || typeof sheetId !== 'string') {
    throw new Error("El ID de la hoja de cálculo es obligatorio y debe ser una cadena de texto");
  }
  
  // Validar código de país
  if (!country) {
    throw new Error("El código de país es obligatorio");
  } else if (!CONFIG.COUNTRY_CODES[country]) {
    throw new Error("Código de país no válido. Valores aceptados: ARG, BRA, MEX");
  }
  
  // Validar código de idioma
  if (!languageCode) {
    throw new Error("El código de idioma es obligatorio");
  } else if (!Object.values(CONFIG.LANGUAGE_CODES).includes(languageCode)) {
    throw new Error("Código de idioma no válido. Valores aceptados: 0A (Español), 21 (Portugués)");
  }
  
  // Validar ID de carpeta temporal
  if (!tempFolderId || typeof tempFolderId !== 'string') {
    throw new Error("El ID de la carpeta temporal es obligatorio y debe ser una cadena de texto");
  }
  
  // Verificar que la hoja de cálculo existe y es accesible
  try {
    SpreadsheetApp.openById(sheetId);
    logMessage("Hoja de cálculo verificada");
  } catch (e) {
    throw new Error("No se puede acceder a la hoja de cálculo. Verifique el ID y los permisos: " + e.message);
  }
  
  // Verificar que la carpeta existe y es accesible
  try {
    DriveApp.getFolderById(tempFolderId);
    logMessage("Carpeta temporal verificada");
  } catch (e) {
    throw new Error("No se puede acceder a la carpeta temporal. Verifique el ID y los permisos: " + e.message);
  }
  
  logMessage("Todos los parámetros validados correctamente");
}

/**
 * Función para probar la conversión con datos de muestra o con una hoja real
 * @param {string} sheetId - ID opcional de una hoja real para pruebas
 */
function testConversion(sheetId) {
  Logger.log("=== INICIANDO PRUEBA DE CONVERSIÓN ===");
  
  try {
    // Activar modo verbose
    setVerbose(true);
    
    // Si se proporciona un ID de hoja, usar esa hoja para la prueba
    if (sheetId) {
      Logger.log(`Usando hoja real para prueba: ${sheetId}`);
      
      // Carpeta para guardar el resultado
      const testFolderId = "1ygkUbBtKKWcvq7WQuFXhc3smcac-0cha";
      
      // Llamar a convertSheetToSTL con la hoja real
      const result = convertSheetToSTL(sheetId, "ARG", "0A", testFolderId, true);
      
      Logger.log("=== PRUEBA COMPLETADA CON HOJA REAL ===");
      Logger.log(`Archivo STL generado: ${result.fileName}`);
      Logger.log(`URL: ${result.fileUrl}`);
      
      return result;
    }
    
    // Si no hay ID de hoja, usar los datos de muestra como antes
    const sampleSubtitles = [
      {
        index: 1,
        timeIn: "01:00:10:15",
        timeOut: "01:00:15:00",
        text: "Línea 1 de subtítulo\nLínea 2 de subtítulo"
      },
      {
        index: 2,
        timeIn: "01:00:20:04",
        timeOut: "01:00:23:12",
        text: "Este es otro subtítulo de prueba"
      },
      {
        index: 3,
        timeIn: "01:00:30:00",
        timeOut: "01:00:35:08",
        text: "Subtítulo con caracteres especiales: áéíóúñ¿?¡!%"
      }
    ];
    
    Logger.log("\n=== PRUEBA 1: VALIDACIÓN DE TIMECODES ===");
    const timecodeTests = [
      "01:02:03:04",               // Formato correcto HH:MM:SS:FF
      "10:20:30.15",               // Con punto en lugar de :
      "01:02:03,25",               // Con coma en lugar de :
      "1:02:03:04",                // Hora sin cero inicial
      "01:2:03:04",                // Minuto sin cero inicial
      new Date(2023, 5, 15, 1, 2, 3), // Objeto Date
      "texto inválido",            // Texto que no es timecode
      "25:70:99:40",               // Valores fuera de rango
      null,                        // null
      undefined,                   // undefined
      ""                           // String vacío
    ];
    
    for (const test of timecodeTests) {
      const result = validateAndFormatTimecode(test, true);
      Logger.log(`- Validando timecode: ${String(test)} -> ${result ? "VÁLIDO (" + result + ")" : "INVÁLIDO"}`);
    }
    
    Logger.log("\n=== PRUEBA 2: VALIDACIÓN DE TEXTO ===");
    const textTests = [
      "Texto normal",
      "Texto con salto\nde línea",
      "Texto con \"comillas especiales\"",
      "Texto muy largo que podría exceder el límite de caracteres por línea y requerir un tratamiento especial para asegurar la correcta visualización",
      "Texto con caracteres especiales como: á é í ó ú ñ Ñ ¿ ? ¡ ! % & $ @ # * ( ) + – — ...",
      null,
      undefined,
      ""
    ];
    
    for (const test of textTests) {
      const result = validateAndFormatText(test);
      Logger.log(`- Validando texto: "${String(test).substring(0, 30)}${test && test.length > 30 ? '...' : ''}" -> ${result ? "VÁLIDO" : "INVÁLIDO"}`);
    }
    
    Logger.log("\n=== PRUEBA 3: CREACIÓN DE BLOQUES STL ===");
    
    // Probar creación de bloque GSI
    Logger.log("Creando bloque GSI...");
    const gsiBlock = createGSIBlock("Prueba", "ARG", "0A", sampleSubtitles.length, true);
    Logger.log(`- Bloque GSI creado: ${gsiBlock ? gsiBlock.length + " bytes" : "ERROR"}`);
    
    // Probar creación de bloques TTI
    Logger.log("Creando bloques TTI...");
    const ttiBlocks = [];
    for (const subtitle of sampleSubtitles) {
      try {
        const ttiBlock = createTTIBlock(subtitle.index, subtitle.timeIn, subtitle.timeOut, subtitle.text, true);
        ttiBlocks.push(ttiBlock);
        Logger.log(`- Bloque TTI #${subtitle.index} creado: ${ttiBlock ? ttiBlock.length + " bytes" : "ERROR"}`);
      } catch (error) {
        Logger.log(`- Error al crear bloque TTI #${subtitle.index}: ${error.message}`);
      }
    }
    
    // Probar combinación de bloques
    Logger.log("Combinando bloques...");
    const stlData = combineBlocks(gsiBlock, ttiBlocks, true);
    Logger.log(`- Datos STL creados: ${stlData ? stlData.length + " bytes" : "ERROR"}`);
    
    // Guardar archivo STL real en Drive
    const testFolderId = "1ygkUbBtKKWcvq7WQuFXhc3smcac-0cha";
    const testFileName = "FN2STL_Test.STL";
    Logger.log("\n=== GUARDANDO ARCHIVO STL REAL ===");
    try {
      const fileId = saveSTLFile(stlData, testFileName, testFolderId, true);
      Logger.log(`- Archivo STL guardado en Drive: ${testFileName}`);
      Logger.log(`- FileId: ${fileId}`);
      Logger.log(`- URL: https://drive.google.com/file/d/${fileId}/view`);
    } catch (error) {
      Logger.log(`- Error al guardar archivo STL: ${error.message}`);
    }
    
    Logger.log("\n=== PRUEBA 4: ANÁLISIS DE ESTRUCTURA DE HOJA ===");
    // Crear una hoja de prueba temporal
    const tempSs = SpreadsheetApp.create("FN2STL_Temp_Test");
    const tempSheet = tempSs.getActiveSheet();
    
    // Agregar encabezados
    tempSheet.getRange("A1:E1").setValues([["#", "Time In", "Spanish Text", "Time Out", "English Text"]]);
    
    // Agregar datos de muestra
    const sampleData = sampleSubtitles.map(subtitle => [
      subtitle.index,
      subtitle.timeIn,
      subtitle.text,
      subtitle.timeOut,
      `English translation for: ${subtitle.text.split('\n')[0]}`
    ]);
    
    tempSheet.getRange(2, 1, sampleData.length, 5).setValues(sampleData);
    
    // Analizar estructura
    Logger.log("Analizando estructura de hoja temporal...");
    const structure = analyzeSheetStructure(tempSheet, 1, true);
    
    Logger.log(`- Estructura detectada: TimeIn=${structure.timeInIndex !== -1 ? String.fromCharCode(65 + structure.timeInIndex) : "NO"}, ` +
              `Text=${structure.textIndex !== -1 ? String.fromCharCode(65 + structure.textIndex) : "NO"}, ` +
              `TimeOut=${structure.timeOutIndex !== -1 ? String.fromCharCode(65 + structure.timeOutIndex) : "NO"}, ` +
              `EnglishText=${structure.englishTextIndex !== -1 ? String.fromCharCode(65 + structure.englishTextIndex) : "NO"}`);
    
    // Limpiar después de la prueba
    DriveApp.getFileById(tempSs.getId()).setTrashed(true);
    
    Logger.log("\n=== PRUEBA 5: VALIDACIÓN DE MAPEO DE CARACTERES ===");
    const specialChars = "áéíóúüñÑ¿?¡!%&$#@~^*()_+-=[]{}|\\:;\"'<>,./ –—•…";
    Logger.log(`- Caracteres originales: ${specialChars}`);
    const mappedChars = mapSpecialCharacters(specialChars);
    Logger.log(`- Caracteres mapeados: ${mappedChars}`);
    
    Logger.log("\n=== PRUEBA COMPLETA ===");
    Logger.log("La prueba de conversión ha finalizado correctamente.");
    
  } catch (error) {
    Logger.log(`ERROR EN LA PRUEBA: ${error.message}`);
    Logger.log(`Stack: ${error.stack}`);
  }
}

/**
 * Analiza la estructura de la hoja de cálculo para depuración
 * @param {string} sheetId - ID de la hoja de cálculo
 */
function analyzeSheetStructure(sheetId) {
  try {
    const spreadsheet = SpreadsheetApp.openById(sheetId);
    const sheet = spreadsheet.getActiveSheet();
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    
    Logger.log("\n=== ANÁLISIS DE ESTRUCTURA DE HOJA ===");
    Logger.log("Nombre de la hoja: " + sheet.getName());
    Logger.log("Total de filas: " + values.length);
    
    // Analizar encabezados (fila 9, índice 8)
    if (values.length > 8) {
      Logger.log("\nEncabezados (Fila 9):");
      for (let i = 0; i < values[8].length; i++) {
        Logger.log(`  Columna ${String.fromCharCode(65 + i)}: "${values[8][i]}"`);
      }
    }
    
    // Analizar primeras filas de datos
    Logger.log("\nPrimeras filas de datos:");
    for (let i = 10; i < Math.min(values.length, 15); i++) {
      let rowLog = `Fila ${i+1}: `;
      for (let j = 0; j < Math.min(values[i].length, 5); j++) {
        rowLog += `[${String.fromCharCode(65 + j)}:"${values[i][j]}"] `;
      }
      Logger.log(rowLog);
    }
    
    // Buscar códigos de tiempo
    Logger.log("\nBúsqueda de códigos de tiempo:");
    for (let i = 10; i < Math.min(values.length, 15); i++) {
      for (let j = 0; j < Math.min(values[i].length, 5); j++) {
        const value = values[i][j];
        if (value && typeof value === 'string' && value.match(/\d{2}:\d{2}:\d{2}/)) {
          Logger.log(`  Encontrado posible timecode en Fila ${i+1}, Columna ${String.fromCharCode(65 + j)}: "${value}"`);
        }
      }
    }
    
  } catch (error) {
    Logger.log("Error al analizar la estructura: " + error.message);
  }
}

/**
 * Función que se ejecuta cuando se accede a la aplicación web
 * @param {object} e - Objeto de evento
 * @return {HtmlOutput} Salida HTML
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Conversor de Subtítulos a STL')
    .setFaviconUrl('https://www.google.com/images/favicon.ico')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Contenido HTML para la interfaz de usuario
 * @return {string} Contenido HTML
 */
function getIndexContent() {
  return `
<!DOCTYPE html>
<html>
  <head>
    <base target="_top">
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <style>
      body {
        font-family: Arial, sans-serif;
        margin: 0;
        padding: 20px;
        background-color: #f5f5f5;
      }
      .container {
        max-width: 800px;
        margin: 0 auto;
        background-color: white;
        padding: 20px;
        border-radius: 8px;
        box-shadow: 0 2px 10px rgba(0,0,0,0.1);
      }
      h1 {
        color: #4285f4;
        margin-top: 0;
      }
      .form-group {
        margin-bottom: 15px;
      }
      label {
        display: block;
        margin-bottom: 5px;
        font-weight: bold;
      }
      input[type="text"] {
        width: 100%;
        padding: 8px;
        border: 1px solid #ddd;
        border-radius: 4px;
        box-sizing: border-box;
      }
      button {
        background-color: #4285f4;
        color: white;
        border: none;
        padding: 10px 15px;
        border-radius: 4px;
        cursor: pointer;
        font-size: 16px;
      }
      button:hover {
        background-color: #3367d6;
      }
      .loading {
        display: none;
        margin-top: 20px;
        text-align: center;
      }
      .result {
        margin-top: 20px;
        padding: 15px;
        border-radius: 4px;
      }
      .success {
        background-color: #d4edda;
        border: 1px solid #c3e6cb;
        color: #155724;
      }
      .error {
        background-color: #f8d7da;
        border: 1px solid #f5c6cb;
        color: #721c24;
      }
    </style>
  </head>
  <body>
    <div class="container">
      <h1>Conversor de Subtítulos a STL</h1>
      <p>Esta herramienta convierte subtítulos desde Google Sheets al formato EBU STL.</p>
      
      <div class="form-group">
        <label for="sheetId">ID de la Hoja de Google Sheets:</label>
        <input type="text" id="sheetId" name="sheetId" placeholder="Ej: 1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms">
      </div>
      
      <div class="form-group">
        <label for="country">Código de País:</label>
        <input type="text" id="country" name="country" value="ARG" placeholder="Ej: ARG">
      </div>
      
      <div class="form-group">
        <label for="languageCode">Código de Idioma:</label>
        <input type="text" id="languageCode" name="languageCode" value="0A" placeholder="Ej: 0A para español">
      </div>
      
      <div class="form-group">
        <label for="folderId">ID de Carpeta de Destino (opcional):</label>
        <input type="text" id="folderId" name="folderId" placeholder="Dejar en blanco para usar la carpeta raíz">
      </div>
      
      <div class="form-group">
        <label>
          <input type="checkbox" id="verbose" name="verbose">
          Modo Verbose (para depuración)
        </label>
      </div>
      
      <button onclick="convertToSTL()">Convertir a STL</button>
      
      <div id="loading" class="loading">
        <p>Procesando... Por favor espere.</p>
      </div>
      
      <div id="result"></div>
    </div>
    
    <script>
      function convertToSTL() {
        // Validar campos
        const sheetId = document.getElementById('sheetId').value.trim();
        if (!sheetId) {
          alert('Por favor ingrese un ID de hoja válido');
          return;
        }
        
        // Mostrar indicador de carga
        document.getElementById('loading').style.display = 'block';
        document.getElementById('result').innerHTML = '';
        
        // Recopilar datos del formulario
        const formData = {
          sheetId: sheetId,
          country: document.getElementById('country').value.trim(),
          languageCode: document.getElementById('languageCode').value.trim(),
          folderId: document.getElementById('folderId').value.trim(),
          verbose: document.getElementById('verbose').checked
        };
        
        // Llamar a la función del servidor
        google.script.run
          .withSuccessHandler(function(result) {
            document.getElementById('loading').style.display = 'none';
            
            if (result.success) {
              let html = '<div class="result success">';
              html += '<h3>Conversión Exitosa</h3>';
              html += '<p>Archivo: ' + result.fileName + '</p>';
              html += '<p><a href="' + result.fileUrl + '" target="_blank">Ver en Google Drive</a></p>';
              html += '</div>';
              
              document.getElementById('result').innerHTML = html;
            } else {
              let html = '<div class="result error">';
              html += '<h3>Error en la Conversión</h3>';
              html += '<p>' + result.error + '</p>';
              html += '</div>';
              
              document.getElementById('result').innerHTML = html;
            }
          })
          .withFailureHandler(function(error) {
            document.getElementById('loading').style.display = 'none';
            
            let html = '<div class="result error">';
            html += '<h3>Error en la Conversión</h3>';
            html += '<p>' + error.message + '</p>';
            html += '</div>';
            
            document.getElementById('result').innerHTML = html;
          })
          .convertFromUI(formData);
      }
    </script>
  </body>
</html>
  `;
}

/**
 * Función para convertir desde la interfaz de usuario
 * @param {object} formData - Datos del formulario
 * @return {object} Resultado de la conversión
 */
function convertFromUI(formData) {
  try {
    // Validar parámetros
    if (!formData.sheetId) {
      return { success: false, error: "El ID de la hoja es requerido" };
    }
    
    // Establecer valores por defecto
    const country = formData.country || "ARG";
    const languageCode = formData.languageCode || "0A";
    const folderId = formData.folderId || null;
    const verboseFlag = formData.verbose === "true" || formData.verbose === true;
    
    // Ejecutar la conversión
    const result = convertSheetToSTL(
      formData.sheetId,
      country,
      languageCode,
      folderId,
      verboseFlag
    );
    
    return result;
    
  } catch (e) {
    Logger.log(`Error en convertFromUI: ${e.message}`);
    return {
      success: false,
      error: e.message
    };
  }
}

/**
 * Genera el contenido del archivo de depuración
 * @param {Uint8Array} stlData - Datos del archivo STL
 * @param {string} title - Título del programa
 * @param {Array} subtitles - Array de subtítulos
 * @return {string} Contenido del archivo de depuración
 */
function generateDebugContent(stlData, title, subtitles) {
  let content = "=== INFORMACIÓN DE DEPURACIÓN FN2STL ===\n\n";
  content += "Título: " + title + "\n";
  content += "Fecha de generación: " + new Date().toISOString() + "\n";
  content += "Número de subtítulos: " + subtitles.length + "\n\n";
  
  content += "=== SUBTÍTULOS ===\n";
  subtitles.forEach((subtitle, index) => {
    content += `\n[${index + 1}]\n`;
    content += `Time In: ${subtitle.timeIn}\n`;
    content += `Time Out: ${subtitle.timeOut}\n`;
    content += `Texto: ${subtitle.text}\n`;
  });
  
  content += "\n=== INFORMACIÓN DEL ARCHIVO STL ===\n";
  content += `Tamaño total: ${stlData.length} bytes\n`;
  content += `GSI Block: ${CONFIG.GSI_BLOCK_SIZE} bytes\n`;
  content += `TTI Blocks: ${subtitles.length} x ${CONFIG.TTI_BLOCK_SIZE} bytes\n`;
  
  return content;
}

/**
 * Función para realizar una prueba de conversión desde la interfaz de usuario
 * @param {string} sheetId - ID de la hoja de Google Sheets a usar para la prueba
 * @return {Object} - Información del archivo generado
 */
function testConversionFromUI(sheetId) {
  try {
    setVerbose(true);
    Logger.log("=== INICIANDO PRUEBA DE CONVERSIÓN DESDE UI ===");
    
    // Abrir la hoja para obtener datos básicos
    const spreadsheet = SpreadsheetApp.openById(sheetId);
    const title = spreadsheet.getName();
    
    Logger.log(`Hoja encontrada: "${title}"`);
    
    // Identificar la carpeta para guardar el resultado
    const testFolderId = "1ygkUbBtKKWcvq7WQuFXhc3smcac-0cha";
    
    // Llamar a la conversión normal pero asegurando que el nombre sea el de la hoja
    return convertSheetToSTL(sheetId, "ARG", "0A", testFolderId, true);
    
  } catch (error) {
    Logger.log("Error en testConversionFromUI: " + error.message);
    throw {
      error: error.message
    };
  }
}

/**
 * Función para probar la conversión con un ID de hoja real
 * @param {string} sheetId - ID de la hoja de Google Sheets a convertir
 * @return {object} Resultado de la conversión
 */
function testConversionReal(sheetId) {
  Logger.log("=== INICIANDO PRUEBA DE CONVERSIÓN CON HOJA REAL ===");
  
  if (!sheetId) {
    Logger.log("Error: Se requiere un ID de hoja válido");
    return { success: false, error: "ID de hoja no proporcionado" };
  }
  
  Logger.log(`Usando hoja con ID: ${sheetId}`);
  
  // Parámetros por defecto
  const country = "ARG";
  const languageCode = "0A";
  const verboseFlag = true;
  
  try {
    // Ejecutar la conversión
    const result = convertSheetToSTL(sheetId, country, languageCode, null, verboseFlag);
    
    Logger.log("=== PRUEBA COMPLETADA ===");
    Logger.log(`Resultado: ${result.success ? "ÉXITO" : "ERROR"}`);
    
    if (result.success) {
      Logger.log(`Archivo STL generado: ${result.fileName}`);
      Logger.log(`URL: ${result.fileUrl}`);
      Logger.log(`Tamaño: ${result.fileSize} bytes`);
      Logger.log(`Subtítulos: ${result.subtitleCount}`);
    } else {
      Logger.log(`Error: ${result.error}`);
    }
    
    return result;
  } catch (e) {
    Logger.log(`Error en la prueba: ${e.message}`);
    return { success: false, error: e.message };
  }
}

/**
 * Función de prueba para detectar la estructura de una hoja específica
 * @param {string} sheetId - ID de la hoja de cálculo
 */
function probarDeteccionEstructura(sheetId) {
  Logger.log("=== PRUEBA DE DETECCIÓN DE ESTRUCTURA ===");
  Logger.log("Usando hoja con ID: " + sheetId);
  
  try {
    // Abrir la hoja
    const spreadsheet = SpreadsheetApp.openById(sheetId);
    const sheet = spreadsheet.getActiveSheet();
    
    Logger.log("Hoja abierta: \"" + spreadsheet.getName() + "\"");
    
    // Detectar estructura con modo verbose
    const verboseFlag = true;
    const structure = analyzeSheetStructure(sheet, verboseFlag);
    
    if (structure.valid) {
      Logger.log("Estructura detectada correctamente:");
      Logger.log("- Columna de número: " + columnToLetter(structure.numberCol) + " (" + structure.numberCol + ")");
      Logger.log("- Columna de tiempo inicial: " + columnToLetter(structure.startTimeCol) + " (" + structure.startTimeCol + ")");
      Logger.log("- Columna de texto: " + columnToLetter(structure.textCol) + " (" + structure.textCol + ")");
      Logger.log("- Columna de tiempo final: " + columnToLetter(structure.endTimeCol) + " (" + structure.endTimeCol + ")");
      Logger.log("- Fila de inicio de datos: " + structure.dataStartRow);
      
      // Obtener algunos ejemplos de datos
      const startRow = structure.dataStartRow;
      const sampleSize = Math.min(3, sheet.getLastRow() - startRow + 1);
      
      if (sampleSize > 0) {
        Logger.log("\nEjemplos de datos:");
        
        const range = sheet.getRange(startRow, 1, sampleSize, Math.max(structure.startTimeCol, structure.textCol, structure.endTimeCol));
        const values = range.getValues();
        
        for (let i = 0; i < values.length; i++) {
          const row = values[i];
          const startTime = row[structure.startTimeCol - 1];
          const text = row[structure.textCol - 1];
          const endTime = row[structure.endTimeCol - 1];
          
          Logger.log(`Subtítulo #${i+1}:`);
          Logger.log(`- Tiempo inicial: ${startTime}`);
          Logger.log(`- Texto: ${text}`);
          Logger.log(`- Tiempo final: ${endTime}`);
        }
      }
    } else {
      Logger.log("Error en la detección de estructura: " + structure.error);
    }
  } catch (error) {
    Logger.log("Error: " + error.message);
  }
  
  Logger.log("=== FIN DE PRUEBA ===");
}

/**
 * Función de prueba para la conversión completa de una hoja específica
 * @param {string} sheetId - ID de la hoja de cálculo
 */
function probarConversionCompleta(sheetId) {
  Logger.log("=== PRUEBA DE CONVERSIÓN COMPLETA ===");
  Logger.log("Usando hoja con ID: " + sheetId);
  
  try {
    // Parámetros de prueba
    const country = "ARG";
    const languageCode = "0A"; // Español
    
    // Crear una carpeta temporal para la prueba
    const folder = DriveApp.createFolder("Prueba_STL_" + new Date().getTime());
    const folderId = folder.getId();
    
    Logger.log("Carpeta temporal creada: " + folder.getName() + " (ID: " + folderId + ")");
    
    // Realizar la conversión con modo verbose
    const verboseFlag = true;
    const result = convertSheetToSTL(sheetId, country, languageCode, folderId, verboseFlag);
    
    // Mostrar resultado
    if (result.success) {
      Logger.log("Conversión completada con éxito:");
      Logger.log("- Archivo: " + result.fileName);
      Logger.log("- URL: " + result.fileUrl);
      if (result.downloadUrl) {
        Logger.log("- URL de descarga: " + result.downloadUrl);
      }
      Logger.log("- Subtítulos: " + result.subtitleCount);
    } else {
      Logger.log("Error en la conversión: " + result.error);
    }
    
  } catch (error) {
    Logger.log("Error en la conversión: " + error.message);
  }
  
  Logger.log("=== FIN DE PRUEBA ===");
}

/**
 * Función para realizar una prueba de las correcciones
 */
function probarCorrecciones() {
  // Configuración de prueba
  const sheetId = "1YlaOCinPTChLLxDF-Ce7mIyu2UAMHqGSKcf4t0yLCM0";
  const country = "ARG";
  const languageCode = "0A";
  
  // Crear carpeta temporal para pruebas
  const folder = DriveApp.createFolder("Prueba_Corregida_STL_" + new Date().getTime());
  
  // Activar modo verbose para ver logs detallados
  setVerbose(true);
  
  // Inicia log de ejecución
  Logger.log("=== PRUEBA DE CORRECCIONES ===");
  Logger.log("Probando con configuraciones corregidas:");
  Logger.log("- SheetId: " + sheetId);
  Logger.log("- País: " + country);
  Logger.log("- Código de idioma: " + languageCode);
  Logger.log("- Carpeta: " + folder.getName() + " (ID: " + folder.getId() + ")");
  
  // Ejecutar la conversión
  const result = convertSheetToSTL(sheetId, country, languageCode, folder.getId());
  
  // Mostrar resultado
  Logger.log("\n¡CONVERSIÓN EXITOSA!");
  Logger.log("- Nombre de archivo: " + result.fileName);
  Logger.log("- URL: " + result.fileUrl);
  Logger.log("- URL de descarga: " + result.downloadUrl);
  Logger.log("- Subtítulos: " + result.subtitleCount);
  Logger.log("=== FIN DE PRUEBA DE CORRECCIONES ===");
  
  return result;
}

function test(){
probarConversionCompleta("1YlaOCinPTChLLxDF-Ce7mIyu2UAMHqGSKcf4t0yLCM0")
}

/**
 * Combina el bloque GSI y los bloques TTI en un solo archivo STL
 * @param {Uint8Array} gsiBlock - Bloque GSI
 * @param {Array<Uint8Array>} ttiBlocks - Array de bloques TTI
 * @param {boolean} verboseFlag - Activar logs detallados
 * @return {Uint8Array} - Archivo STL completo
 */
function combineBlocks(gsiBlock, ttiBlocks, verboseFlag) {
  if (verboseFlag) Logger.log(`Combinando bloques: 1 GSI + ${ttiBlocks.length} TTI`);
  
  try {
    // Actualizar el número de bloques TTI en el bloque GSI
    const ttiCountStr = ttiBlocks.length.toString().padStart(5, '0');
    writeStringToBuffer(gsiBlock, ttiCountStr, 238, 5);
    writeStringToBuffer(gsiBlock, ttiCountStr, 243, 5);
    
    // Calcular el tamaño total del archivo
    const totalSize = gsiBlock.length + (ttiBlocks.length * 128);
    
    // Crear un nuevo array para el archivo completo
    const stlFile = new Uint8Array(totalSize);
    
    // Copiar el bloque GSI
    stlFile.set(gsiBlock, 0);
    
    // Copiar cada bloque TTI
    for (let i = 0; i < ttiBlocks.length; i++) {
      stlFile.set(ttiBlocks[i], gsiBlock.length + (i * 128));
    }
    
    if (verboseFlag) Logger.log(`Archivo STL creado: ${totalSize} bytes`);
    
    return stlFile;
  } catch (error) {
    Logger.log(`Error al combinar bloques: ${error.message}`);
    throw error;
  }
}