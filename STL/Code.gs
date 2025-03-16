/**
 * Conversor de Excel a STL (EBU Subtitle Data Exchange format)
 * Creado para procesar archivos de subtítulos
 */

// Función para crear la interfaz de usuario web
function doGet(e) {
  try {
    // Registrar los parámetros recibidos completos para depuración
    Logger.log('Solicitud recibida completa: ' + JSON.stringify(e || {}));
    
    // Detectar el idioma del navegador a partir de los encabezados HTTP
    var language = 'es'; // Español por defecto
    
    // Verificar si se ha especificado un idioma en los parámetros de la URL
    if (e && e.parameter) {
      // Registrar todos los parámetros recibidos
      Object.keys(e.parameter).forEach(function(key) {
        Logger.log('Parámetro recibido: ' + key + ' = ' + e.parameter[key]);
      });
      
      // Verificar parámetro de idioma
      if (e.parameter.lang) {
        // Validar el idioma: solo permitir 'es' o 'pt-BR'
        if (e.parameter.lang === 'pt-BR') {
          language = 'pt-BR';
        }
      }
    
    // Verificar si se solicita la página de ayuda
      if (e.parameter.page === 'help') {
        // Crear la página de ayuda en el idioma seleccionado
        var helpTemplate = HtmlService.createTemplateFromFile(language === 'pt-BR' ? 'Ajuda' : 'Ayuda');
        helpTemplate.language = language;
        helpTemplate.getScriptUrl = getScriptUrl;
        
        var html = helpTemplate.evaluate()
          .setTitle(language === 'pt-BR' ? 'Ajuda - Conversor Excel para STL' : 'Ayuda - Conversor Excel a STL')
            .setSandboxMode(HtmlService.SandboxMode.IFRAME)
            .addMetaTag('viewport', 'width=device-width, initial-scale=1')
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
        
        return html;
      }
      
      // Verificar si se solicita la herramienta de diagnóstico
      if (e.parameter.page === 'diagnostic') {
        // Devolver la herramienta de diagnóstico
        return showSTLDiagnosticTool();
      }
    }
    
    // Crear la interfaz principal
      var template = HtmlService.createTemplateFromFile('Index');
      template.language = language;
      template.logoUrl = "https://mediaaccesscompany.com/wp-content/uploads/2022/12/logo_banner_home.png";
    template.getScriptUrl = getScriptUrl;
      
    var html = template.evaluate()
        .setTitle(language === 'pt-BR' ? 'Conversor Excel para STL' : 'Conversor Excel a STL')
        .setSandboxMode(HtmlService.SandboxMode.IFRAME)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      
    // Registrar finalización exitosa
    Logger.log('Interfaz de usuario generada correctamente');
    
    return html;
    } catch (error) {
    // Registrar el error
    Logger.log('Error al generar interfaz: ' + error.toString());
    
    // Devolver una página de error
    var errorHtml = HtmlService.createHtmlOutput('<h1>Error</h1><p>Ocurrió un error al cargar la aplicación: ' + error.message + '</p>')
        .setTitle('Error - Conversor Excel a STL')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
    
    return errorHtml;
  }
}

/**
 * Función auxiliar para obtener la URL del script actual
 * @return {String} URL del script
 */
function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}

/**
 * Procesa el archivo Excel cargado y lo convierte a formato STL
 * @param {Blob} excelBlob - El archivo Excel en formato Blob
 * @return {String} URL de descarga del archivo STL
 */
function processExcel(excelBlob) {
  try {
    // Crear una carpeta temporal en el Drive del usuario
    var timestamp = new Date().getTime();
    var userName = Session.getActiveUser().getEmail().split('@')[0]; // Obtener nombre de usuario
    
    // Crear una carpeta específica para este proceso
    var tempFolder = DriveApp.createFolder("STL_Temp_" + userName + "_" + timestamp);
    var tempFolderId = tempFolder.getId();
    
    // Guardar el ID de la carpeta temporal para limpieza posterior
    PropertiesService.getUserProperties().setProperty('lastTempFolderId', tempFolderId);
    
    // Guardar el archivo Excel en la carpeta temporal
    var excelFile = tempFolder.createFile(excelBlob);
    
    // Abrir el archivo como hoja de cálculo
    var spreadsheet = SpreadsheetApp.openById(excelFile.getId());
    var sheet = spreadsheet.getSheets()[0];
    
    // Obtener configuración
    var excelConfig = getConfig('excel');
    var startRow = excelConfig.startRow;
    
    // Obtener datos comenzando desde la fila configurada
    var lastRow = sheet.getLastRow();
    var data = sheet.getRange(startRow, 1, lastRow - startRow + 1, 4).getValues();
    
    // Extraer información del título
    var metadataConfig = excelConfig.metadata;
    var titleInfo = {
      englishTitle: sheet.getRange(metadataConfig.titleEnglish, 3, 1, 1).getValue() || "",
      spanishTitle: sheet.getRange(metadataConfig.titleSpanish, 3, 1, 1).getValue() || "",
      episodeNumber: sheet.getRange(metadataConfig.episodeNumber, 3, 1, 1).getValue() || ""
    };
    
    // Generar el archivo STL
    var stlContent = generateSTL(data, titleInfo);
    
    // Generar nombre para el archivo STL
    var fileName = (titleInfo.spanishTitle.replace(/[^\w\s]/gi, '') || 'subtitulos') + '.stl';
    
    // Crear el archivo STL en la carpeta temporal
    var stlBlob = Utilities.newBlob(stlContent, 'application/octet-stream', fileName);
    var stlFile = tempFolder.createFile(stlBlob);
    
    // Configurar para compartir - simplificado para evitar problemas de serialización
    stlFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    // IMPORTANTE: Solo devolvemos un string, no objetos complejos
    return "https://drive.google.com/uc?export=download&id=" + stlFile.getId();
    
  } catch (error) {
    Logger.log('Error: ' + error.toString() + '\n' + error.stack);
    throw new Error('Error al procesar: ' + error.message);
  }
}

/**
 * Procesa el contenido de un archivo Excel en Base64
 * @param {String} base64Content - Contenido del archivo en Base64
 * @param {String} fileName - Nombre del archivo original
 * @return {Object|String} Resultado con URL de descarga o mensaje de error
 */
function processExcelBase64(base64Content, fileName) {
  try {
    Logger.log("Procesando archivo: " + fileName);
    var result = { success: false };
    
    // Verificar formato del archivo
    if (!fileName || (!fileName.endsWith('.xlsx') && !fileName.endsWith('.xls'))) {
      throw new Error('El archivo debe ser un archivo Excel (.xls o .xlsx)');
    }
    
    // Convertir base64 a Blob
    var blob = Utilities.newBlob(Utilities.base64Decode(base64Content), 'application/octet-stream', fileName);
    var stlFileName = fileName.replace(/\.xlsx?$/, '.stl');
    
    // Importar datos del Excel (convertido a Google Sheets)
    var excelData = importExcel(blob);
    
    // Verificar que tenemos datos válidos
    if (excelData && excelData.subtitles && excelData.subtitles.length > 0) {
      Logger.log("Datos obtenidos: " + excelData.subtitles.length + " subtítulos");
      
      // Generar STL usando los subtítulos extraídos del Google Sheets
      var stlBlob = generateExactReferenceSTL(excelData.subtitles, excelData.metadata);
      
      // Guardar archivo STL en Drive temporalmente
      var tempFolder = getOrCreateTempFolder();
      var stlFile = tempFolder.createFile(stlBlob);
      stlFile.setName(stlFileName);
      
      // Hacer archivo accesible por URL
      stlFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      result = {
        success: true,
        downloadUrl: stlFile.getDownloadUrl(),
        fileName: stlFileName
      };
      
      Logger.log("STL generado correctamente: " + stlFileName);
    } else {
      throw new Error('No se pudieron extraer datos del archivo Excel');
    }
    
    // Retornar información de resultado
    return result;
  } catch (error) {
    Logger.log('Error al procesar Excel: ' + error.toString());
    throw new Error('Error al procesar Excel: ' + error.message);
  }
}

/**
 * Genera un archivo en formato STL (EBU Subtitle Data Exchange Format) según la especificación EBU Tech 3264
 * @param {Array} data - Array de subtítulos con timecodes y texto
 * @param {Object} metadata - Metadatos del programa
 * @return {Blob} Contenido en formato STL como un Blob binario
 */
function generateSTL(data, metadata) {
  try {
    Logger.log("Generando archivo STL con " + data.length + " subtítulos");
    // Crear los bytes del archivo STL
    var stlBytes = [];
    
    // Configuración STL
    var stlConfig = getConfig('stl');
    
    // Crear el encabezado GSI (1024 bytes)
    var gsiBytes = createSTLHeader(metadata);
    
    // Contar subtítulos
    var subtitleCount = 0;
    var totalCharacters = 0;
    var columnConfig = getConfig('excel').columns;
    var maxSubtitles = getConfig('limits').maxSubtitles;
    
    // Crear array para almacenar bloques TTI
    var ttiBlocks = [];
    
    // Crear primer bloque TTI con metadatos especiales
    var metadataTTI = createMetadataTTI(metadata);
    ttiBlocks.push(metadataTTI);
    subtitleCount++;
    
    // Crear bloques TTI para cada subtítulo
    for (var i = 0; i < data.length && subtitleCount < maxSubtitles; i++) {
      // Verificar formato de los datos
      var timecodeIn, dialogueLine, timecodeOut;
      
      // Si los datos tienen 4 columnas (formato normal desde Excel)
      if (data[i].length >= 4) {
        timecodeIn = data[i][columnConfig.timeIn];
        dialogueLine = data[i][columnConfig.text];
        timecodeOut = data[i][columnConfig.timeOut];
      } 
      // Si los datos tienen 3 columnas (formato alternativo o datos de muestra)
      else if (data[i].length >= 3) {
        timecodeIn = data[i][0];
        dialogueLine = data[i][1];
        timecodeOut = data[i][2];
      }
      else {
        // Saltar esta fila si no tiene suficientes datos
        continue;
      }
      
      // Verificar si los datos son válidos
      if (timecodeIn && timecodeOut && dialogueLine) {
        subtitleCount++;
        // Determinar justificación (centrado por defecto)
        var justification = stlConfig.justification.center; // 2=centro
        
        // Calcular número de caracteres para estadísticas
        totalCharacters += dialogueLine.length;
        
        var ttiBlock = createTTIBlock(subtitleCount, timecodeIn, timecodeOut, dialogueLine, justification);
        ttiBlocks.push(ttiBlock);
      }
    }
    
    // Actualizar campos en el encabezado GSI relacionados con estadísticas
    // TNS: Total Number of Subtitle Groups (bytes 194-198)
    writeToGSI(gsiBytes, 194, padNumber(1, 5)); // Siempre 1 grupo de subtítulos
    
    // TNB: Total Number of TTI blocks (bytes 199-203)
    writeToGSI(gsiBytes, 199, padNumber(subtitleCount, 5));
    
    // TCS: Total Character Count (bytes 204-208)
    writeToGSI(gsiBytes, 204, padNumber(totalCharacters, 5));
    
    // TND: Total Number of Disks (bytes 209-210)
    writeToGSI(gsiBytes, 209, "01");
    
    // DSN: Disk Sequence Number (bytes 211-212)
    writeToGSI(gsiBytes, 211, "01");
    
    // CO: Country of Origin (bytes 213-214)
    var countryCode = "ES"; // Default
    if (metadata && metadata.language) {
      countryCode = stlConfig.countryCode[metadata.language] || "ES";
    }
    writeToGSI(gsiBytes, 213, countryCode);
    
    // PUB: Publisher (bytes 215-230)
    writeToGSI(gsiBytes, 215, "MEDIA ACCESS CO.");
    
    // EN: Editor Name (bytes 231-262)
    writeToGSI(gsiBytes, 231, "EXCEL STL CONVERTER");
    
    // ECD: Editor Contact Details (bytes 263-294)
    writeToGSI(gsiBytes, 263, "appsheet@mediaaccesscompany.com");
    
    // Agregar los bytes del GSI al array de bytes STL
    stlBytes = gsiBytes;
    
    // Agregar cada bloque TTI
    for (var j = 0; j < ttiBlocks.length; j++) {
      stlBytes = stlBytes.concat(ttiBlocks[j]);
    }
    
    // Convertir el array de bytes a un Blob binario
    var byteArray = new Uint8Array(stlBytes);
    var blob = Utilities.newBlob(byteArray, 'application/octet-stream', 'subtitles.stl');
    
    Logger.log("Archivo STL generado correctamente con " + subtitleCount + " subtítulos y " + totalCharacters + " caracteres");
    return blob;
  } catch (error) {
    Logger.log('Error al generar STL: ' + error.toString());
    throw new Error('Error al generar archivo STL: ' + error.message);
  }
}

/**
 * Escribe texto en un buffer GSI en la posición especificada
 * @param {Array} buffer - Array de bytes del GSI
 * @param {Number} offset - Posición donde escribir
 * @param {String} text - Texto a escribir
 */
function writeToGSI(buffer, offset, text) {
  for (var i = 0; i < text.length; i++) {
    buffer[offset + i] = text.charCodeAt(i);
  }
}

/**
 * Crea el encabezado GSI para el archivo STL según la especificación EBU Tech 3264
 * @param {Object} metadata - Metadatos del programa
 * @return {Array} Array de bytes del encabezado GSI (1024 bytes)
 */
function createSTLHeader(metadata) {
  // Configuración STL
  var stlConfig = getConfig('stl');
  
  // Crear un buffer de 1024 bytes rellenado con espacios (ASCII 32)
  var gsiBytes = new Array(stlConfig.gsiSize);
  for (var i = 0; i < stlConfig.gsiSize; i++) {
    gsiBytes[i] = 32; // Espacio (ASCII 32)
  }
  
  // Obtener fecha actual
  var now = new Date();
  
  // CPN: Code Page Number (bytes 0-2) - "850" para Latin-1
  writeToGSI(gsiBytes, 0, stlConfig.codePage);
  
  // DFC: Disk Format Code (bytes 3-10) - "STL25.01" para PAL (25fps)
  writeToGSI(gsiBytes, 3, stlConfig.format);
  
  // DSC: Display Standard Code (bytes 11-13) - Normalmente "1" para nivel 1 teletext
  writeToGSI(gsiBytes, 11, "1");
  
  // CCT: Character Code Table (bytes 14-15) - "00" para Latin
  writeToGSI(gsiBytes, 14, stlConfig.charTable);
  
  // LC: Language Code (bytes 16-17)
  var langCode = stlConfig.ebuLanguageCode['es']; // Valor por defecto
  if (metadata && metadata.language && stlConfig.ebuLanguageCode[metadata.language]) {
    langCode = stlConfig.ebuLanguageCode[metadata.language];
  }
  writeToGSI(gsiBytes, 16, langCode);
  
  // OPT: Original Programme Title (bytes 18-49)
  var originalTitle = "";
  if (metadata && metadata.englishTitle) {
    originalTitle = metadata.englishTitle.substring(0, 32);
  }
  writeToGSI(gsiBytes, 18, originalTitle);
  
  // OEP: Original Episode Title (bytes 50-81)
  var episodeCode = "";
  if (metadata && metadata.episodeNumber) {
    episodeCode = metadata.episodeNumber.substring(0, 32);
  }
  writeToGSI(gsiBytes, 50, episodeCode);
  
  // TPT: Translated Programme Title (bytes 82-113)
  var translatedTitle = "";
  if (metadata && metadata.spanishTitle) {
    translatedTitle = metadata.spanishTitle.substring(0, 32);
  }
  writeToGSI(gsiBytes, 82, translatedTitle);
  
  // TET: Translated Episode Title (bytes 114-145)
  // Usar el mismo código de episodio
  writeToGSI(gsiBytes, 114, episodeCode);
  
  // TN: Translator Name (bytes 146-177)
  writeToGSI(gsiBytes, 146, "MEDIA ACCESS");
  
  // TCD: Translator Contact Details (bytes 178-209)
  writeToGSI(gsiBytes, 178, "appsheet@mediaaccesscompany.com");
  
  // SLR: Subtitle List Reference (bytes 210-224)
  // Usar título + episodio como referencia
  var slr = (translatedTitle + "_" + episodeCode).substring(0, 15);
  writeToGSI(gsiBytes, 210, slr);
  
  // CD: Creation Date (bytes 225-232) - Format: YYYYMMDD
  var dateStr = formatDateYYYYMMDD(now);
  writeToGSI(gsiBytes, 225, dateStr);
  
  // RD: Revision Date (bytes 233-240) - Format: YYYYMMDD
  writeToGSI(gsiBytes, 233, dateStr);
  
  // RN: Revision Number (bytes 241-242)
  writeToGSI(gsiBytes, 241, "01");
  
  // TNB, TCF, TND, DSN se actualizarán después al conocer el número total de subtítulos
  
  // CO: Country of Origin (bytes 274-275)
  var countryCode = "ES"; // Default
  if (metadata && metadata.language) {
    countryCode = stlConfig.countryCode[metadata.language] || "ES";
  }
  writeToGSI(gsiBytes, 274, countryCode);
  
  // PUB: Publisher (bytes 276-291)
  writeToGSI(gsiBytes, 276, "MEDIA ACCESS CO.");
  
  // EN: Editor Name (bytes 292-323)
  writeToGSI(gsiBytes, 292, "EXCEL STL CONVERTER");
  
  // ECD: Editor Contact Details (bytes 324-355)
  writeToGSI(gsiBytes, 324, "appsheet@mediaaccesscompany.com");
  
  return gsiBytes;
}

/**
 * Crea un bloque TTI de metadatos especiales (primer bloque)
 * @param {Object} metadata - Metadatos del programa
 * @return {Array} Array de bytes del bloque TTI (128 bytes)
 */
function createMetadataTTI(metadata) {
  // Crear un buffer de 128 bytes rellenado con 0xFF
  var ttiBytes = new Array(128);
  for (var i = 0; i < 128; i++) {
    ttiBytes[i] = 0xFF;
  }
  
  // SGN: Subtitle Group Number (bytes 0-1) - Siempre 0
  ttiBytes[0] = 0;
  ttiBytes[1] = 0;
  
  // SN: Subtitle Number (bytes 2-3) - Little endian (1)
  ttiBytes[2] = 1;
  ttiBytes[3] = 0;
  
  // EBN: Extension Block Number (byte 4) - 0xFF para no usar
  ttiBytes[4] = 0xFF;
  
  // CS: Cumulative Status (byte 5) - 0 para no acumulativo
  ttiBytes[5] = 0;
  
  // Timecode IN (bytes 6-9) - Zeros para metadata
  for (var i = 6; i <= 9; i++) {
    ttiBytes[i] = 0;
  }
  
  // Timecode OUT (bytes 10-13) - Zeros para metadata
  for (var i = 10; i <= 13; i++) {
    ttiBytes[i] = 0;
  }
  
  // VP: Vertical Position (byte 14) - 0 para metadatos
  ttiBytes[14] = 0;
  
  // JC: Justification Code (byte 15) - 0 para metadatos
  ttiBytes[15] = 0;
  
  // CF: Comment Flag (byte 16) - 1 para comentario
  ttiBytes[16] = 1;
  
  // TF: Text Field (bytes 17-127)
  // Formatear texto con información completa del programa
  var metadataText = "";
  
  // Título y episodio
  if (metadata) {
    var title = metadata.spanishTitle || "";
    var episode = metadata.episodeNumber || "";
    var engTitle = metadata.englishTitle || "";
    
    metadataText = "STORY:" + title;
    if (episode) {
      metadataText += " " + episode;
    }
    
    // Añadir idioma
    var langCode = metadata.language === 'pt' ? "POR" : "SPA";
    metadataText += "\u008ALANG:" + langCode;
    
    // Añadir título original
    if (engTitle) {
      metadataText += "\u008AORIGIN:" + engTitle;
    }
  }
  
  // Escribir el texto al buffer usando la codificación EBU
  writeTextWithEBUEncoding(ttiBytes, 17, metadataText);
  
  return ttiBytes;
}

/**
 * Escribe texto en un buffer TTI usando la codificación EBU
 * @param {Array} buffer - Array de bytes del bloque TTI
 * @param {Number} offset - Posición inicial donde escribir
 * @param {String} text - Texto a escribir
 * @return {Number} Posición final después de escribir
 */
function writeTextWithEBUEncoding(buffer, offset, text) {
  var pos = offset;
  
  // Confirmar que tenemos espacio suficiente (máximo 112 bytes de texto)
  var maxLength = Math.min(text.length, 112);
  
  for (var i = 0; i < maxLength; i++) {
    var char = text.charAt(i);
    
    // Manejar saltos de línea
    if (char === '\n' || char === '\u008A') {
      buffer[pos++] = 0x8A; // Código EOL (End Of Line) para EBU
      continue;
    }
    
    // Manejar caracteres ASCII estándar
    var charCode = char.charCodeAt(0);
    if (charCode >= 0x20 && charCode <= 0x7E) {
      // Caracteres imprimibles ASCII
      buffer[pos++] = charCode;
      continue;
    }
    
    // Manejar caracteres especiales según la tabla EBU
    var ebuCode = getEBUCodeForChar(char);
    buffer[pos++] = ebuCode;
    
    // Verificar si nos acercamos al límite (dejando espacio para ETX)
    if (pos >= offset + 111) {
      Logger.log("Advertencia: Texto truncado por exceder el límite de 112 bytes");
      break;
    }
  }
  
  // Marcar el final del texto con ETX (End of Text, 0x03)
  // Este es el estándar correcto para EBU STL
  buffer[pos++] = 0x03;
  
  // Rellenar el resto con 0xFF (espacio no usado)
  // Este es un estándar en editores profesionales
  for (var i = pos; i < offset + 112; i++) {
    buffer[i] = 0xFF;
  }
  
  return pos;
}

/**
 * Obtiene el código EBU correspondiente para un carácter especial
 * Mapeo basado en la página de código Latin del estándar EBU
 * @param {String} char - Carácter a convertir
 * @return {Number} Código EBU correspondiente
 */
function getEBUCodeForChar(char) {
  // Mapa de caracteres latinos a códigos EBU según la norma EBU Tech 3264-E
  const charMap = {
    // Vocales acentuadas minúsculas
    'á': 0xE1, 'é': 0xE9, 'í': 0xED, 'ó': 0xF3, 'ú': 0xFA,
    'à': 0xE0, 'è': 0xE8, 'ì': 0xEC, 'ò': 0xF2, 'ù': 0xF9,
    'ä': 0xE4, 'ë': 0xEB, 'ï': 0xEF, 'ö': 0xF6, 'ü': 0xFC,
    'â': 0xE2, 'ê': 0xEA, 'î': 0xEE, 'ô': 0xF4, 'û': 0xFB,
    'ã': 0xE3, 'õ': 0xF5, 'ñ': 0xF1, 'ç': 0xE7,
    
    // Vocales acentuadas mayúsculas
    'Á': 0xC1, 'É': 0xC9, 'Í': 0xCD, 'Ó': 0xD3, 'Ú': 0xDA,
    'À': 0xC0, 'È': 0xC8, 'Ì': 0xCC, 'Ò': 0xD2, 'Ù': 0xD9,
    'Ä': 0xC4, 'Ë': 0xCB, 'Ï': 0xCF, 'Ö': 0xD6, 'Ü': 0xDC,
    'Â': 0xC2, 'Ê': 0xCA, 'Î': 0xCE, 'Ô': 0xD4, 'Û': 0xDB,
    'Ã': 0xC3, 'Õ': 0xD5, 'Ñ': 0xD1, 'Ç': 0xC7,
    
    // Símbolos especiales y puntuación
    '¿': 0xBF, '¡': 0xA1, '°': 0xB0, '€': 0xA4, '£': 0xA3, '¥': 0xA5,
    '§': 0xA7, '÷': 0xF7, '×': 0xD7, '¼': 0xBC, '½': 0xBD, '¾': 0xBE,
    '²': 0xB2, '³': 0xB3, '±': 0xB1, 'µ': 0xB5, '¶': 0xB6, '·': 0xB7,
    '¢': 0xA2, '¦': 0xA6, '©': 0xA9, '®': 0xAE, '™': 0x54, // ™ no está en EBU, usar T
    
    // Comillas y guiones específicos del formato EBU
    '"': 0x22, '"': 0x22, '"': 0x22, // Comillas dobles
    "'": 0x27, "'": 0x27, '´': 0x27, // Comillas simples/apostrofes
    '«': 0xAB, '»': 0xBB, // Comillas angulares
    '–': 0x2D, '—': 0x2D, '-': 0x2D, // Guiones
    
    // Caracteres especiales de control
    '\n': 0x8A, // Salto de línea (EOL)
    
    // Otros caracteres comunes en español/portugués
    'ª': 0xAA, 'º': 0xBA,
    
    // Caracteres específicos del portugués
    'ũ': 0xFB, 'Ũ': 0xDB, // u con tilde (aproximación, usar circunflejo)
    'ẽ': 0xEA, 'Ẽ': 0xCA, // e con tilde (aproximación, usar circunflejo)
    'ĩ': 0xEE, 'Ĩ': 0xCE, // i con tilde (aproximación, usar circunflejo)
    
    // Símbolos matemáticos y otros que podrían usarse
    '<': 0x3C, '>': 0x3E, '|': 0x7C, '\\': 0x5C, '^': 0x5E, '~': 0x7E,
    '{': 0x7B, '}': 0x7D, '[': 0x5B, ']': 0x5D, '`': 0x60
  };
  
  // Verificar si el carácter está en el mapa
  if (charMap[char] !== undefined) {
    return charMap[char];
  }
  
  // Obtener el código de carácter Unicode
  const charCode = char.charCodeAt(0);
  
  // Si es un carácter ASCII básico, usar ese código
  if (charCode >= 32 && charCode <= 126) {
    return charCode;
  }
  
  // Para caracteres no reconocidos, reemplazar con un espacio
  Logger.log('Carácter no soportado en STL: ' + char + ' (Unicode: ' + charCode + ')');
  return 32; // Espacio (ASCII 32)
}

/**
 * Crea un bloque TTI para un subtítulo
 * @param {Number} subtitleNumber - Número de subtítulo
 * @param {String} timecodeIn - Código de tiempo de entrada
 * @param {String} timecodeOut - Código de tiempo de salida
 * @param {String} text - Texto del subtítulo
 * @param {Number} justification - Código de justificación (1=left, 2=center, 3=right)
 * @return {Array} Array de bytes del bloque TTI
 */
function createTTIBlock(subtitleNumber, timecodeIn, timecodeOut, text, justification) {
  Logger.log('Creando bloque TTI para subtítulo: ' + subtitleNumber);
  Logger.log('Texto del subtítulo: ' + text);
  // Crear un buffer de 128 bytes rellenado con 0xFF
  var ttiBytes = new Array(128);
  for (var i = 0; i < 128; i++) {
    ttiBytes[i] = 0xFF;
  }
  
  // SGN: Subtitle Group Number (bytes 0-1) - Siempre 0
  ttiBytes[0] = 0;
  ttiBytes[1] = 0;
  
  // SN: Subtitle Number (bytes 2-3) - Little endian
  ttiBytes[2] = subtitleNumber & 0xFF;
  ttiBytes[3] = (subtitleNumber >> 8) & 0xFF;
  
  // EBN: Extension Block Number (byte 4) - 0xFF para no usar
  ttiBytes[4] = 0xFF;
  
  // CS: Cumulative Status (byte 5) - 0 para no acumulativo
  ttiBytes[5] = 0;
  
  // Timecode IN (bytes 6-9)
  var tcIn = parseTimecode(timecodeIn);
  ttiBytes[6] = tcIn.hours;
  ttiBytes[7] = tcIn.minutes;
  ttiBytes[8] = tcIn.seconds;
  ttiBytes[9] = tcIn.frames;
  
  // Timecode OUT (bytes 10-13)
  var tcOut = parseTimecode(timecodeOut);
  ttiBytes[10] = tcOut.hours;
  ttiBytes[11] = tcOut.minutes;
  ttiBytes[12] = tcOut.seconds;
  ttiBytes[13] = tcOut.frames;
  
  // VP: Vertical Position (byte 14)
  // Calcular posición vertical según el número de líneas
  var lines = text.split('\n').length;
  ttiBytes[14] = calculateVerticalPosition(lines);
  
  // JC: Justification Code (byte 15)
  ttiBytes[15] = justification || 2; // 2=centro por defecto
  
  // CF: Comment Flag (byte 16) - 0 para subtítulo normal
  ttiBytes[16] = 0;
  
  // TF: Text Field (bytes 17-127)
  // Escribir el texto con codificación EBU
  writeTextWithEBUEncoding(ttiBytes, 17, text);
  
  // Asegurar que el texto termine con ETX (0x03) y rellenar con 0xFF
  var position = 17;
  for (var i = 17; i < 127; i++) {
    if (ttiBytes[i] === 0x03) {
      position = i + 1;
      break;
    }
  }
  if (position < 128) {
    ttiBytes[position++] = 0x03;
    for (var i = position; i < 128; i++) {
      ttiBytes[i] = 0xFF;
    }
  }
  
  return ttiBytes;
}

/**
 * Calcula la posición vertical óptima para un subtítulo según el número de líneas
 * @param {Number} lines - Número de líneas del subtítulo
 * @return {Number} Posición vertical según la especificación EBU
 */
function calculateVerticalPosition(lines) {
  var stlConfig = getConfig('stl');
  
  // Para PAL/625 líneas, valores típicos según configuración
  if (lines === 1) {
    return stlConfig.verticalPosition.singleLine; // Una línea centrada
  } else {
    return stlConfig.verticalPosition.twoLinesFirst; // Primera línea de un subtítulo de dos líneas
  }
}

/**
 * Analiza un código de tiempo en formato 'HH:MM:SS:FF' a componentes
 * @param {String} timecode - Código de tiempo a analizar
 * @return {Object} Objeto con componentes del timecode
 */
function parseTimecode(timecode) {
  // Objeto resultado por defecto
  var result = {
    hours: 0,
    minutes: 0,
    seconds: 0,
    frames: 0
  };
  
  try {
    // Validar el tipo de entrada
    if (typeof timecode !== 'string' || !timecode) {
      Logger.log('Warning: Timecode inválido, usando 00:00:00:00');
      return result;
    }
    
    // Normalizamos el formato (soportamos HH:MM:SS:FF, HH:MM:SS.FF, HH:MM:SS;FF)
    var normalizedTC = timecode.replace(/[\.;]/g, ':');
    
    // Diferentes formatos posibles
    var regex = {
      fourParts: /^(\d{1,2}):(\d{1,2}):(\d{1,2}):(\d{1,2})$/, // HH:MM:SS:FF
      threeParts: /^(\d{1,2}):(\d{1,2}):(\d{1,2})$/, // MM:SS:FF (sin horas)
      dropFrame: /^(\d{1,2}):(\d{1,2}):(\d{1,2});(\d{1,2})$/ // HH:MM:SS;FF (formato drop frame)
    };
    
    // Intentar extraer componentes
    var match = normalizedTC.match(regex.fourParts);
    
    if (match) {
      // Formato completo HH:MM:SS:FF
      result.hours = parseInt(match[1], 10);
      result.minutes = parseInt(match[2], 10);
      result.seconds = parseInt(match[3], 10);
      result.frames = parseInt(match[4], 10);
    } else if (match = normalizedTC.match(regex.threeParts)) {
      // Formato MM:SS:FF (sin horas)
      result.minutes = parseInt(match[1], 10);
      result.seconds = parseInt(match[2], 10);
      result.frames = parseInt(match[3], 10);
  } else {
      // Formato no reconocido, usar valores predeterminados
      Logger.log('Warning: Formato de timecode no reconocido: ' + timecode);
    }
    
    // Validar límites (en formato STL PAL 25fps)
    if (result.hours > 23) result.hours = 23;
    if (result.minutes > 59) result.minutes = 59;
    if (result.seconds > 59) result.seconds = 59;
    if (result.frames > 24) result.frames = 24; // 0-24 para 25fps
    
  } catch (error) {
    Logger.log('Error al parsear timecode: ' + error.message);
  }
  
  return result;
}

/**
 * Formatea una fecha como YYYYMMDD
 * @param {Date} date - Fecha a formatear
 * @return {String} Fecha formateada
 */
function formatDateYYYYMMDD(date) {
  return date.getFullYear().toString() + 
         padNumber(date.getMonth() + 1, 2) + 
         padNumber(date.getDate(), 2);
}

/**
 * Rellena un número con ceros a la izquierda
 * @param {Number} num - Número a rellenar
 * @param {Number} size - Tamaño deseado
 * @return {String} Número rellenado con ceros
 */
function padNumber(num, size) {
  var s = num.toString();
  while (s.length < size) s = "0" + s;
  return s;
}

/**
 * Sanitiza un nombre de archivo para evitar caracteres no permitidos
 * @param {String} fileName - Nombre de archivo a sanitizar
 * @return {String} Nombre de archivo sanitizado
 */
function sanitizeFileName(fileName) {
  if (!fileName) return "";
  
  // Reemplazar caracteres no permitidos en nombres de archivo
  return fileName.replace(/[\\/:*?"<>|&%#=+,.;$^()[\]{}~`@]/g, '_')
                .replace(/\s+/g, '_')  // Reemplazar espacios con guiones bajos
                .trim();  // Eliminar espacios en blanco al inicio y final
}

/**
 * Rellena una cadena con el carácter especificado a la derecha
 * @param {String} text - Texto a rellenar
 * @param {Number} length - Longitud deseada
 * @param {String} char - Carácter de relleno
 * @return {String} Texto rellenado
 */
function padRight(text, length, char) {
  return (text + char.repeat(length)).substr(0, length);
}

/**
 * Limpia las carpetas temporales para liberar espacio en Drive
 * Esta función se puede ejecutar manualmente desde el editor de Apps Script
 * o automáticamente mediante un trigger
 */
function cleanupTempFolders() {
  try {
    // Limpiar la última carpeta temporal del usuario actual (si existe)
    var folderId = PropertiesService.getUserProperties().getProperty('lastTempFolderId');
    if (folderId) {
      try {
        DriveApp.getFolderById(folderId).setTrashed(true);
        PropertiesService.getUserProperties().deleteProperty('lastTempFolderId');
        Logger.log('Carpeta temporal personal eliminada: ' + folderId);
      } catch (err) {
        Logger.log('Error al eliminar carpeta personal: ' + err.toString());
      }
    }
    
    // Limpiar carpetas temporales antiguas en el Drive del usuario (mayores a 24 horas)
    var folders = DriveApp.getRootFolder().getFolders();
    var oneDayAgo = new Date();
    oneDayAgo.setDate(oneDayAgo.getDate() - 1);
    
    var countDeleted = 0;
    while (folders.hasNext()) {
      var folder = folders.next();
      // Solo limpiar las carpetas que tienen el prefijo STL_Temp_
      if (folder.getName().startsWith("STL_Temp_") && 
          folder.getDateCreated() < oneDayAgo) {
        folder.setTrashed(true);
        countDeleted++;
      }
    }
    
    Logger.log('Limpieza completada. Carpetas eliminadas: ' + countDeleted);
    return "Limpieza completada con éxito. Se eliminaron " + countDeleted + " carpetas temporales antiguas.";
  } catch (e) {
    Logger.log('Error al limpiar carpetas temporales: ' + e.toString());
    return "Error al limpiar: " + e.toString();
  }
}

/**
 * Configura un trigger para ejecutar la limpieza de carpetas temporales 
 * automáticamente todas las noches a la 1:00 AM
 * Esta función debe ejecutarse manualmente una sola vez
 * por un administrador para configurar el trigger
 */
function setupNightlyCleanupTrigger() {
  try {
    // Eliminar triggers existentes para evitar duplicados
    var triggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === 'cleanupTempFolders') {
        ScriptApp.deleteTrigger(triggers[i]);
      }
    }
    
    // Crear un nuevo trigger que se ejecute todos los días a la 1:00 AM
    ScriptApp.newTrigger('cleanupTempFolders')
      .timeBased()
      .atHour(1)
      .everyDays(1)
      .create();
    
    return "Trigger de limpieza nocturna configurado exitosamente. Se ejecutará todos los días a la 1:00 AM.";
  } catch (e) {
    Logger.log('Error al configurar el trigger: ' + e.toString());
    return "Error al configurar el trigger: " + e.toString();
  }
}

/**
 * Permite a un usuario limpiar sus propios archivos temporales
 * Esta función puede ser llamada desde la interfaz de usuario
 */
function cleanupMyTempFiles() {
  try {
    var userEmail = Session.getActiveUser().getEmail();
    var userPrefix = userEmail.split('@')[0];
    
    // Detectar idioma del usuario
    var language = 'es'; // Español por defecto
    try {
      var userLocale = Session.getActiveUserLocale();
      if (userLocale && (userLocale.startsWith('pt') || userLocale === 'pt-BR')) {
        language = 'pt-BR';
      }
    } catch(err) {
      // Si hay error, mantener el español
      Logger.log('Error al detectar el idioma: ' + err.toString());
    }
    
    // Buscar en el root del Drive del usuario las carpetas que comienzan con STL_Temp_
    var folders = DriveApp.getRootFolder().getFolders();
    
    var countDeleted = 0;
    var totalFiles = 0;
    while (folders.hasNext()) {
      var folder = folders.next();
      // Solo limpiar las carpetas que pertenecen al usuario actual
      if (folder.getName().startsWith("STL_Temp_" + userPrefix + "_")) {
        // Contar el número de archivos en la carpeta
        var files = folder.getFiles();
        var fileCount = 0;
        while (files.hasNext()) {
          files.next();
          fileCount++;
        }
        totalFiles += fileCount;
        
        // Eliminar la carpeta y su contenido
        folder.setTrashed(true);
        countDeleted++;
      }
    }
    
    // Limpiar también la última carpeta temporal del usuario (si existe)
    var folderId = PropertiesService.getUserProperties().getProperty('lastTempFolderId');
    if (folderId) {
      try {
        var lastFolder = DriveApp.getFolderById(folderId);
        
        // Contar archivos en la última carpeta
        var lastFiles = lastFolder.getFiles();
        var lastFolderFileCount = 0;
        while (lastFiles.hasNext()) {
          lastFiles.next();
          lastFolderFileCount++;
        }
        totalFiles += lastFolderFileCount;
        
        // Eliminar la carpeta
        lastFolder.setTrashed(true);
        countDeleted++;
        
        PropertiesService.getUserProperties().deleteProperty('lastTempFolderId');
      } catch (err) {
        Logger.log('Error al eliminar última carpeta personal: ' + err.toString());
      }
    }
    
    Logger.log('Limpieza personal completada. Carpetas eliminadas: ' + countDeleted + ', Archivos eliminados: ' + totalFiles);
    
    // Mensajes según el idioma
    if (language === 'pt-BR') {
      // Mensaje en portugués
      if (countDeleted === 0) {
        return "Não foram encontradas pastas temporárias para excluir.";
      } else if (countDeleted === 1) {
        return "1 pasta temporária com " + totalFiles + " arquivo(s) foi excluída.";
      } else {
        return countDeleted + " pastas temporárias com um total de " + totalFiles + " arquivo(s) foram excluídas.";
      }
    } else {
      // Mensaje en español (predeterminado)
      if (countDeleted === 0) {
        return "No se encontraron carpetas temporales para eliminar.";
      } else if (countDeleted === 1) {
        return "Se ha eliminado 1 carpeta temporal con " + totalFiles + " archivo(s).";
      } else {
        return "Se han eliminado " + countDeleted + " carpetas temporales con un total de " + totalFiles + " archivo(s).";
      }
    }
  } catch (e) {
    Logger.log('Error al limpiar carpetas personales: ' + e.toString());
    
    // Mensaje de error según el idioma
    if (language === 'pt-BR') {
      return "Erro ao limpar suas pastas: " + e.message;
    } else {
      return "Error al limpiar tus carpetas: " + e.message;
    }
  }
}

/**
 * Crea un menú personalizado en la interfaz de Google Sheets al abrir el archivo
 * Esta función se ejecuta automáticamente cuando un usuario abre la hoja de cálculo
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  
  // Crear el menú principal
  ui.createMenu('STL - Administración')
    .addItem('Configurar limpieza automática nocturna', 'setupNightlyCleanupTrigger')
    .addItem('Limpiar todas las carpetas temporales antiguas', 'cleanupTempFolders')
    .addItem('Limpiar mis archivos temporales', 'cleanupMyTempFiles')
    .addSeparator()
    .addItem('Acerca de la aplicación', 'showAboutDialog')
    .addToUi();
}

/**
 * Muestra un diálogo con información sobre la aplicación
 */
function showAboutDialog() {
  var ui = SpreadsheetApp.getUi();
  var message = 
      "Conversor Excel a STL\n" +
      "Versión: 1.0.0\n\n" +
      "Desarrollado para Media Access Co.\n" +
      "La aplicación permite convertir archivos Excel con formato de subtítulos al formato STL.\n\n" +
      "Características:\n" +
      "- Conversión de Excel a STL\n" +
      "- Soporte para español y portugués\n" +
      "- Gestión automática de archivos temporales\n\n" +
      "© " + new Date().getFullYear() + " - Media Access Co.";
  
  ui.alert("Acerca de la aplicación", message, ui.ButtonSet.OK);
}

/**
 * Gestiona la descarga de un archivo STL
 * @param {Array} data - Matriz de datos extraídos del Excel
 * @param {Object} metadata - Metadatos del programa
 * @return {Object} Resultado de la operación
 */
function downloadSTL(data, metadata) {
  try {
    // Crear una carpeta temporal para este usuario y proceso
    var userEmail = Session.getActiveUser().getEmail();
    var userName = userEmail.split('@')[0];
    var timestamp = new Date().getTime();
    var folderName = "STL_Temp_" + userName + "_" + timestamp;
    var tempFolder = DriveApp.createFolder(folderName);
    
    // Guardar el ID de la carpeta temporal más reciente para este usuario
    var userProperties = PropertiesService.getUserProperties();
    userProperties.setProperty('lastTempFolderName', folderName);
    userProperties.setProperty('lastTempFolderId', tempFolder.getId());
    
    // Si hay un nombre original en los metadatos, usarlo como base para el nombre STL
    var fileName;
    if (metadata.originalFileName) {
      // Usar el nombre original con extensión .stl
      fileName = metadata.originalFileName.replace(/\.(xls|xlsx|csv)$/i, '');
    } else {
      // Generar nombre de archivo basado en metadatos
      fileName = metadata.programCode;
      if (metadata.spanishTitle) {
        fileName = sanitizeFileName(metadata.spanishTitle);
      }
      fileName = fileName.toUpperCase();
      
      // Agregar sufijos según idioma y tipo
      if (metadata.language === 'es') {
        fileName += '_SPA_SUB_FN';
      } else if (metadata.language === 'pt') {
        fileName += '_POR_SUB_FN';
      }
      
      // Agregar el título del episodio si existe
      if (metadata.episodeTitle) {
        fileName += '_' + sanitizeFileName(metadata.episodeTitle);
      }
      
      // Agregar el número de episodio si existe
      if (metadata.episodeNumber) {
        fileName += '_Ep' + metadata.episodeNumber;
      }
    }
    
    // Asegurar que el nombre de archivo es único
    var existingFiles = tempFolder.getFilesByName(fileName + '.stl');
    if (existingFiles.hasNext()) {
      // Agregar un sufijo con la versión
      var version = 2;
      var versionedName = fileName + '_V' + version;
      while (tempFolder.getFilesByName(versionedName + '.stl').hasNext()) {
        version++;
        versionedName = fileName + '_V' + version;
      }
      fileName = versionedName;
    }
    
    // Generar el contenido STL
    var stlBlob = generateSTL(data, metadata);
    
    // Guardar el archivo en Google Drive
    var file = tempFolder.createFile(stlBlob.setName(fileName + '.stl'));
    
    // Generar URL de descarga
    var downloadUrl = file.getDownloadUrl();
    
    return {
      success: true,
      message: 'Archivo STL generado exitosamente.',
      downloadUrl: downloadUrl,
      fileName: fileName + '.stl',
      fileId: file.getId(),
      folderId: tempFolder.getId()
    };
  } catch (error) {
    Logger.log('Error al descargar STL: ' + error.toString());
    return {
      success: false,
      message: 'Error al generar archivo STL: ' + error.message
    };
  }
}

/**
 * Genera un archivo Excel con los datos de subtítulos
 * @param {Array} data - Matriz de datos extraídos del Excel original
 * @param {Object} metadata - Metadatos del programa
 * @return {Object} Resultado de la operación
 */
function generateExcel(data, metadata) {
  try {
    // Obtener el ID de la carpeta compartida
    var sharedFolderId = getConfig('folders').tempSharedFolderId;
    var sharedFolder = DriveApp.getFolderById(sharedFolderId);
    
    // Crear una carpeta temporal si no existe una reciente
    var userProperties = PropertiesService.getUserProperties();
    var lastTempFolderId = userProperties.getProperty('lastTempFolderId');
    var tempFolder;
    
    if (!lastTempFolderId) {
      // Crear una nueva carpeta temporal
      var userEmail = Session.getActiveUser().getEmail();
      var userName = userEmail.split('@')[0];
      var timestamp = new Date().getTime();
      var folderName = userName + '_' + timestamp;
      tempFolder = sharedFolder.createFolder(folderName);
      
      // Guardar el ID de la carpeta temporal
      userProperties.setProperty('lastTempFolderName', folderName);
      userProperties.setProperty('lastTempFolderId', tempFolder.getId());
    } else {
      // Usar la carpeta temporal existente
      tempFolder = DriveApp.getFolderById(lastTempFolderId);
    }
    
    // Generar nombre de archivo
    var fileName = "Subtitulos_";
    if (metadata.spanishTitle) {
      fileName += sanitizeFileName(metadata.spanishTitle);
    } else {
      fileName += "Programa";
    }
    
    // Agregar el título del episodio si existe
    if (metadata.episodeTitle) {
      fileName += "_" + sanitizeFileName(metadata.episodeTitle);
    }
    
    // Agregar el número de episodio si existe
    if (metadata.episodeNumber) {
      fileName += "_Ep" + metadata.episodeNumber;
    }
    
    // Crear una hoja de cálculo temporal
    var spreadsheet = SpreadsheetApp.create(fileName);
    var sheet = spreadsheet.getActiveSheet();
    
    // Configurar encabezados según la estructura definida en config
    var columnConfig = getConfig('excel').columns;
    var headers = [];
    
    // Lista fija de encabezados en el orden deseado
    headers[columnConfig.tcIn] = "TC IN";
    headers[columnConfig.tcOut] = "TC OUT";
    headers[columnConfig.text] = "TEXT";
    
    // Establecer encabezados en la primera fila
    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
         .setFontWeight("bold").setBackground("#f3f3f3");
    
    // Insertar datos
    if (data && data.length > 0) {
      sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
    }
    
    // Ajustar anchos de columna
    sheet.setColumnWidth(columnConfig.tcIn + 1, 100);
    sheet.setColumnWidth(columnConfig.tcOut + 1, 100);
    sheet.setColumnWidth(columnConfig.text + 1, 600);
    
    // Guardar la hoja de cálculo como Excel
    var spreadsheetId = spreadsheet.getId();
    var url = "https://docs.google.com/spreadsheets/d/" + spreadsheetId + "/export?format=xlsx";
    var token = ScriptApp.getOAuthToken();
    var response = UrlFetchApp.fetch(url, {
      headers: {
        'Authorization': 'Bearer ' + token
      }
    });
    
    // Guardar el archivo Excel en la carpeta temporal
    var xlsxBlob = response.getBlob().setName(fileName + ".xlsx");
    var file = tempFolder.createFile(xlsxBlob);
    
    // Eliminar la hoja de cálculo temporal
    DriveApp.getFileById(spreadsheetId).setTrashed(true);
    
    return {
      success: true,
      message: 'Archivo Excel generado exitosamente.',
      downloadUrl: file.getDownloadUrl(),
      fileName: fileName + ".xlsx",
      fileId: file.getId(),
      folderId: tempFolder.getId()
    };
  } catch (error) {
    Logger.log('Error al generar Excel: ' + error.toString());
    return {
      success: false,
      message: 'Error al generar archivo Excel: ' + error.message
    };
  }
}

/**
 * Incluye contenido HTML desde otro archivo
 * @param {String} filename - Nombre del archivo a incluir
 * @return {String} Contenido del archivo HTML
 */
function include(filename) {
  try {
    // Intentar obtener el contenido del archivo
    var content = HtmlService.createHtmlOutputFromFile(filename).getContent();
    
    // Registrar éxito de inclusión
    Logger.log('Archivo incluido correctamente: ' + filename);
    
    return content;
  } catch (error) {
    // Registrar error
    Logger.log('Error al incluir archivo ' + filename + ': ' + error.toString());
    
    // Devolver un comentario HTML como fallback
    return '<!-- Error al incluir ' + filename + ': ' + error.message + ' -->';
  }
}

/**
 * Genera un archivo STL idéntico al modelo de referencia
 * @param {Array} data - Array de subtítulos con timecodes y texto
 * @param {Object} metadata - Metadatos del programa
 * @return {Blob} Contenido en formato STL como un Blob binario
 */
function generateLegacySTL(data, metadata) {
  try {
    Logger.log("Generando archivo STL compatible con " + data.length + " subtítulos");
    
    // Crear los bytes del archivo STL
    var stlBytes = [];
    
    // Crear el encabezado GSI con los valores exactos del modelo
    var gsiBytes = createLegacySTLHeader(metadata);
    
    // Crear array para almacenar bloques TTI
    var ttiBlocks = [];
    
    // Leer configuración
    var columnConfig = getConfig('excel').columns;
    var maxSubtitles = getConfig('limits').maxSubtitles;
    
    // Crear un subtítulo especial para el título (subtitleNumber = 0)
    // Este es el primer subtítulo en el archivo de referencia
    var titleText = metadata.spanishTitle || "Sin título";
    if (metadata.episodeNumber) {
      titleText += ":" + metadata.episodeNumber;
    }
    
    var titleSubtitle = {
      number: 0,
      timecodeIn: "01:00:00:15",
      timecodeOut: "01:00:04:15",
      text: titleText
    };
    
    // Añadir el subtítulo del título como bloque 0
    var titleTTI = createFixedLegacyTTIBlock(0, titleSubtitle.timecodeIn, titleSubtitle.timecodeOut, titleSubtitle.text, 2);
    ttiBlocks.push(titleTTI);
    
    // Crear bloques TTI para cada subtítulo desde Excel (a partir del número 1)
    var subtitleNumber = 1;
    
    for (var i = 0; i < data.length && subtitleNumber < maxSubtitles; i++) {
      // Extraer datos del Excel
      var timecodeIn, subtitleText, timecodeOut;
      
      // Si tiene al menos 4 columnas (formato Excel normal)
      if (data[i].length >= 4) {
        timecodeIn = data[i][columnConfig.timeIn];
        subtitleText = data[i][columnConfig.text];
        timecodeOut = data[i][columnConfig.timeOut];
      } 
      // Formato alternativo con 3 columnas
      else if (data[i].length >= 3) {
        timecodeIn = data[i][0];
        subtitleText = data[i][1];
        timecodeOut = data[i][2];
      }
      else {
        // Saltar esta fila
        continue;
      }
      
      // Verificar datos mínimos
      if (!timecodeIn || !timecodeOut || !subtitleText) {
        continue;
      }
      
      // Crear un bloque TTI para este subtítulo
      var ttiBlock = createFixedLegacyTTIBlock(subtitleNumber, timecodeIn, timecodeOut, subtitleText, 2);
        ttiBlocks.push(ttiBlock);
      subtitleNumber++;
    }
    
    // Agregar los bytes del GSI al array final
    stlBytes = gsiBytes;
    
    // Agregar cada bloque TTI
    for (var j = 0; j < ttiBlocks.length; j++) {
      stlBytes = stlBytes.concat(ttiBlocks[j]);
    }
    
    // Convertir el array a un Blob binario
    var byteArray = new Uint8Array(stlBytes);
    var blob = Utilities.newBlob(byteArray, 'application/octet-stream', 'subtitles.stl');
    
    Logger.log("Archivo STL generado correctamente con " + ttiBlocks.length + " subtítulos");
    return blob;
  } catch (error) {
    Logger.log('Error al generar STL: ' + error.toString());
    throw new Error('Error al generar archivo STL: ' + error.message);
  }
}

/**
 * Crea un encabezado GSI exactamente como el archivo de referencia
 * @param {Object} metadata - Metadatos del programa
 * @return {Array} Array de bytes del encabezado GSI (1024 bytes)
 */
function createLegacySTLHeader(metadata) {
  // Crear un buffer de 1024 bytes para el GSI
  var buffer = new Array(1024);
  
  // Llenar el buffer con espacios (ASCII 32)
  for (var i = 0; i < buffer.length; i++) {
    buffer[i] = 32;
  }
  
  // Obtener la configuración STL
  var stlConfig = getConfig('stl');
  
  // Determinar el idioma desde los metadatos (por defecto español)
  var language = (metadata && metadata.language) ? metadata.language : 'es';
  
  // CPN - Code Page Number (bytes 0-2) - Establecer a "850" según la configuración
  writeBlockToGSI(buffer, 0, stlConfig.codePage);
  
  // DFC - Disk Format Code (bytes 3-10) - Establecer a "STL23.01" según la configuración
  writeBlockToGSI(buffer, 3, "STL23.01");
  
  // DSC - Display Standard Code (byte 11) - Establecer a '0' (48 en ASCII para '0')
  buffer[11] = 48; // '0' para Open subtitling
  
  // CCT - Character Code Table (bytes 12-13) - Establecer a "00" para Latin
  writeBlockToGSI(buffer, 12, stlConfig.charTable);
  
  // LC - Language Code (bytes 14-16) - Establecer según el idioma configurado
  var langCode = stlConfig.ebuLanguageCode[language] || stlConfig.ebuLanguageCode['es'];
  writeBlockToGSI(buffer, 14, langCode);
  
  // OPT - Original Programme Title (bytes 24-55)
  var title = metadata.englishTitle || "Descendants_ShuffleOfLove_EP01_L";
  title = title.substring(0, 32);
  writeBlockToGSI(buffer, 24, title);
  
  // OET - Original Episode Title (bytes 56-87)
  var episodeCode = metadata.episodeNumber || "";
  episodeCode = episodeCode.substring(0, 32);
  writeBlockToGSI(buffer, 56, episodeCode);
  
  // TPT - Translated Programme Title (bytes 88-119)
  var translatedTitle = metadata.spanishTitle || "";
  translatedTitle = translatedTitle.substring(0, 32);
  writeBlockToGSI(buffer, 88, translatedTitle);
  
  // TET - Translated Episode Title (bytes 120-151)
  // Dejar en blanco o con valores por defecto
  
  // RN - Revision Number (bytes 241-242)
  writeBlockToGSI(buffer, 241, "01");
  
  // TNB - Total Number of TTI Blocks - se actualizará más tarde
  
  // TNS - Total Number of Subtitles - se actualizará más tarde
  
  // Max Characters Per Row (no especificado en el estándar pero usado por editores)
  // Posición 226-228 para algunos editores, establecer a "040" (40 caracteres)
  writeBlockToGSI(buffer, 226, "040");
  
  // Max Number of Rows (no especificado en el estándar pero usado por editores)
  // Posición 229-231 para algunos editores, establecer a "023" (23 filas)
  writeBlockToGSI(buffer, 229, "023");
  
  // DSN - Disk Sequence Number (bytes 355-356)
  writeBlockToGSI(buffer, 355, "1");
  
  // TND - Total Number of Disks (bytes 381-382)
  writeBlockToGSI(buffer, 381, "1");
  
  // CO - Country of Origin (bytes 208-210)
  var countryCode = stlConfig.countryCode[language] || stlConfig.countryCode['es'];
  writeBlockToGSI(buffer, 208, countryCode === 'BR' ? 'BRA' : (countryCode === 'ES' ? 'ARG' : countryCode));
  
  // PUB - Publisher (bytes 215-230)
  writeToGSI(gsiBytes, 215, "MEDIA ACCESS CO.");
  
  // EN: Editor Name (bytes 231-262)
  writeToGSI(gsiBytes, 231, "EXCEL STL CONVERTER");
  
  // ECD: Editor Contact Details (bytes 263-294)
  writeToGSI(gsiBytes, 263, "appsheet@mediaaccesscompany.com");
  
  return gsiBytes;
}

/**
 * Escribe un bloque de texto en el GSI a partir de la posición especificada
 * @param {Array} buffer - Array de bytes del GSI
 * @param {Number} offset - Posición donde escribir
 * @param {String} text - Texto a escribir
 */
function writeBlockToGSI(buffer, offset, text) {
  if (!text) return;
  
  for (var i = 0; i < text.length; i++) {
    buffer[offset + i] = text.charCodeAt(i);
  }
}

/**
 * Crea bloques TTI exactamente como el archivo de referencia
 * 
 * @param {Object} subtitle - Objeto con la información del subtítulo
 * @param {Number} subtitleNumber - Número de subtítulo
 * @param {Object} metadata - Metadatos adicionales
 * @return {Array} Array de bytes representando el bloque TTI
 */
function createFixedTTIBlock(subtitle, subtitleNumber) {
  // Crear un buffer de 128 bytes
  var buffer = new Array(128);
  
  // Inicializar todos los bytes con 0x00 (nulos) para evitar que los editores de subtítulos
  // interpreten incorrectamente los bytes después del ETX como caracteres adicionales
  for (var i = 0; i < buffer.length; i++) {
    buffer[i] = 0x00;
  }
  
  // SGN - Subtitle Group Number (bytes 0-1) - Little-endian
  buffer[0] = 0x00;
  buffer[1] = 0x00;
  
  // SN - Subtitle Number (bytes 2-3) - En formato little-endian
  // IMPORTANTE: El número de subtítulo debe estar en formato little-endian
  // El primer byte es el menos significativo (LSB)
  buffer[2] = subtitleNumber & 0xFF;         // LSB (byte menos significativo)
  buffer[3] = (subtitleNumber >> 8) & 0xFF;  // MSB (byte más significativo)
  
  // EBN - Extension Block Number (byte 4)
  buffer[4] = 0xFF;
  
  // CS - Cumulative Status (byte 5)
  buffer[5] = 0x00;
  
  // Procesar timecode IN (bytes 6-9)
  var tcIn = parseTimecode(subtitle.timecodeIn);
  buffer[6] = tcIn.hours;
  buffer[7] = tcIn.minutes;
  buffer[8] = tcIn.seconds;
  buffer[9] = tcIn.frames;
  
  // Procesar timecode OUT (bytes 10-13)
  var tcOut = parseTimecode(subtitle.timecodeOut);
  buffer[10] = tcOut.hours;
  buffer[11] = tcOut.minutes;
  buffer[12] = tcOut.seconds;
  buffer[13] = tcOut.frames;
  
  // VP - Vertical Position (byte 14)
  // Calcular posición vertical basada en el número de líneas
  var text = subtitle.text || "";
  var lines = text.split('\n').length;
  buffer[14] = calculateVerticalPosition(lines);
  
  // JC - Justification Code (byte 15)
  // IMPORTANTE: Usar el valor correcto para la justificación (2 = centrado)
  buffer[15] = 0x02; // Centrado por defecto
  
  // CF - Comment Flag (byte 16)
  buffer[16] = 0x00;
  
  // TF - Text Field (bytes 17-127)
  // Procesar el texto del subtítulo
  
  // Procesar etiquetas HTML como <br>
  text = text.replace(/<br\s*\/?>/gi, "\n");
  text = text.replace(/<[^>]*>/g, ""); // Eliminar otras etiquetas HTML
  
  // Posición actual en el buffer para escribir texto
  var position = 17;
  
  // Procesar cada carácter del texto
  for (var i = 0; i < text.length && position < 127; i++) {
    var char = text.charAt(i);
    
      if (char === '\n') {
      // Salto de línea como EOL en EBU
      buffer[position++] = 0x8A;
    } else {
      // Usar la función getEBUCodeForChar para mapear correctamente los caracteres
      buffer[position++] = getEBUCodeForChar(char);
    }
  }
  
  // Añadir terminador ETX (End of Text, 0x03) siempre después del último carácter
  if (position < 128) {
    buffer[position++] = 0x03; // ETX - End of Text
    for (var i = position; i < 128; i++) {
      buffer[i] = 0xFF;
    }
  }
  
  return buffer;
}

/**
 * Detecta y asigna la posición vertical basada en el texto
 * Analiza el texto para identificar etiquetas de posición \anX
 * @param {String} text - Texto del subtítulo
 * @return {Number} Código de posición vertical para el STL
 */
function detectVerticalPosition(text) {
  // Valores predeterminados según número de líneas
  var lines = text.split('\n').length;
  var defaultPosition = lines === 1 ? 16 : 14;
  
  // Buscar etiquetas de posición \anX
  var positionRegex = /\\an([1-9])/;
  var match = text.match(positionRegex);
  
  if (match) {
    // Convertir la posición \anX a valores verticales de STL
    var positionCode = parseInt(match[1]);
    switch (positionCode) {
      case 1: case 4: case 7: return 18; // Superior
      case 2: case 5: case 8: return 14; // Medio
      case 3: case 6: case 9: return 10; // Inferior
      default: return defaultPosition;
    }
  }
  
  return defaultPosition;
}

/**
 * Preprocesa el texto de subtítulo para manejar etiquetas de formato
 * Transforma etiquetas como <font color="Red"> en códigos de control STL
 * @param {String} text - Texto original del subtítulo
 * @return {Array} Array mixto de cadenas y códigos numéricos EBU
 */
function preprocessSTLText(text) {
  var result = [];
  
  // Reemplazar etiquetas \anX ya que serán manejadas por la posición vertical
  text = text.replace(/\\an[1-9]/g, '');
  
  // Dividir el texto por etiquetas de font
  var parts = text.split(/<\/?font[^>]*>/);
  var tags = text.match(/<\/?font[^>]*>/g) || [];
  
  // Códigos de color según especificación EBU
  var colorCodes = {
    "Red": 0x81,
    "Green": 0x82, 
    "Yellow": 0x83,
    "Blue": 0x84,
    "Magenta": 0x85,
    "Cyan": 0x86,
    "White": 0x87,
    "Black": 0x88
  };
  
  // Si no hay etiquetas, devolver el texto como está
  if (tags.length === 0) {
    for (var i = 0; i < text.length; i++) {
      result.push(text[i]);
    }
    return result;
  }
  
  // Procesar cada parte del texto con sus etiquetas correspondientes
  for (var i = 0; i < parts.length; i++) {
    // Agregar la parte actual del texto
    for (var j = 0; j < parts[i].length; j++) {
      result.push(parts[i][j]);
    }
    
    // Procesar la etiqueta si existe
    if (i < tags.length) {
      var tag = tags[i];
      var colorMatch = tag.match(/color="([^"]+)"/i);
      
      if (colorMatch) {
        var color = colorMatch[1];
        // Insertar código de control de color si está en nuestro mapa
        if (colorCodes[color]) {
          result.push(colorCodes[color]);
        }
      } else if (tag === '</font>') {
        // Cerrar etiqueta de fuente - volver al color predeterminado
        result.push(0x87); // Blanco como color predeterminado
      }
    }
  }
  
  return result;
}

/**
 * Formatea una fecha en formato YYMMDD (formato requerido por STL)
 * @param {Date} date - Fecha a formatear
 * @return {String} Fecha formateada como YYMMDD
 */
function formatDate(date) {
  var day = padNumber(date.getDate(), 2);
  var month = padNumber(date.getMonth() + 1, 2);
  var year = date.getFullYear().toString().substring(2);
  return year + month + day;
}

/**
 * Obtiene o crea una carpeta temporal para guardar los archivos STL
 * @return {GoogleAppsScript.Drive.Folder} Carpeta temporal
 */
function getOrCreateTempFolder() {
  try {
    // Nombre de la carpeta temporal con información del usuario para evitar conflictos
    var userEmail = Session.getActiveUser().getEmail();
    var userName = userEmail ? userEmail.split('@')[0] : "anonymous";
    const folderName = 'STL_Temp_Files_' + userName;
    
    // Intentar buscar la carpeta en Drive personal (no compartido)
    // Nota: No usamos getRootFolder() porque podría apuntar a unidades compartidas
    var folders = DriveApp.getFoldersByName(folderName);
    
    // Si la carpeta existe, devolverla
    if (folders.hasNext()) {
      return folders.next();
    }
    
    // Si no existe, crear una nueva carpeta en Drive personal
    const newFolder = DriveApp.createFolder(folderName);
    Logger.log('Nueva carpeta temporal creada: ' + newFolder.getName());
    
    // Guardar registro de la carpeta en las propiedades del usuario
    try {
      PropertiesService.getUserProperties().setProperty('stl_temp_folder_id', newFolder.getId());
    } catch (propError) {
      Logger.log('No se pudo guardar referencia a la carpeta: ' + propError.message);
    }
    
    return newFolder;
  } catch (error) {
    // Si ocurre un error, registrarlo y crear una carpeta con timestamp único
    Logger.log('Error al acceder a carpeta compartida: ' + error.message);
    Logger.log('Creando carpeta alternativa en Drive personal...');
    
    // Crear nombre único con timestamp
    var timestamp = new Date().getTime();
    return DriveApp.createFolder('STL_Temp_Files_' + timestamp);
  }
}

/**
 * Limpia los archivos temporales antiguos
 * @return {String} Mensaje con el resultado de la limpieza
 */
function cleanupMyTempFiles() {
  const folder = getOrCreateTempFolder();
  const files = folder.getFiles();
  const retentionDays = getConfig('folders').tempRetentionDays || 7;
  const cutoffDate = new Date();
  cutoffDate.setDate(cutoffDate.getDate() - retentionDays);
  
  let deletedCount = 0;
  
  while (files.hasNext()) {
    const file = files.next();
    const fileDate = file.getDateCreated();
    
    if (fileDate < cutoffDate) {
      file.setTrashed(true);
      deletedCount++;
    }
  }
  
  return deletedCount + ' archivos temporales eliminados.';
}

/**
 * Importa datos desde un archivo Excel (XLS/XLSX)
 * @param {Blob} blob - Blob del archivo Excel a importar
 * @return {Object} Objeto con datos y metadatos extraídos del Excel
 */
function importExcel(blob) {
  var tempFolder = null;
  var tempFile = null;
  var convertedSheet = null;
  
  try {
    Logger.log('Importando Excel: ' + blob.getName());
    
    // Crear una carpeta temporal en el Drive personal del usuario (no en unidad compartida)
    try {
      tempFolder = DriveApp.createFolder('STL_Excel_Temp_' + new Date().getTime());
      Logger.log('Carpeta temporal creada: ' + tempFolder.getName());
    } catch (folderError) {
      Logger.log('Error al crear carpeta temporal: ' + folderError.message);
      // Intentar con carpeta de usuario
      tempFolder = getOrCreateTempFolder();
    }
    
    // Guardar el blob como archivo en la carpeta temporal
    try {
      tempFile = tempFolder.createFile(blob);
      Logger.log('Archivo temporal creado: ' + tempFile.getName() + ' (ID: ' + tempFile.getId() + ')');
    } catch (fileError) {
      Logger.log('Error al crear archivo temporal: ' + fileError.message);
      throw new Error('No se pudo guardar el archivo Excel temporalmente: ' + fileError.message);
    }
    
    // Variables para almacenar datos
    var subtitles = [];
    var metadata = {
      englishTitle: 'Untitled',
      spanishTitle: 'Sin título',
      episodeNumber: '01',
      language: 'es'
    };
    
    // Intentar procesar el archivo con diferentes métodos
    var success = false;
    
    // MÉTODO 1: Convertir el archivo Excel a Google Sheets y procesarlo
    try {
      Logger.log('Convirtiendo Excel a Google Sheets...');
      // Usar la API de Drive para convertir el archivo
      var fileId = tempFile.getId();
      var fileName = tempFile.getName();
      
      // Crear una copia como Google Sheets
      convertedSheet = convertExcelToGoogleSheet(fileId, "Convertido_" + fileName);
      
      if (convertedSheet) {
        // Abrir Google Sheets
        var sheet = SpreadsheetApp.openById(convertedSheet.getId()).getActiveSheet();
        
        // Usar nuestra función para procesar la hoja
        var result = processGoogleSheet(sheet);
        
        if (result && result.success && result.subtitles && result.subtitles.length > 0) {
          subtitles = result.subtitles;
          metadata = result.metadata || metadata; // Usar los metadatos del resultado o los predeterminados
          success = true;
          Logger.log('Éxito procesando con Google Sheets convertido');
        } else {
          Logger.log('No se pudieron extraer subtítulos del Google Sheets: ' + JSON.stringify(result));
        }
      } else {
        Logger.log('No se pudo convertir el archivo a Google Sheets');
      }
    } catch (gsError) {
      Logger.log('Error al procesar con Google Sheets: ' + gsError.message);
    }
    
    // Limpiar archivos temporales
    cleanupTempFiles(tempFile, convertedSheet, tempFolder);
    
    // Verificar si tuvimos éxito
    if (!success || !subtitles || subtitles.length === 0) {
      throw new Error('No se pudo procesar el archivo Excel o no contiene subtítulos válidos');
    }
    
    // Registrar los datos extraídos para depuración
    Logger.log('Subtítulos extraídos: ' + subtitles.length);
    Logger.log('Metadatos: ' + JSON.stringify(metadata));
    
    return {
      subtitles: subtitles,
      metadata: metadata
    };
  } catch (error) {
    Logger.log('Error al importar Excel: ' + error.toString());
    
    // Asegurarnos de limpiar cualquier archivo temporal que haya quedado
    cleanupTempFiles(tempFile, convertedSheet, tempFolder);
    
    throw new Error('Error al importar Excel: ' + error.message);
  }
}

/**
 * Convierte un archivo Excel a formato Google Sheets
 * @param {String} fileId - ID del archivo Excel en Drive
 * @param {String} newName - Nombre para el nuevo archivo Google Sheets
 * @return {File} Archivo de Google Sheets convertido
 */
function convertExcelToGoogleSheet(fileId, newName) {
  try {
    // Obtener el archivo desde Drive
    var file = DriveApp.getFileById(fileId);
    if (!file) {
      throw new Error('No se pudo encontrar el archivo con ID: ' + fileId);
    }
    
    // Verificar que es un archivo Excel
    var mimeType = file.getMimeType();
    if (mimeType !== 'application/vnd.ms-excel' && 
        mimeType !== 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') {
      throw new Error('El archivo no es un Excel válido. MimeType: ' + mimeType);
    }
    
    // Usar la API avanzada de Drive para convertir
    var resource = {
      name: newName,
      mimeType: 'application/vnd.google-apps.spreadsheet'
    };
    
    // Intentar método 1: Usar Drive API v2 (disponible en GAS)
    try {
      var driveService = Drive.Files.copy(resource, fileId);
      return DriveApp.getFileById(driveService.id);
    } catch (apiError) {
      Logger.log('Error en Drive API v2: ' + apiError.message);
      
      // Intentar método 2: Crear una copia y renombrarla
      var fileBlob = file.getBlob();
      var tempFolder = file.getParents().next();
      var newFile = tempFolder.createFile(fileBlob);
      newFile.setName(newName);
      
      // Convertir a Google Sheets (esto solo funciona si Drive está configurado para convertir automáticamente)
      var id = newFile.getId();
      var convertedFile = Drive.Files.update({
        mimeType: 'application/vnd.google-apps.spreadsheet'
      }, id);
      
      return DriveApp.getFileById(convertedFile.id);
    }
  } catch (error) {
    Logger.log('Error al convertir Excel a Google Sheets: ' + error.toString());
    throw new Error('No se pudo convertir el archivo Excel a Google Sheets: ' + error.message);
  }
}

/**
 * Limpia los archivos temporales creados durante el procesamiento
 * @param {File} tempFile - Archivo temporal original
 * @param {File} convertedSheet - Hoja de cálculo convertida
 * @param {Folder} tempFolder - Carpeta temporal
 */
function cleanupTempFiles(tempFile, convertedSheet, tempFolder) {
  try {
    // Limpiar archivo original
    if (tempFile) {
      try {
        tempFile.setTrashed(true);
        Logger.log('Archivo temporal original eliminado');
      } catch (e) {
        Logger.log('No se pudo eliminar archivo temporal original: ' + e.message);
      }
    }
    
    // Limpiar archivo convertido
    if (convertedSheet) {
      try {
        DriveApp.getFileById(convertedSheet.getId()).setTrashed(true);
        Logger.log('Archivo Google Sheets convertido eliminado');
      } catch (e) {
        Logger.log('No se pudo eliminar archivo convertido: ' + e.message);
      }
    }
    
    // Limpiar carpeta temporal
    if (tempFolder && tempFolder.getName().startsWith('STL_Excel_Temp_')) {
      try {
        tempFolder.setTrashed(true);
        Logger.log('Carpeta temporal eliminada');
      } catch (e) {
        Logger.log('No se pudo eliminar carpeta temporal: ' + e.message);
      }
    }
  } catch (error) {
    Logger.log('Error al limpiar archivos temporales: ' + error.message);
    // No bloqueamos la ejecución por errores en la limpieza
  }
}

/**
 * Procesa una hoja de cálculo para extraer subtítulos y metadatos
 * @param {Sheet} sheet - Hoja de Google Sheets a procesar
 * @return {Object} Objeto con subtítulos y metadatos
 */
function processSpreadsheet(sheet) {
  // Configuración de Excel desde config
  var config = getConfig('excel');
  var startRow = config.startRow || 11;  // Fila donde comienzan los subtítulos (1-indexado)
  var timeInCol = config.columns.timeIn || 1;  // Columna B = 1 (0-indexado)
  var textCol = config.columns.text || 2;  // Columna C = 2 (0-indexado)
  var timeOutCol = config.columns.timeOut || 3;  // Columna D = 3 (0-indexado)
  
  // Extraer metadatos
  var metadata = {};
  try {
    var titleEnglishRow = config.metadata.titleEnglish || 3;
    var titleSpanishRow = config.metadata.titleSpanish || 4;
    var episodeRow = config.metadata.episodeNumber || 5;
    
    // Usar getValue para obtener cada celda individual
    var englishTitle = sheet.getRange(titleEnglishRow, 3).getValue();
    var spanishTitle = sheet.getRange(titleSpanishRow, 3).getValue();
    var episodeNumber = sheet.getRange(episodeRow, 3).getValue();
    
    // Asegurarnos de que son strings y limpiarlos
    englishTitle = (englishTitle || '').toString().replace(/^.*English Title:?\s*/i, '').trim();
    spanishTitle = (spanishTitle || '').toString().replace(/^.*Spanish Title:?\s*/i, '').trim();
    episodeNumber = (episodeNumber || '').toString().replace(/^.*Episode.*?:?\s*/i, '').trim();
    
    metadata.englishTitle = englishTitle || 'Untitled';
    metadata.spanishTitle = spanishTitle || 'Sin título';
    metadata.episodeNumber = episodeNumber || '01';
    metadata.language = 'es';  // Default language
    
    Logger.log('Metadatos extraídos: ' + JSON.stringify(metadata));
  } catch (metaError) {
    Logger.log('Error al extraer metadatos: ' + metaError.message);
    // Establecer valores predeterminados si no se pueden extraer metadatos
    metadata = {
      englishTitle: 'Untitled',
      spanishTitle: 'Sin título',
      episodeNumber: '01',
      language: 'es'
    };
  }
  
  // Extraer datos de subtítulos
  var subtitles = [];
  var lastRow = sheet.getLastRow();
  var columnConfig = config.columns;
  
  // Columnas (ajustar según la configuración)
  var timeInCol = columnConfig.timeIn || 1;  // 0-indexado (B)
  var textCol = columnConfig.text || 2;      // 0-indexado (C)
  var timeOutCol = columnConfig.timeOut || 3; // 0-indexado (D)
  
  for (var i = startRow; i <= lastRow; i++) {
    // Obtener los valores de timecode
    var timeIn = sheet.getRange(i, timeInCol + 1).getValue(); // +1 porque getRange es 1-indexado
    if (!timeIn) continue; // Saltar filas vacías
    
    var timeOut = sheet.getRange(i, timeOutCol + 1).getValue();
    if (!timeOut) continue;
    
    // Intentar obtener el formato del texto
    var textCell = sheet.getRange(i, textCol + 1);
    var richText = textCell.getRichTextValue();
    var text = richText ? richText.getText() : textCell.getValue();
    
    // Si es un valor de texto, asegurarnos de que es string
    text = text ? text.toString() : "";
    
    // IMPORTANTE: Preservar formato original - no eliminar HTML aquí
    // Convertir saltos de línea a <br> para preservarlos
    text = text.replace(/\n/g, '<br>');
    
    // Añadir el subtítulo si todos los datos están presentes
    if (timeIn && timeOut && text) {
      // Formatear los timecodes correctamente
      timeIn = formatTimecode(timeIn.toString());
      timeOut = formatTimecode(timeOut.toString());
      
      subtitles.push({
        timecodeIn: timeIn,
        timecodeOut: timeOut,
        text: text
      });
    }
  }
  
  Logger.log('Subtítulos extraídos: ' + subtitles.length);
  
  // Devolver los resultados
  return {
    metadata: metadata,
    subtitles: subtitles
  };
}

/**
 * Verifica si un string tiene formato de timecode válido de manera más flexible
 * Acepta varios formatos comunes como HH:MM:SS:FF, HH:MM:SS.FF, etc.
 * @param {String} timecode - String a verificar
 * @return {Boolean} true si el formato es válido
 */
function isValidTimecodeFlexible(timecode) {
  if (!timecode) return false;
  
  // Formato estándar HH:MM:SS:FF
  var regex1 = /^(\d{1,2}):(\d{1,2}):(\d{1,2}):(\d{1,2})$/;
  
  // Formato con punto HH:MM:SS.FF
  var regex2 = /^(\d{1,2}):(\d{1,2}):(\d{1,2})\.(\d{1,2})$/;
  
  // Formato horas opcionales MM:SS:FF
  var regex3 = /^(\d{1,2}):(\d{1,2}):(\d{1,2})$/;
  
  return regex1.test(timecode) || regex2.test(timecode) || regex3.test(timecode);
}

/**
 * Genera un archivo STL con mensaje de error para informar al usuario
 * @param {String} errorMessage - Mensaje de error a incluir en el archivo
 * @return {Blob} Blob del archivo STL con mensaje de error
 */
function generateErrorSTL(errorMessage) {
  // Crear un conjunto de subtítulos con el mensaje de error
  var errorSubtitles = [
    ["00:00:05:00", "ERROR EN LA CONVERSIÓN", "00:00:10:00"],
    ["00:00:10:00", "Se produjo un error al procesar el archivo.", "00:00:15:00"],
    ["00:00:15:00", errorMessage, "00:00:20:00"],
    ["00:00:20:00", "Por favor, verifique el formato del archivo Excel", "00:00:25:00"],
    ["00:00:25:00", "y asegúrese de que cumple con las especificaciones.", "00:00:30:00"]
  ];
  
  // Crear metadatos de error
  var errorMetadata = {
    englishTitle: "ERROR - EXCEL TO STL",
    spanishTitle: "ERROR - CONVERSIÓN DE EXCEL A STL",
    episodeNumber: "N/A",
    language: "es"
  };
  
  // Generar el STL con los datos de error
  return generateLegacySTL(errorSubtitles, errorMetadata);
}

/**
 * Función de diagnóstico para analizar archivos STL
 * @param {Blob} stlFile - Archivo STL a analizar
 * @return {Object} Diagnóstico detallado del archivo STL
 */
function analyzeLegacySTL(stlFile) {
  try {
    // Convertir el blob a array de bytes
    var stlBytes = stlFile.getBytes();
    var result = {
      fileSize: stlBytes.length,
      gsiHeader: {},
      gsiTechnicalDetails: {}, // Nuevo: Detalles técnicos específicos del GSI
      ttiBlocks: [],
      ttiTechnicalDetails: [], // Nuevo: Detalles técnicos de los bloques TTI
      diagnostics: [],
      trailingCharacters: [],
      mergedSubtitlesRisk: false,
      cumulativeStatusAnalysis: {} // Nuevo: Análisis del estado acumulativo
    };
    
    // Analizar GSI header (primeros 1024 bytes)
    if (stlBytes.length >= 1024) {
      // Información básica del GSI
      result.gsiHeader = {
        codePageNumber: String.fromCharCode.apply(null, stlBytes.slice(0, 3)),
        diskFormatCode: String.fromCharCode.apply(null, stlBytes.slice(3, 11)),
        displayStandardCode: String.fromCharCode(stlBytes[11]),
        characterCodeTableNumber: String.fromCharCode.apply(null, stlBytes.slice(12, 14)),
        languageCode: String.fromCharCode.apply(null, stlBytes.slice(14, 17)),
        originalProgrammeTitle: String.fromCharCode.apply(null, stlBytes.slice(24, 56)).trim(),
        originalEpisodeTitle: String.fromCharCode.apply(null, stlBytes.slice(56, 88)).trim(),
        translatedProgrammeTitle: String.fromCharCode.apply(null, stlBytes.slice(88, 120)).trim(),
        totalNumberOfTTI: readWordLittleEndian(stlBytes, 236),
        totalNumberOfSubtitles: readWordLittleEndian(stlBytes, 230),
        maximumNumberCharacters: String.fromCharCode.apply(null, stlBytes.slice(252, 255)).trim(), // MNC
        maximumNumberRows: String.fromCharCode.apply(null, stlBytes.slice(255, 258)).trim(), // MNR
        revision: stlBytes[946]
      };
      
      // NUEVO: Detalles técnicos específicos del GSI como bytes hexadecimales
      result.gsiTechnicalDetails = {
        CPN: bytesToHexString(stlBytes.slice(0, 4)),
        DFC: bytesToHexString(stlBytes.slice(4, 12)),
        DSC: "0x" + stlBytes[12].toString(16).padStart(2, '0').toUpperCase(),
        CCT: bytesToHexString(stlBytes.slice(12, 14)),
        LC: bytesToHexString(stlBytes.slice(14, 16)),
        TNB_bytes: bytesToHexString(stlBytes.slice(236, 241)),
        TNS_bytes: bytesToHexString(stlBytes.slice(241, 246)),
        TNG_bytes: bytesToHexString(stlBytes.slice(246, 251)),
        MNC_bytes: bytesToHexString(stlBytes.slice(252, 255)),
        MNR_bytes: bytesToHexString(stlBytes.slice(255, 258)),
        TCS_bytes: bytesToHexString(stlBytes.slice(258, 259)),
      };
      
      // Verificar integridad del GSI
      if (result.gsiHeader.codePageNumber !== "STL") {
        result.diagnostics.push("WARNING: GSI header no comienza con 'STL'");
      }
      
      // Identificar otras posibles inconsistencias
      result.diagnostics.push("Formato de disco: " + result.gsiHeader.diskFormatCode);
      result.diagnostics.push("Estándar de display: " + result.gsiHeader.displayStandardCode);
      result.diagnostics.push("Máximo número de caracteres por línea (MNC): " + result.gsiHeader.maximumNumberCharacters);
      result.diagnostics.push("Máximo número de líneas (MNR): " + result.gsiHeader.maximumNumberRows);
    } else {
      result.diagnostics.push("ERROR: Archivo STL demasiado pequeño para tener GSI header válido");
    }
    
    // Analizar bloques TTI
    var ttiCount = 0;
    var totalTextLength = 0;
    var hasZeroBytes = false;
    var cumulativeStatusCount = {
      "0": 0, // No acumulativo
      "1": 0, // Acumulativo (puede causar problemas)
      otros: 0
    };
    
    for (var offset = 1024; offset + 128 <= stlBytes.length; offset += 128) {
      var ttiBlock = analyzeBlock(stlBytes, offset);
      
      // Analizar el estado acumulativo (CS)
      if (ttiBlock.cumulativeStatus === 0) {
        cumulativeStatusCount["0"]++;
      } else if (ttiBlock.cumulativeStatus === 1) {
        cumulativeStatusCount["1"]++;
      } else {
        cumulativeStatusCount.otros++;
      }
      
      // Verificar si hay caracteres después del ETX
      if (ttiBlock.etxPosition > -1 && ttiBlock.hasTrailingChars) {
        result.trailingCharacters.push({
          subtitleNumber: ttiCount,
          subtitleId: ttiBlock.subtitleNumber,
          etxPosition: ttiBlock.etxPosition - 16, // Posición relativa al inicio del texto
          trailingBytes: ttiBlock.trailingBytes,
          count: ttiBlock.trailingCharCount,
          bytesAfterETX: ttiBlock.bytesAfterETX // Nuevo: bytes específicos después del ETX
        });
        
        // Añadir diagnóstico
        result.diagnostics.push("WARNING: Subtítulo #" + ttiCount + " tiene " + 
                              ttiBlock.trailingCharCount + " caracteres adicionales después del ETX");
                              
        // Verificar si los caracteres después de ETX son bytes 0x00 (posible problema de fusión)
        if (ttiBlock.hasZeroTrailingBytes) {
          hasZeroBytes = true;
        }
      }
      
      // Acumular longitud total de texto para análisis de fusión
      totalTextLength += ttiBlock.textChars.length;
      
      // Agregar detalles técnicos completos para este bloque TTI
      if (ttiCount < 5) { // Limitar a los primeros 5 para no sobrecargar
        result.ttiTechnicalDetails.push({
          blockNumber: ttiCount,
          subtitleNumber: ttiBlock.subtitleNumber,
          byteDetails: {
            SGN: "0x" + ttiBlock.subtitleGroupNumber.toString(16).padStart(2, '0').toUpperCase(),
            SN: ttiBlock.subtitleNumber,
            SN_bytes: bytesToHexString([stlBytes[offset+1], stlBytes[offset+2]]),
            EBN: "0x" + ttiBlock.extendedBlockNumber.toString(16).padStart(2, '0').toUpperCase(),
            CS: "0x" + ttiBlock.cumulativeStatus.toString(16).padStart(2, '0').toUpperCase() + " (" + ttiBlock.cumulativeStatus + ")",
            TC_IN: formatTcBytes(ttiBlock.timecodeIn),
            TC_OUT: formatTcBytes(ttiBlock.timecodeOut),
            VP: "0x" + ttiBlock.verticalPosition.toString(16).padStart(2, '0').toUpperCase() + " (" + ttiBlock.verticalPosition + ")",
            JC: "0x" + ttiBlock.justificationCode.toString(16).padStart(2, '0').toUpperCase() + " (" + ttiBlock.justificationCode + ")",
            CF: "0x" + ttiBlock.commentFlag.toString(16).padStart(2, '0').toUpperCase()
          },
          textHexDump: ttiBlock.textHexDump || "No disponible"
        });
      }
      
      result.ttiBlocks.push(ttiBlock);
      ttiCount++;
    }
    
    // Análisis del estado acumulativo (CS)
    result.cumulativeStatusAnalysis = {
      noAcumulativo: cumulativeStatusCount["0"],
      acumulativo: cumulativeStatusCount["1"],
      otros: cumulativeStatusCount.otros,
      porcentajeNoAcumulativo: Math.round((cumulativeStatusCount["0"] / ttiCount) * 100),
      riesgoMerge: cumulativeStatusCount["1"] > 0
    };
    
    // Añadir diagnóstico sobre el estado acumulativo
    if (cumulativeStatusCount["1"] > 0) {
      result.diagnostics.push("ALERTA: " + cumulativeStatusCount["1"] + 
                            " subtítulos tienen estado acumulativo (CS=1), lo que puede causar problemas de fusión en algunos editores.");
    }
    
    // Verificar consistencia de número de bloques TTI
    if (result.gsiHeader.totalNumberOfTTI !== ttiCount) {
      result.diagnostics.push("WARNING: El número declarado de TTI (" + 
                            result.gsiHeader.totalNumberOfTTI + 
                            ") no coincide con el número real (" + ttiCount + ")");
    }
    
    // Detectar posible riesgo de subtítulos fusionados
    var avgSubtitleLength = totalTextLength / ttiCount;
    if (avgSubtitleLength > 70 && hasZeroBytes) {
      result.mergedSubtitlesRisk = true;
      result.diagnostics.push("ALERTA: Posible fusión de subtítulos en una sola línea. La longitud media de los subtítulos (" + 
                            Math.round(avgSubtitleLength) + 
                            " caracteres) es inusualmente alta y hay bytes 0x00 después de ETX.");
      result.diagnostics.push("SOLUCIÓN: Esto puede ocurrir cuando los editores de subtítulos no interpretan correctamente los bloques TTI. " +
                            "Recomendamos revisar el archivo en un editor de texto hexadecimal.");
    }
    
    // Agregar análisis de compatibilidad con editores
    result.editorCompatibility = {
      eztitles: getTtiCompatibilityScore(result.ttiBlocks, "eztitles"),
      fabSubtitler: getTtiCompatibilityScore(result.ttiBlocks, "fab"),
      wincaps: getTtiCompatibilityScore(result.ttiBlocks, "wincaps"),
      subtitleEdit: getTtiCompatibilityScore(result.ttiBlocks, "subtitle_edit")
    };
    
    return result;
  } catch (error) {
  return {
      error: error.toString(),
      diagnostics: ["ERROR en el análisis: " + error.message]
    };
  }
  
  // Función auxiliar para formatear bytes de timecode a formato legible
  function formatTcBytes(tc) {
    return "0x" + tc.hours.toString(16).padStart(2, '0').toUpperCase() + " " +
           "0x" + tc.minutes.toString(16).padStart(2, '0').toUpperCase() + " " +
           "0x" + tc.seconds.toString(16).padStart(2, '0').toUpperCase() + " " +
           "0x" + tc.frames.toString(16).padStart(2, '0').toUpperCase() + 
           " (" + padZero(tc.hours) + ":" + padZero(tc.minutes) + ":" + 
           padZero(tc.seconds) + ":" + padZero(tc.frames) + ")";
  }
  
  function padZero(num) {
    return (num < 10 ? "0" : "") + num;
  }
  
  // Función auxiliar para convertir bytes a string hexadecimal
  function bytesToHexString(bytes) {
    var hexString = "";
    for (var i = 0; i < bytes.length; i++) {
      hexString += "0x" + bytes[i].toString(16).padStart(2, '0').toUpperCase() + " ";
    }
    return hexString.trim();
  }
  
  // Función auxiliar para analizar un bloque TTI
  function analyzeBlock(bytes, offset) {
    var block = {
      subtitleGroupNumber: bytes[offset],
      subtitleNumber: readWordLittleEndian(bytes, offset + 1),
      extendedBlockNumber: bytes[offset + 3],
      cumulativeStatus: bytes[offset + 4],
      timecodeIn: {
        hours: bytes[offset + 5],
        minutes: bytes[offset + 6],
        seconds: bytes[offset + 7],
        frames: bytes[offset + 8]
      },
      timecodeOut: {
        hours: bytes[offset + 9],
        minutes: bytes[offset + 10],
        seconds: bytes[offset + 11],
        frames: bytes[offset + 12]
      },
      verticalPosition: bytes[offset + 13],
      justificationCode: bytes[offset + 14],
      commentFlag: bytes[offset + 15],
      textChars: [],
      hasTrailingChars: false,
      trailingBytes: [],
      trailingCharCount: 0,
      hasZeroTrailingBytes: false,
      etxPosition: -1,
      bytesAfterETX: [], // Nuevo: array de bytes después del ETX
      textHexDump: "" // Nuevo: dump hexadecimal completo del texto
    };
    
    // Extraer caracteres de texto hasta encontrar ETX (0x03) o espacio no usado (0xFF)
    var textData = [];
    var foundETX = false;
    var etxPosition = -1;
    var hexDump = "";
    
    for (var i = 0; i < 112; i++) {
      var charByte = bytes[offset + 16 + i];
      
      // Añadir al dump hexadecimal
      if (i % 16 === 0) {
        hexDump += "\n" + (i).toString(16).padStart(4, '0') + ": ";
      }
      hexDump += charByte.toString(16).padStart(2, '0') + " ";
      
      if (charByte === 0x03) { // ETX - End of Text
        etxPosition = offset + 16 + i;
        foundETX = true;
        break;
      }
      
      if (charByte === 0xFF) { // Fin del texto (espacio no usado)
        break;
      }
      
      textData.push({
        byte: charByte,
        hex: "0x" + charByte.toString(16).padStart(2, '0').toUpperCase()
      });
    }
    
    block.textChars = textData;
    block.textHexDump = hexDump;
    
    // Si se encontró un ETX, verificar caracteres adicionales después
    if (foundETX) {
      block.etxPosition = etxPosition;
      
      // Verificar bytes después del ETX
      var trailingBytes = [];
      var allZeros = true;
      var bytesAfterETX = [];
      
      for (var i = etxPosition + 1; i < offset + 128; i++) {
        var charByte = bytes[i];
        trailingBytes.push("0x" + charByte.toString(16).padStart(2, '0').toUpperCase());
        bytesAfterETX.push(charByte);
        if (charByte !== 0x00) {
          allZeros = false;
        }
      }
      
      if (trailingBytes.length > 0) {
        block.hasTrailingChars = true;
        block.trailingBytes = trailingBytes.join(', ');
        block.trailingCharCount = trailingBytes.length;
        block.hasZeroTrailingBytes = allZeros;
        block.bytesAfterETX = bytesAfterETX; // Guardar los bytes después del ETX
      }
    }
    
    // Extraer texto legible (aproximado)
    var textBytes = [];
    for (var i = 0; i < textData.length; i++) {
      textBytes.push(textData[i].byte);
    }
    block.textPreview = bytesToString(textBytes);
    
    return block;
  }
}

/**
 * Calcula un puntaje de compatibilidad para un editor específico basándose en los bloques TTI
 * @param {Array} ttiBlocks - Array de bloques TTI analizados
 * @param {String} editorType - Tipo de editor (eztitles, fab, wincaps, subtitle_edit)
 * @return {Object} Puntaje de compatibilidad y comentarios
 */
function getTtiCompatibilityScore(ttiBlocks, editorType) {
  var score = 100;
  var comments = [];
  
  // Contar problemas
  var cumulativeStatusProblems = 0;
  var etxPositionProblems = 0;
  var trailingBytesProblems = 0;
  var verticalPositionProblems = 0;
  
  for (var i = 0; i < ttiBlocks.length; i++) {
    var block = ttiBlocks[i];
    
    // Verificar CS (Cumulative Status)
    if (block.cumulativeStatus !== 0) {
      cumulativeStatusProblems++;
    }
    
    // Verificar si falta ETX
    if (block.etxPosition === -1) {
      etxPositionProblems++;
    }
    
    // Verificar si hay bytes no-zero después de ETX
    if (block.hasTrailingChars && !block.hasZeroTrailingBytes) {
      trailingBytesProblems++;
    }
    
    // Verificar la posición vertical
    if (editorType === "eztitles" && block.verticalPosition !== 14 && block.verticalPosition !== 15) {
      verticalPositionProblems++;
    }
  }
  
  // Calcular porcentajes de problemas
  var totalBlocks = ttiBlocks.length;
  
  if (totalBlocks > 0) {
    var csPercentage = (cumulativeStatusProblems / totalBlocks) * 100;
    var etxPercentage = (etxPositionProblems / totalBlocks) * 100;
    var bytesPercentage = (trailingBytesProblems / totalBlocks) * 100;
    var vpPercentage = (verticalPositionProblems / totalBlocks) * 100;
    
    // Ajustar puntaje según el editor y los problemas
    if (editorType === "eztitles") {
      if (csPercentage > 0) {
        score -= Math.min(25, csPercentage);
        comments.push("EZTitles no maneja bien subtítulos con CS=1 (acumulativo)");
      }
      if (etxPercentage > 0) {
        score -= Math.min(20, etxPercentage);
        comments.push("ETX faltante en " + etxPositionProblems + " bloques");
      }
      if (vpPercentage > 0) {
        score -= Math.min(10, vpPercentage);
        comments.push("Posición vertical no óptima para EZTitles (ideal: 14-15)");
      }
    } else if (editorType === "fab") {
      if (csPercentage > 0) {
        score -= Math.min(20, csPercentage);
        comments.push("FAB Subtitler es sensible al campo CS");
      }
      if (bytesPercentage > 0) {
        score -= Math.min(15, bytesPercentage);
        comments.push("FAB Subtitler es sensible a bytes no-cero después del ETX");
      }
    } else if (editorType === "wincaps") {
      if (csPercentage > 0) {
        score -= Math.min(30, csPercentage);
        comments.push("WinCAPS es muy sensible al campo CS");
      }
    } else if (editorType === "subtitle_edit") {
      // Subtitle Edit es bastante tolerante
      if (csPercentage > 50) {
        score -= 10;
        comments.push("Incluso Subtitle Edit puede tener problemas con muchos CS=1");
      }
    }
  }
  
  // Asegurarse de que el puntaje esté entre 0-100
  score = Math.max(0, Math.min(100, Math.round(score)));
  
  return {
    score: score,
    comments: comments
  };
}

/**
 * Compara dos archivos STL byte por byte para analizar diferencias
 * @param {Blob} refBlob - Archivo STL de referencia
 * @param {Blob} compBlob - Archivo STL a comparar
 * @return {Object} Objeto con las diferencias encontradas
 */
function compareSTLFiles(refBlob, compBlob) {
  try {
    // Convertir los blobs a bytes
    var refBytes = refBlob.getBytes();
    var compBytes = compBlob.getBytes();
    
    // Inicializar objeto de respuesta
    var result = {
      filesValid: true,
      differences: {
        general: [],
        gsiHeader: {},
        ttiBlocks: []
      },
      trailingCharsAnalysis: [] // Añadir análisis de caracteres adicionales
    };
    
    // Comparar tamaños de archivo
    if (refBytes.length !== compBytes.length) {
      result.differences.general.push("Tamaños de archivo diferentes: " + refBytes.length + " vs " + compBytes.length);
    }
    
    // Analizar GSI Header (primeros 1024 bytes)
    var gsiRefBytes = refBytes.slice(0, 1024);
    var gsiCompBytes = compBytes.slice(0, 1024);
    
    // Comparar campos específicos del GSI Header
    
    // 1. CPN (Code Page Number) - bytes 0-2
    var cpnRef = bytesToString(gsiRefBytes.slice(0, 3));
    var cpnComp = bytesToString(gsiCompBytes.slice(0, 3));
    if (cpnRef !== cpnComp) {
      result.differences.gsiHeader.codePageNumber = {
        reference: cpnRef,
        compared: cpnComp
      };
    }
    
    // 2. DFC (Disk Format Code) - bytes 3-10
    var dfcRef = bytesToString(gsiRefBytes.slice(3, 11));
    var dfcComp = bytesToString(gsiCompBytes.slice(3, 11));
    if (dfcRef !== dfcComp) {
      result.differences.gsiHeader.diskFormatCode = {
        reference: dfcRef,
        compared: dfcComp
      };
    }
    
    // 3. DSC (Display Standard Code) - byte 11
    var dscRef = gsiRefBytes[11];
    var dscComp = gsiCompBytes[11];
    if (dscRef !== dscComp) {
      result.differences.gsiHeader.displayStandardCode = {
        reference: String.fromCharCode(dscRef),
        compared: String.fromCharCode(dscComp)
      };
    }
    
    // 4. CCT (Character Code Table) - bytes 12-13
    var cctRef = bytesToString(gsiRefBytes.slice(12, 14));
    var cctComp = bytesToString(gsiCompBytes.slice(12, 14));
    if (cctRef !== cctComp) {
      result.differences.gsiHeader.characterCodeTableNumber = {
        reference: cctRef,
        compared: cctComp
      };
    }
    
    // 5. LC (Language Code) - bytes 14-16
    var lcRef = bytesToString(gsiRefBytes.slice(14, 17));
    var lcComp = bytesToString(gsiCompBytes.slice(14, 17));
    if (lcRef !== lcComp) {
      result.differences.gsiHeader.languageCode = {
        reference: lcRef,
        compared: lcComp
      };
    }
    
    // Analizar bloques TTI (cada 128 bytes después del GSI)
    var refTTICount = Math.floor((refBytes.length - 1024) / 128);
    var compTTICount = Math.floor((compBytes.length - 1024) / 128);
    
    // Comparar número de bloques TTI
    if (refTTICount !== compTTICount) {
      result.differences.general.push("Número diferente de bloques TTI: " + refTTICount + " vs " + compTTICount);
    }
    
    // Comparar cada bloque TTI disponible
    var maxTTI = Math.min(refTTICount, compTTICount);
    for (var i = 0; i < maxTTI; i++) {
      var ttiDiff = compareTTIBlock(
        refBytes.slice(1024 + i * 128, 1024 + (i + 1) * 128),
        compBytes.slice(1024 + i * 128, 1024 + (i + 1) * 128),
        i
      );
      
      // Si hay caracteres adicionales, agregarlos al análisis específico
      if (ttiDiff.trailingChars) {
        result.trailingCharsAnalysis.push({
          subtitleNumber: i,
          etxPosition: ttiDiff.trailingChars.etxPosition,
          count: ttiDiff.trailingChars.count,
          bytes: ttiDiff.trailingChars.bytes
        });
      }
      
      if (Object.keys(ttiDiff).length > 1) { // Si hay más que solo subtitleNumber
        result.differences.ttiBlocks.push(ttiDiff);
      }
    }
    
    return result;
  } catch (error) {
    Logger.log("Error al comparar archivos STL: " + error);
  return {
      filesValid: false,
      error: error.toString()
    };
  }
}

/**
 * Compara un bloque TTI entre dos archivos STL
 * @param {Uint8Array} refTTI - Bloque TTI de referencia
 * @param {Uint8Array} compTTI - Bloque TTI a comparar
 * @param {Number} index - Índice del bloque TTI
 * @return {Object} Diferencias encontradas en el bloque TTI
 */
function compareTTIBlock(refTTI, compTTI, index) {
  var result = {
    subtitleNumber: index
  };
  
  // 1. Extraer número de subtítulo (bytes 1-2, little-endian)
  var snRef = refTTI[1] + (refTTI[2] << 8);
  var snComp = compTTI[1] + (compTTI[2] << 8);
  if (snRef !== snComp) {
    result.subtitleNumberValue = {
      reference: snRef,
      compared: snComp
    };
  }
  
  // 2. Comparar Timecode IN (bytes 5-8)
  var tcInRef = {
    hours: refTTI[5],
    minutes: refTTI[6],
    seconds: refTTI[7],
    frames: refTTI[8]
  };
  
  var tcInComp = {
    hours: compTTI[5],
    minutes: compTTI[6],
    seconds: compTTI[7],
    frames: compTTI[8]
  };
  
  if (JSON.stringify(tcInRef) !== JSON.stringify(tcInComp)) {
    result.timecodeIn = {};
    if (tcInRef.hours !== tcInComp.hours) {
      result.timecodeIn.hours = {
        reference: tcInRef.hours,
        compared: tcInComp.hours
      };
    }
    if (tcInRef.minutes !== tcInComp.minutes) {
      result.timecodeIn.minutes = {
        reference: tcInRef.minutes,
        compared: tcInComp.minutes
      };
    }
    if (tcInRef.seconds !== tcInComp.seconds) {
      result.timecodeIn.seconds = {
        reference: tcInRef.seconds,
        compared: tcInComp.seconds
      };
    }
    if (tcInRef.frames !== tcInComp.frames) {
      result.timecodeIn.frames = {
        reference: tcInRef.frames,
        compared: tcInComp.frames
      };
    }
  }
  
  // 3. Comparar Timecode OUT (bytes 9-12)
  var tcOutRef = {
    hours: refTTI[9],
    minutes: refTTI[10],
    seconds: refTTI[11],
    frames: refTTI[12]
  };
  
  var tcOutComp = {
    hours: compTTI[9],
    minutes: compTTI[10],
    seconds: compTTI[11],
    frames: compTTI[12]
  };
  
  if (JSON.stringify(tcOutRef) !== JSON.stringify(tcOutComp)) {
    result.timecodeOut = {};
    if (tcOutRef.hours !== tcOutComp.hours) {
      result.timecodeOut.hours = {
        reference: tcOutRef.hours,
        compared: tcOutComp.hours
      };
    }
    if (tcOutRef.minutes !== tcOutComp.minutes) {
      result.timecodeOut.minutes = {
        reference: tcOutRef.minutes,
        compared: tcOutComp.minutes
      };
    }
    if (tcOutRef.seconds !== tcOutComp.seconds) {
      result.timecodeOut.seconds = {
        reference: tcOutRef.seconds,
        compared: tcOutComp.seconds
      };
    }
    if (tcOutRef.frames !== tcOutComp.frames) {
      result.timecodeOut.frames = {
        reference: tcOutRef.frames,
        compared: tcOutComp.frames
      };
    }
  }
  
  // 4. Comparar posición vertical (byte 13)
  if (refTTI[13] !== compTTI[13]) {
    result.verticalPosition = {
      reference: refTTI[13],
      compared: compTTI[13]
    };
  }
  
  // 5. Comparar justificación (byte 14)
  if (refTTI[14] !== compTTI[14]) {
    result.justificationCode = {
      reference: refTTI[14],
      compared: compTTI[14]
    };
  }
  
  // 6. Comparar texto del subtítulo (bytes 16-127)
  // Primero determinamos la longitud real del texto en cada caso y la posición del ETX
  var refTextLength = 0;
  var compTextLength = 0;
  var refEtxPosition = -1;
  var compEtxPosition = -1;
  
  for (var i = 16; i < 128; i++) {
    if (refTTI[i] === 0x03) { // ETX
      refTextLength = i - 16;
      refEtxPosition = i;
      break;
    }
    if (refTTI[i] === 0xFF) { // Espacio no usado
      refTextLength = i - 16;
      break;
    }
    if (i === 127) {
      refTextLength = 112; // Todo el espacio disponible
    }
  }
  
  for (var i = 16; i < 128; i++) {
    if (compTTI[i] === 0x03) { // ETX
      compTextLength = i - 16;
      compEtxPosition = i;
      break;
    }
    if (compTTI[i] === 0xFF) { // Espacio no usado
      compTextLength = i - 16;
      break;
    }
    if (i === 127) {
      compTextLength = 112; // Todo el espacio disponible
    }
  }
  
  // Comparar longitudes
  if (refTextLength !== compTextLength) {
    result.textLength = {
      reference: refTextLength,
      compared: compTextLength
    };
  }
  
  // Extraer previsualizaciones de texto
  var refTextPreview = bytesToString(refTTI.slice(16, 16 + Math.min(refTextLength, 40)));
  var compTextPreview = bytesToString(compTTI.slice(16, 16 + Math.min(compTextLength, 40)));
  
  // Si los textos son diferentes, incluir previsualizaciones
  if (refTextPreview !== compTextPreview) {
    result.textPreview = {
      reference: refTextPreview,
      compared: compTextPreview
    };
    
    // Incluir diferencias de bytes específicas (hasta 5 primeras diferencias)
    var textDiffs = [];
    var maxComparableLength = Math.min(refTextLength, compTextLength);
    var diffCount = 0;
    
    for (var i = 0; i < maxComparableLength; i++) {
      if (refTTI[16 + i] !== compTTI[16 + i] && diffCount < 5) {
        textDiffs.push({
          position: i,
          reference: "0x" + refTTI[16 + i].toString(16).toUpperCase(),
          compared: "0x" + compTTI[16 + i].toString(16).toUpperCase()
        });
        diffCount++;
      }
    }
    
    if (textDiffs.length > 0) {
      result.textDiffs = textDiffs;
    }
  }
  
  // Detectar caracteres adicionales después del ETX en el archivo comparado
  if (compEtxPosition > -1 && compEtxPosition < 127) {
    // Verificar si hay caracteres no-0xFF después del ETX
    var hasTrailingChars = false;
    var trailingBytes = [];
    
    for (var i = compEtxPosition + 1; i < 128; i++) {
      if (compTTI[i] !== 0xFF) {
        hasTrailingChars = true;
        trailingBytes.push("0x" + compTTI[i].toString(16).toUpperCase());
      }
    }
    
    if (hasTrailingChars) {
      result.trailingChars = {
        etxPosition: compEtxPosition - 16, // Posición relativa al inicio del texto
        count: trailingBytes.length,
        bytes: trailingBytes.join(', ')
      };
    }
  }
  
  return result;
}

/**
 * Convierte un array de bytes a una cadena de texto
 * @param {Uint8Array} bytes - Array de bytes
 * @return {String} Cadena de texto resultante
 */
function bytesToString(bytes) {
  var result = "";
  for (var i = 0; i < bytes.length; i++) {
    // Solo incluir caracteres imprimibles (>= 32)
    if (bytes[i] >= 32 && bytes[i] <= 126) {
      result += String.fromCharCode(bytes[i]);
    } else if (bytes[i] === 0x8A) {
      result += "\n"; // Representar EOL como salto de línea
    }
  }
  return result;
}

/**
 * Muestra una interfaz para diagnosticar problemas en archivos STL
 * Esta función crea una página HTML para analizar y comparar archivos STL
 * @return {HtmlOutput} Página HTML con la herramienta de diagnóstico
 */
function showSTLDiagnosticTool() {
  // Crear la interfaz HTML
  var htmlTemplate = HtmlService.createTemplateFromFile('Diagnostico');
  
  // Pasar variables al template
  htmlTemplate.scriptUrl = getScriptUrl();
  
  var html = htmlTemplate.evaluate()
      .setTitle('Herramienta de Diagnóstico STL')
      .setWidth(900)
      .setHeight(700);
  
  return html;
}

/**
 * Endpoint para recibir y analizar archivos STL subidos por el usuario
 * @param {Object} formData - Datos del formulario con el archivo STL
 * @return {Object} Resultado del análisis
 */
function processSTLDiagnostic(formData) {
  try {
    // Obtener el archivo a partir de la información Base64
    var fileData = formData.fileContent;
    var fileBytes = Utilities.base64Decode(fileData.split(',')[1]);
    var stlBlob = Utilities.newBlob(fileBytes, 'application/octet-stream', formData.fileName);
    
    // Realizar el análisis
    var analysis = analyzeLegacySTL(stlBlob);
    
    return {
      success: true,
      analysis: analysis
    };
  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Endpoint para comparar dos archivos STL
 * @param {Object} formData - Datos del formulario con ambos archivos STL
 * @return {Object} Resultado de la comparación
 */
function compareSTLDiagnostic(formData) {
  try {
    // Obtener el primer archivo (referencia)
    var fileData1 = formData.fileContent1;
    var fileBytes1 = Utilities.base64Decode(fileData1.split(',')[1]);
    var stlBlob1 = Utilities.newBlob(fileBytes1, 'application/octet-stream', formData.fileName1);
    
    // Obtener el segundo archivo (a comparar)
    var fileData2 = formData.fileContent2;
    var fileBytes2 = Utilities.base64Decode(fileData2.split(',')[1]);
    var stlBlob2 = Utilities.newBlob(fileBytes2, 'application/octet-stream', formData.fileName2);
    
    // Realizar la comparación
    var comparison = compareSTLFiles(stlBlob1, stlBlob2);
    
    return {
      success: true,
      comparison: comparison
    };
  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Crea un bloque TTI estándar optimizado sin caracteres adicionales después del ETX
 * 
 * @param {Number} subtitleNumber - Número de subtítulo
 * @param {String} timecodeIn - Código de tiempo de entrada
 * @param {String} timecodeOut - Código de tiempo de salida
 * @param {String} text - Texto del subtítulo
 * @param {Number} justification - Código de justificación (1=izq, 2=centro, 3=der)
 * @return {Array} Array de bytes representando el bloque TTI
 */
function createFixedLegacyTTIBlock(subtitleNumber, timecodeIn, timecodeOut, text, justification) {
  // Crear un buffer de 128 bytes inicializado con 0x00 (en lugar de 0xFF) para evitar problemas con editores de subtítulos
  var buffer = new Array(128);
  for (var i = 0; i < buffer.length; i++) {
    buffer[i] = 0x00;
  }
  
  // SGN - Subtitle Group Number (bytes 0-1) - Little-endian
  buffer[0] = 0x00;
  buffer[1] = 0x00;
  
  // SN - Subtitle Number (bytes 2-3) - Little-endian
  // IMPORTANTE: El número de subtítulo debe estar en formato little-endian
  // El primer byte es el menos significativo (LSB)
  buffer[2] = subtitleNumber & 0xFF;         // LSB (byte menos significativo)
  buffer[3] = (subtitleNumber >> 8) & 0xFF;  // MSB (byte más significativo)
  
  // EBN - Extension Block Number (byte 4)
  buffer[4] = 0xFF;
  
  // CS - Cumulative Status (byte 5)
  buffer[5] = 0x00;
  
  // Procesar timecode IN (bytes 6-9)
  var tcIn = parseTimecode(timecodeIn);
  buffer[6] = tcIn.hours;
  buffer[7] = tcIn.minutes;
  buffer[8] = tcIn.seconds;
  buffer[9] = tcIn.frames;
  
  // Procesar timecode OUT (bytes 10-13)
  var tcOut = parseTimecode(timecodeOut);
  buffer[10] = tcOut.hours;
  buffer[11] = tcOut.minutes;
  buffer[12] = tcOut.seconds;
  buffer[13] = tcOut.frames;
  
  // VP - Vertical Position (byte 14)
  // Calcular posición vertical según el número de líneas
  var lines = text.split('\n').length;
  buffer[14] = calculateVerticalPosition(lines);
  
  // JC - Justification Code (byte 15)
  // IMPORTANTE: Usar el valor correcto para la justificación (2 = centrado)
  buffer[15] = justification || 2; // 2=centro por defecto
  
  // CF - Comment Flag (byte 16)
  buffer[16] = 0x00;
  
  // TF - Text Field (bytes 17-127)
  // Procesar etiquetas HTML como <br>
  var processedText = text.replace(/<br\s*\/?>/gi, "\n");
  processedText = processedText.replace(/<[^>]*>/g, ""); // Eliminar otras etiquetas HTML
  
  // Posición actual en el buffer para escribir texto
  var position = 17;
  
  // Procesar cada carácter del texto
  for (var i = 0; i < processedText.length && position < 127; i++) {
    var char = processedText.charAt(i);
    
      if (char === '\n') {
      // Salto de línea como EOL en EBU
      buffer[position++] = 0x8A;
    } else {
      // Usar la función getEBUCodeForChar para mapear correctamente los caracteres
      buffer[position++] = getEBUCodeForChar(char);
    }
  }
  
  // Añadir terminador ETX (End of Text, 0x03) SIEMPRE después del último carácter
  if (position < 128) {
    buffer[position++] = 0x03; // ETX - End of Text
    for (var i = position; i < 128; i++) {
      buffer[i] = 0x00; // Usar 0x00 en lugar de 0xFF para compatibilidad con todos los editores
    }
  }
  
  return buffer;
} 

/**
 * Genera un archivo STL basado en un modelo de referencia exacto,
 * pero utilizando los datos reales del Excel
 * 
 * @param {Array} data - Array con datos de subtítulos extraídos del Excel
 * @param {Object} metadata - Objeto con metadatos del programa
 * @return {Blob} Blob binario con el archivo STL generado
 */
function generateExactReferenceSTL(data, metadata) {
  try {
    Logger.log("Generando archivo STL con modelo de referencia exacto usando datos reales");
    Logger.log("Datos recibidos: " + (data ? data.length : 0) + " subtítulos");
    Logger.log("Metadatos: " + JSON.stringify(metadata));
    
    // Array para almacenar bytes del STL
    var stlBytes = [];
    
    // IMPORTANTE: Utilizar el header exacto del modelo de referencia
    var gsiHeader = createExactGSIHeader(metadata);
    stlBytes = stlBytes.concat(gsiHeader);
    
    // El número de subtítulos está limitado por el número de entradas en el Excel
    var subtitleCount = 0;
    
    // Crear un texto para el título (subtítulo 0)
    var titleText = '';
    
    // Si tenemos metadatos, incluir información del título
    if (metadata) {
      if (metadata.spanishTitle || metadata.titleSpanish) {
        titleText = (metadata.spanishTitle || metadata.titleSpanish).toUpperCase();
      } else if (metadata.englishTitle || metadata.titleEnglish) {
        titleText = (metadata.englishTitle || metadata.titleEnglish).toUpperCase();
      } else if (metadata.episodeNumber) {
        titleText = 'EPISODIO ' + metadata.episodeNumber;
      } else {
        titleText = 'SUBTITULOS';
      }
    } else {
      titleText = 'SUBTITULOS';
    }
    
    // Limitar longitud del título
    if (titleText.length > 40) {
      titleText = titleText.substring(0, 40);
    }
    
    // ENFOQUE ALTERNATIVO: Dividir el título en dos bloques separados
    if (titleText.length > 20) {
      // Encontrar un punto natural para dividir el título
      var parts = titleText.split(/\s+/);
      var firstPart = "";
      var secondPart = "";
      var middleIndex = Math.floor(parts.length / 2);
      
      for (var i = 0; i < parts.length; i++) {
        if (i < middleIndex) {
          firstPart += parts[i] + " ";
        } else {
          secondPart += parts[i] + " ";
        }
      }
      
      firstPart = firstPart.trim();
      secondPart = secondPart.trim();
      
      Logger.log("DEBUG - Título dividido manualmente: " + firstPart + "¶" + secondPart);
      
      // NUEVO ENFOQUE: Crear un TTI block con dos líneas reales usando el separador ¶
      var combinedText = firstPart + "\n" + secondPart;
      var titleBlock = createExactTTIBlock(subtitleCount, {
        timecodeIn: "01:00:00:15",
        timecodeOut: "01:00:04:15",
        text: combinedText
      });
      
      stlBytes = stlBytes.concat(titleBlock);
      subtitleCount++;
      
      Logger.log("Título (bloque " + (subtitleCount-1) + ") - Final: " + debugBuffer(titleBlock));
    } else {
      // Si el título es corto, usar un solo bloque
      var titleBlock = createExactTTIBlock(subtitleCount, {
      timecodeIn: "01:00:00:15",
      timecodeOut: "01:00:04:15",
      text: titleText
    });
      
    stlBytes = stlBytes.concat(titleBlock);
      subtitleCount++;
    }
    
    // Procesamos los subtítulos
    if (data && data.length > 0) {
      Logger.log("Tipo de datos: " + (data[0].timecodeIn !== undefined ? "Objetos de subtítulos" : "Arrays"));
      
      // Si data contiene objetos con propiedades de subtítulo
      if (data[0] && data[0].timecodeIn !== undefined) {
        // Este es el nuevo formato de datos
        for (var i = 0; i < data.length; i++) {
          var subtitle = data[i];
          
          // Verificar datos mínimos
          if (!subtitle.timecodeIn || !subtitle.timecodeOut || !subtitle.text) {
            Logger.log("Saltando subtítulo " + i + " por datos incompletos");
            continue;
          }
          
          // Log para depuración
          Logger.log("Subtítulo " + (subtitleCount) + ": " + subtitle.timecodeIn + " -> " + subtitle.timecodeOut + ": " + subtitle.text.substring(0, 30) + (subtitle.text.length > 30 ? "..." : ""));
          
          // ENFOQUE ALTERNATIVO: Crear dos bloques TTI separados para cada subtítulo
          // Primero analizamos el texto para ver si es una o dos líneas
          var processedText = preprocessSubtitleText(subtitle.text);
          var lines = processedText.split('\n');
          
          // Si el subtítulo ya tiene dos líneas, usamos el enfoque normal
          if (lines.length >= 2) {
            var ttiBlock = createExactTTIBlock(subtitleCount, {
              timecodeIn: subtitle.timecodeIn,
              timecodeOut: subtitle.timecodeOut,
              text: processedText
            });
            
            stlBytes = stlBytes.concat(ttiBlock);
            subtitleCount++;
          } 
          // Si solo tiene una línea, creamos ARTIFICIALMENTE un subtítulo de DOS LÍNEAS
          else if (lines.length === 1 && processedText.length > 10) {
            // Encontrar un punto natural para dividir el texto (como una coma, punto o espacio)
            var middlePoint = Math.floor(processedText.length / 2);
            var breakPoint = -1;
            
            // Buscar puntuación cerca del punto medio
            for (var j = middlePoint - 5; j <= middlePoint + 5 && j < processedText.length; j++) {
              if (j > 0 && /[,.;:!?]/.test(processedText.charAt(j-1))) {
                breakPoint = j;
                break;
              }
            }
            
            // Si no hay puntuación, buscar un espacio
            if (breakPoint === -1) {
              for (var j = middlePoint; j > 0 && j > middlePoint - 10; j--) {
                if (processedText.charAt(j) === ' ') {
                  breakPoint = j;
                  break;
                }
              }
              
              // Si no encontramos espacio hacia atrás, buscar hacia adelante
              if (breakPoint === -1) {
                for (var j = middlePoint; j < processedText.length && j < middlePoint + 10; j++) {
                  if (processedText.charAt(j) === ' ') {
                    breakPoint = j;
                    break;
                  }
                }
              }
            }
            
            // Si encontramos un punto de división natural, dividimos el texto
            if (breakPoint !== -1) {
              var firstLine = processedText.substring(0, breakPoint).trim();
              var secondLine = processedText.substring(breakPoint).trim();
              
              // Combinar ambas líneas con un salto de línea real
              var combinedText = firstLine + "\n" + secondLine;
              
              var ttiBlock = createExactTTIBlock(subtitleCount, {
                timecodeIn: subtitle.timecodeIn,
                timecodeOut: subtitle.timecodeOut,
                text: combinedText // Texto con salto de línea explícito
              });
              
              stlBytes = stlBytes.concat(ttiBlock);
              subtitleCount++;
            } else {
              // Si no podemos encontrar un punto natural, simplemente dividimos por la mitad
              var halfLength = Math.floor(processedText.length / 2);
              var firstLine = processedText.substring(0, halfLength).trim();
              var secondLine = processedText.substring(halfLength).trim();
              
              // Combinar ambas líneas con un salto de línea real
              var combinedText = firstLine + "\n" + secondLine;
              
              var ttiBlock = createExactTTIBlock(subtitleCount, {
                timecodeIn: subtitle.timecodeIn,
                timecodeOut: subtitle.timecodeOut,
                text: combinedText // Texto con salto de línea explícito
              });
              
              stlBytes = stlBytes.concat(ttiBlock);
              subtitleCount++;
            }
          } else {
            // Para textos muy cortos, simplemente los dejamos como están
            var ttiBlock = createExactTTIBlock(subtitleCount, {
              timecodeIn: subtitle.timecodeIn,
              timecodeOut: subtitle.timecodeOut,
              text: processedText
            });
            
            stlBytes = stlBytes.concat(ttiBlock);
            subtitleCount++;
          }
        }
      } else {
        // Formato antiguo (array de arrays)
        // Configurar columnas (según el formato estándar)
        var columnConfig = {
          timeIn: getConfig('excel').columns.timeIn,
          text: getConfig('excel').columns.text,
          timeOut: getConfig('excel').columns.timeOut
        };
        
    for (var i = 0; i < data.length; i++) {
      // Extraer datos
      var timecodeIn, subtitleText, timecodeOut;
      
      // Si tiene al menos 4 columnas (formato Excel normal)
      if (data[i].length >= 4) {
        timecodeIn = data[i][columnConfig.timeIn];
        subtitleText = data[i][columnConfig.text];
        timecodeOut = data[i][columnConfig.timeOut];
      } 
      // Formato alternativo con 3 columnas
      else if (data[i].length >= 3) {
        timecodeIn = data[i][0];
        subtitleText = data[i][1];
        timecodeOut = data[i][2];
      }
      else {
        // Saltar esta fila si no tiene suficientes datos
        continue;
      }
      
      // Verificar datos mínimos
      if (!timecodeIn || !timecodeOut || !subtitleText) {
        continue;
      }
      
          // Log para depuración
          Logger.log("Subtítulo " + (subtitleCount) + ": " + timecodeIn + " -> " + timecodeOut + ": " + subtitleText.substring(0, 30) + (subtitleText.length > 30 ? "..." : ""));
          
          // ENFOQUE ALTERNATIVO: Crear dos bloques TTI separados para cada subtítulo
          // Primero analizamos el texto para ver si es una o dos líneas
          var processedText = preprocessSubtitleText(subtitleText);
          var lines = processedText.split('\n');
          
          // Si el subtítulo ya tiene dos líneas, usamos el enfoque normal
          if (lines.length >= 2) {
            var ttiBlock = createExactTTIBlock(subtitleCount, {
        timecodeIn: timecodeIn,
        timecodeOut: timecodeOut,
              text: processedText
      });
      
            stlBytes = stlBytes.concat(ttiBlock);
      subtitleCount++;
          } 
          // Si solo tiene una línea, creamos ARTIFICIALMENTE un subtítulo de DOS LÍNEAS
          else if (lines.length === 1 && processedText.length > 10) {
            // Encontrar un punto natural para dividir el texto (como una coma, punto o espacio)
            var middlePoint = Math.floor(processedText.length / 2);
            var breakPoint = -1;
            
            // Buscar puntuación cerca del punto medio
            for (var j = middlePoint - 5; j <= middlePoint + 5 && j < processedText.length; j++) {
              if (j > 0 && /[,.;:!?]/.test(processedText.charAt(j-1))) {
                breakPoint = j;
                break;
              }
            }
            
            // Si no hay puntuación, buscar un espacio
            if (breakPoint === -1) {
              for (var j = middlePoint; j > 0 && j > middlePoint - 10; j--) {
                if (processedText.charAt(j) === ' ') {
                  breakPoint = j;
                  break;
                }
              }
              
              // Si no encontramos espacio hacia atrás, buscar hacia adelante
              if (breakPoint === -1) {
                for (var j = middlePoint; j < processedText.length && j < middlePoint + 10; j++) {
                  if (processedText.charAt(j) === ' ') {
                    breakPoint = j;
                    break;
                  }
                }
              }
            }
            
            // Si encontramos un punto de división natural, dividimos el texto
            if (breakPoint !== -1) {
              var firstLine = processedText.substring(0, breakPoint).trim();
              var secondLine = processedText.substring(breakPoint).trim();
              
              // Combinar ambas líneas con un salto de línea real
              var combinedText = firstLine + "\n" + secondLine;
              
              var ttiBlock = createExactTTIBlock(subtitleCount, {
                timecodeIn: timecodeIn,
                timecodeOut: timecodeOut,
                text: combinedText // Texto con salto de línea explícito
              });
              
              stlBytes = stlBytes.concat(ttiBlock);
              subtitleCount++;
            } else {
              // Si no podemos encontrar un punto natural, simplemente dividimos por la mitad
              var halfLength = Math.floor(processedText.length / 2);
              var firstLine = processedText.substring(0, halfLength).trim();
              var secondLine = processedText.substring(halfLength).trim();
              
              // Combinar ambas líneas con un salto de línea real
              var combinedText = firstLine + "\n" + secondLine;
              
              var ttiBlock = createExactTTIBlock(subtitleCount, {
                timecodeIn: timecodeIn,
                timecodeOut: timecodeOut,
                text: combinedText // Texto con salto de línea explícito
              });
              
              stlBytes = stlBytes.concat(ttiBlock);
              subtitleCount++;
            }
          } else {
            // Para textos muy cortos, simplemente los dejamos como están
            var ttiBlock = createExactTTIBlock(subtitleCount, {
              timecodeIn: timecodeIn,
              timecodeOut: timecodeOut,
              text: processedText
            });
            
            stlBytes = stlBytes.concat(ttiBlock);
            subtitleCount++;
          }
        }
      }
    }
    
    // Actualizar el número de subtítulos en el GSI header (bytes 242-246)
    var tnbString = subtitleCount.toString().padStart(5, '0');
    for (var i = 0; i < 5; i++) {
      stlBytes[242 + i] = tnbString.charCodeAt(i);
    }
    
    // DEBUG: Verificación final del STL
    Logger.log("DEBUG - Tamaño total del STL: " + stlBytes.length + " bytes");
    Logger.log("DEBUG - Número total de subtítulos: " + subtitleCount);
    
    // Convertir el array a un Blob binario
    var byteArray = new Uint8Array(stlBytes);
    var blob = Utilities.newBlob(byteArray, 'application/octet-stream', 'subtitles.stl');
    
    Logger.log("Archivo STL generado correctamente con " + subtitleCount + " subtítulos");
    return blob;
  } catch (error) {
    Logger.log('Error al generar STL: ' + error.toString());
    throw new Error('Error al generar archivo STL: ' + error.message);
  }
}

/**
 * Analiza el contenido original del subtítulo para verificar si contiene etiquetas HTML
 * @param {String} text - Texto original del subtítulo
 * @return {Object} Resultado del análisis
 */
function analyzeSubtitleText(text) {
  var result = {
    containsHtml: false,
    colorTags: 0,
    brTags: 0,
    otherTags: 0,
    tagTypes: [],
    htmlPattern: null,
    recommendation: ""
  };
  
  // Verificar si contiene etiquetas HTML
  if (/<[^>]+>/.test(text)) {
    result.containsHtml = true;
    
    // Contar etiquetas de color
    var colorMatches = text.match(/<font\s+color=['"][^'"]*['"][^>]*>/gi);
    result.colorTags = colorMatches ? colorMatches.length : 0;
    
    // Contar etiquetas BR
    var brMatches = text.match(/<br\s*\/?>/gi);
    result.brTags = brMatches ? brMatches.length : 0;
    
    // Contar otras etiquetas
    var otherMatches = text.match(/<(?!font|br)[a-z]+[^>]*>/gi);
    result.otherTags = otherMatches ? otherMatches.length : 0;
    
    // Identificar tipos de etiquetas
    var tagRegex = /<([a-z]+)[^>]*>/gi;
    var match;
    var tagTypes = new Set();
    
    while ((match = tagRegex.exec(text)) !== null) {
      tagTypes.add(match[1].toLowerCase());
    }
    
    result.tagTypes = Array.from(tagTypes);
    
    // Determinar el patrón HTML
    if (result.colorTags > 0 && result.colorTags % 2 === 0) {
      // Buscar patrones específicos que causan problemas
      if (text.match(/<font color=['"][^'"]+['"]><\/font>\s*<font color=/gi)) {
        result.htmlPattern = "EMPTY_COLOR_TAGS";
        result.recommendation = "Convertir etiquetas de color vacías a saltos de línea";
      } else if (text.match(/<\/font><font\s+color=/gi)) {
        result.htmlPattern = "ADJACENT_COLOR_TAGS";
        result.recommendation = "Insertar saltos de línea entre etiquetas de color adyacentes";
      } else {
        result.htmlPattern = "NORMAL_COLOR_TAGS";
        result.recommendation = "Preservar etiquetas de color o convertir a formato STL";
      }
    } else if (result.brTags > 0) {
      result.htmlPattern = "BR_TAGS";
      result.recommendation = "Convertir <br> a saltos de línea reales";
    } else {
      result.htmlPattern = "MIXED_HTML";
      result.recommendation = "Limpiar HTML y preservar formato significativo";
    }
  }
  
  return result;
}

/**
 * Pre-procesa el texto del subtítulo para manejar correctamente las etiquetas HTML
 * @param {String} text - Texto del subtítulo con posibles etiquetas HTML
 * @return {String} Texto limpio sin etiquetas HTML
 */
function preprocessSubtitleText(text) {
  if (!text) return "";
  
  // Registrar el texto original para debugging
  Logger.log("Texto original: " + text);
  
  // SIMPLIFICACIÓN RADICAL: Limpieza completa de HTML
  
  // 1. Eliminar etiquetas \an y similares
  text = text.replace(/\{\\an\d+\}/g, "");
  text = text.replace(/<\\an\d+>/g, "");
  text = text.replace(/\\an\d+/g, "");
  
  // 2. Convertir <br> a saltos de línea
  text = text.replace(/<br\s*\/?>/gi, "\n");
  
  // 3. Eliminar etiquetas font con su contenido (preservando el texto dentro)
  text = text.replace(/<font[^>]*>(.*?)<\/font>/gi, "$1");
  
  // 4. Eliminar cualquier etiqueta HTML restante
  text = text.replace(/<[^>]*>/g, "");
  
  // 5. Limpiar espacios en blanco excesivos pero preservar saltos de línea
  text = text.replace(/[ \t]+/g, " ");
  text = text.split("\n").map(line => line.trim()).filter(line => line).join("\n");
  
  // Registro final del texto procesado
  Logger.log("Texto procesado: " + text.replace(/\n/g, "¶"));
  
  return text;
}

/**
 * Función de depuración para imprimir un buffer en formato hexadecimal
 * @param {Array} buffer - Buffer a imprimir
 * @param {String} nombre - Nombre identificativo para el log
 */
function debugBuffer(buffer, nombre) {
  var hexString = "";
  var asciiString = "";
  
  for (var i = 0; i < buffer.length; i++) {
    // Imprimir índice cada 16 bytes
    if (i % 16 === 0) {
      if (i > 0) {
        Logger.log(hexString + " | " + asciiString);
        hexString = "";
        asciiString = "";
      }
      hexString = i.toString(16).padStart(4, '0') + ": ";
    }
    
    // Añadir byte en hexadecimal
    hexString += buffer[i].toString(16).padStart(2, '0') + " ";
    
    // Añadir caracter ASCII si es imprimible
    if (buffer[i] >= 32 && buffer[i] <= 126) {
      asciiString += String.fromCharCode(buffer[i]);
    } else if (buffer[i] === 0x8A) {
      // Representar salto de línea EBU con un símbolo especial
      asciiString += "¶"; // símbolo de párrafo para representar salto de línea
    } else {
      asciiString += "."; // Punto para bytes no imprimibles
    }
  }
  
  // Imprimir la última línea
  if (hexString) {
    Logger.log(nombre + " - Final: " + hexString + " | " + asciiString);
  }
}

/**
 * Crea un bloque TTI exacto siguiendo el modelo de referencia
 * Esta función crea un bloque TTI idéntico al formato esperado por Subtitle Edit,
 * pero procesa correctamente las etiquetas HTML de formato
 * 
 * @param {Number} subtitleNumber - Número del subtítulo (0 para el título, 1+ para subtítulos)
 * @param {Object} subtitle - Objeto con información del subtítulo (timecodeIn, timecodeOut, text)
 * @return {Array} Array de bytes representando el bloque TTI exacto
 */
function createExactTTIBlock(subtitleNumber, subtitle) {
  var buffer = new Array(128);
  
  // Inicializar todo el buffer con 0x00
  for (var i = 0; i < 128; i++) {
    buffer[i] = 0x00;
  }
  
  // SGN - Subtitle Group Number (bytes 0-1)
  buffer[0] = 0x00;
  buffer[1] = 0x00;
  
  // SN - Número de subtítulo (bytes 2-3) - formato little-endian
  // IMPORTANTE: Asegurarnos de que cada subtítulo tenga un número único
  buffer[2] = subtitleNumber & 0xFF;
  buffer[3] = (subtitleNumber >> 8) & 0xFF;
  
  // EBN - Extension Block Number (byte 4)
  buffer[4] = 0xFF;
  
  // CS - Cumulative Status (byte 5) - CRÍTICO: 0 para no acumulativo
  // Establecer explícitamente a 0 para que cada subtítulo sea independiente
  buffer[5] = 0x00;
  
  // Timecode IN (bytes 6-9)
  var tcIn = parseTimecode(subtitle.timecodeIn);
  buffer[6] = tcIn.hours;
  buffer[7] = tcIn.minutes;
  buffer[8] = tcIn.seconds;
  buffer[9] = tcIn.frames;
  
  // Timecode OUT (bytes 10-13)
  var tcOut = parseTimecode(subtitle.timecodeOut);
  buffer[10] = tcOut.hours;
  buffer[11] = tcOut.minutes;
  buffer[12] = tcOut.seconds;
  buffer[13] = tcOut.frames;
  
  // Asegurarnos de que los timecodes sean válidos y diferentes
  if (tcIn.hours === tcOut.hours && 
      tcIn.minutes === tcOut.minutes && 
      tcIn.seconds === tcOut.seconds && 
      tcIn.frames === tcOut.frames) {
    // Si los timecodes son iguales, hacerlos diferentes
    Logger.log("ADVERTENCIA: Timecodes iguales, ajustando timeOut");
    buffer[13] = (tcOut.frames + 10) % 25; // Añadir 10 frames
    if (buffer[13] < tcOut.frames) {
      // Si pasamos al siguiente segundo
      buffer[12] = (tcOut.seconds + 1) % 60;
      if (buffer[12] < tcOut.seconds) {
        // Si pasamos al siguiente minuto
        buffer[11] = (tcOut.minutes + 1) % 60;
        if (buffer[11] < tcOut.minutes) {
          // Si pasamos a la siguiente hora
          buffer[10] = (tcOut.hours + 1) % 24;
        }
      }
    }
  }
  
  // VP - Vertical Position (byte 14) - SIEMPRE usar 0x0E (14) para forzar múltiples líneas
  buffer[14] = 0x0E; // Valor 14 (0x0E) para múltiples líneas
  
  // JC - Justification Code (byte 15) - Centrado (2)
  buffer[15] = 0x02;
  
  // CF - Comment Flag (byte 16) - Normal subtitle
  buffer[16] = 0x00;
  
  // Text Field (bytes 17-127)
  // SIMPLIFICACIÓN: Usar directamente preprocessSubtitleText 
  var text = preprocessSubtitleText(subtitle.text);
  var position = 17;
  
  // Verificar si el texto tiene múltiples líneas
  var lines = text.split('\n');
  
  // Log para depuración
  Logger.log("DEBUG - Subtítulo #" + subtitleNumber + " - Líneas: " + lines.length);
  for (var i = 0; i < lines.length; i++) {
    Logger.log("DEBUG - Línea " + (i+1) + ": [" + lines[i] + "]");
  }
  
  // Asegurar que siempre tenemos al menos dos líneas para compatibilidad
  if (lines.length === 1 && text.length > 5) {
    // Buscar punto medio natural
    var middlePoint = Math.floor(text.length / 2);
    var breakPoint = -1;
    
    // Buscar espacio cercano al medio
    for (var i = middlePoint; i > 0; i--) {
      if (text.charAt(i) === ' ') {
        breakPoint = i;
        break;
      }
    }
    
    // Si no encontramos antes, buscar después
    if (breakPoint === -1) {
      for (var i = middlePoint; i < text.length; i++) {
        if (text.charAt(i) === ' ') {
          breakPoint = i;
          break;
        }
      }
    }
    
    // Si encontramos punto de división, aplicarlo
    if (breakPoint !== -1) {
      var firstLine = text.substring(0, breakPoint).trim();
      var secondLine = text.substring(breakPoint).trim();
      text = firstLine + "\n" + secondLine;
      lines = [firstLine, secondLine];
    } else {
      // Si no hay espacios, dividir por la mitad
      var halfLength = Math.floor(text.length / 2);
      var firstLine = text.substring(0, halfLength).trim();
      var secondLine = text.substring(halfLength).trim();
      text = firstLine + "\n" + secondLine;
      lines = [firstLine, secondLine];
    }
    
    Logger.log("DEBUG - Inserción FORZADA de salto de línea: [" + text.replace("\n", "¶") + "]");
  }
  
  // Registrar texto final
  Logger.log("DEBUG - Texto final a codificar: [" + text.replace(/\n/g, "¶") + "]");
  
  // Escribir el texto con códigos EBU
  for (var i = 0; i < text.length && position < 127; i++) {
    var char = text.charAt(i);
    
    if (char === '\n') {
      Logger.log("DEBUG - Insertando salto de línea (0x8A) en posición " + position);
      buffer[position++] = 0x8A; // EBU Latin Break
    } else {
      buffer[position++] = getEBUCodeForChar(char);
    }
  }
  
  // Terminar con ETX (0x03)
  if (position < 127) {
    Logger.log("DEBUG - Insertando ETX (0x03) en posición " + position);
    buffer[position++] = 0x03; // ETX
    
    // Rellenar el resto con 0x00
    for (var i = position; i < 128; i++) {
      buffer[i] = 0x00;
    }
  }
  
  // Debug información completa
  debugBuffer(buffer, "TTI Block " + subtitleNumber);
  
  return buffer;
}

/**
 * Crea un encabezado GSI exacto siguiendo la especificación EBU
 * 
 * @param {Object} metadata - Objeto con metadatos para el encabezado
 * @return {Array} Array de bytes representando el encabezado GSI exacto
 */
function createExactGSIHeader(metadata) {
  var buffer = new Array(1024);
  var config = getConfig();
  
  // Inicializamos el buffer con 0x20 (espacio) como indica la norma EBU
  for (var i = 0; i < 1024; i++) {
    buffer[i] = 0x20;
  }
  
  // 0-3: Code Page Number (CPN)
  writeASCIIToBuffer(buffer, 0, "850 ", 4);
  
  // 4-11: Disk Format Code (DFC)
  writeASCIIToBuffer(buffer, 4, "STL25.01", 8);
  
  // 12: Display Standard Code (DSC)
  // CRÍTICO: Usar código 0x30 (ASCII '0') para PAL 625
  buffer[12] = 0x30;
  
  // 13: Character Code Table (CCT)
  buffer[13] = 0x00;
  
  // 14-15: Language Code (LC) - Se usa el código ISO 639-2 del idioma
  writeASCIIToBuffer(buffer, 14, getLanguageCode(metadata.language), 2);
  
  // 16-47: Original Programme Title (OPT)
  var englishTitle = (metadata.englishTitle || "").slice(0, 32);
  writeASCIIToBuffer(buffer, 16, englishTitle, 32);
  
  // 48-79: Original Episode Title (OET)
  var spanishTitle = (metadata.spanishTitle || "").slice(0, 32);
  writeASCIIToBuffer(buffer, 48, spanishTitle, 32);
  
  // 80-111: Translated Programme Title (TPT)
  writeASCIIToBuffer(buffer, 80, spanishTitle, 32);
  
  // 112-143: Translated Episode Title (TET)
  var episodeTitle = (metadata.episodeNumber || "").slice(0, 32);
  writeASCIIToBuffer(buffer, 112, episodeTitle, 32);
  
  // 144-159: Translator Name (TN)
  writeASCIIToBuffer(buffer, 144, "MEDIA ACCESS CO.", 16);
  
  // 160-175: Translator Contact Details (TCD)
  writeASCIIToBuffer(buffer, 160, "DUBAPP", 16);
  
  // 176-207: Subtitle List Reference Code (SLR)
  var slr = "DUBAPP-STL-" + new Date().getTime().toString(16).slice(-8);
  writeASCIIToBuffer(buffer, 176, slr, 32);
  
  // 208-223: Creation Date (CD)
  var today = new Date();
  var cd = today.getFullYear().toString().slice(-2) + 
           ("0" + (today.getMonth() + 1)).slice(-2) + 
           ("0" + today.getDate()).slice(-2) + 
           ("0" + today.getHours()).slice(-2) + 
           ("0" + today.getMinutes()).slice(-2);
  writeASCIIToBuffer(buffer, 208, cd, 16);
  
  // 224-229: Revision Date (RD)
  writeASCIIToBuffer(buffer, 224, cd.slice(0, 6), 6);
  
  // 230-235: Revision Number (RN)
  writeASCIIToBuffer(buffer, 230, "01", 2);
  
  // 236-237: Total Number of TTI Blocks (TNB)
  // CRÍTICO: Indicar el número correcto de bloques TTI
  var totalBlocks = 100; // Valor predeterminado alto para asegurar que todos los subtítulos sean leídos
  writeASCIIToBuffer(buffer, 236, "00", 2);
  
  // 238-243: Total Number of Subtitles (TNS)
  // Dejamos esto en blanco o con un valor alto para que el editor procese todos los subtítulos
  writeASCIIToBuffer(buffer, 238, "005", 3);
  
  // 244-251: Total Number of Subtitle Groups (TNG)
  writeASCIIToBuffer(buffer, 244, "00", 2);
  
  // 252-253: Maximum Number of Displayable Characters (MNC)
  // CRÍTICO: Establecer un número bajo para forzar múltiples líneas
  writeASCIIToBuffer(buffer, 252, "030", 3);
  
  // 255-256: Maximum Number of Displayable Rows (MNR)
  // CRÍTICO: Establecer a 002 para forzar máximo 2 líneas
  writeASCIIToBuffer(buffer, 255, "002", 3);
  
  // 258-259: Time Code: Status (TCS)
  writeASCIIToBuffer(buffer, 258, "1", 1);
  
  // 260-273: Time Code: Start-of-Programme (TCP)
  writeASCIIToBuffer(buffer, 260, "00000000", 8);
  
  // 274-287: Time Code: First In-Cue (TCF)
  writeASCIIToBuffer(buffer, 274, "00000000", 8);
  
  // 288-299: Disk Format Code (DFC) de nuevo
  writeASCIIToBuffer(buffer, 288, "STL25.01", 8);
  
  // 300-1023: Campo reservado (se deja en espacios)
  
  // Establecer MSP (Maximum Subtitle Position) a 14 para posición vertical fija
  writeASCIIToBuffer(buffer, 400, "14", 2);
  
  // Establecer DSC (Display Standard Code) en una posición alternativa para mayor compatibilidad
  buffer[500] = 0x30; // ASCII '0' para PAL 625
  
  // Identificar claramente este STL como generado por DubApp
  writeASCIIToBuffer(buffer, 600, "DUBAPP STL CONVERTER V3.0", 24);
  
  // Establecer explícitamente el campo de TCF (Time Code Format) para mayor compatibilidad
  writeASCIIToBuffer(buffer, 700, "STL25.01", 8);
  
  // DEBUG: Mostrar las partes importantes del encabezado GSI
  Logger.log("DEBUG - Encabezado GSI: CPN=" + String.fromCharCode.apply(null, buffer.slice(0, 4)) + 
             " DFC=" + String.fromCharCode.apply(null, buffer.slice(4, 12)) +
             " DSC=" + String.fromCharCode.apply(null, [buffer[12]]) +
             " LC=" + String.fromCharCode.apply(null, buffer.slice(14, 16)) +
             " MNC=" + String.fromCharCode.apply(null, buffer.slice(252, 255)) +
             " MNR=" + String.fromCharCode.apply(null, buffer.slice(255, 258)));
  
  return buffer;
}

function processGoogleSheet(sheet) {
  var result = {};
  
  try {
    // Configuración
    var config = getConfig('excel');
    var startRow = config.startRow || 11;
    
    // Extraer metadatos
    var metadata = {};
    try {
      var titleEnglishRow = config.metadata.titleEnglish || 3;
      var titleSpanishRow = config.metadata.titleSpanish || 4;
      var episodeRow = config.metadata.episodeNumber || 5;
      
      // Obtener los valores
      var englishTitle = sheet.getRange(titleEnglishRow, 3).getValue();
      var spanishTitle = sheet.getRange(titleSpanishRow, 3).getValue();
      var episodeNumber = sheet.getRange(episodeRow, 3).getValue();
      
      // Limpiar los valores
      englishTitle = (englishTitle || '').toString().replace(/^.*English Title:?\s*/i, '').trim();
      spanishTitle = (spanishTitle || '').toString().replace(/^.*Spanish Title:?\s*/i, '').trim();
      episodeNumber = (episodeNumber || '').toString().replace(/^.*Episode.*?:?\s*/i, '').trim();
      
      // Guardar metadatos
      metadata.englishTitle = englishTitle || 'Untitled';
      metadata.spanishTitle = spanishTitle || 'Sin título';
      metadata.episodeNumber = episodeNumber || '01';
      metadata.language = 'es';
      
      Logger.log('Metadatos extraídos: ' + JSON.stringify(metadata));
    } catch (metadataError) {
      Logger.log('Error al extraer metadatos: ' + metadataError);
      // Valores por defecto
      metadata = {
        englishTitle: 'Untitled',
        spanishTitle: 'Sin título',
        episodeNumber: '01',
        language: 'es'
      };
    }
    
    // Extraer subtítulos
    var subtitles = [];
    var lastRow = sheet.getLastRow();
    var columnConfig = config.columns;
    
    // Columnas (ajustar según la configuración)
    var timeInCol = columnConfig.timeIn || 1;  // 0-indexado (B)
    var textCol = columnConfig.text || 2;      // 0-indexado (C)
    var timeOutCol = columnConfig.timeOut || 3; // 0-indexado (D)
    
    for (var i = startRow; i <= lastRow; i++) {
      try {
        // Obtener los valores de timecode
        var timeIn = sheet.getRange(i, timeInCol + 1).getValue(); // +1 porque getRange es 1-indexado
        if (!timeIn) continue; // Saltar filas vacías
        
        var timeOut = sheet.getRange(i, timeOutCol + 1).getValue();
        if (!timeOut) continue;
        
        // Obtener el texto 
        var textCell = sheet.getRange(i, textCol + 1);
        var text = textCell.getDisplayValue() || textCell.getValue();
        
        // Si es un valor de texto, asegurarnos de que es string
        text = text ? text.toString() : "";
        
        // Depuración: Registrar el texto original
        Logger.log("TEXTO ORIGINAL [Fila " + i + "]: " + text);
        
        // Usar preprocessSubtitleText para limpiar HTML en lugar de hacerlo aquí
        text = preprocessSubtitleText(text);
        
        // 6. Verificar si necesitamos dividir en dos líneas
        var lines = text.split('\n');
        if (lines.length === 1 && text.length > 10) {
          // Encontrar un punto natural para dividir
          var middlePoint = Math.floor(text.length / 2);
          var breakPoint = -1;
          
          // Intentar encontrar punto, coma, etc.
          for (var j = middlePoint - 10; j <= middlePoint + 10; j++) {
            if (j > 0 && j < text.length && ".,:;?!".indexOf(text.charAt(j)) !== -1) {
              breakPoint = j + 1;
              break;
            }
          }
          
          // Si no hay puntuación, buscar espacio cercano al medio
          if (breakPoint === -1) {
            var spaceBeforeMid = text.lastIndexOf(' ', middlePoint);
            var spaceAfterMid = text.indexOf(' ', middlePoint);
            
            if (spaceBeforeMid !== -1 && spaceAfterMid !== -1) {
              breakPoint = (middlePoint - spaceBeforeMid < spaceAfterMid - middlePoint) 
                ? spaceBeforeMid : spaceAfterMid;
            } else if (spaceBeforeMid !== -1) {
              breakPoint = spaceBeforeMid;
            } else if (spaceAfterMid !== -1) {
              breakPoint = spaceAfterMid;
            }
          }
          
          // Aplicar división si encontramos punto de ruptura
          if (breakPoint !== -1) {
            var firstLine = text.substring(0, breakPoint).trim();
            var secondLine = text.substring(breakPoint).trim();
            text = firstLine + "\n" + secondLine;
            
            Logger.log("TEXTO DIVIDIDO [Fila " + i + "]: " + text.replace("\n", " | "));
          }
        }
        
        // Depuración: Registrar el texto procesado final
        Logger.log("TEXTO PROCESADO [Fila " + i + "]: " + text.replace(/\n/g, " | "));
        
        // Formatear los timecodes
        var formattedTimeIn = formatTimecode(timeIn.toString());
        var formattedTimeOut = formatTimecode(timeOut.toString());
        
        // Añadir el subtítulo al array
        subtitles.push({
          timecodeIn: formattedTimeIn,
          timecodeOut: formattedTimeOut,
          text: text
        });
        
        // Depuración: Registrar el subtítulo añadido
        Logger.log("SUBTÍTULO AÑADIDO [" + subtitles.length + "]: " + 
                 formattedTimeIn + " -> " + formattedTimeOut + ": " + 
                 text.substring(0, 30) + (text.length > 30 ? "..." : ""));
                 
      } catch (rowError) {
        Logger.log("ERROR al procesar fila " + i + ": " + rowError.message);
      }
    }
    
    Logger.log('Subtítulos extraídos: ' + subtitles.length);
    
    result.success = true;
    result.metadata = metadata;
    result.subtitles = subtitles;
    
  } catch (error) {
    Logger.log('Error al procesar Google Sheet: ' + error.message);
    result.success = false;
    result.error = error.message;
  }
  
  return result;
}

/**
 * Formatea un código de tiempo para asegurar que tenga el formato correcto para STL (HH:MM:SS:FF)
 * @param {String} timecode - Timecode en diferentes formatos posibles
 * @return {String} Timecode formateado como HH:MM:SS:FF
 */
function formatTimecode(timecode) {
  if (!timecode) return "00:00:00:00";
  
  // Remover espacios
  timecode = timecode.trim();
  
  // Reemplazar puntos por dos puntos en el frame portion (00:00:00.00 → 00:00:00:00)
  timecode = timecode.replace(/\.(\d{1,2})$/, ":$1");
  
  // Añadir horas si no están presentes (00:00:00 → 00:00:00:00)
  if (timecode.split(':').length === 3) {
    timecode = "00:" + timecode;
  }
  
  // Si solo tiene dos partes, asumir MM:SS y añadir horas y frames (MM:SS → 00:MM:SS:00)
  if (timecode.split(':').length === 2) {
    timecode = "00:" + timecode + ":00";
  }
  
  // Asegurar que tenga 4 partes: HH:MM:SS:FF
  var parts = timecode.split(':');
  if (parts.length !== 4) {
    return "00:00:00:00"; // Si no podemos interpretar, devolver un timecode por defecto
  }
  
  // Asegurar que todos los valores sean números de dos dígitos
  for (var i = 0; i < 4; i++) {
    var val = parseInt(parts[i], 10);
    if (isNaN(val)) val = 0;
    parts[i] = val.toString().padStart(2, '0');
  }
  
  // Devolver el formato correcto
  return parts.join(':');
}

/**
 * Función de diagnóstico para visualizar buffer en formato hexadecimal
 * @param {Array} buffer - Buffer de bytes a visualizar
 * @param {String} label - Etiqueta para identificar la visualización
 * @return {String} Representación hexadecimal formateada
 */
function debugBuffer(buffer, label) {
  var hexDump = "";
  var asciiDump = "";
  var lineHex = "";
  var lineAscii = "| ";
  
  for (var i = 0; i < buffer.length; i++) {
    // Formato hexadecimal con padding de ceros
    var hex = buffer[i].toString(16).padStart(2, '0');
    lineHex += hex + " ";
    
    // Representación ASCII (usar punto para caracteres no imprimibles)
    var ascii = (buffer[i] >= 32 && buffer[i] <= 126) ? 
                String.fromCharCode(buffer[i]) : '.';
    
    // Caracteres especiales
    if (buffer[i] === 0x8A) ascii = '¶'; // Salto de línea
    if (buffer[i] === 0x03) ascii = '.'; // ETX
    
    lineAscii += ascii;
    
    // Formatear en líneas de 16 bytes
    if ((i + 1) % 16 === 0 || i === buffer.length - 1) {
      // Rellenar la última línea si es necesario
      while ((i + 1) % 16 !== 0) {
        lineHex += "   ";
        i++;
      }
      
      // Añadir la línea al dump
      var lineNum = (i - 15).toString(16).padStart(4, '0') + ": ";
      Logger.log(lineNum + lineHex + lineAscii);
      
      hexDump += lineNum + lineHex + lineAscii + "\n";
      lineHex = "";
      lineAscii = "| ";
    }
  }
  
  // Registrar mensaje final
  Logger.log(label + " - Final: " + hexDump.split("\n")[hexDump.split("\n").length - 2]);
  
  return hexDump;
}

/**
 * Función para analizar las características del texto del subtítulo
 * y detectar patrones específicos como etiquetas HTML, múltiples líneas, etc.
 * 
 * @param {String} text - Texto del subtítulo a analizar
 * @return {Object} Objeto con análisis detallado del texto
 */
function analyzeSubtitleText(text) {
  if (!text) return { containsHtml: false, lineCount: 0, length: 0 };
  
  var result = {
    containsHtml: false,
    lineCount: 1,
    length: text.length,
    containsColorTags: false,
    containsFormattingTags: false,
    containsBrTags: false,
    containsAnTags: false
  };
  
  // Detectar etiquetas HTML
  if (/<[^>]+>/.test(text)) {
    result.containsHtml = true;
    
    // Detectar etiquetas específicas
    if (/<font\s+color/i.test(text)) {
      result.containsColorTags = true;
      
      // Contar etiquetas de color
      var colorMatches = text.match(/<font\s+color=['"][^'"]*['"][^>]*>/gi);
      result.colorTagsCount = colorMatches ? colorMatches.length : 0;
    }
    
    if (/<(b|i|u|strong|em)>/i.test(text)) {
      result.containsFormattingTags = true;
    }
    
    if (/<br\s*\/?>/i.test(text)) {
      result.containsBrTags = true;
      // Contar saltos de línea
      var brMatches = text.match(/<br\s*\/?>/gi);
      if (brMatches) {
        result.lineCount += brMatches.length;
      }
    }
  }
  
  // Detectar etiquetas \anX (posición)
  if (/\\an[1-9]/.test(text) || /<\\an[1-9]>/.test(text)) {
    result.containsAnTags = true;
  }
  
  // Contar saltos de línea literales
  var newlineMatches = text.match(/\n/g);
  if (newlineMatches) {
    result.lineCount += newlineMatches.length;
  }
  
  // Detectar secuencias de etiquetas de color que podría indicar cambios de línea
  var colorTagSequences = text.match(/<\/font>\s*<font color/g);
  if (colorTagSequences) {
    result.colorTagSequences = colorTagSequences.length;
  }
  
  return result;
}

/**
 * Escribe una cadena ASCII en un buffer en la posición especificada
 * 
 * @param {Array} buffer - El buffer donde escribir
 * @param {Number} position - La posición inicial en el buffer
 * @param {String} text - El texto a escribir
 * @param {Number} maxLength - Longitud máxima a escribir
 */
function writeASCIIToBuffer(buffer, position, text, maxLength) {
  if (!text) return;
  
  text = text.toString();
  var length = Math.min(text.length, maxLength);
  
  for (var i = 0; i < length; i++) {
    buffer[position + i] = text.charCodeAt(i);
  }
}

/**
 * Obtiene el código de idioma ISO 639-2 para el encabezado STL
 * 
 * @param {String} language - Código de idioma (es, pt-BR, etc.)
 * @return {String} Código de idioma de dos caracteres
 */
function getLanguageCode(language) {
  if (!language) return "SP"; // Español por defecto
  
  language = language.toLowerCase();
  
  switch (language) {
    case "es":
    case "es-es":
    case "es-ar":
    case "es-mx":
      return "SP"; // Español
    case "pt":
    case "pt-br":
      return "PO"; // Portugués
    case "en":
    case "en-us":
    case "en-gb":
      return "EN"; // Inglés
    case "fr":
      return "FR"; // Francés
    case "de":
      return "GE"; // Alemán
    case "it":
      return "IT"; // Italiano
    default:
      return "SP"; // Español como fallback
  }
}

/**
 * Muestra el contenido de un buffer en formato hexadecimal para depuración
 * 
 * @param {Array} buffer - Buffer a mostrar
 * @param {String} label - Etiqueta para identificar el buffer en los logs
 */
function debugBuffer(buffer, label) {
  var hexLines = [];
  var lineLen = 16;
  
  for (var i = 0; i < buffer.length; i += lineLen) {
    var hex = '';
    var ascii = '| ';
    
    for (var j = 0; j < lineLen && i + j < buffer.length; j++) {
      var byte = buffer[i + j];
      hex += ('0' + byte.toString(16)).slice(-2) + ' ';
      
      // Para ASCII imprimible (32-126), mostrar el carácter; para otros, mostrar un punto
      ascii += (byte >= 32 && byte <= 126) ? String.fromCharCode(byte) : '.';
    }
    
    // Rellenar con espacios si no completamos la línea
    while (hex.length < lineLen * 3) {
      hex += '   ';
    }
    
    var offset = ('0000' + i.toString(16)).slice(-4);
    hexLines.push(offset + ': ' + hex + ascii);
  }
  
  // Mostrar cada línea en los logs
  for (var i = 0; i < hexLines.length; i++) {
    Logger.log(hexLines[i]);
  }
  
  // Mostrar un resumen al final
  Logger.log(label + " - Final: " + hexLines[hexLines.length - 1]);
}
