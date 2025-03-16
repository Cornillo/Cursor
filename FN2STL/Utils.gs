/**
 * Módulo de Acceso a Datos
 * Funciones para leer datos de Google Sheets y guardar archivos en Google Drive
 */

// No es necesario declarar estas variables en Google Apps Script
// ya que están disponibles globalmente

/**
 * Lee los datos de la hoja de cálculo especificada
 * @param {string} sheetId - ID de la hoja de cálculo de Google
 * @return {Object} Datos de la hoja de cálculo
 */
function readSpreadsheetData(sheetId) {
  logMessage("Abriendo hoja de cálculo: " + sheetId);
  
  try {
    const spreadsheet = SpreadsheetApp.openById(sheetId);
    const sheet = spreadsheet.getActiveSheet();
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    
    logMessage("Hoja de cálculo leída: " + values.length + " filas encontradas");
    
    return {
      values: values,
      spreadsheet: spreadsheet,
      sheet: sheet
    };
  } catch (error) {
    logMessage("Error al leer la hoja de cálculo: " + error.message);
    throw new Error("No se pudo leer la hoja de cálculo: " + error.message);
  }
}

/**
 * Extrae el título del programa de la hoja de cálculo
 * @param {Object} sheetData - Datos de la hoja de cálculo
 * @return {string} Título del programa
 */
function extractTitle(sheetData) {
  try {
    const values = sheetData.values;
    
    // Buscar el título en diferentes ubicaciones comunes
    const possibleTitleLocations = [
      { row: 4, col: 2 },  // Fila 5, columna C
      { row: 2, col: 2 },  // Fila 3, columna C
      { row: 0, col: 2 }   // Fila 1, columna C
    ];
    
    for (const loc of possibleTitleLocations) {
      if (values.length > loc.row && 
          values[loc.row] && 
          values[loc.row].length > loc.col && 
          values[loc.row][loc.col]) {
        const title = values[loc.row][loc.col].toString().trim();
        if (title) {
          logMessage("Título encontrado en fila " + (loc.row + 1) + ": " + title);
          return title;
        }
      }
    }
    
    // Si no se encuentra el título, usar el nombre del archivo
    logMessage("No se encontró el título en las posiciones esperadas. Usando nombre del archivo");
    return sheetData.spreadsheet.getName();
  } catch (error) {
    logMessage("Error al extraer el título: " + error.message + ". Usando nombre del archivo");
    return sheetData.spreadsheet.getName();
  }
}

/**
 * Extrae los datos de subtítulos de una hoja de Google Sheets
 * @param {Sheet} sheet - Objeto de hoja de Google Sheets
 * @param {object} structure - Estructura detectada con columnas y fila de inicio
 * @param {boolean} verboseFlag - Activar logs detallados
 * @return {Array} Array de objetos con los datos de subtítulos
 */
function extractSubtitleData(sheet, structure, verboseFlag) {
  if (verboseFlag) Logger.log(`Extrayendo datos de subtítulos desde la fila ${structure.dataStartRow}...`);
  
  try {
    // Obtener el rango de datos
    const lastRow = sheet.getLastRow();
    const numRows = lastRow - structure.dataStartRow + 1;
    
    if (numRows <= 0) {
      if (verboseFlag) Logger.log("No se encontraron filas de datos");
      return [];
    }
    
    // Determinar el número máximo de columnas necesarias
    const maxCol = Math.max(
      structure.numberCol,
      structure.startTimeCol,
      structure.endTimeCol,
      structure.textCol
    );
    
    // Obtener todos los datos de una vez
    const dataRange = sheet.getRange(structure.dataStartRow, 1, numRows, maxCol);
    const values = dataRange.getValues();
    
    const subtitles = [];
    let validCount = 0;
    let invalidCount = 0;
    
    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      
      // Obtener valores de las columnas (ajustando índices)
      const number = row[structure.numberCol - 1] || i + 1;
      const startTime = row[structure.startTimeCol - 1];
      const endTime = row[structure.endTimeCol - 1];
      const text = row[structure.textCol - 1];
      
      // Verificar si la fila tiene datos válidos
      if (!startTime && !endTime && !text) {
        continue; // Fila vacía, saltar
      }
      
      // Validar y formatear los códigos de tiempo
      const formattedStartTime = validateAndFormatTimecode(startTime, verboseFlag);
      const formattedEndTime = validateAndFormatTimecode(endTime, verboseFlag);
      
      // Verificar si los códigos de tiempo son válidos
      if (!formattedStartTime || !formattedEndTime) {
        if (verboseFlag) {
          Logger.log(`Fila ${structure.dataStartRow + i} inválida: Códigos de tiempo incorrectos`);
          Logger.log(`- Tiempo inicial: ${startTime} -> ${formattedStartTime || 'INVÁLIDO'}`);
          Logger.log(`- Tiempo final: ${endTime} -> ${formattedEndTime || 'INVÁLIDO'}`);
        }
        invalidCount++;
        continue;
      }
      
      // Verificar si hay texto
      if (!text) {
        if (verboseFlag) Logger.log(`Fila ${structure.dataStartRow + i} inválida: No hay texto`);
        invalidCount++;
        continue;
      }
      
      // Agregar subtítulo válido
      subtitles.push({
        number: Number(number),
        startTime: formattedStartTime,
        endTime: formattedEndTime,
        text: String(text).trim()
      });
      
      validCount++;
    }
    
    if (verboseFlag) {
      Logger.log(`Subtítulos extraídos: ${validCount} válidos, ${invalidCount} inválidos`);
      
      // Mostrar algunos ejemplos
      for (let i = 0; i < Math.min(3, subtitles.length); i++) {
        Logger.log(`Ejemplo #${i+1}: ${subtitles[i].startTime} - ${subtitles[i].endTime} "${subtitles[i].text.substring(0, 30)}${subtitles[i].text.length > 30 ? '...' : ''}"`);
      }
    }
    
    return subtitles;
    
  } catch (e) {
    Logger.log(`Error al extraer datos de subtítulos: ${e.message}`);
    return [];
  }
}

/**
 * Verifica si una cadena tiene formato de código de tiempo
 * @param {string} value - Valor a verificar
 * @return {boolean} true si tiene formato de código de tiempo
 */
function isTimecodeFormat(value) {
  if (!value) return false;
  
  // Convertir a string si no lo es
  const str = String(value).trim();
  
  // Verificar formato HH:MM:SS:FF (o con ; o . como separadores)
  return /^\d{1,2}[:;.]\d{1,2}[:;.]\d{1,2}[:;.]\d{1,2}$/.test(str);
}

/**
 * Añade segundos a un timecode y devuelve el nuevo timecode
 * @param {string} timecode - Timecode en formato HH:MM:SS:FF
 * @param {number} seconds - Segundos a añadir
 * @return {string} - Nuevo timecode
 */
function addSecondsToTimecode(timecode, seconds) {
  // Extraer componentes
  const parts = timecode.split(':');
  
  // Convertir a segundos totales
  let hours = parseInt(parts[0], 10);
  let minutes = parseInt(parts[1], 10);
  let secs = parseInt(parts[2], 10);
  let frames = parts.length > 3 ? parseInt(parts[3], 10) : 0;
  
  // Añadir segundos
  secs += seconds;
  
  // Ajustar
  while (secs >= 60) {
    secs -= 60;
    minutes++;
  }
  
  while (minutes >= 60) {
    minutes -= 60;
    hours++;
  }
  
  // Formatear
  const result = `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}:${secs.toString().padStart(2, '0')}:${frames.toString().padStart(2, '0')}`;
  
  return result;
}

/**
 * Valida y formatea un código de tiempo al formato estándar HH:MM:SS:FF
 * @param {any} value - Valor a validar y formatear
 * @param {boolean} verboseFlag - Activar logs detallados
 * @return {string|null} Código de tiempo formateado o null si es inválido
 */
function validateAndFormatTimecode(value, verboseFlag) {
  if (!value) return null;
  
  try {
    // Si ya tiene formato de timecode, verificar y normalizar
    if (isTimecodeFormat(value)) {
      if (verboseFlag) Logger.log(`Valor ya tiene formato de timecode: ${value}`);
      
      // Normalizar separadores a ':' (Subtitle Edit espera este formato)
      let timecode = String(value).trim().replace(/[;.]/g, ':');
      
      // Extraer componentes
      const parts = timecode.split(':');
      let hours = parseInt(parts[0], 10);
      let minutes = parseInt(parts[1], 10);
      let seconds = parseInt(parts[2], 10);
      let frames = parseInt(parts[3], 10);
      
      // Validar rangos
      if (hours < 0 || hours > 23) {
        if (verboseFlag) Logger.log(`Horas fuera de rango (${hours}), ajustando...`);
        hours = Math.max(0, Math.min(hours, 23));
      }
      
      if (minutes < 0 || minutes > 59) {
        if (verboseFlag) Logger.log(`Minutos fuera de rango (${minutes}), ajustando...`);
        minutes = Math.max(0, Math.min(minutes, 59));
      }
      
      if (seconds < 0 || seconds > 59) {
        if (verboseFlag) Logger.log(`Segundos fuera de rango (${seconds}), ajustando...`);
        seconds = Math.max(0, Math.min(seconds, 59));
      }
      
      // Convertir frames a valores apropiados para STL25.01 (25fps)
      // En la visualización, frames debe estar entre 0 y 24
      if (frames < 0 || frames > 24) {
        if (verboseFlag) Logger.log(`Frames fuera de rango (${frames}), ajustando...`);
        frames = Math.max(0, Math.min(frames, 24));
      }
      
      // Formatear con ceros a la izquierda
      const formattedTimecode = `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}:${seconds.toString().padStart(2, '0')}:${frames.toString().padStart(2, '0')}`;
      
      if (verboseFlag) Logger.log(`Timecode normalizado: ${timecode} -> ${formattedTimecode}`);
      
      return formattedTimecode;
    }
    
    // Si es un objeto Date, convertir a timecode
    if (value instanceof Date) {
      if (verboseFlag) Logger.log(`Convirtiendo Date a timecode: ${value}`);
      
      const hours = value.getHours();
      const minutes = value.getMinutes();
      const seconds = value.getSeconds();
      // Para 25fps, convertimos milisegundos a frames
      const ms = value.getMilliseconds();
      const frames = Math.min(Math.floor(ms / 40), 24); // 1000ms / 25fps = 40ms por frame
      
      return `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}:${seconds.toString().padStart(2, '0')}:${frames.toString().padStart(2, '0')}`;
    }
    
    // Si es un número, podría ser un timestamp o valor de Excel
    if (typeof value === 'number') {
      if (verboseFlag) Logger.log(`Convirtiendo número a timecode: ${value}`);
      
      // Si es un valor de Excel (fracción de día)
      if (value < 1) {
        const totalSeconds = Math.round(value * 86400); // 24*60*60
        const hours = Math.floor(totalSeconds / 3600);
        const minutes = Math.floor((totalSeconds % 3600) / 60);
        const seconds = totalSeconds % 60;
        
        return `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}:${seconds.toString().padStart(2, '0')}:00`;
      }
      
      // Si es un timestamp en milisegundos
      if (value > 1000000) {
        const date = new Date(value);
        return validateAndFormatTimecode(date, verboseFlag);
      }
    }
    
    // Si es un string con formato diferente, intentar convertir
    if (typeof value === 'string') {
      const str = value.trim();
      
      // Formato HH:MM:SS (sin frames)
      if (/^\d{1,2}:\d{1,2}:\d{1,2}$/.test(str)) {
        return str + ':00';
      }
      
      // Formato HH:MM:SS,FFF o HH:MM:SS.FFF (con milisegundos)
      const msMatch = str.match(/^(\d{1,2}):(\d{1,2}):(\d{1,2})[,.](\d{1,3})$/);
      if (msMatch) {
        const hours = parseInt(msMatch[1], 10);
        const minutes = parseInt(msMatch[2], 10);
        const seconds = parseInt(msMatch[3], 10);
        const ms = parseInt(msMatch[4], 10);
        
        // Convertir milisegundos a frames para 25fps (formato PAL)
        const frames = Math.min(Math.floor(ms / 40), 24); // 1000ms / 25fps = 40ms por frame
        
        return `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}:${seconds.toString().padStart(2, '0')}:${frames.toString().padStart(2, '0')}`;
      }
    }
    
    if (verboseFlag) Logger.log(`No se pudo convertir a timecode: ${value}`);
    return null;
    
  } catch (e) {
    if (verboseFlag) Logger.log(`Error al validar timecode: ${e.message}`);
    return null;
  }
}

/**
 * Valida y formatea el texto del subtítulo
 * @param {any} value - Valor a validar
 * @return {string|null} Texto formateado o null si es inválido
 */
function validateAndFormatText(value) {
  // Si es null, undefined o vacío, retornar null
  if (value === null || value === undefined) return null;
  
  // Convertir a string y limpiar espacios al inicio y final
  let text = String(value).trim();
  if (text.length === 0) return null;
  
  // Reemplazar caracteres no válidos
  text = text.replace(/[""]/g, '"')
             .replace(/['']/g, "'")
             .replace(/…/g, "...")
             .replace(/[—–]/g, "-");
  
  // Validar longitud máxima por línea con advertencia pero sin rechazar
  const lines = text.split(/\r?\n/);
  
  for (const line of lines) {
    if (line.length > CONFIG.DEFAULT_MAX_CHARS_PER_ROW) {
      if (verboseFlag) {
        Logger.log(`Advertencia: Línea excede el máximo de caracteres (${line.length} > ${CONFIG.DEFAULT_MAX_CHARS_PER_ROW}): "${line}"`);
      }
    }
  }
  
  // Validar número máximo de líneas con advertencia pero sin rechazar
  if (lines.length > CONFIG.DEFAULT_MAX_ROWS) {
    if (verboseFlag) {
      Logger.log(`Advertencia: Subtítulo excede el máximo de líneas (${lines.length} > ${CONFIG.DEFAULT_MAX_ROWS})`);
    }
  }
  
  return text;
}

/**
 * Obtiene el nombre de la hoja de cálculo
 * @param {string} sheetId - ID de la hoja de cálculo
 * @return {string} Nombre de la hoja de cálculo
 */
function getSpreadsheetName(sheetId) {
  try {
    const spreadsheet = SpreadsheetApp.openById(sheetId);
    return spreadsheet.getName();
  } catch (error) {
    logMessage("Error al obtener el nombre de la hoja: " + error.message);
    return "Subtitles_" + new Date().getTime();
  }
}

/**
 * Guarda un archivo STL en Google Drive
 * @param {Uint8Array} data - Datos del archivo STL
 * @param {string} fileName - Nombre del archivo
 * @param {string} folderId - ID de la carpeta donde guardar el archivo
 * @param {boolean} verboseFlag - Activar logs detallados
 * @return {Object} - Información del archivo guardado
 */
function saveSTLFile(data, fileName, folderId, verboseFlag) {
  if (verboseFlag) Logger.log(`Guardando archivo STL: ${fileName}`);
  
  try {
    // Asegurar que el nombre termine con .STL (en mayúsculas)
    if (!fileName.toUpperCase().endsWith('.STL')) {
      fileName = fileName.replace(/\.[^/.]+$/, "") + '.STL'; // Eliminar cualquier extensión y añadir .STL
      if (verboseFlag) Logger.log(`Nombre de archivo normalizado: ${fileName}`);
    }
    
    // Convertir el array de bytes a un blob
    const blob = Utilities.newBlob(data, 'application/octet-stream', fileName);
    
    // Obtener la carpeta (o usar la raíz si no se especifica)
    let folder;
    
    if (folderId) {
      try {
        folder = DriveApp.getFolderById(folderId);
        if (verboseFlag) Logger.log(`Carpeta encontrada: ${folder.getName()}`);
      } catch (e) {
        if (verboseFlag) Logger.log(`Error al obtener carpeta: ${e}. Usando carpeta raíz.`);
        folder = DriveApp.getRootFolder();
      }
    } else {
      if (verboseFlag) Logger.log("No se especificó carpeta, usando raíz");
      folder = DriveApp.getRootFolder();
    }
    
    // Crear el archivo en Drive
    const file = folder.createFile(blob);
    
    // Configurar permisos para que cualquiera con el enlace pueda ver
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    if (verboseFlag) {
      Logger.log(`Archivo STL guardado correctamente`);
      Logger.log(`- Nombre: ${file.getName()}`);
      Logger.log(`- ID: ${file.getId()}`);
      Logger.log(`- Tamaño: ${file.getSize()} bytes`);
      Logger.log(`- URL: ${file.getUrl()}`);
    }
    
    // Devolver información del archivo
    return {
      id: file.getId(),
      name: file.getName(),
      url: file.getUrl(),
      size: file.getSize(),
      downloadUrl: `https://drive.google.com/uc?id=${file.getId()}&export=download`
    };
    
  } catch (error) {
    Logger.log(`Error al guardar archivo STL: ${error.message}`);
    throw error;
  }
}

/**
 * Analiza la estructura de una hoja de cálculo para identificar las columnas relevantes
 * @param {Sheet} sheet - Objeto de hoja de Google Sheets
 * @param {boolean} verboseFlag - Activar logs detallados
 * @return {object} Estructura detectada con columnas y fila de inicio
 */
function analyzeSheetStructure(sheet, verboseFlag) {
  if (verboseFlag) Logger.log("Analizando estructura de la hoja...");
  
  try {
    const numRows = sheet.getLastRow();
    const numCols = sheet.getLastColumn();
    
    if (numRows < 2 || numCols < 3) {
      return {
        valid: false,
        error: "La hoja no tiene suficientes datos para analizar"
      };
    }
    
    // Obtener los encabezados (primeras 10 filas como máximo)
    const headerRows = Math.min(10, numRows);
    const headerRange = sheet.getRange(1, 1, headerRows, numCols);
    const headerValues = headerRange.getValues();
    
    // Buscar columnas por nombres comunes en los encabezados
    let numberCol = -1;
    let startTimeCol = -1;
    let endTimeCol = -1;
    let textCol = -1;
    let dataStartRow = -1;
    
    // Patrones para buscar en los encabezados
    const numberPattern = /^(n[úu]m(ero)?|#|number|index|id)$/i;
    const startTimePattern = /^(in|inicio|start|tc\s*in|entrada|cue\s*in|time\s*in|time code in)$/i;
    const endTimePattern = /^(out|fin|end|tc\s*out|salida|cue\s*out|time\s*out|time code out)$/i;
    const textPattern = /^(texto|text|subtitle|subt[íi]tulo|content|contenido|di[áa]logo|dialog|spanish|español|portugues|português)$/i;
    
    // Verificar específicamente la fila 9 (índice 8) que es donde suelen estar los encabezados en el formato estándar
    if (headerRows > 8) {
      const standardHeaderRow = 8; // Fila 9 (índice 8)
      let foundStandardHeader = false;
      
      for (let col = 0; col < numCols; col++) {
        const cellValue = String(headerValues[standardHeaderRow][col] || "").trim().toLowerCase();
        
        if (!cellValue) continue;
        
        if (verboseFlag) Logger.log(`Examinando celda en fila 9, columna ${col+1}: "${cellValue}"`);
        
        // Buscar encabezados estándar: "Time code in", texto (idioma), "Time code out", etc.
        if (startTimeCol === -1 && (cellValue.includes("time code in") || cellValue.includes("time in"))) {
          startTimeCol = col + 1;
          foundStandardHeader = true;
          if (verboseFlag) Logger.log(`Encabezado estándar: columna de tiempo inicial encontrada en columna ${startTimeCol}`);
        }
        else if (endTimeCol === -1 && (cellValue.includes("time code out") || cellValue.includes("time out"))) {
          endTimeCol = col + 1;
          foundStandardHeader = true;
          if (verboseFlag) Logger.log(`Encabezado estándar: columna de tiempo final encontrada en columna ${endTimeCol}`);
        }
        else if (textCol === -1 && (cellValue.includes("spanish") || cellValue.includes("español") || 
                                   cellValue.includes("português") || cellValue.includes("portugues"))) {
          textCol = col + 1;
          foundStandardHeader = true;
          if (verboseFlag) Logger.log(`Encabezado estándar: columna de texto encontrada en columna ${textCol}`);
        }
      }
      
      // Si encontramos encabezados estándar en la fila 9, establecer la fila de inicio de datos en 11
      if (foundStandardHeader) {
        dataStartRow = 11;
        if (verboseFlag) Logger.log("Detectado formato estándar con encabezados en fila 9, datos a partir de fila 11");
      }
    }
    
    // Si no encontramos el formato estándar, continuar con la búsqueda general
    if (startTimeCol === -1 || endTimeCol === -1 || textCol === -1) {
      for (let row = 0; row < headerRows; row++) {
        let foundHeader = false;
        
        for (let col = 0; col < numCols; col++) {
          const cellValue = String(headerValues[row][col] || "").trim();
          
          if (!cellValue) continue;
          
          if (numberCol === -1 && numberPattern.test(cellValue)) {
            numberCol = col + 1;
            foundHeader = true;
            if (verboseFlag) Logger.log(`Columna de número encontrada: ${cellValue} (columna ${numberCol})`);
          }
          else if (startTimeCol === -1 && startTimePattern.test(cellValue)) {
            startTimeCol = col + 1;
            foundHeader = true;
            if (verboseFlag) Logger.log(`Columna de tiempo inicial encontrada: ${cellValue} (columna ${startTimeCol})`);
          }
          else if (endTimeCol === -1 && endTimePattern.test(cellValue)) {
            endTimeCol = col + 1;
            foundHeader = true;
            if (verboseFlag) Logger.log(`Columna de tiempo final encontrada: ${cellValue} (columna ${endTimeCol})`);
          }
          else if (textCol === -1 && textPattern.test(cellValue)) {
            textCol = col + 1;
            foundHeader = true;
            if (verboseFlag) Logger.log(`Columna de texto encontrada: ${cellValue} (columna ${textCol})`);
          }
        }
        
        // Si encontramos al menos un encabezado, la siguiente fila es donde comienzan los datos
        if (foundHeader && dataStartRow === -1) {
          dataStartRow = row + 2; // +1 para convertir índice a número de fila, +1 para ir a la siguiente fila
        }
      }
    }
    
    // Si no se encontró una fila de inicio de datos, usar la fila 2 por defecto
    if (dataStartRow === -1) {
      dataStartRow = 2;
      if (verboseFlag) Logger.log("No se detectó fila de inicio de datos, usando fila 2 por defecto");
    }
    
    // Intentar detectar columnas por contenido si no se encontraron por encabezados
    if (startTimeCol === -1 || endTimeCol === -1 || textCol === -1) {
      if (verboseFlag) Logger.log("Intentando detectar columnas por contenido...");
      
      // Obtener una muestra de datos
      const sampleStartRow = dataStartRow;
      const sampleRange = sheet.getRange(sampleStartRow, 1, Math.min(10, numRows - sampleStartRow + 1), numCols);
      const sampleValues = sampleRange.getValues();
      
      // Array para contar ocurrencias de formatos de timecode por columna
      const timecodeColumns = [];
      
      for (let col = 0; col < numCols; col++) {
        let timeCodeCount = 0;
        let textCount = 0;
        
        for (let row = 0; row < sampleValues.length; row++) {
          const cellValue = String(sampleValues[row][col] || "").trim();
          
          if (!cellValue) continue;
          
          // Detectar códigos de tiempo (formato HH:MM:SS:FF o similar)
          if (/^\d{1,2}[:;.]\d{1,2}[:;.]\d{1,2}[:;.]\d{1,2}$/.test(cellValue)) {
            timeCodeCount++;
            // Añadir esta columna al array de columnas de timecode si no está ya
            if (!timecodeColumns.includes(col + 1)) {
              timecodeColumns.push(col + 1);
            }
          }
          // Detectar texto (más de 5 caracteres y no es un timecode)
          else if (cellValue.length > 5 && !/^\d+[:;.]/.test(cellValue)) {
            textCount++;
          }
        }
        
        // Asignar columna de texto basada en el contenido
        if (textCol === -1 && textCount >= 2) {
          textCol = col + 1;
          if (verboseFlag) Logger.log(`Columna de texto detectada por contenido: columna ${textCol}`);
        }
      }
      
      // Asignar columnas de tiempo basadas en el order de aparición
      // La primera columna con timecodes es generalmente el inicio
      // La segunda columna con timecodes es generalmente el fin
      if (timecodeColumns.length >= 2) {
        if (startTimeCol === -1) {
          startTimeCol = timecodeColumns[0];
          if (verboseFlag) Logger.log(`Columna de tiempo inicial detectada por contenido: columna ${startTimeCol}`);
        }
        
        if (endTimeCol === -1) {
          // Si hay más de una columna de timecode, usar la segunda
          // Pero asegurarse de que sea diferente de la columna de inicio
          for (let i = 0; i < timecodeColumns.length; i++) {
            if (timecodeColumns[i] !== startTimeCol) {
              endTimeCol = timecodeColumns[i];
              if (verboseFlag) Logger.log(`Columna de tiempo final detectada por contenido: columna ${endTimeCol}`);
              break;
            }
          }
        }
      }
      // Si solo hay una columna de timecode, intentar inferir usando posiciones relativas
      else if (timecodeColumns.length === 1 && textCol !== -1) {
        const timeCol = timecodeColumns[0];
        
        // Si la columna de tiempo está antes de la de texto, asumimos que es tiempo inicial
        if (timeCol < textCol && startTimeCol === -1) {
          startTimeCol = timeCol;
          if (verboseFlag) Logger.log(`Columna de tiempo inicial inferida: columna ${startTimeCol}`);
        }
        // Si la columna de tiempo está después de la de texto, asumimos que es tiempo final
        else if (timeCol > textCol && endTimeCol === -1) {
          endTimeCol = timeCol;
          if (verboseFlag) Logger.log(`Columna de tiempo final inferida: columna ${endTimeCol}`);
        }
      }
    }
    
    // Si se detectó una columna de tiempo pero no la otra, buscar en columnas adyacentes
    if (startTimeCol !== -1 && endTimeCol === -1) {
      // Buscar en la columna siguiente a la de tiempo inicial
      if (startTimeCol + 2 <= numCols) {
        endTimeCol = startTimeCol + 2;
        if (verboseFlag) Logger.log(`Columna de tiempo final inferida: columna ${endTimeCol} (2 columnas después de inicio)`);
      }
    } else if (startTimeCol === -1 && endTimeCol !== -1) {
      // Buscar en la columna anterior a la de tiempo final
      if (endTimeCol - 2 >= 1) {
        startTimeCol = endTimeCol - 2;
        if (verboseFlag) Logger.log(`Columna de tiempo inicial inferida: columna ${startTimeCol} (2 columnas antes de fin)`);
      }
    }
    
    // Verificar que se encontraron las columnas necesarias
    if (startTimeCol === -1 || endTimeCol === -1 || textCol === -1) {
      // Último intento: si estamos en formato A=in, B=text, C=out (formato común)
      if (textCol === 2) {
        if (startTimeCol === -1) startTimeCol = 1;
        if (endTimeCol === -1) endTimeCol = 3;
        if (verboseFlag) Logger.log(`Usando formato estándar A=in, B=text, C=out por columna de texto en B`);
      }
      // O si tenemos A=inicio, B=texto, falta final, asumir que C=final
      else if (startTimeCol === 1 && textCol === 2 && endTimeCol === -1 && numCols >= 3) {
        endTimeCol = 3;
        if (verboseFlag) Logger.log(`Inferencia final: columna C como tiempo final`);
      }
      // Si aún así no detectamos todas las columnas, reportar error
      else {
        return {
          valid: false,
          error: "No se pudieron detectar todas las columnas necesarias"
        };
      }
    }
    
    // Si no se encontró la columna de número, usar la primera columna
    if (numberCol === -1) {
      numberCol = 1;
      if (verboseFlag) Logger.log("No se detectó columna de número, usando columna 1 por defecto");
    }
    
    return {
      valid: true,
      numberCol: numberCol,
      startTimeCol: startTimeCol,
      endTimeCol: endTimeCol,
      textCol: textCol,
      dataStartRow: dataStartRow
    };
    
  } catch (e) {
    if (verboseFlag) Logger.log(`Error al analizar estructura: ${e.message}`);
    return {
      valid: false,
      error: `Error al analizar estructura: ${e.message}`
    };
  }
}

/**
 * Convierte un índice de columna a letra (1=A, 2=B, etc.)
 * @param {number} column - Índice de columna (1-based)
 * @return {string} - Letra de columna
 */
function columnToLetter(column) {
  let temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

/**
 * Busca un encabezado en un array de encabezados basado en palabras clave
 * @param {Array} headers - Array de encabezados
 * @param {Array} keywords - Palabras clave a buscar
 * @return {number} - Índice del encabezado encontrado o -1 si no se encuentra
 */
function findColumnIndex(headers, keywords) {
  for (let i = 0; i < headers.length; i++) {
    const header = String(headers[i] || "").toLowerCase();
    for (let j = 0; j < keywords.length; j++) {
      if (header.includes(keywords[j].toLowerCase())) {
        return i;
      }
    }
  }
  return -1;
}
