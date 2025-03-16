/**
 * Módulo para la creación de bloques TTI (Text and Timing Information)
 * Implementa la creación de los bloques de subtítulos según la especificación EBU Tech 3264-1991
 */

/**
 * Crea un bloque TTI (Text and Timing Information) para un subtítulo
 * @param {Object} options - Opciones para el bloque TTI
 * @param {number} options.subtitleNumber - Número de subtítulo (1-based)
 * @param {string} options.timecodeIn - Código de tiempo de entrada (formato HH:MM:SS:FF)
 * @param {string} options.timecodeOut - Código de tiempo de salida (formato HH:MM:SS:FF)
 * @param {string} options.text - Texto del subtítulo
 * @param {boolean} options.verboseFlag - Activar logs detallados
 * @return {Uint8Array} - Bloque TTI formateado (128 bytes)
 */
function createTTIBlock(options) {
  const {
    subtitleNumber,
    timecodeIn,
    timecodeOut,
    text,
    verboseFlag = false
  } = options || {};
  
  if (verboseFlag) {
    Logger.log(`=== Creando bloque TTI #${subtitleNumber} ===`);
    Logger.log(`Entrada: ${timecodeIn}`);
    Logger.log(`Salida: ${timecodeOut}`);
    Logger.log(`Texto: "${text}"`);
  }
  
  // Crear un buffer de 128 bytes (tamaño del bloque TTI) con valores 0x20 (espacio)
  const ttiBlock = new Uint8Array(128).fill(0x20);
  
  try {
    // 1. Subtítulo Group Number (SGN) - 1 byte - Usamos 1 por defecto
    ttiBlock[0] = 0x00;
    
    // 2. Subtítulo Number (SN) - 2 bytes - BCD
    const snBytes = convertNumberToBCD(subtitleNumber, 2);
    ttiBlock[1] = snBytes[0];
    ttiBlock[2] = snBytes[1];
    
    // 3. Extension Block Number (EBN) - 1 byte - Usamos 0 (sin extensión)
    ttiBlock[3] = 0xFF;
    
    // 4. Cumulative Status (CS) - 1 byte - Usamos 0 (no acumulativo)
    ttiBlock[4] = 0x00;
    
    // 5. Time Code In (TCI) - 4 bytes - BCD
    const tciBytes = convertTimecodeToBytes(timecodeIn, verboseFlag);
    for (let i = 0; i < 4; i++) {
      ttiBlock[5 + i] = tciBytes[i];
    }
    
    // 6. Time Code Out (TCO) - 4 bytes - BCD
    const tcoBytes = convertTimecodeToBytes(timecodeOut, verboseFlag);
    for (let i = 0; i < 4; i++) {
      ttiBlock[9 + i] = tcoBytes[i];
    }
    
    // 7. Vertical Position (VP) - 1 byte - Posición vertical en pantalla (0-23)
    ttiBlock[13] = 0x16; // Valor típico para subtítulos en parte inferior
    
    // 8. Justification Code (JC) - 1 byte - Centrado (0x02)
    ttiBlock[14] = 0x02;
    
    // 9. Comment Flag (CF) - 1 byte - No es comentario (0)
    ttiBlock[15] = 0x00;
    
    // 10. Texto del subtítulo - 112 bytes
    // Usar nuestra función optimizada para el manejo de caracteres especiales
    const processedText = processTextForSTL(text, verboseFlag);
    
    // Escribir el texto procesado en el bloque TTI
    for (let i = 0; i < processedText.length && i < 112; i++) {
      ttiBlock[16 + i] = processedText[i];
    }
    
    if (verboseFlag) {
      let hexOutput = "";
      for (let i = 0; i < 32; i++) {
        hexOutput += ttiBlock[i].toString(16).padStart(2, '0') + " ";
      }
      Logger.log(`TTI Block Header (hex): ${hexOutput}`);
      Logger.log(`Bloque TTI #${subtitleNumber} creado correctamente`);
    }
    
    return ttiBlock;
    
  } catch (error) {
    if (verboseFlag) {
      Logger.log(`Error al crear bloque TTI: ${error.message}`);
    }
    // Devolver un bloque vacío en caso de error
    return new Uint8Array(128).fill(0x20);
  }
}

/**
 * Convierte un número a formato BCD (Binary Coded Decimal)
 * @param {number} number - Número a convertir
 * @param {number} byteLength - Longitud del resultado en bytes
 * @return {Uint8Array} - Representación BCD del número
 */
function convertNumberToBCD(number, byteLength) {
  const result = new Uint8Array(byteLength);
  const str = number.toString().padStart(byteLength * 2, '0');
  
  for (let i = 0; i < byteLength; i++) {
    const highNibble = parseInt(str.charAt(i * 2), 10);
    const lowNibble = parseInt(str.charAt(i * 2 + 1), 10);
    result[i] = (highNibble << 4) | lowNibble;
  }
  
  return result;
}

/**
 * Convierte un código de tiempo al formato de bytes BCD para STL
 * Optimizado para 24fps
 * @param {string} timecode - Código de tiempo en formato HH:MM:SS:FF
 * @param {boolean} verboseFlag - Activar logs detallados
 * @return {Uint8Array} - Representación en bytes del código de tiempo
 */
function convertTimecodeToBytes(timecode, verboseFlag) {
  // Resultado: 4 bytes para HH:MM:SS:FF en formato BCD
  const result = new Uint8Array(4);
  
  try {
    if (!timecode) {
      if (verboseFlag) Logger.log("Timecode vacío o nulo, retornando 00:00:00:00");
      return result; // Devolver 00:00:00:00
    }
    
    if (verboseFlag) Logger.log(`Convirtiendo timecode a bytes: ${timecode}`);
    
    // Normalizar el formato y extraer componentes
    const parts = timecode.replace(/[;.,]/g, ':').split(':');
    
    // Debe tener exactamente 4 partes: HH:MM:SS:FF
    if (parts.length !== 4) {
      if (verboseFlag) Logger.log(`Formato de timecode incorrecto: ${timecode}, se esperan 4 componentes`);
      return result; // Devolver 00:00:00:00
    }
    
    // Convertir cada componente a entero
    const hours = parseInt(parts[0], 10);
    const minutes = parseInt(parts[1], 10);
    const seconds = parseInt(parts[2], 10);
    let frames = parseInt(parts[3], 10);
    
    // Validar rangos (adaptado para 24fps)
    if (hours < 0 || hours > 23) {
      if (verboseFlag) Logger.log(`Horas fuera de rango (0-23): ${hours}`);
      return result;
    }
    
    if (minutes < 0 || minutes > 59) {
      if (verboseFlag) Logger.log(`Minutos fuera de rango (0-59): ${minutes}`);
      return result;
    }
    
    if (seconds < 0 || seconds > 59) {
      if (verboseFlag) Logger.log(`Segundos fuera de rango (0-59): ${seconds}`);
      return result;
    }
    
    // Para 24fps, frames debe estar entre 0-23
    if (frames < 0 || frames > 23) {
      if (verboseFlag) Logger.log(`Frames fuera de rango para 24fps (0-23): ${frames}, ajustando...`);
      frames = Math.max(0, Math.min(frames, 23));
    }
    
    // Convertir cada componente a BCD
    result[0] = (Math.floor(hours / 10) << 4) | (hours % 10);
    result[1] = (Math.floor(minutes / 10) << 4) | (minutes % 10);
    result[2] = (Math.floor(seconds / 10) << 4) | (seconds % 10);
    result[3] = (Math.floor(frames / 10) << 4) | (frames % 10);
    
    if (verboseFlag) {
      const hexOutput = Array.from(result).map(b => b.toString(16).padStart(2, '0')).join(' ');
      Logger.log(`Timecode convertido a bytes BCD: ${hexOutput}`);
    }
    
    return result;
    
  } catch (error) {
    if (verboseFlag) Logger.log(`Error al convertir timecode: ${error.message}`);
    return result; // Devolver 00:00:00:00 en caso de error
  }
}

/**
 * Procesa texto para formato STL con mejor manejo de caracteres especiales
 * @param {string} text - Texto del subtítulo
 * @param {boolean} verboseFlag - Activar logs detallados
 * @return {Uint8Array} - Array de bytes procesados
 */
function processTextForSTL(text, verboseFlag) {
  if (!text) return new Uint8Array(0);
  if (verboseFlag) Logger.log(`Procesando texto para STL: "${text}"`);
  
  // Separar en líneas si hay saltos de línea
  const lines = text.split(/\r?\n/).filter(line => line.trim().length > 0);
  const maxLines = 2; // Máximo 2 líneas por subtítulo
  
  if (verboseFlag && lines.length > maxLines) {
    Logger.log(`Advertencia: Más de ${maxLines} líneas detectadas, se truncará a ${maxLines} líneas`);
  }
  
  // Tomar solo las primeras 2 líneas de subtítulos
  const processedLines = lines.slice(0, maxLines);
  
  // Procesar cada línea y convertir caracteres especiales
  let processedText = "";
  for (let i = 0; i < processedLines.length; i++) {
    // Añadir la línea procesada
    processedText += processedLines[i];
    
    // Añadir retorno de carro si no es la última línea
    if (i < processedLines.length - 1) {
      processedText += "\r\n";
    }
  }
  
  // Procesar caracteres especiales y convertir a bytes según CP437
  const bytes = processSpecialCharsForSTL(processedText, verboseFlag);
  
  // Añadir carácter de fin de texto (0x8F según EBU)
  const result = new Uint8Array(bytes.length + 1);
  result.set(bytes, 0);
  result[bytes.length] = 0x8F; // Carácter de fin de texto
  
  if (verboseFlag) {
    Logger.log(`Texto procesado con ${result.length} bytes (incluyendo EOT)`);
  }
  
  return result;
}

/**
 * Procesa caracteres especiales para STL usando CP437
 * @param {string} text - Texto a procesar
 * @param {boolean} verboseFlag - Activar logs detallados
 * @return {Uint8Array} - Array de bytes de caracteres mapeados
 */
function processSpecialCharsForSTL(text, verboseFlag) {
  if (!text) return new Uint8Array(0);
  
  const bytes = new Uint8Array(text.length);
  
  for (let i = 0; i < text.length; i++) {
    const char = text.charAt(i);
    const code = mapCharToCP437(char);
    bytes[i] = code;
    
    if (verboseFlag && code !== char.charCodeAt(0)) {
      Logger.log(`Mapeando carácter especial '${char}' a byte ${code} (0x${code.toString(16).padStart(2, '0')})`);
    }
  }
  
  return bytes;
}

/**
 * Mapea un carácter individual a su representación en CP437
 * basado en la tabla de caracteres de Subtitle Edit
 * @param {string} char - Carácter a mapear
 * @return {number} - Valor de byte mapeado
 */
function mapCharToCP437(char) {
  // Tabla de mapeo para caracteres especiales en español
  const charMap = {
    // Letras acentuadas minúsculas
    'á': 0xA0, // á
    'é': 0x82, // é
    'í': 0xA1, // í
    'ó': 0xA2, // ó
    'ú': 0xA3, // ú
    'ü': 0x81, // ü
    'ñ': 0xA4, // ñ
    
    // Letras acentuadas mayúsculas
    'Á': 0xB5, // Á
    'É': 0x90, // É
    'Í': 0xD6, // Í
    'Ó': 0xE0, // Ó
    'Ú': 0xE9, // Ú
    'Ü': 0x9A, // Ü
    'Ñ': 0xA5, // Ñ
    
    // Signos de puntuación específicos
    '¿': 0xA8, // ¿
    '¡': 0xAD, // ¡
    
    // Caracteres de control
    '\r': 0x0D, // CR (Carriage Return)
    '\n': 0x0A, // LF (Line Feed)
    ' ': 0x20, // Espacio
    
    // Otros caracteres comunes
    '«': 0xAE, // comilla angular izquierda
    '»': 0xAF, // comilla angular derecha
    '"': 0x22, // comillas dobles
    '"': 0x22, // comillas dobles inglesas
    '"': 0x22, // comillas dobles inglesas
    "'": 0x27, // comilla simple
    // Caracteres de guión
    '–': 0x2D, // guión
    '—': 0x2D  // guión largo
  };
  
  // Si el carácter está en el mapa, devolver su valor mapeado
  if (charMap[char] !== undefined) {
    return charMap[char];
  }
  
  // Para otros caracteres, usar su valor ASCII/Unicode si está en el rango CP437
  const code = char.charCodeAt(0);
  if (code >= 0x20 && code <= 0x7F) { // Rango ASCII imprimible
    return code;
  }
  
  // Para caracteres no mapeados, usar espacio como fallback
  return 0x20; // Espacio (32 decimal)
} 