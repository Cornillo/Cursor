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
 * Optimizado para 24fps - Formato EBU STL compatible con Subtitle Edit
 * Según especificación EBU Tech 3264-1991
 * 
 * @param {string} timecode - Código de tiempo en formato HH:MM:SS:FF
 * @param {boolean} verboseFlag - Activar logs detallados
 * @return {Uint8Array} - Representación en bytes del código de tiempo (4 bytes BCD)
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
    
    // Normalizar el formato: asegurarnos que cualquier separador (;.,) sea reemplazado por :
    // Si viene en formato con milisegundos HH:MM:SS,mmm o HH:MM:SS.mmm, convertir a frames
    const normalizedTC = timecode.replace(/[;]/g, ':');
    
    let hours, minutes, seconds, frames;
    
    // Comprobar si el formato es HH:MM:SS.mmm o HH:MM:SS,mmm (con milisegundos)
    if (normalizedTC.includes(',') || normalizedTC.includes('.')) {
      // Dividir en componentes
      const regex = /^(\d{1,2}):(\d{1,2}):(\d{1,2})[,.](\d{1,3})$/;
      const matches = normalizedTC.match(regex);
      
      if (matches && matches.length === 5) {
        hours = parseInt(matches[1], 10);
        minutes = parseInt(matches[2], 10);
        seconds = parseInt(matches[3], 10);
        const milliseconds = parseInt(matches[4], 10);
        
        // Convertir milisegundos a frames (para 24fps)
        // 1 segundo = 24 frames, por lo que 1 frame = 41.666... ms
        frames = Math.round(milliseconds / 41.666);
      } else {
        if (verboseFlag) Logger.log(`Formato de timecode inválido: ${timecode}`);
        return result;
      }
    } else {
      // Formato estándar HH:MM:SS:FF
      const parts = normalizedTC.split(':');
      
      // Debe tener exactamente 4 partes: HH:MM:SS:FF
      if (parts.length !== 4) {
        if (verboseFlag) Logger.log(`Formato de timecode incorrecto: ${timecode}, se esperan 4 componentes`);
        return result; // Devolver 00:00:00:00
      }
      
      // Convertir cada componente a entero
      hours = parseInt(parts[0], 10);
      minutes = parseInt(parts[1], 10);
      seconds = parseInt(parts[2], 10);
      frames = parseInt(parts[3], 10);
    }
    
    // Validar rangos según especificación EBU para 24fps
    if (hours < 0 || hours > 23) {
      if (verboseFlag) Logger.log(`Horas fuera de rango (0-23): ${hours}`);
      hours = Math.max(0, Math.min(hours, 23));
    }
    
    if (minutes < 0 || minutes > 59) {
      if (verboseFlag) Logger.log(`Minutos fuera de rango (0-59): ${minutes}`);
      minutes = Math.max(0, Math.min(minutes, 59));
    }
    
    if (seconds < 0 || seconds > 59) {
      if (verboseFlag) Logger.log(`Segundos fuera de rango (0-59): ${seconds}`);
      seconds = Math.max(0, Math.min(seconds, 59));
    }
    
    // Para 24fps, frames debe estar entre 0-23
    if (frames < 0 || frames > 23) {
      if (verboseFlag) Logger.log(`Frames fuera de rango para 24fps (0-23): ${frames}, ajustando a valor válido`);
      frames = Math.max(0, Math.min(frames, 23));
    }
    
    // Convertir cada componente a BCD (Binary Coded Decimal) según especificación EBU
    // Cada byte almacena dos dígitos decimales: el dígito de las decenas en los 4 bits altos
    // y el dígito de las unidades en los 4 bits bajos
    result[0] = (Math.floor(hours / 10) << 4) | (hours % 10);
    result[1] = (Math.floor(minutes / 10) << 4) | (minutes % 10);
    result[2] = (Math.floor(seconds / 10) << 4) | (seconds % 10);
    result[3] = (Math.floor(frames / 10) << 4) | (frames % 10);
    
    if (verboseFlag) {
      const hexOutput = Array.from(result).map(b => b.toString(16).padStart(2, '0')).join(' ');
      Logger.log(`Timecode convertido: ${hexOutput} (hex)`);
    }
    
    return result;
    
  } catch (error) {
    if (verboseFlag) Logger.log(`Error al convertir timecode: ${error.message}`);
    return result; // Devolver 00:00:00:00 en caso de error
  }
}

/**
 * Procesa texto para formato STL con mejor manejo de caracteres especiales
 * Optimizado para CP437 y formato 24fps compatible con Subtitle Edit
 * @param {string} text - Texto del subtítulo
 * @param {boolean} verboseFlag - Activar logs detallados
 * @return {Uint8Array} - Array de bytes procesados
 */
function processTextForSTL(text, verboseFlag) {
  if (!text) return new Uint8Array(0);
  if (verboseFlag) Logger.log(`Procesando texto para STL: "${text}"`);
  
  // Pre-normalización para garantizar un mapeo consistente
  text = text
    // Acentos específicos que han causado problemas
    .replace(/í/g, String.fromCharCode(0xED)) // Normalizar í (Unicode estándar)
    .replace(/ó/g, String.fromCharCode(0xF3)) // Normalizar ó (Unicode estándar)
    .replace(/ñ/g, String.fromCharCode(0xF1)) // Normalizar ñ (Unicode estándar)
    
    // Otros caracteres que podrían causar problemas
    .replace(/['']/g, "'") // Comillas simples estándar
    .replace(/[""]/g, '"') // Comillas dobles estándar
    .replace(/…/g, "...") // Puntos suspensivos
    .replace(/–|—/g, '-'); // Guiones
  
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
  
  // Procesar caracteres especiales y convertir a bytes según CP850
  const bytes = processSpecialCharsForSTL(processedText, verboseFlag);
  
  // Añadir carácter de fin de texto (0x8F según EBU)
  const result = new Uint8Array(bytes.length + 1);
  result.set(bytes, 0);
  result[bytes.length] = 0x8F; // Carácter de fin de texto
  
  if (verboseFlag) {
    Logger.log(`Texto procesado con ${result.length} bytes (incluyendo EOT)`);
    
    // Mostrar representación hexadecimal para debugging
    let hexOutput = "";
    for (let i = 0; i < Math.min(result.length, 20); i++) {
      hexOutput += result[i].toString(16).padStart(2, '0') + " ";
    }
    if (result.length > 20) hexOutput += "...";
    Logger.log(`Primeros bytes del texto procesado (hex): ${hexOutput}`);
  }
  
  return result;
}

/**
 * Procesa caracteres especiales para STL usando CP437 (DOS Latin US)
 * @param {string} text - Texto a procesar
 * @param {boolean} verboseFlag - Activar logs detallados
 * @return {Uint8Array} - Array de bytes de caracteres mapeados
 */
function processSpecialCharsForSTL(text, verboseFlag) {
  if (!text) return new Uint8Array(0);
  
  const bytes = new Uint8Array(text.length);
  let specialCharsFound = 0;
  
  for (let i = 0; i < text.length; i++) {
    const char = text.charAt(i);
    const code = mapCharToCP437(char, verboseFlag);
    bytes[i] = code;
    
    // Detectar caracteres especiales para logging
    if (code !== char.charCodeAt(0) || code > 127) {
      specialCharsFound++;
      if (verboseFlag) {
        Logger.log(`Mapeando carácter especial '${char}' (Unicode ${char.charCodeAt(0)}) a byte ${code} (0x${code.toString(16).padStart(2, '0')})`);
      }
    }
  }
  
  if (verboseFlag && specialCharsFound > 0) {
    Logger.log(`Total ${specialCharsFound} caracteres especiales mapeados en el texto`);
  }
  
  return bytes;
}

/**
 * Mapea caracteres individuales a su representación en CP437 (DOS Latin US)
 * Valores verificados con tabla de referencia CP437
 * @param {string} char - Carácter individual a mapear
 * @param {boolean} verboseFlag - Activar logs detallados
 * @return {number} - Valor CP437 para el carácter
 */
function mapCharToCP437(char, verboseFlag) {
  // Tabla de mapeo para caracteres especiales en español usando CP437 (DOS Latin US)
  const charMap = {
    'á': 0xA0, // 160
    'é': 0x82, // 130
    'í': 0xA1, // 161 - Valor correcto para 'í' en CP437
    'ó': 0xA2, // 162 - Valor correcto para 'ó' en CP437
    'ú': 0xA3, // 163
    'ñ': 0xA4, // 164
    'Ñ': 0xA5, // 165
    'ü': 0x81, // 129
    'Ü': 0x9A, // 154
    '¡': 0xAD, // 173
    '¿': 0xBF, // 191
    '°': 0xF8, // 248
    '®': 0xAE, // 174
    'ª': 0xA6, // 166
    'º': 0xA7, // 167
    '¬': 0xAA, // 170
    // Otros caracteres especiales
    // Valores de mapeo optimizados para CP437 y verificados con tabla de referencia
    '«': 0xAE, // 174
    '»': 0xAF, // 175
    '"': 0x22, // 34
    "'": 0x27, // 39
    '–': 0x2D, // 45
    '—': 0x2D,  // 45
    
    // Letras acentuadas mayúsculas
    'Á': 0xB5, // 181
    'É': 0x90, // 144
    'Í': 0xD6, // 214
    'Ó': 0xE0, // 224
    'Ú': 0xE9, // 233
    
    // Caracteres de control
    '\r': 0x0D, // CR
    '\n': 0x0A, // LF
    ' ': 0x20, // Espacio
  };
  
  // Obtener el valor mapeado o usar el código ASCII si no está en el mapa
  if (charMap[char] !== undefined) {
    return charMap[char];
  } else if (char.charCodeAt(0) <= 127) {
    return char.charCodeAt(0); // ASCII estándar (0-127)
  } else {
    // Para caracteres no mapeados, usar '?' (63)
    if (verboseFlag) Logger.log(`Carácter no mapeado: "${char}" (${char.charCodeAt(0)})`);
    return 63; // Signo de interrogación para caracteres no mapeados
  }
} 