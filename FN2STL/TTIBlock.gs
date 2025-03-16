/**
 * Módulo para la creación de bloques TTI (Text and Timing Information)
 * Implementa la creación de los bloques de subtítulos según la especificación EBU Tech 3264-1991
 */

/**
 * Crea un bloque TTI (Text and Timing Information) para un subtítulo
 * @param {number} subtitleNumber - Número de subtítulo
 * @param {string} startTime - Código de tiempo de inicio (HH:MM:SS:FF)
 * @param {string} endTime - Código de tiempo de fin (HH:MM:SS:FF)
 * @param {string} text - Texto del subtítulo
 * @param {boolean} verboseFlag - Activar logs detallados
 * @return {Uint8Array} - Bloque TTI como array de bytes
 */
function createTTIBlock(subtitleNumber, startTime, endTime, text, verboseFlag) {
  try {
    if (verboseFlag) {
      Logger.log(`Creando bloque TTI para subtítulo #${subtitleNumber}`);
      Logger.log(`- Timecode inicio: ${startTime}`);
      Logger.log(`- Timecode fin: ${endTime}`);
      Logger.log(`- Texto: "${text}"`);
    }
    
    // Crear un buffer de 128 bytes (tamaño del bloque TTI)
    const ttiBlock = new Uint8Array(128);
    
    // Inicializar con espacios (código ASCII 32)
    for (let i = 0; i < ttiBlock.length; i++) {
      ttiBlock[i] = 32; // Espacio en ASCII
    }
    
    // 1. Número de subtítulo (SGN: Subtitle Group Number) - 1 byte
    // Siempre 0 para nuestro caso
    ttiBlock[0] = 0;
    
    // 2. Número de subtítulo (SN: Subtitle Number) - 2 bytes
    // Validar que el número de subtítulo esté en el rango correcto (1-65535)
    if (subtitleNumber < 1 || subtitleNumber > 65535) {
      if (verboseFlag) Logger.log(`Error: Número de subtítulo fuera de rango: ${subtitleNumber}`);
      subtitleNumber = Math.max(1, Math.min(subtitleNumber, 65535));
    }
    
    // Escribir el número de subtítulo en formato little-endian
    ttiBlock[1] = subtitleNumber & 0xFF;
    ttiBlock[2] = (subtitleNumber >> 8) & 0xFF;
    
    // 3. Posición de extensión (EBN: Extension Block Number) - 1 byte
    // Siempre 255 (0xFF) para indicar que no hay extensión
    ttiBlock[3] = 0xFF;
    
    // 4. Número de cuadro (CS: Cumulative Status) - 1 byte
    // Siempre 0 para nuestro caso
    ttiBlock[4] = 0;
    
    // 5-8. Código de tiempo de inicio (TCi: Time Code In) - 4 bytes
    // Convertir el código de tiempo a formato BCD
    const startTimeBytes = convertTimecodeToBytes(startTime, verboseFlag);
    ttiBlock[5] = startTimeBytes[0]; // Horas
    ttiBlock[6] = startTimeBytes[1]; // Minutos
    ttiBlock[7] = startTimeBytes[2]; // Segundos
    ttiBlock[8] = startTimeBytes[3]; // Frames
    
    // 9-12. Código de tiempo de fin (TCo: Time Code Out) - 4 bytes
    const endTimeBytes = convertTimecodeToBytes(endTime, verboseFlag);
    ttiBlock[9] = endTimeBytes[0]; // Horas
    ttiBlock[10] = endTimeBytes[1]; // Minutos
    ttiBlock[11] = endTimeBytes[2]; // Segundos
    ttiBlock[12] = endTimeBytes[3]; // Frames
    
    // 13. Formato vertical (VP: Vertical Position) - 1 byte
    // Posición 18 (cerca de la parte inferior de la pantalla)
    ttiBlock[13] = 18;
    
    // 14. Justificación (JC: Justification Code) - 1 byte
    // 2 = centrado
    ttiBlock[14] = 2;
    
    // 15. Comentario (CF: Comment Flag) - 1 byte
    // 0 = no es un comentario
    ttiBlock[15] = 0;
    
    // 16-128. Texto del subtítulo (TF: Text Field) - 112 bytes
    // Procesar el texto para el formato STL
    const processedText = processTextForSTL(text, verboseFlag);
    
    if (verboseFlag) {
      Logger.log(`Texto procesado: "${processedText}"`);
    }
    
    // Escribir el texto procesado en el bloque TTI
    for (let i = 0; i < processedText.length && i < 112; i++) {
      ttiBlock[16 + i] = processedText.charCodeAt(i);
    }
    
    if (verboseFlag) {
      // Mostrar los primeros 20 bytes y los últimos 8 bytes del bloque TTI para depuración
      let blockDebug = "Primeros 20 bytes: ";
      for (let i = 0; i < 20; i++) {
        blockDebug += ttiBlock[i].toString(16).padStart(2, '0') + " ";
      }
      blockDebug += "... Últimos 8 bytes: ";
      for (let i = 120; i < 128; i++) {
        blockDebug += ttiBlock[i].toString(16).padStart(2, '0') + " ";
      }
      Logger.log(blockDebug);
    }
    
    return ttiBlock;
  } catch (e) {
    Logger.log(`Error al crear bloque TTI: ${e.message}`);
    // Devolver un bloque vacío en caso de error
    return new Uint8Array(128);
  }
}

/**
 * Convierte un código de tiempo en formato HH:MM:SS:FF a un array de 4 bytes
 * siguiendo la especificación del formato EBU STL
 * 
 * @param {string} timecode - Código de tiempo en formato HH:MM:SS:FF
 * @param {boolean} verboseFlag - Activar logs detallados
 * @return {Uint8Array} Array de 4 bytes con el código de tiempo
 */
function convertTimecodeToBytes(timecode, verboseFlag) {
  // Array de 4 bytes para almacenar el timecode en formato BCD (Binary Coded Decimal)
  const bytes = new Uint8Array(4);
  
  // Validación básica del timecode
  if (!timecode || typeof timecode !== 'string') {
    Logger.log(`Error: Timecode inválido (no es string): ${timecode}`);
    return bytes; // Devolver 00:00:00:00 en caso de error
  }
  
  try {
    // Normalizar formato del timecode para asegurar HH:MM:SS:FF
    let normalizedTimecode = timecode;
    
    // Reemplazar puntos o comas por dos puntos en los primeros grupos
    normalizedTimecode = normalizedTimecode.replace(/[.,]/g, ':');
    
    // Separar por dos puntos
    const parts = normalizedTimecode.split(':');
    
    if (parts.length < 3) {
      Logger.log(`Error: Formato de timecode incorrecto (no suficientes partes): ${timecode}`);
      return bytes;
    }
    
    // Extraer componentes: horas, minutos, segundos, frames
    let hours = parseInt(parts[0], 10);
    let minutes = parseInt(parts[1], 10);
    let seconds = parseInt(parts[2], 10);
    let frames = parts.length > 3 ? parseInt(parts[3], 10) : 0;
    
    // Validar rangos (según especificación EBU STL)
    if (isNaN(hours)) hours = 0;
    if (hours > 23) hours = 23;
    
    if (isNaN(minutes)) minutes = 0;
    if (minutes > 59) minutes = 59;
    
    if (isNaN(seconds)) seconds = 0;
    if (seconds > 59) seconds = 59;
    
    if (isNaN(frames)) frames = 0;
    
    // Para STL25.01 (PAL):
    // - Los frames en el archivo deben estar entre 0-24 
    // - Subtitle Edit los muestra convertidos a milisegundos (1 frame = 40ms)
    if (frames > 24) frames = 24;
    
    // Para debug: mostrar el timecode normalizado
    if (verboseFlag) {
      Logger.log(`Convertido timecode ${timecode} a componentes: ${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}:${seconds.toString().padStart(2, '0')}:${frames.toString().padStart(2, '0')}`);
    }
    
    // Codificación BCD estricta según EBU STL
    // Cada byte contiene dos dígitos en formato BCD
    // Primera parte: decenas en los bits 7-4, unidades en los bits 3-0
    bytes[0] = (Math.floor(hours / 10) << 4) | (hours % 10);
    bytes[1] = (Math.floor(minutes / 10) << 4) | (minutes % 10);
    bytes[2] = (Math.floor(seconds / 10) << 4) | (seconds % 10);
    bytes[3] = (Math.floor(frames / 10) << 4) | (frames % 10);
    
    // El último bit del tercer byte (segundos) representa si es drop-frame
    // Para PAL (25fps), siempre debe ser 0
    bytes[2] &= 0x7F; // Asegurar drop frame bit = 0
    
    if (verboseFlag) {
      Logger.log(`Bytes de timecode (BCD): [${bytes[0].toString(16).padStart(2, '0')}, ${bytes[1].toString(16).padStart(2, '0')}, ${bytes[2].toString(16).padStart(2, '0')}, ${bytes[3].toString(16).padStart(2, '0')}]`);
    }
    
    return bytes;
  } catch (error) {
    Logger.log(`Error al convertir timecode: ${error.message} para ${timecode}`);
    return new Uint8Array(4);
  }
}

/**
 * Procesa el texto para el formato STL
 * @param {string} text - Texto del subtítulo
 * @param {boolean} verboseFlag - Activar logs detallados
 * @return {string} - Texto procesado
 */
function processTextForSTL(text, verboseFlag) {
  if (!text) return "";
  
  try {
    // 1. Reemplazar caracteres especiales
    let processedText = mapSpecialCharacters(text);
    
    // 2. Manejar saltos de línea
    // En STL, el salto de línea se representa con el carácter 0x8A
    processedText = processedText.replace(/\r\n|\r|\n/g, String.fromCharCode(0x8A));
    
    // 3. Limitar a 2 líneas como máximo
    const lines = processedText.split(String.fromCharCode(0x8A));
    if (lines.length > 2) {
      if (verboseFlag) Logger.log(`Advertencia: El subtítulo tiene más de 2 líneas. Se truncará.`);
      processedText = lines.slice(0, 2).join(String.fromCharCode(0x8A));
    }
    
    // Ya no truncamos el texto a 40 caracteres por línea
    // Dejamos que el programa que lee el STL maneje esto según sus propios estándares
    
    // 5. Terminar el texto con un carácter de fin de texto (0x8F)
    processedText += String.fromCharCode(0x8F);
    
    return processedText;
  } catch (e) {
    if (verboseFlag) Logger.log(`Error al procesar texto para STL: ${e.message}`);
    return "";
  }
}

/**
 * Mapea caracteres especiales a sus equivalentes en el estándar Latin/ISO-8859-1
 * @param {string} text - Texto a procesar
 * @return {string} - Texto con caracteres especiales mapeados
 */
function mapSpecialCharacters(text) {
  if (!text) return "";
  
  if (verboseFlag) {
    Logger.log(`Mapeando caracteres especiales de: "${text}"`);
  }
  
  // Tabla de mapeo para caracteres especiales usando Latin/ISO-8859-1 (CCT=01)
  // Esta tabla garantiza compatibilidad con Subtitle Edit
  const charMap = {
    // Vocales acentuadas en español - ISO-8859-1
    'á': String.fromCharCode(0xE1), // á (ISO-8859-1: 225)
    'é': String.fromCharCode(0xE9), // é (ISO-8859-1: 233)
    'í': String.fromCharCode(0xED), // í (ISO-8859-1: 237)
    'ó': String.fromCharCode(0xF3), // ó (ISO-8859-1: 243)
    'ú': String.fromCharCode(0xFA), // ú (ISO-8859-1: 250)
    'Á': String.fromCharCode(0xC1), // Á (ISO-8859-1: 193)
    'É': String.fromCharCode(0xC9), // É (ISO-8859-1: 201)
    'Í': String.fromCharCode(0xCD), // Í (ISO-8859-1: 205)
    'Ó': String.fromCharCode(0xD3), // Ó (ISO-8859-1: 211)
    'Ú': String.fromCharCode(0xDA), // Ú (ISO-8859-1: 218)
    
    // Otros caracteres especiales españoles
    'ñ': String.fromCharCode(0xF1), // ñ (ISO-8859-1: 241)
    'Ñ': String.fromCharCode(0xD1), // Ñ (ISO-8859-1: 209)
    'ü': String.fromCharCode(0xFC), // ü (ISO-8859-1: 252)
    'Ü': String.fromCharCode(0xDC), // Ü (ISO-8859-1: 220)
    
    // Signos de puntuación españoles
    '¿': String.fromCharCode(0xBF), // ¿ (ISO-8859-1: 191)
    '¡': String.fromCharCode(0xA1), // ¡ (ISO-8859-1: 161)
    
    // Caracteres especiales portugueses
    'ç': String.fromCharCode(0xE7), // ç (ISO-8859-1: 231)
    'Ç': String.fromCharCode(0xC7), // Ç (ISO-8859-1: 199)
    'ã': String.fromCharCode(0xE3), // ã (ISO-8859-1: 227)
    'Ã': String.fromCharCode(0xC3), // Ã (ISO-8859-1: 195)
    'õ': String.fromCharCode(0xF5), // õ (ISO-8859-1: 245)
    'Õ': String.fromCharCode(0xD5), // Õ (ISO-8859-1: 213)
    'â': String.fromCharCode(0xE2), // â (ISO-8859-1: 226)
    'Â': String.fromCharCode(0xC2), // Â (ISO-8859-1: 194)
    'ê': String.fromCharCode(0xEA), // ê (ISO-8859-1: 234)
    'Ê': String.fromCharCode(0xCA), // Ê (ISO-8859-1: 202)
    'ô': String.fromCharCode(0xF4), // ô (ISO-8859-1: 244)
    'Ô': String.fromCharCode(0xD4), // Ô (ISO-8859-1: 212)
    
    // Símbolos comunes
    '°': String.fromCharCode(0xB0), // grados (ISO-8859-1: 176)
    '®': String.fromCharCode(0xAE), // registered (ISO-8859-1: 174)
    '©': String.fromCharCode(0xA9), // copyright (ISO-8859-1: 169)
    
    // Caracteres que necesitan sustitución simple
    '…': '...',
    '—': '-',
    '–': '-',
    '"': '"',
    '"': '"',
    "'": "'",
    "'": "'",
    '€': 'EUR',
    '£': '#'
  };
  
  // Reemplazar caracteres especiales
  let result = text;
  
  // Convertir el texto a representación binaria para depuración
  if (verboseFlag) {
    let bytesString = "";
    for (let i = 0; i < text.length; i++) {
      bytesString += text.charCodeAt(i).toString(16).padStart(2, '0') + ' ';
    }
    Logger.log(`Bytes originales: ${bytesString}`);
  }
  
  // Reemplazar cada carácter especial
  for (const [original, replacement] of Object.entries(charMap)) {
    // Usar una expresión regular con la bandera 'g' para reemplazar todas las ocurrencias
    const regex = new RegExp(original.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&'), 'g');
    result = result.replace(regex, replacement);
  }
  
  // Mostrar el resultado final para depuración
  if (verboseFlag) {
    let bytesString = "";
    for (let i = 0; i < result.length; i++) {
      bytesString += result.charCodeAt(i).toString(16).padStart(2, '0') + ' ';
    }
    Logger.log(`Bytes mapeados: ${bytesString}`);
    Logger.log(`Resultado del mapeo: "${result}"`);
  }
  
  return result;
} 