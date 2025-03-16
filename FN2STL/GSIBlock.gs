/**
 * Módulo para la creación del bloque GSI (General Subtitle Information)
 * Implementa la creación del encabezado del archivo STL según la especificación EBU Tech 3264-1991
 */

/**
 * Crea el bloque GSI (General Subtitle Information) para el archivo STL
 * con formato 23.976fps y mejor compatibilidad con Subtitle Edit
 * 
 * @param {Object} options - Opciones para el bloque GSI
 * @param {string} options.programName - Nombre del programa (max 32 chars)
 * @param {string} options.countryOrigin - Código ISO del país de origen (3 chars)
 * @param {string} options.languageCode - Código de idioma (2 chars)
 * @param {number} options.subtitleCount - Número de subtítulos en el archivo
 * @param {boolean} options.verboseFlag - Habilitar logs detallados
 * @return {Uint8Array} - Bloque GSI formateado (1024 bytes)
 */
function createGSIBlock(options) {
  const {
    programName = "Default Program",
    countryOrigin = "ARG",
    languageCode = "0A", // Español
    subtitleCount = 0,
    verboseFlag = false
  } = options || {};
  
  if (verboseFlag) {
    Logger.log("=== Creando Bloque GSI ===");
    Logger.log(`Programa: ${programName}`);
    Logger.log(`País: ${countryOrigin}`);
    Logger.log(`Idioma: ${languageCode}`);
    Logger.log(`Subtítulos: ${subtitleCount}`);
  }
  
  // Crear un buffer de 1024 bytes lleno de espacios (0x20)
  const buffer = new Uint8Array(1024).fill(0x20);
  
  // ===== ENCABEZADO DE ARCHIVO =====
  
  // Código de página de caracteres (CPN) - byte 0-2 - "850" (DOS Latin Extended)
  writeStringToBuffer(buffer, 0, "850");
  
  // Código de formato de disco (DFC) - bytes 3-10 - "STL24.01" (24fps exacto)
  writeStringToBuffer(buffer, 3, "STL24.01");
  
  // Código estándar de pantalla (DSC) - byte 11 - "0" (Open subtitling)
  buffer[11] = 0x30; // '0' (Open subtitling)
  
  // Código de tabla de caracteres (CCT) - bytes 12-13 - "00" (Latin/CP437)
  writeStringToBuffer(buffer, 12, "00");
  
  // Código de idioma (LC) - bytes 14-15 - "0A" (Español)
  writeStringToBuffer(buffer, 14, languageCode);
  
  // ===== INFORMACIÓN DEL PROGRAMA =====
  
  // Título original del programa - bytes 16-47
  writeStringToBuffer(buffer, 16, programName.substring(0, 32).padEnd(32, ' '));
  
  // Título original de episodio - bytes 48-79
  writeStringToBuffer(buffer, 48, "");
  
  // Título traducido - bytes 80-111
  writeStringToBuffer(buffer, 80, "");
  
  // Información adicional - bytes 112-143
  writeStringToBuffer(buffer, 112, "");
  
  // ===== INFORMACIÓN TÉCNICA =====
  
  // Código de país de origen - bytes 274-276
  writeStringToBuffer(buffer, 274, countryOrigin);
  
  // Fecha de creación (YYMMDD) - bytes 304-309
  const now = new Date();
  const year = String(now.getFullYear()).substring(2);
  const month = String(now.getMonth() + 1).padStart(2, '0');
  const day = String(now.getDate()).padStart(2, '0');
  writeStringToBuffer(buffer, 304, `${year}${month}${day}`);
  
  // Número de subtítulos en el archivo - bytes 330-334
  const subtitleCountStr = String(subtitleCount).padStart(5, '0');
  writeStringToBuffer(buffer, 330, subtitleCountStr);
  
  // Código de inicio de TTI blocks - bytes 1012-1023
  writeStringToBuffer(buffer, 1020, "STL");
  buffer[1023] = 0x31; // '1'
  
  if (verboseFlag) {
    Logger.log("Bloque GSI creado con éxito");
    Logger.log(`Código de página: 850`);
    Logger.log(`Formato de disco: STL24.01`);
    Logger.log(`Estándar de pantalla: 0 (Open subtitling)`);
    Logger.log(`Tabla de caracteres: 00 (Latin/CP437)`);
    Logger.log(`Fecha de creación: ${year}${month}${day}`);
    Logger.log(`Número de subtítulos: ${subtitleCountStr}`);
  }
  
  return buffer;
}

/**
 * Escribe una cadena en el bloque GSI
 * 
 * @param {Uint8Array} block - Bloque GSI
 * @param {number} offset - Posición de inicio
 * @param {string} value - Valor a escribir
 */
function writeStringToBuffer(block, offset, value) {
  // Convertir el valor a string y limitar a la longitud máxima
  const strValue = String(value).substring(0, 32);
  
  // Escribir cada carácter en el bloque
  for (let i = 0; i < 32; i++) {
    if (i < strValue.length) {
      // Si hay un carácter en esta posición, escribirlo
      block[offset + i] = strValue.charCodeAt(i);
    }
    // Si no hay carácter, el espacio ya está establecido por el relleno inicial
  }
} 