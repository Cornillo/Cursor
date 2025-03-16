/**
 * Módulo para la creación del bloque GSI (General Subtitle Information)
 * Implementa la creación del encabezado del archivo STL según la especificación EBU Tech 3264-1991
 */

/**
 * Crea el bloque GSI (General Subtitle Information) para el archivo STL
 * @param {string} title - Título del programa
 * @param {string} country - Código de país (3 caracteres)
 * @param {string} languageCode - Código de idioma (2 caracteres)
 * @param {number} totalSubtitles - Número total de subtítulos
 * @param {boolean} verboseFlag - Activar logs detallados
 * @return {Uint8Array} - Bloque GSI como array de bytes
 */
function createGSIBlock(title, country, languageCode, totalSubtitles, verboseFlag) {
  if (verboseFlag) Logger.log(`Creando bloque GSI para "${title}"`);
  
  // Crear un buffer de 1024 bytes (tamaño del bloque GSI)
  const gsiBlock = new Uint8Array(1024);
  
  // Inicializar con espacios (código ASCII 32)
  for (let i = 0; i < gsiBlock.length; i++) {
    gsiBlock[i] = 32; // Espacio en ASCII
  }
  
  // 1. Código de Página de Caracteres (CPN: Code Page Number)
  // Usar "865" (Nordic) que tiene mejor compatibilidad con caracteres españoles en Subtitle Edit
  writeStringToBuffer(gsiBlock, "865", 0, 3);
  
  // 2. Código de Disco (DFC: Disk Format Code)
  // "STL25.01" para 25 fps (formato PAL) - Compatible con Subtitle Edit
  writeStringToBuffer(gsiBlock, "STL25.01", 3, 8);
  
  // 3. Código de Estándar de Visualización (DSC: Display Standard Code)
  // 1 = nivel 1 teletext - mejor para Subtitle Edit
  writeStringToBuffer(gsiBlock, "1", 11, 1);
  
  // 4. Código de Tabla de Caracteres (CCT: Character Code Table)
  // 01 = Latin/ISO 8859-1 (mejor para caracteres españoles)
  writeStringToBuffer(gsiBlock, "01", 12, 2);
  
  // 5. Código de Idioma (LC: Language Code)
  writeStringToBuffer(gsiBlock, languageCode, 14, 2);
  
  // 6. Título del Programa Original (OPT: Original Programme Title)
  const truncatedTitle = title.substring(0, Math.min(title.length, 32));
  writeStringToBuffer(gsiBlock, truncatedTitle, 16, 32);
  
  // 7. Título del Episodio Original (OET: Original Episode Title)
  writeStringToBuffer(gsiBlock, "", 48, 32);
  
  // 8. Título del Programa Traducido (TPT: Translated Programme Title)
  writeStringToBuffer(gsiBlock, truncatedTitle, 80, 32);
  
  // 9. Título del Episodio Traducido (TET: Translated Episode Title)
  writeStringToBuffer(gsiBlock, "", 112, 32);
  
  // 10. Nombre del Traductor (TN: Translator's Name)
  writeStringToBuffer(gsiBlock, "DUBAPP", 144, 32);
  
  // 11. Detalles de Contacto del Traductor (TCD: Translator's Contact Details)
  writeStringToBuffer(gsiBlock, "", 176, 32);
  
  // 12. Referencia de la Lista de Subtítulos (SLR: Subtitle List Reference)
  writeStringToBuffer(gsiBlock, "", 208, 16);
  
  // 13. Fecha de Creación (CD: Creation Date)
  const now = new Date();
  const year = now.getFullYear().toString().slice(-2);
  const month = (now.getMonth() + 1).toString().padStart(2, '0');
  const day = now.getDate().toString().padStart(2, '0');
  writeStringToBuffer(gsiBlock, `${year}${month}${day}`, 224, 6);
  
  // 14. Fecha de Revisión (RD: Revision Date)
  writeStringToBuffer(gsiBlock, `${year}${month}${day}`, 230, 6);
  
  // 15. Número de Revisión (RN: Revision Number)
  writeStringToBuffer(gsiBlock, "01", 236, 2);
  
  // 16. Número Total de Bloques TTI (TNB: Total Number of TTI Blocks)
  const ttiCountStr = totalSubtitles.toString().padStart(5, '0');
  writeStringToBuffer(gsiBlock, ttiCountStr, 238, 5);
  
  // 17. Número Total de Subtítulos (TNS: Total Number of Subtitles)
  // Igual que TNB en nuestro caso
  writeStringToBuffer(gsiBlock, ttiCountStr, 243, 5);
  
  // 18. Número Total de Grupos de Subtítulos (TNG: Total Number of Subtitle Groups)
  writeStringToBuffer(gsiBlock, "001", 248, 3);
  
  // 19. Número Máximo de Caracteres (MNC: Maximum Number of Characters)
  writeStringToBuffer(gsiBlock, "40", 251, 2);
  
  // 20. Número Máximo de Filas (MNR: Maximum Number of Rows)
  writeStringToBuffer(gsiBlock, "02", 253, 2);
  
  // 21. Estado del Código de Tiempo (TCS: Time Code Status)
  // 1 = Time code relates to the program
  writeStringToBuffer(gsiBlock, "1", 255, 1);
  
  // 22. Código de Tiempo de Inicio (TCP: Time Code: Start-of-Programme)
  writeStringToBuffer(gsiBlock, "00000000", 256, 8);
  
  // 23. Código de Tiempo de Fin (TCF: Time Code: First In-Cue)
  writeStringToBuffer(gsiBlock, "00000000", 264, 8);
  
  // 24. Código de País (CO: Country of Origin)
  writeStringToBuffer(gsiBlock, country.toUpperCase().padEnd(3, ' '), 272, 3);
  
  // 25. Tipo de Subtítulos (TND: Type of Subtitling)
  // 0 = Undefined
  writeStringToBuffer(gsiBlock, "0", 275, 1);
  
  if (verboseFlag) {
    Logger.log(`Bloque GSI creado correctamente para ${totalSubtitles} subtítulos`);
    
    // Log de los campos principales
    Logger.log(`- CPN: ${String.fromCharCode.apply(null, gsiBlock.slice(0, 3))}`);
    Logger.log(`- DFC: ${String.fromCharCode.apply(null, gsiBlock.slice(3, 11))}`);
    Logger.log(`- DSC: ${String.fromCharCode.apply(null, gsiBlock.slice(11, 12))}`);
    Logger.log(`- CCT: ${String.fromCharCode.apply(null, gsiBlock.slice(12, 14))}`);
    Logger.log(`- LC: ${String.fromCharCode.apply(null, gsiBlock.slice(14, 16))}`);
    Logger.log(`- OPT: ${String.fromCharCode.apply(null, gsiBlock.slice(16, 48)).trim()}`);
    Logger.log(`- TNB: ${String.fromCharCode.apply(null, gsiBlock.slice(238, 243))}`);
    Logger.log(`- CO: ${String.fromCharCode.apply(null, gsiBlock.slice(272, 275))}`);
  }
  
  return gsiBlock;
}

/**
 * Escribe una cadena en el bloque GSI
 * 
 * @param {Uint8Array} block - Bloque GSI
 * @param {string} value - Valor a escribir
 * @param {number} offset - Posición de inicio
 * @param {number} length - Longitud del campo
 */
function writeStringToBuffer(block, value, offset, length) {
  // Si el valor es nulo o undefined, usar cadena vacía
  value = value || "";
  
  // Convertir el valor a string y limitar a la longitud máxima
  const strValue = String(value).substring(0, length);
  
  // Escribir cada carácter en el bloque
  for (let i = 0; i < length; i++) {
    if (i < strValue.length) {
      // Si hay un carácter en esta posición, escribirlo
      block[offset + i] = strValue.charCodeAt(i);
    }
    // Si no hay carácter, el espacio ya está establecido por el relleno inicial
  }
} 