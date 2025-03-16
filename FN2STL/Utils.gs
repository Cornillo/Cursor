/**
 * Módulo de Utilidades para FN2STL
 * Funciones auxiliares para combinar bloques y otras operaciones
 */

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
    writeStringToBuffer(gsiBlock, ttiCountStr, 100, 5);
    writeStringToBuffer(gsiBlock, ttiCountStr, 105, 5);
    
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

/**
 * Valida la estructura del archivo STL generado
 * 
 * @param {Uint8Array} stlData - Datos del archivo STL
 * @return {boolean} true si la estructura es válida, false en caso contrario
 */
function validateSTLStructure(stlData) {
  logMessage("Validando estructura del archivo STL...");
  
  // Verificar tamaño mínimo (al menos el bloque GSI)
  if (stlData.length < CONFIG.GSI_BLOCK_SIZE) {
    logMessage("Error: Tamaño de archivo insuficiente");
    return false;
  }
  
  // Verificar identificador STL en el bloque GSI
  const dfc = String.fromCharCode.apply(null, stlData.slice(3, 11));
  if (!dfc.startsWith("STL")) {
    logMessage("Error: Identificador STL no encontrado");
    return false;
  }
  
  // Verificar que el número de bloques TTI coincide con lo indicado en el GSI
  const tnbStr = String.fromCharCode.apply(null, stlData.slice(238, 243));
  const tnb = parseInt(tnbStr, 10);
  
  const expectedSize = CONFIG.GSI_BLOCK_SIZE + (tnb * CONFIG.TTI_BLOCK_SIZE);
  if (stlData.length !== expectedSize) {
    logMessage("Error: Tamaño del archivo no coincide con el número de bloques TTI declarados");
    return false;
  }
  
  logMessage("Estructura del archivo STL validada correctamente");
  return true;
}

/**
 * Crea un archivo de texto para depuración con el contenido del archivo STL
 * Solo se usa cuando verboseFlag es true
 * 
 * @param {Uint8Array} stlData - Datos del archivo STL
 * @param {string} fileName - Nombre del archivo
 * @param {string} folderId - ID de la carpeta donde guardar el archivo
 */
function createDebugFile(stlData, fileName, folderId) {
  if (!verboseFlag) return;
  
  logMessage("Creando archivo de depuración...");
  
  try {
    // Extraer información del GSI
    const gsiInfo = extractGSIInfo(stlData);
    
    // Extraer información de los TTI
    const ttiInfo = extractTTIInfo(stlData);
    
    // Crear contenido del archivo de depuración
    let debugContent = "FN2STL - ARCHIVO DE DEPURACIÓN\n";
    debugContent += "================================\n\n";
    
    // Información GSI
    debugContent += "BLOQUE GSI:\n";
    debugContent += "-----------\n";
    for (const [key, value] of Object.entries(gsiInfo)) {
      debugContent += `${key}: ${value}\n`;
    }
    
    // Información TTI
    debugContent += "\nBLOQUES TTI:\n";
    debugContent += "------------\n";
    for (let i = 0; i < ttiInfo.length; i++) {
      debugContent += `\nSubtítulo #${i+1}:\n`;
      for (const [key, value] of Object.entries(ttiInfo[i])) {
        debugContent += `  ${key}: ${value}\n`;
      }
    }
    
    // Guardar archivo de depuración
    const folder = DriveApp.getFolderById(folderId);
    const debugFile = folder.createTextFile(fileName + ".debug.txt", debugContent);
    
    logMessage("Archivo de depuración creado: " + debugFile.getName());
  } catch (error) {
    logMessage("Error al crear archivo de depuración: " + error.message);
  }
}

/**
 * Extrae información del bloque GSI para depuración
 * 
 * @param {Uint8Array} stlData - Datos del archivo STL
 * @return {Object} Información del bloque GSI
 */
function extractGSIInfo(stlData) {
  const gsiBlock = stlData.slice(0, CONFIG.GSI_BLOCK_SIZE);
  
  return {
    "CPN (Code Page Number)": String.fromCharCode.apply(null, gsiBlock.slice(0, 3)),
    "DFC (Disk Format Code)": String.fromCharCode.apply(null, gsiBlock.slice(3, 11)),
    "DSC (Display Standard Code)": String.fromCharCode.apply(null, gsiBlock.slice(11, 12)),
    "LC (Language Code)": String.fromCharCode.apply(null, gsiBlock.slice(14, 16)),
    "OPT (Original Programme Title)": String.fromCharCode.apply(null, gsiBlock.slice(16, 48)).trim(),
    "TNB (Total Number of TTI Blocks)": String.fromCharCode.apply(null, gsiBlock.slice(238, 243)),
    "CO (Country of Origin)": String.fromCharCode.apply(null, gsiBlock.slice(274, 277))
  };
}

/**
 * Extrae información de los bloques TTI para depuración
 * 
 * @param {Uint8Array} stlData - Datos del archivo STL
 * @return {Array} Array con información de los bloques TTI
 */
function extractTTIInfo(stlData) {
  const tnbStr = String.fromCharCode.apply(null, stlData.slice(238, 243));
  const tnb = parseInt(tnbStr, 10);
  
  const ttiInfo = [];
  
  for (let i = 0; i < tnb; i++) {
    const offset = CONFIG.GSI_BLOCK_SIZE + (i * CONFIG.TTI_BLOCK_SIZE);
    const ttiBlock = stlData.slice(offset, offset + CONFIG.TTI_BLOCK_SIZE);
    
    // Extraer información básica del bloque TTI
    ttiInfo.push({
      "SN (Subtitle Number)": (ttiBlock[1] << 8) | ttiBlock[2],
      "TCI (Time Code In)": formatTimecodeFromBytes(ttiBlock.slice(5, 9)),
      "TCO (Time Code Out)": formatTimecodeFromBytes(ttiBlock.slice(9, 13)),
      "Text": extractTextFromTTI(ttiBlock)
    });
  }
  
  return ttiInfo;
}

/**
 * Extrae el texto de un bloque TTI
 * 
 * @param {Uint8Array} ttiBlock - Bloque TTI
 * @return {string} Texto extraído
 */
function extractTextFromTTI(ttiBlock) {
  // El texto comienza en el byte 16 y puede ocupar hasta 112 bytes
  const textField = ttiBlock.slice(16, 128);
  
  // Convertir a string hasta encontrar un espacio o fin del campo
  let text = "";
  for (let i = 0; i < textField.length; i++) {
    if (textField[i] === 32 && textField.slice(i).every(b => b === 32)) {
      break; // Detener si encontramos solo espacios hasta el final
    }
    
    // Convertir byte a carácter
    // 0x8A es el código para nueva línea en STL
    if (textField[i] === 0x8A) {
      text += "\n";
    } else {
      text += String.fromCharCode(textField[i]);
    }
  }
  
  return text;
}

/**
 * Convierte un array de 4 bytes a un código de tiempo en formato HH:MM:SS:FF
 * @param {Uint8Array} bytes - Array de 4 bytes con el código de tiempo
 * @return {string} - Código de tiempo en formato HH:MM:SS:FF
 */
function formatTimecodeFromBytes(bytes) {
  if (!bytes || bytes.length < 4) {
    return "00:00:00:00";
  }
  
  try {
    // Extraer los dígitos de cada byte
    const hours = ((bytes[0] >> 4) & 0x0F) * 10 + (bytes[0] & 0x0F);
    const minutes = ((bytes[1] >> 4) & 0x0F) * 10 + (bytes[1] & 0x0F);
    const seconds = ((bytes[2] >> 4) & 0x0F) * 10 + (bytes[2] & 0x0F);
    const frames = ((bytes[3] >> 4) & 0x0F) * 10 + (bytes[3] & 0x0F);
    
    // Formatear como HH:MM:SS:FF
    return `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}:${seconds.toString().padStart(2, '0')}:${frames.toString().padStart(2, '0')}`;
  } catch (error) {
    Logger.log(`Error al formatear timecode desde bytes: ${error.message}`);
    return "00:00:00:00";
  }
}

/**
 * Convierte un código de tiempo (HH:MM:SS:FF) a formato BCD para STL
 * @param {string} timecode - Código de tiempo en formato HH:MM:SS:FF
 * @param {boolean} verboseFlag - Activar logs detallados
 * @return {Uint8Array} - Array de 4 bytes con el código de tiempo en formato BCD
 */
function convertTimecodeToBytes(timecode, verboseFlag) {
  if (!timecode) return new Uint8Array([0, 0, 0, 0]);
  
  try {
    // Extraer componentes del código de tiempo
    const parts = timecode.split(':');
    if (parts.length !== 4) {
      if (verboseFlag) Logger.log(`Error: Formato de timecode inválido: ${timecode}`);
      return new Uint8Array([0, 0, 0, 0]);
    }
    
    const hours = parseInt(parts[0], 10);
    const minutes = parseInt(parts[1], 10);
    const seconds = parseInt(parts[2], 10);
    const frames = parseInt(parts[3], 10);
    
    // Validar rangos y ajustar si es necesario
    const validHours = Math.max(0, Math.min(hours, 23));
    const validMinutes = Math.max(0, Math.min(minutes, 59));
    const validSeconds = Math.max(0, Math.min(seconds, 59));
    
    // Para formato STL30.01 (30fps), el máximo de frames es 29
    const validFrames = Math.max(0, Math.min(frames, 29));
    
    // Convertir a BCD (Binary Coded Decimal)
    // En BCD, cada dígito decimal se codifica en 4 bits
    // Por ejemplo, el número 25 se codifica como 0010 0101 (0x25)
    const hoursBCD = ((Math.floor(validHours / 10) << 4) | (validHours % 10)) & 0xFF;
    const minutesBCD = ((Math.floor(validMinutes / 10) << 4) | (validMinutes % 10)) & 0xFF;
    const secondsBCD = ((Math.floor(validSeconds / 10) << 4) | (validSeconds % 10)) & 0xFF;
    const framesBCD = ((Math.floor(validFrames / 10) << 4) | (validFrames % 10)) & 0xFF;
    
    if (verboseFlag) {
      Logger.log(`Convertido timecode ${timecode} a componentes: ${validHours.toString().padStart(2, '0')}:${validMinutes.toString().padStart(2, '0')}:${validSeconds.toString().padStart(2, '0')}:${validFrames.toString().padStart(2, '0')}`);
      Logger.log(`Convertido timecode ${timecode} a bytes BCD:`);
      Logger.log(`- Horas: ${validHours} -> ${hoursBCD.toString(16).padStart(2, '0')}`);
      Logger.log(`- Minutos: ${validMinutes} -> ${minutesBCD.toString(16).padStart(2, '0')}`);
      Logger.log(`- Segundos: ${validSeconds} -> ${secondsBCD.toString(16).padStart(2, '0')}`);
      Logger.log(`- Frames: ${validFrames} -> ${framesBCD.toString(16).padStart(2, '0')}`);
      Logger.log(`Bytes de timecode (BCD): [${hoursBCD.toString(16).padStart(2, '0')}, ${minutesBCD.toString(16).padStart(2, '0')}, ${secondsBCD.toString(16).padStart(2, '0')}, ${framesBCD.toString(16).padStart(2, '0')}]`);
    }
    
    return new Uint8Array([hoursBCD, minutesBCD, secondsBCD, framesBCD]);
  } catch (e) {
    if (verboseFlag) Logger.log(`Error al convertir timecode a bytes: ${e.message}`);
    return new Uint8Array([0, 0, 0, 0]);
  }
}

/**
 * Guarda los datos del archivo STL en Google Drive
 * @param {Uint8Array} stlData - Datos binarios del archivo STL
 * @param {string} fileName - Nombre del archivo
 * @param {string} folderId - ID de la carpeta destino
 * @param {boolean} verboseFlag - Indicador para mostrar logs detallados
 * @return {object} Información del archivo creado
 */
function saveSTLFile(stlData, fileName, folderId, verboseFlag) {
  if (verboseFlag) Logger.log(`Guardando archivo STL: ${fileName}`);
  
  try {
    // Asegurarnos de que el nombre del archivo tenga la extensión .STL
    if (!fileName.toLowerCase().endsWith('.stl')) {
      fileName += '.STL';
    }
    
    // Convertir Uint8Array a Blob
    // En Google Apps Script necesitamos convertir Uint8Array a Array normal primero
    // para que Utilities.newBlob funcione correctamente
    const regularArray = [].slice.call(stlData);
    const blob = Utilities.newBlob(regularArray, 'application/octet-stream', fileName);
    
    // Obtener la carpeta de destino
    const folder = DriveApp.getFolderById(folderId);
    
    // Guardar el archivo
    const file = folder.createFile(blob);
    
    // Establecer permisos de acceso (cualquiera con el enlace puede ver)
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    if (verboseFlag) {
      Logger.log(`Archivo STL guardado con éxito. ID: ${file.getId()}`);
      Logger.log(`URL: ${file.getUrl()}`);
    }
    
    // Devolver objeto con información del archivo creado
    return {
      fileId: file.getId(),
      fileName: file.getName(),
      fileUrl: file.getUrl(),
      downloadUrl: `https://drive.google.com/uc?id=${file.getId()}&export=download`
    };
    
  } catch (error) {
    Logger.log(`Error al guardar archivo STL: ${error.message}`);
    throw error;
  }
} 