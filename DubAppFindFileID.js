/**
 * Mueve un archivo de Google Drive entre carpetas
 * @param {string} sourceFolderId - ID de la carpeta origen en GDrive
 * @param {string} targetFolderId - ID de la carpeta destino en GDrive
 * @param {string} fileName - Nombre del archivo a buscar
 * @return {string} ID del archivo movido
 */
function moveAndConfigureFile(sourceFolderId, targetFolderId, fileName) {
  try {
    // 1. Buscar el archivo por nombre en la carpeta origen
    const sourceFolder = DriveApp.getFolderById(sourceFolderId);
    const files = sourceFolder.getFilesByName(fileName);
    
    if (!files.hasNext()) {
      throw new Error('No se encontró el archivo: ' + fileName);
    }
    
    // Obtener el primer archivo que coincida con el nombre
    const file = files.next();
    const fileId = file.getId();
    
    // 2. Mover el archivo a la carpeta destino
    const targetFolder = DriveApp.getFolderById(targetFolderId);
    file.moveTo(targetFolder);
    
    // 3. Retornar el ID del archivo
    return fileId;
    
  } catch (error) {
    console.error('Error en moveAndConfigureFile: ' + error.message);
    throw error;
  }
}

/**
 * Función de prueba 
 */
function testMoveFile() {
  try {
    const sourceId = '1cbymeAzgSLyakBymKvzpZDRVjPo2JTDy';
    const targetId = '1x5Q_9ab3OHD3RL69k_U_OWHjUp0ULuse';
    const filename = 'QCIn - Shifting gears - S1 - Episode F003.pdf';
    
    const fileId = moveAndConfigureFile(sourceId, targetId, filename);
    Logger.log('Archivo movido con éxito. ID: ' + fileId);
    return fileId;
  } catch (error) {
    Logger.log('Error: ' + error.message);
    return null;
  }
}
