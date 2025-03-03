/**
 * Crea una planilla para un usuario específico y retorna su ID
 * @param {string} userEmail - Email del usuario
 * @returns {Object} Objeto con el ID y URL de la planilla
 */
function obtenerPlanillaUsuario(userEmail) {
  try {
    // Obtener acceso a DubAppNoTrack
    const allIDs = databaseID.getID();
    
    // Crear nueva planilla desde el template
    const templateID = allIDs['userSheettTemplateID'];
    const template = DriveApp.getFileById(templateID);
    
    // Extraer el nombre del usuario del email
    const userName = userEmail.split('@')[0];
    const nombrePlanilla = `DubApp Edición externa / ${userName}`;
    
    // Crear copia en el folder específico
    const userTempFolder = DriveApp.getFolderById(allIDs["userTempID"]);
    const nuevaPlanilla = template.makeCopy(nombrePlanilla, userTempFolder);
    
    // Dar permisos al usuario
    nuevaPlanilla.setSharing(
      DriveApp.Access.ANYONE_WITH_LINK,
      DriveApp.Permission.EDIT
    );
    nuevaPlanilla.addEditor(userEmail);
    
    Logger.log('Nueva planilla creada para usuario: ' + userEmail);
    return nuevaPlanilla.getId();
    
  } catch (error) {
    Logger.log('Error en obtenerPlanillaUsuario: ' + error.toString());
    return null;
  }
}

/**
 * Actualiza una planilla específica con los datos del ProjectID
 */
function actualizarPlanillaConProjectID(projectID, sheetId) {
  try {
    Logger.log('Iniciando actualización con ProjectID: ' + projectID);
    
    // Abrir la planilla específica
    const targetSpreadsheet = SpreadsheetApp.openById(sheetId);
    
    // Temporalmente establecer como activa
    SpreadsheetApp.setActiveSpreadsheet(targetSpreadsheet);
    
    // Procesar el ProjectID usando la función existente
    DubAppOpenEditEvents.procesarProjectID(projectID);
    
    Logger.log('Actualización completada exitosamente');
    return {
      success: true,
      message: 'Planilla actualizada correctamente'
    };
    
  } catch (error) {
    Logger.log('Error en actualizarPlanillaConProjectID: ' + error.toString());
    return {
      success: false,
      message: 'Error: ' + error.message
    };
  }
}
