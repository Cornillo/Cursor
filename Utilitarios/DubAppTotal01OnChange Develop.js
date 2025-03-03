function registrarActivacionTrigger(e) {
  try {
    // Obtener información básica del evento
    const source = e.source;
    const activeSheet = source.getActiveSheet();
    const sheetName = activeSheet.getName();
    
    // Excluir cambios en la hoja TriggerLog
    if (sheetName === 'TriggerLog') {
      console.log('Cambio en TriggerLog ignorado');
      return;
    }
    
    // Obtener la fecha y hora actual
    const timestamp = new Date().toISOString();
    const changeType = e.changeType;
    const user = Session.getActiveUser().getEmail();
    
    // Variables para almacenar detalles del cambio
    let range = 'N/A';
    let detallesCambio = '';
    
    // Determinar el tipo de cambio y capturar detalles específicos
    switch(changeType) {
      case 'EDIT':
        detallesCambio = 'Edición de contenido';
        break;
      case 'INSERT_ROW':
        detallesCambio = `Fila insertada en posición: ${activeSheet.getActiveRange().getRow()}`;
        break;
      case 'INSERT_COLUMN':
        detallesCambio = `Columna insertada en posición: ${activeSheet.getActiveRange().getColumn()}`;
        break;
      case 'REMOVE_ROW':
        detallesCambio = 'Fila eliminada';
        break;
      case 'REMOVE_COLUMN':
        detallesCambio = 'Columna eliminada';
        break;
      case 'INSERT_GRID':
        detallesCambio = 'Nueva hoja insertada';
        break;
      case 'REMOVE_GRID':
        detallesCambio = 'Hoja eliminada';
        break;
      case 'FORMAT':
        detallesCambio = 'Cambio de formato';
        break;
      case 'OTHER':
        detallesCambio = 'Otro tipo de cambio';
        break;
    }
    
    // Intentar obtener el rango activo si existe
    const activeRange = activeSheet.getActiveRange();
    if (activeRange) {
      range = activeRange.getA1Notation();
    }
    
    // Guardar registro en una hoja específica
    const logSheet = source.getSheetByName('TriggerLog');
    if (!logSheet) {
      throw new Error('No se encontró la hoja TriggerLog');
    }
    
    // Obtener la última fila y agregar el nuevo registro
    const lastRow = logSheet.getLastRow();
    logSheet.getRange(lastRow + 1, 1, 1, 6).setValues([[
      timestamp,
      changeType,
      detallesCambio,
      user,
      sheetName,
      range
    ]]);
    
    console.log('Registro de activación guardado correctamente');
    
  } catch (error) {
    console.error('Error en registrarActivacionTrigger:', error);
    console.error('Stack:', error.stack);
  }
}

function crearTriggerOnChange() {
  // Eliminar triggers existentes para evitar duplicados
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'registrarActivacionTrigger') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // Crear nuevo trigger OnChange
  ScriptApp.newTrigger('registrarActivacionTrigger')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onChange()
    .create();
    
  console.log('Trigger OnChange creado exitosamente');
}

// Asignar esta función al trigger OnChange
