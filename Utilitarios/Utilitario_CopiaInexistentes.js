function compararWorksheets(idWorksheet1, idWorksheet2) {
  // Abrir los spreadsheets
  var ss1 = SpreadsheetApp.openById(idWorksheet1);
  var ss2 = SpreadsheetApp.openById(idWorksheet2);
  
  // Obtener todas las hojas
  var sheets1 = ss1.getSheets();
  var sheets2 = ss2.getSheets();
  
  // Procesar cada hoja del primer spreadsheet
  sheets1.forEach(function(sheet1) {
    var nombreHoja = sheet1.getName();
    
    // Ignorar las hojas especificadas
    if (nombreHoja === 'DWO-LogLabels' || nombreHoja === 'Index') {
      return;
    }
    
    // Buscar la hoja correspondiente en el segundo spreadsheet
    var sheet2 = ss2.getSheetByName(nombreHoja);
    if (!sheet2) {
      Logger.log('La hoja ' + nombreHoja + ' no existe en el segundo spreadsheet');
      return;
    }
    
    // Obtener datos
    var datos1 = sheet1.getDataRange().getValues();
    var datos2 = sheet2.getDataRange().getValues();
    
    if (datos1.length <= 1) return; // Ignorar si solo tiene encabezados
    
    // Determinar la columna clave según el nombre de la hoja
    var colClave = 0; // Por defecto columna A
    if (nombreHoja === 'DWO' || nombreHoja === 'DWO-ChannelEventType' || nombreHoja === 'DWO-Series') {
      // Para estas hojas, usar todas las columnas como clave
      var casosUnicos = new Set();
      var casosAInsertar = [];
      
      // Convertir filas a string para comparación
      datos1.slice(1).forEach(function(fila) {
        var clave = fila[colClave];
        if (!casosUnicos.has(clave)) {
          casosUnicos.add(clave);
          var existeEnDatos2 = datos2.slice(1).some(function(fila2) {
            return fila2[colClave] === clave;
          });
          if (!existeEnDatos2) {
            casosAInsertar.push(fila);
          }
        }
      });
      
      // Insertar casos faltantes
      if (casosAInsertar.length > 0) {
        sheet2.getRange(sheet2.getLastRow() + 1, 1, casosAInsertar.length, casosAInsertar[0].length)
              .setValues(casosAInsertar);
        
        // Log solo de la clave insertada
        casosAInsertar.forEach(function(caso) {
          Logger.log('Hoja: ' + nombreHoja + ' - Nueva clave insertada: ' + caso[colClave]);
        });
      }
      
    } else {
      // Para el resto de hojas, usar columna A como clave
      var valoresHoja2 = datos2.slice(1).map(function(fila) { return fila[colClave]; });
      var casosUnicos = new Set();
      var casosAInsertar = [];
      
      datos1.slice(1).forEach(function(fila) {
        var clave = fila[colClave];
        if (!casosUnicos.has(clave) && !valoresHoja2.includes(clave)) {
          casosUnicos.add(clave);
          casosAInsertar.push(fila);
        }
      });
      
      // Insertar casos faltantes
      if (casosAInsertar.length > 0) {
        sheet2.getRange(sheet2.getLastRow() + 1, 1, casosAInsertar.length, casosAInsertar[0].length)
              .setValues(casosAInsertar);
        
        // Log solo de la clave insertada
        casosAInsertar.forEach(function(caso) {
          Logger.log('Hoja: ' + nombreHoja + ' - Nueva clave insertada: ' + caso[colClave]);
        });
      }
    }
    
    Logger.log('Procesada hoja: ' + nombreHoja + '. Total casos insertados: ' + casosAInsertar.length);
  });
}

function ejecutarComparacion() {
  // IDs de los Google Spreadsheets a comparar
  var idWorksheet1 = '1tDAZAYjaS8sBqsiVMGTz60qp-La2h3u4ZegRf3Ye9d8'; // Reemplazar con el ID del primer spreadsheet
  var idWorksheet2 = '1UH_Fia--oOafojVbm1ahC94gdToxtGMXkP7ebe1yTGc'; // Reemplazar con el ID del segundo spreadsheet
  
  // Registrar inicio de la ejecución
  Logger.log('Iniciando comparación de spreadsheets');
  Logger.log('Spreadsheet 1 ID: ' + idWorksheet1);
  Logger.log('Spreadsheet 2 ID: ' + idWorksheet2);
  
  try {
    // Llamar a la función de comparación
    compararWorksheets(idWorksheet1, idWorksheet2);
    Logger.log('Comparación finalizada exitosamente');
  } catch (error) {
    Logger.log('Error durante la comparación: ' + error.toString());
    throw error;
  }
}


