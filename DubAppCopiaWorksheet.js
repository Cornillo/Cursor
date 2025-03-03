function call(){
  //    const nuevaSS = crearCopiaVacia("1d258eRpy3NovzQ1wKX_WVRCfXhxc2u2pwpYUuTDyfIw", "DubAppActive02");
      const nuevaSS = crearCopiaVacia("17HuNNBf4ZzM7kn3zC8VjszdCtiMBEdnUXHLe9E3vK00", "DubAppActive02");
      Logger.log('Enlace a la nueva copia: ' + nuevaSS.getUrl());
  }
   

/**
 * Crea una copia de una spreadsheet manteniendo solo los encabezados y las hojas especiales intactas
 * @param {string} spreadsheetId - ID de la spreadsheet a copiar
 * @param {string} nombreCopia - Nombre que tendrá la nueva spreadsheet
 * @return {Spreadsheet} Nueva spreadsheet creada
 */
function crearCopiaVacia(spreadsheetId, nombreCopia) {
  const ssOrigen = SpreadsheetApp.openById(spreadsheetId);
  const ssNueva = SpreadsheetApp.create(nombreCopia);
  const hojasOrigen = ssOrigen.getSheets();
  const primeraHojaNueva = ssNueva.getSheets()[0];
  
  hojasOrigen.forEach((hojaOrigen, index) => {
    const nombreHoja = hojaOrigen.getName();
    
    if (nombreHoja === 'Index' || nombreHoja === 'DWO-LogLabels') {
      // Copiar toda la hoja exactamente como está
      const hojaCopiada = hojaOrigen.copyTo(ssNueva);
      
      // Si existe una hoja con el mismo nombre, eliminarla
      const hojaExistente = ssNueva.getSheetByName(nombreHoja);
      if (hojaExistente && hojaExistente.getName() !== 'Sheet1') {
        ssNueva.deleteSheet(hojaExistente);
      }
      
      hojaCopiada.setName(nombreHoja);
    } else {
      const hojaNueva = index === 0 
        ? primeraHojaNueva.setName(nombreHoja)
        : ssNueva.insertSheet(nombreHoja);
      
      const numColumnas = hojaOrigen.getLastColumn();
      
      // Copiar las dos primeras filas
      const dosFilas = hojaOrigen.getRange(1, 1, 2, numColumnas);
      const destinoDosFilas = hojaNueva.getRange(1, 1, 2, numColumnas);
      
      // Copiar valores y formatos manualmente
      destinoDosFilas.setValues(dosFilas.getValues());
      destinoDosFilas.setBackgrounds(dosFilas.getBackgrounds());
      destinoDosFilas.setTextStyles(dosFilas.getTextStyles());
      destinoDosFilas.setDataValidations(dosFilas.getDataValidations());
      destinoDosFilas.setNumberFormats(dosFilas.getNumberFormats());
      destinoDosFilas.setHorizontalAlignments(dosFilas.getHorizontalAlignments());
      destinoDosFilas.setVerticalAlignments(dosFilas.getVerticalAlignments());
      destinoDosFilas.setWraps(dosFilas.getWraps());
      
      // Limpiar solo los valores de la segunda fila
      const destinoFila2 = hojaNueva.getRange(2, 1, 1, numColumnas);
      const valoresVacios = Array(numColumnas).fill('');
      destinoFila2.setValues([valoresVacios]);
      
      // Copiar anchos de columna
      for (let col = 1; col <= numColumnas; col++) {
        hojaNueva.setColumnWidth(col, hojaOrigen.getColumnWidth(col));
      }
    }
  });
  
  // Eliminar Sheet1 si aún existe
  const sheet1 = ssNueva.getSheetByName('Sheet1');
  if (sheet1) {
    ssNueva.deleteSheet(sheet1);
  }
  
  // Después de copiar todas las hojas, eliminar filas y columnas en blanco
  const hojasNuevas = ssNueva.getSheets();
  hojasNuevas.forEach(hoja => eliminarFilasColumnasVacias(hoja));
  
  // Obtener y registrar el enlace de la nueva spreadsheet
  const urlNuevaCopia = ssNueva.getUrl();
  Logger.log('Nueva spreadsheet creada: ' + urlNuevaCopia);
  
  return ssNueva;
}

/**
 * Elimina las filas y columnas en blanco de una hoja, preservando las dos primeras filas
 * @param {Sheet} hoja - Hoja a procesar
 */
function eliminarFilasColumnasVacias(hoja) {
  const maxFilas = hoja.getMaxRows();
  const maxCols = hoja.getMaxColumns();
  const datosUsados = hoja.getDataRange();
  const filasUsadas = Math.max(datosUsados.getNumRows(), 2); // Asegurar que mantenemos al menos 2 filas
  const colsUsadas = datosUsados.getNumColumns();
  
  // Eliminar filas vacías al final, empezando desde la fila 3
  if (maxFilas > filasUsadas) {
    hoja.deleteRows(filasUsadas + 1, maxFilas - filasUsadas);
  }
  
  // Eliminar columnas vacías al final
  if (maxCols > colsUsadas) {
    hoja.deleteColumns(colsUsadas + 1, maxCols - colsUsadas);
  }
}
