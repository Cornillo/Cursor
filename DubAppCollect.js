function collect(e, install) {
    //OnChange trigger for tables DubAppActive01 / DubAppTotal01
    //Version dev 15/3/24
    //Database architecture doc https://docs.google.com/document/d/1vtoE9m8mkgFHjs-C27YjS55kobLsNK01086UE4s0h38/edit
  
    const allIDs = databaseID.getID();
    const dubAppControlID = allIDs["controlID"];
    const timezone = allIDs["timezone"];
    const timestamp_format = allIDs["timestamp_format"];
  
    //Capture On change info
    const actionType = e.changeType;
    
    const MAX_RETRIES = 3;
    const RETRY_DELAY = 1000; // 1 segundo

    //Error capture
    try {
      
      //Process only Inserts and edit
      if(actionType!="INSERT_ROW" && actionType!="EDIT" ) {return;}
  
      var SheetChanged = e.source.getActiveSheet();
      var sheet2track = SheetChanged.getName();
  
      //If tables inhibited of logging in Worksheet
      if(sheet2track=="Index") {return;}
  
      // Eliminar variables duplicadas y moverlas arriba
      let SheetLabelRow, SheetLabelActionRow, keyCurrent, lastUserColCurrent;
      
      /*Connect label*/
      const labelData = retryOperation(() => {
          const SheetSheetLabelRow = getLabelData();
          const SheetLabelNDX = SheetSheetLabelRow.map(r => r[0]);
          const labelRow = SheetLabelNDX.indexOf(sheet2track, 0);
          
          if(labelRow === -1) return null; // Si la tabla no es rastreable
  
          SheetLabelRow = SheetSheetLabelRow[labelRow];
          SheetLabelActionRow = SheetSheetLabelRow[labelRow + 1];
          keyCurrent = SheetLabelActionRow.indexOf("K") - 1;
          lastUserColCurrent = SheetLabelRow.indexOf("Last user") - 1;
  
          return {
              rangeCurrent: SheetChanged.getActiveRange(),
              firstRowCurrent: SheetChanged.getActiveRange().getRow(),
              lastRowCurrent: SheetChanged.getActiveRange().getLastRow(),
              lastColCurrent: SheetChanged.getLastColumn()
          };
      });
  
      if (!labelData) return; // Si la tabla no es rastreable
  
      /*Obtain Current values*/
      const rangeCurrent = labelData.rangeCurrent;
      const firstRowCurrent = labelData.firstRowCurrent;
      const lastRowCurrent = labelData.lastRowCurrent;
      const lastColCurrent = labelData.lastColCurrent;
      
      // Cargar solo CON-Control para verificar verbose
      const controlData = retryOperation(() => {
          const ss = SpreadsheetApp.openById(dubAppControlID);
          const conControl = ss.getSheetByName("CON-Control");
          return conControl.getRange('A2:M2').getValues();
      });
  
      const verboseFlag = controlData[0][11];
  
      // Array para almacenar todos los buffers
      const buffers = [];

      for (var i = firstRowCurrent; i <= lastRowCurrent; i++) {
        /*If change is in label 1 = labels, continue*/
        if (i == 1) {continue;} 
  
        //Obtain values
        var SheetChangedData = SheetChanged.getRange(i, 1, 1, lastColCurrent); 
        var SheetChangedValue = SheetChangedData.getValues();
        var logKey = SheetChangedValue[0][keyCurrent]+"";
  
        //Discard if deleted in Appsheet
        if (logKey == "") {continue;} 
  
        var logUser = SheetChangedValue[0][lastUserColCurrent];
        
        if(verboseFlag === true) {
            console.log("Table: "+sheet2track+" / LogKey: "+logKey+"/ User: "+logUser+" / "+Utilities.formatDate(new Date(), timezone, timestamp_format));
        }

        // Crear buffer y añadirlo al array
        const buffer = [
            sheet2track,
            "'" + logKey.toString(),
            Utilities.formatDate(new Date(), timezone, timestamp_format),
            install,
            actionType,
            logUser,
            "01 Pending",
            ""
        ];

        if(verboseFlag === true) {
            console.log("Buffer logKey value: " + buffer[1] + " (tipo: " + typeof buffer[1] + ")");
        }

        buffers.push(buffer);
      }

      // Si hay datos para grabar, hacerlo de una sola vez
      if (buffers.length > 0) {
          retryOperation(() => {
              const ss = SpreadsheetApp.openById(dubAppControlID);
              const conTask = ss.getSheetByName("CON-TaskCurrent");
              conTask.getRange(conTask.getLastRow() + 1, 1, buffers.length, buffers[0].length).setValues(buffers);
          });
      }
  
    } catch (error) {
        // Los errores siempre se deben mostrar, independientemente del verboseFlag
        console.error("ERROR / Table: " + sheet2track + 
                     " / Time: " + Utilities.formatDate(new Date(), timezone, timestamp_format) +
                     " / Error: " + error.toString());
        throw error; // Re-throw para que el error sea visible en los logs
    }
}

// Función auxiliar para reintentos
function retryOperation(operation, maxRetries = 3, delay = 1000) {
    let attempts = 0;
    while (attempts < maxRetries) {
        try {
            return operation();
        } catch (error) {
            attempts++;
            if (attempts === maxRetries) {
                console.error("Error después de " + maxRetries + " intentos: " + error.toString());
                throw error; // Considera agregar más información al error
            }
            Utilities.sleep(delay);
        }
    }
}

// Función para obtener los datos de etiquetas
function getLabelData() {
    const cache = PropertiesService.getScriptProperties();
    let labelData = cache.getProperty('labelDataCache');
    
    if (!labelData) {
        // Si no está en caché, leer de la hoja
        const SheetLabel = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DWO-LogLabels');
        const lastRow = SheetLabel.getLastRow();
        const SheetLabelData = SheetLabel.getRange(2, 1, lastRow - 1, 102).getValues();
        
        // Guardar en caché por 6 horas (21600 segundos)
        cache.setProperty('labelDataCache', JSON.stringify(SheetLabelData));
        return SheetLabelData;
    }
    
    return JSON.parse(labelData);
}

// Función para limpiar el caché (ejecutar cuando se actualicen las etiquetas)
function clearLabelCache() {
    PropertiesService.getScriptProperties().deleteProperty('labelDataCache');
}