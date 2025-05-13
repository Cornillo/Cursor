/**
 * Árbol de llamadas de funciones:
 * 
 * checkColFormats()
 * ├── Función principal que inicia el proceso de validación para todas las bases de datos
 * ├── Verifica que no sea fin de semana (sábado o domingo)
 * ├── Obtiene IDs de las bases desde databaseID.getID()
 * │   └── Retorna objeto con IDs de todas las bases de datos
 * ├── Ejecuta validación para activeID, totalID, noTrackID y logsID
 * └── Consolida y muestra todos los logs al final
 *     
 * validarTiposColumnas(id, dbType)
 * ├── Procesa todas las hojas de una base de datos específica
 * ├── Parámetros:
 * │   ├── id: ID del spreadsheet a procesar
 * │   └── dbType: Identificador de la base (activeID/totalID/noTrackID/logsID)
 * ├── Excluye hojas:
 * │   ├── 'Index'
 * │   ├── 'DWO-LogLabels'
 * │   └── Hojas que terminan en 'Log'
 * ├── Detecta tipos de datos en cada columna
 * │   ├── string: formato @
 * │   ├── number: formato #,##0 o #,##0.00
 * │   ├── date: formato dd/mm/yyyy o dd/mm/yyyy hh:mm:ss
 * │   └── duration: formato [hh]:mm:ss
 * └── Retorna resultado y logs de la operación
 * 
 * @databaseID.js/getID()
 * └── Retorna objeto con todos los IDs de las bases de datos del sistema
 *     ├── activeID: Base de datos activa
 *     ├── totalID: Base de datos total
 *     └── noTrackID: Base de datos sin seguimiento
 */

// Variable global de debug
const DEBUG = true;

function checkColFormats() {
  // Verificar si es fin de semana (excepto en modo debug)
  if (!DEBUG) {
    var today = new Date();
    var day = today.getDay();
    if (day === 0 || day === 6) {
      Logger.log("Esta función no se ejecuta los fines de semana");
      return;
    }
  }

  // Obtener IDs de las bases de datos
  const allIDs = databaseID.getID();
  const databases = {
    'activeID': allIDs.activeID,
    'totalID': allIDs.totalID,
    'noTrackID': allIDs.noTrackID,
    'logsID': allIDs.logsID
  };
  
  let allLogs = [];
  
  for (let dbName in databases) {
    let result = validarTiposColumnas(databases[dbName], dbName);
    allLogs = allLogs.concat(result.logs);
  }
  
  // Imprimir todos los logs al final
  allLogs.forEach(log => Logger.log(log));
}

function validarTiposColumnas(id, dbType) {
  let logs = [];
  logs.push(`Iniciando validación para base de datos: ${dbType}`);
  
  try {
    var spreadsheet = SpreadsheetApp.openById(id);
    var sheets = spreadsheet.getSheets();
    var hojasExcluidas = ['Index', 'DWO-LogLabels'];
    var resultados = [];

    // Procesar cada hoja
    for (var s = 0; s < sheets.length; s++) {
      var sheet = sheets[s];
      var sheetName = sheet.getName();
      
      // Excluir hojas específicas y las que terminan en Log
      if (hojasExcluidas.includes(sheetName) || 
          sheetName.endsWith('Log') || 
          sheetName.includes('MultiLog')) {
        logs.push(`Saltando hoja excluida: ${sheetName}`);
        continue;
      }

      logs.push(`Procesando hoja: ${sheetName}`);
      var data = sheet.getDataRange().getValues();
      
      if (data.length < 2) {
        resultados.push({
          hoja: sheetName,
          valido: true,
          mensaje: "La hoja tiene menos de 2 filas"
        });
        continue;
      }

      var tiposReferencia = [];
      var formatosReferencia = [];
      
      for (var col = 0; col < data[1].length; col++) {
        var valor = data[1][col];
        
        if (valor === "" || valor === null) {
          for (var row = 2; row < data.length; row++) {
            valor = data[row][col];
            if (valor !== "" && valor !== null) {
              break;
            }
          }
        }
        
        if (valor !== "" && valor !== null) {
          if (valor instanceof Date) {
            if (valor.getTime() < 24 * 60 * 60 * 1000) {
              tiposReferencia[col] = 'duration';
              formatosReferencia[col] = '[hh]:mm:ss';
            } else {
              tiposReferencia[col] = 'date';
              var hours = valor.getHours();
              var minutes = valor.getMinutes();
              var seconds = valor.getSeconds();
              if (hours === 0 && minutes === 0 && seconds === 0) {
                formatosReferencia[col] = 'dd/mm/yyyy';
              } else {
                formatosReferencia[col] = 'dd/mm/yyyy hh:mm:ss';
              }
            }
          }
          else if (typeof valor === 'number') {
            if (valor < 1 && valor > 0) {
              tiposReferencia[col] = 'duration';
              formatosReferencia[col] = '[hh]:mm:ss';
            } else if (Number.isInteger(valor)) {
              tiposReferencia[col] = 'number';
              formatosReferencia[col] = '#,##0';
            } else {
              tiposReferencia[col] = 'number';
              formatosReferencia[col] = '#,##0.00';
            }
          }
          else {
            tiposReferencia[col] = 'string';
            formatosReferencia[col] = '@';
          }
        } else {
          tiposReferencia[col] = 'string';
          formatosReferencia[col] = '@';
        }
      }
      
      logs.push(`${dbType} - ${sheetName} - Tipos encontrados: ${JSON.stringify(tiposReferencia)}`);
      logs.push(`${dbType} - ${sheetName} - Formatos encontrados: ${JSON.stringify(formatosReferencia)}`);

      // Aplicar tipos y formatos a columnas completas
      for (var col = 0; col < tiposReferencia.length; col++) {
        var rango = sheet.getRange(2, col + 1, data.length - 1, 1);
        rango.setNumberFormat(formatosReferencia[col]);
      }

      resultados.push({
        hoja: sheetName,
        valido: true,
        mensaje: "Tipos y formatos aplicados exitosamente"
      });
    }

    logs.push(`Proceso completado para ${dbType}`);
    return {
      valido: true,
      mensaje: "Proceso completado",
      resultados: resultados,
      logs: logs
    };

  } catch(e) {
    logs.push(`Error en ${dbType}: ${e.message}`);
    return {
      valido: false,
      mensaje: `Error al procesar ${dbType}: ${e.message}`,
      resultados: resultados,
      logs: logs
    };
  }
}
