/**
 * Identifica valores únicos de columna C donde hay repeticiones en columna B
 * excluyendo registros que tienen valores en columna N
 * @return {Array} Lista de valores de columna C con repeticiones en columna B
 */
function encontrarColumnaCConRepeticionesB() {
    try {
      console.log("Buscando valores de columna C con repeticiones en columna B (excluyendo registros con columna N)...");
      
      // Obtener ID de base de datos activa desde databaseID.js
      const allIDs = databaseID.getID();
      const activeSpreadsheet = SpreadsheetApp.openById(allIDs.activeID);
      
      // Buscar la hoja por su nombre
      const sheet = activeSpreadsheet.getSheetByName('DWO_CharacterProduction');
      if (!sheet) {
        console.error("No se encontró la hoja DWO_CharacterProduction");
        return [];
      }
      
      // Obtener datos de las columnas B, C y N
      const datos = sheet.getDataRange().getValues();
      if (datos.length <= 1) {
        console.log("La hoja está vacía o solo tiene encabezados");
        return [];
      }
      
      // Mapeo de valores de columna: B=1, C=2, N=13
      const colB = 1;
      const colC = 2;
      const colN = 13;
      
      // Agrupar por valor de columna C para analizar repeticiones
      const repeticionesPorC = {};
      
      // Procesar todas las filas (excepto encabezados)
      for (let i = 1; i < datos.length; i++) {
        const fila = datos[i];
        const valorB = fila[colB]; // Columna B
        const valorC = fila[colC]; // Columna C
        const valorN = fila[colN]; // Columna N
        
        // Ignorar filas con valores vacíos en B o C, o con valores en N
        if (!valorB || !valorC || valorN) continue;
        
        // Normalizar valores
        const valorBStr = String(valorB).trim();
        const valorCStr = String(valorC).trim();
        
        // Inicializar registro para este valor C si no existe
        if (!repeticionesPorC[valorCStr]) {
          repeticionesPorC[valorCStr] = {
            valoresB: new Set(),
            filas: 0
          };
        }
        
        // Agregar el valor B y contar esta fila
        repeticionesPorC[valorCStr].valoresB.add(valorBStr);
        repeticionesPorC[valorCStr].filas++;
      }
      
      // Filtrar solo los valores C donde hay repeticiones de valores B
      const valoresCConRepeticiones = [];
      
      for (const valorC in repeticionesPorC) {
        const info = repeticionesPorC[valorC];
        
        // Si el número de valores B únicos es menor que el total de filas, hay repeticiones
        if (info.valoresB.size < info.filas) {
          valoresCConRepeticiones.push(valorC);
        }
      }
      
      // Ordenar resultados
      valoresCConRepeticiones.sort();
      
      // Mostrar resultados por log
      console.log(`Se encontraron ${valoresCConRepeticiones.length} valores de columna C con repeticiones en columna B (excluyendo registros con columna N).`);
      
      if (valoresCConRepeticiones.length > 0) {
        console.log("Lista de valores de columna C con repeticiones en columna B:");
        console.log(valoresCConRepeticiones.join(", "));
      } else {
        console.log("No se encontraron repeticiones que cumplan con los criterios.");
      }
      
      return valoresCConRepeticiones;
    } catch (error) {
      console.error("Error: " + error.message);
      return [];
    }
  }