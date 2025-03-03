/**
 * DubApp Consistency Checks
 * 
 * Descripción:
 * Script para detectar y manejar registros huérfanos en la base de datos DubApp.
 * Un registro huérfano es aquel que tiene una referencia a un registro padre que no existe.
 * 
 * Árbol de llamadas:
 * detectarHuerfanos()
 * ├── LazyLoad(ssAux, sheetNameAux)         // Carga eficiente de hojas de cálculo
 * ├── columnaAIndice(columna)               // Convierte letra de columna a índice
 * ├── obtenerNivelPadre(nivel, hierarchy)   // Encuentra el nivel padre en la jerarquía
 * └── SendEmail.AppSendEmailX()             // Envía notificación por correo si hay huérfanos
 * 
 * Jerarquía de tablas:
 * - Nivel 1 (DWO)
 *   ├── Nivel 11 (DWO-Production)
 *   │   ├── Nivel 111 (DWO-Event)
 *   │   │   ├── Nivel 1111 (DWO_Files)
 *   │   │   │   ├── Nivel 11111 (DWO_FilesLines)
 *   │   │   │   └── Nivel 11112 (DWO_FilesCharacter)
 *   │   │   ├── Nivel 1112 (DWO-MixAndEdit)
 *   │   │   └── Nivel 1113 (DWO-Observation)
 *   │   ├── Nivel 113 (DWO_Song)
 *   │   │   └── Nivel 1131 (DWO_SongDetail)
 *   │   └── Nivel 114 (DWO-SynopsisProduction)
 *   ├── Nivel 12 (DWO_Character)
 *   │   └── Nivel 112 (DWO_CharacterProduction)
 *   └── Nivel 13 (DWO-SynopsisProject)
 */

// Variables globales
var eliminarHuerfanos = false;
var allIDs = databaseID.getID(); // Obtener la configuración
var auxToday = Utilities.formatDate(new Date(), allIDs.timezone, "dd/MM/yy");
var auxInstalation = "Arg";
var auxHTML = "";
var auxSendMail = true;

// Variables necesarias para LazyLoad
let containerSheet, containerValues, containerNDX, containerNDX2;
let ssActive, ssTotal, ssLogs, ssNoTrack;
let labelValues, userValues, labelNDX, userNDX;
let verboseFlag;

function detectarHuerfanos() {
  try {
    // Inicializar LazyLoad con la hoja principal
    LazyLoad('DubAppActive01', 'DWO');
    
    if (!containerSheet || !containerValues) {
      Logger.log('Error: No se pudo cargar la hoja principal');
      return;
    }

    // Validar que la librería databaseID esté disponible
    if (typeof databaseID === 'undefined') {
      Logger.log('Error: La librería databaseID no está disponible');
      return;
    }

    // Contadores totales
    let totalHuerfanosDetectados = 0;
    let totalHuerfanosLimpiados = 0;

    // Inicializar tabla HTML solo si se va a enviar mail
    let huerfanosEncontrados = false;
    if (auxSendMail) {
      auxHTML = "<table border='1' style='border-collapse: collapse; width: 100%;'>";
      auxHTML += "<tr style='background-color: #f2f2f2;'><th>Hoja</th><th>Clave</th><th>Clave padre no encontrada</th></tr>";
    }

    const hierarchy = {
      'Nivel 1': { sheetName: 'DWO', keyColumn: 'B', dependentSheets: ['Nivel 11', 'Nivel 12', 'Nivel 13'] },
      'Nivel 11': { sheetName: 'DWO-Production', keyColumn: 'A', parentColumn: 'B', dependentSheets: ['Nivel 111', 'Nivel 113', 'Nivel 114'] },
      'Nivel 111': { sheetName: 'DWO-Event', keyColumn: 'A', parentColumn: 'B', dependentSheets: ['Nivel 1111', 'Nivel 1112', 'Nivel 1113'] },
      'Nivel 12': { sheetName: 'DWO_Character', keyColumn: 'A', parentColumn: 'B', dependentSheets: ['Nivel 112'] },
      'Nivel 112': { sheetName: 'DWO_CharacterProduction', keyColumn: 'A', parentColumn: 'B', dependentSheets: [] },
      'Nivel 1111': { sheetName: 'DWO_Files', keyColumn: 'A', parentColumn: 'D', dependentSheets: ['Nivel 11111', 'Nivel 11112'] },
      'Nivel 1112': { sheetName: 'DWO-MixAndEdit', keyColumn: 'A', parentColumn: 'B', dependentSheets: [] },
      'Nivel 1113': { sheetName: 'DWO-Observation', keyColumn: 'A', parentColumn: 'B', dependentSheets: [] },
      'Nivel 11111': { sheetName: 'DWO_FilesLines', keyColumn: 'A', parentColumn: 'Q', dependentSheets: [] },
      'Nivel 11112': { sheetName: 'DWO_FilesCharacter', keyColumn: 'A', parentColumn: 'B', dependentSheets: [] },
      'Nivel 113': { sheetName: 'DWO_Song', keyColumn: 'A', parentColumn: 'B', dependentSheets: ['Nivel 1131'] },
      'Nivel 1131': { sheetName: 'DWO_SongDetail', keyColumn: 'A', parentColumn: 'B', dependentSheets: [] },
      'Nivel 13': { sheetName: 'DWO-SynopsisProject', keyColumn: 'A', parentColumn: 'A', dependentSheets: [] },
      'Nivel 114': { sheetName: 'DWO-SynopsisProduction', keyColumn: 'A', parentColumn: 'A', dependentSheets: [] }
    };

    // Crear un mapa de llaves para cada nivel
    const keysMap = {};
    for (const nivel in hierarchy) {
      const sheetInfo = hierarchy[nivel];
      try {
        LazyLoad('DubAppActive01', sheetInfo.sheetName);
        if (!containerValues) {
          Logger.log(`Hoja no encontrada: ${sheetInfo.sheetName}`);
          continue;
        }
        
        const keyColIndex = columnaAIndice(sheetInfo.keyColumn);
        keysMap[nivel] = new Set();
        
        // Empezar desde 0 ya que LazyLoad ya omite la fila de encabezado
        for (let i = 0; i < containerValues.length; i++) {
          const key = String(containerValues[i][keyColIndex - 1] || '');
          if (key.trim()) {
            keysMap[nivel].add(key.trim());
          }
        }
      } catch (e) {
        Logger.log(`Error al procesar la hoja ${sheetInfo.sheetName}: ${e.toString()}`);
        continue;
      }
    }

    // Detectar huerfanos
    for (const nivel in hierarchy) {
      const sheetInfo = hierarchy[nivel];
      if (!sheetInfo.parentColumn) continue;
      
      const nivelPadre = obtenerNivelPadre(nivel, hierarchy);
      if (!nivelPadre) continue;

      const clavesPadre = keysMap[nivelPadre];
      if (!clavesPadre) continue;

      try {
        LazyLoad('DubAppActive01', sheetInfo.sheetName);
        if (!containerValues) {
          Logger.log(`Hoja no encontrada: ${sheetInfo.sheetName}`);
          continue;
        }

        const keyColIndex = columnaAIndice(sheetInfo.keyColumn);
        const parentColIndex = columnaAIndice(sheetInfo.parentColumn);
        
        // Almacenar filas a eliminar
        const filasAEliminar = [];
        let huerfanosEnHoja = 0;
        
        // Detectar huérfanos
        for (let i = 0; i < containerValues.length; i++) {
          const clavePadre = String(containerValues[i][parentColIndex - 1] || '');
          const claveActual = String(containerValues[i][keyColIndex - 1] || '');
          
          // Asegurarse que las claves no estén vacías
          if (!clavePadre.trim() || !claveActual.trim()) {
            continue;
          }

          // Comparar las claves exactamente como strings, sin conversión numérica
          if (!clavesPadre.has(clavePadre.trim())) {
            huerfanosEnHoja++;
            totalHuerfanosDetectados++;
            huerfanosEncontrados = true;
            
            if (auxSendMail) {
              auxHTML += `<tr><td>${sheetInfo.sheetName}</td><td>${claveActual.trim()}</td><td>${clavePadre.trim()}</td></tr>`;
            }
            
            Logger.log(`Huerfano detectado en hoja: ${sheetInfo.sheetName}, Clave: ${claveActual.trim()}, Clave padre no encontrada: ${clavePadre.trim()}`);
            
            if (eliminarHuerfanos) {
              filasAEliminar.push(i + 2); // +2 porque i empieza en 0 y hay que considerar el encabezado
              keysMap[nivel].delete(claveActual.trim());
            }
          }
        }

        // Limpiar filas huérfanas
        if (eliminarHuerfanos && filasAEliminar.length > 0) {
          filasAEliminar.sort((a, b) => b - a);
          for (const fila of filasAEliminar) {
            containerSheet.getRange(fila, 1, 1, containerSheet.getLastColumn()).clearContent();
            Logger.log(`Fila ${fila} limpiada en hoja: ${sheetInfo.sheetName}`);
            totalHuerfanosLimpiados++;
          }
          Logger.log(`Total de filas limpiadas en ${sheetInfo.sheetName}: ${filasAEliminar.length}`);
        }

        if (huerfanosEnHoja > 0) {
          Logger.log(`Total de huérfanos en ${sheetInfo.sheetName}: ${huerfanosEnHoja}`);
        }
      } catch (e) {
        Logger.log(`Error al procesar la hoja ${sheetInfo.sheetName}: ${e.toString()}`);
        continue;
      }
    }

    // Cerrar tabla HTML
    if (auxSendMail) {
      auxHTML += "</table>";
    }

    // Resumen final
    Logger.log('=== RESUMEN DE HUÉRFANOS ===');
    Logger.log(`Total de huérfanos detectados: ${totalHuerfanosDetectados}`);
    if (eliminarHuerfanos) {
      Logger.log(`Total de huérfanos limpiados: ${totalHuerfanosLimpiados}`);
    }
    Logger.log('=========================');

    // Enviar email solo si auxSendMail es true y se encontraron huérfanos
    if (auxSendMail && huerfanosEncontrados) {
      SendEmail.AppSendEmailX(
        "appsheet@mediaaccesscompany.com",
        "",
        "",
        "1r0YxzUZSkda9PZYbP2V6fXELIQN5u0ipvFz1B0Z5-GI",
        "",
        "DubApp: Detección de inconsistencia",
        "Detalle::" + auxHTML + "||Fecha::" + auxToday + "||Instalacion::" + auxInstalation + "||Tipo::Detección Huérfanos",
        "",
        ""
      );
    }
  } catch (e) {
    Logger.log('Error general en detectarHuerfanos: ' + e.toString());
    return;
  }
}

// Función para convertir columna letra a índice
function columnaAIndice(columna) {
  let indice = 0;
  for (let i = 0; i < columna.length; i++) {
    indice *= 26;
    indice += columna.charCodeAt(i) - 64;
  }
  return indice;
}

// Función para obtener el nivel padre
function obtenerNivelPadre(nivel, hierarchy) {
  for (const nivelPadre in hierarchy) {
    if (hierarchy[nivelPadre].dependentSheets.includes(nivel)) {
      return nivelPadre;
    }
  }
  return null;
}
