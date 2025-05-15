/**
 * Script para realizar una revisión y comparación de los loops calculados
 * para un archivo de script reprocesado, basándose en la lógica actual.
 * Compara los loops calculados con los registrados previamente en DWO_CharacterProduction
 * y graba las diferencias en columnas específicas (Y o Z) según la versión del archivo.
 * Diseñado para uso único.
 */

// --- Variables globales reutilizadas y configuraciones ---
var ssActive;
var sheetID; // ID de la hoja de cálculo principal
var folderId; // ID de la carpeta de Google Drive "Uploaded/"
var TIMEZONE; // Configuración de timezone
var TIMESTAMP_FORMAT; // Configuración de formato de timestamp
var isNow; // Timestamp de la ejecución
var auxSheet; // Usado por OpenSheet/OpenSht
var auxValues; // Usado por OpenSheet/OpenSht
var auxNDX; // Usado por OpenSheet/OpenSht
var auxNDX2; // Usado por OpenSheet/OpenSht
var auxFilteredValues; // Usado por OpenSheet/OpenSht
var auxRow; // Usado por OpenSheet/OpenSht
var containerSheet; // Usado por LazyLoad
var containerValues; // Usado por LazyLoad
var sheetCache = { initialized: false, sheets: new Map() }; // Usado por LazyLoad
var inhibited = ["MAIN TITLE", "GRAPHICS INSERTS", "PRINCIPAL PHOTOGRAPHY", "NONE", "GRAFICA", "BURNEDIN SUBS", "EPISODIC TITLE"]; // Lista de tipos ignorados
var rateSecondsWord = 1.6; // Usado en EstimatedTimecode - Copiado de ScriptBreakdown.js
var timecodeOutRevisionFlag = false; // Usado en EstimatedTimecode - Copiado de ScriptBreakdown.js
var fasttrack = false; // Usado en ExtractDialogLine - Copiado de ScriptBreakdown.js


// --- Funciones Helper copiadas directamente desde DubAppScriptBreakdown.js ---

function initializeGlobals() {
    // Obtener IDs de spreadsheet, folder, timezone, etc.
    // Asume que databaseID es un objeto global accesible con el método getID()
    try {
        if (typeof databaseID === 'undefined' || typeof databaseID.getID !== 'function') {
             console.error("databaseID object or getID function not found.");
             // Lanza un error para detener la ejecución si no se pueden obtener los IDs
             throw new Error("Missing configuration: databaseID not available.");
        }
        const allIDs = databaseID.getID();
        sheetID = allIDs["activeID"];
        folderId = allIDs["uploaded"]; // ID de la carpeta de subidas
        // controlID = allIDs["controlID"]; // No se usa processControlChanges
        TIMEZONE = allIDs.timezone;
        TIMESTAMP_FORMAT = allIDs.timestamp_format;

        ssActive = SpreadsheetApp.openById(sheetID);
        console.log(`Globals initialized. Sheet ID: ${sheetID}, Upload Folder ID: ${folderId}`);
    } catch (e) {
         console.error("Failed to initialize globals: " + e.message);
         throw e; // Re-lanza el error para detener el script si la inicialización falla
    }
}

function OpenSheet(sheetNameAux, ndxCol, key, ndxCol2, ss) {
    // Copiada de DubAppScriptBreakdown.js
    LazyLoad("DubAppActive01", sheetNameAux); // "DubAppActive01" es el nombre de la librería/contenedor
    auxSheet = containerSheet;
    var lastRow = auxSheet.getLastRow();
    if (lastRow === 1) { auxRow = -1; auxValues = []; auxNDX = []; auxNDX2 = []; auxFilteredValues = []; return; } // Asegurarse de inicializar auxFilteredValues también
    auxValues = containerValues;
    if (ndxCol2 != 0) {
        auxNDX2 = auxValues.map(function (r) { return r[ndxCol2 - 1].toString(); });
    }
    if (ndxCol > 0) {
        auxNDX = auxValues.map(function (r) { return r[ndxCol - 1].toString(); });
        if (key != "") {
            if (ndxCol2 != 0) {
                var auxcase = auxNDX2.indexOf(key);
            } else {
                var auxcase = auxNDX.indexOf(key);
            }
            auxFilteredValues = [];
            while (auxcase !== -1) {
                auxRow = auxcase;
                auxFilteredValues.push(auxValues[auxcase]);

                if (ndxCol2 != 0) {
                    auxcase = auxNDX2.indexOf(key, auxcase + 1);
                } else {
                    auxcase = auxNDX.indexOf(key, auxcase + 1);
                }
            }
             if(auxFilteredValues.length === 0) auxRow = -1; // Si no se encontraron resultados, auxRow debe ser -1
        }
    }
     if (!auxFilteredValues && key === "") { // Si no se filtró por key, auxFilteredValues es auxValues completo
        auxFilteredValues = auxValues;
     }
     console.log(`OpenSheet: ${sheetNameAux}, ndxCol:${ndxCol}, key:"${key}", ndxCol2:${ndxCol2}. Results: ${auxFilteredValues ? auxFilteredValues.length : 0}`);
}

function OpenSht(sheetNameAux, ndxColValues, keyCol, keyValue, ndxColFiltered, ss) {
    // Copiada de DubAppScriptBreakdown.js
    LazyLoad("DubAppActive01", sheetNameAux); // "DubAppActive01" es el nombre de la librería/contenedor
    auxSheet = containerSheet;
    var lastRow = auxSheet.getLastRow();

    if (lastRow === 1) {
        auxRow = -1; auxValues = []; auxNDX = []; auxNDX2 = []; auxFilteredValues = [];
        console.log(`OpenSht: ${sheetNameAux} is empty.`);
        return;
    }
    auxValues = containerValues;

    if (ndxColValues > 0) {
        auxNDX = auxValues.map(r => r[ndxColValues - 1].toString());
    } else {
        auxNDX = []; // Inicializar si no se usa
    }

    if (keyCol !== 0 && keyValue != "") {
        var filteredValues = auxValues
            .map((row, index) => ({ row: row, index: index }))
            .filter(item => item.row[keyCol - 1].toString() === keyValue);

        auxFilteredValues = filteredValues.map(item => {
            var newRow = item.row.slice();
            newRow.push(item.index); // Agregar el número de fila (0-based index in auxValues)
            return newRow;
        });

        if (ndxColFiltered > 0) {
            auxNDX2 = auxFilteredValues.map(r => r[ndxColFiltered - 1].toString());
        } else {
            auxNDX2 = []; // Inicializar si no se usa
        }
         if(auxFilteredValues.length === 0) auxRow = -1; // Si no se encontraron resultados, auxRow debe ser -1

    } else {
        // Si no se filtra, auxFilteredValues es auxValues completo (sin añadir índice de fila)
        auxFilteredValues = auxValues;
        auxNDX2 = []; // Inicializar si no se usa
         auxRow = auxValues.length > 0 ? 0 : -1; // Simular auxRow para coherencia, aunque no se usará para un solo resultado
    }
    console.log(`OpenSht: ${sheetNameAux}, keyCol:${keyCol}, keyValue:"${keyValue}", ndxColFiltered:${ndxColFiltered}. Results: ${auxFilteredValues ? auxFilteredValues.length : 0}`);
}

function LazyLoad(containerName, sheetNameAux) {
    // Copiada de DubAppScriptBreakdown.js
    // Adaptada para no usar containerSheetByName ya que se basa en ID principal ssActive
    if (!ssActive) {
         console.error("ssActive is not initialized in LazyLoad.");
         throw new Error("Spreadsheet not initialized.");
    }

    if (!sheetCache.initialized) {
        // Cargar todas las hojas necesarias en el cache la primera vez
        // Puedes agregar otros nombres de hojas si son de uso frecuente aquí
        const sheetsToCache = ["DWO_Files", "DWO_Character", "DWO_CharacterProduction", "DWO-Production", "DWO", "App-User", "CON-TaskCurrent"];
        sheetsToCache.forEach(name => {
            try {
                const sheet = ssActive.getSheetByName(name);
                if (sheet) {
                     // Excluir la primera fila (encabezado) al cargar valores si es grande
                     const range = sheet.getDataRange();
                     const values = range.getValues();
                     // Opcional: quitar encabezado si la hoja no está vacía
                     const dataValues = values.length > 1 ? values.slice(1) : [];

                    sheetCache.sheets.set(name, { sheet: sheet, values: dataValues });
                    console.log(`Cached sheet: ${name} with ${dataValues.length} rows.`);
                } else {
                    console.warn(`Sheet not found in ssActive: ${name}`);
                     // Cachear con sheet nula y valores vacíos si no existe
                     sheetCache.sheets.set(name, { sheet: null, values: [] });
                }
            } catch(e) {
                 console.error(`Error caching sheet ${name}: ${e.message}`);
                 // Cachear con sheet nula y valores vacíos en caso de error
                 sheetCache.sheets.set(name, { sheet: null, values: [] });
            }
        });
        sheetCache.initialized = true;
        console.log("Sheet cache initialized.");
    }

    if (sheetCache.sheets.has(sheetNameAux)) {
        const cached = sheetCache.sheets.get(sheetNameAux);
        containerSheet = cached.sheet;
        containerValues = cached.values;
        //console.log(`LazyLoad: Loaded ${sheetNameAux} from cache.`);
    } else {
        // Esto no debería pasar si cacheamos todas las hojas necesarias al inicio
        // Pero como fallback, cargar solo esta hoja si no está en cache
         console.warn(`Sheet ${sheetNameAux} not in initial cache. Loading directly.`);
        try {
             const sheet = ssActive.getSheetByName(sheetNameAux);
             if (sheet) {
                 const range = sheet.getDataRange();
                 const values = range.getValues();
                 const dataValues = values.length > 1 ? values.slice(1) : []; // Quitar encabezado
                 containerSheet = sheet;
                 containerValues = dataValues;
                  // Opcional: añadir a cache para futuras llamadas (aunque este script es one-time)
                 sheetCache.sheets.set(sheetNameAux, { sheet: sheet, values: dataValues });
                 console.log(`LazyLoad: Loaded ${sheetNameAux} directly.`);
             } else {
                 console.error(`Sheet not found directly: ${sheetNameAux}`);
                 containerSheet = null;
                 containerValues = [];
             }
        } catch(e) {
             console.error(`Error loading sheet directly ${sheetNameAux}: ${e.message}`);
             containerSheet = null;
             containerValues = [];
        }
    }
}


function String2Seconds(cadenaDuracion) {
    // Copiada de DubAppScriptBreakdown.js
    if (!cadenaDuracion || typeof cadenaDuracion !== 'string') return 0;
    var partes = cadenaDuracion.split(":");
     if (partes.length !== 3) {
        //console.warn(`Invalid timecode format: ${cadenaDuracion}`);
        return 0; // Manejar formatos inválidos
    }
    var horas = parseInt(partes[0], 10) || 0;
    var minutos = parseInt(partes[1], 10) || 0;
    var segundos = parseInt(partes[2], 10) || 0;
    return (horas * 3600 + minutos * 60 + segundos);
}

function Time2Seconds(fecha) {
    // Copiada de DubAppScriptBreakdown.js
    if (!fecha) return 0; // Cambiado de fecha a 0 para consistencia en cálculos

    if (typeof fecha === 'string') {
        return String2Seconds(fecha);
    }

    if (fecha instanceof Date) {
        // Obtener componentes de tiempo en la zona horaria del script si es posible,
        // o simplemente usar UTC si la zona horaria no está configurada correctamente.
        // Dado que estamos en GAS, asume que getHours/getMinutes/getSeconds
        // operan en la zona horaria del proyecto si está configurada.
        const hora = fecha.getHours();
        const minutos = fecha.getMinutes();
        const segundos = fecha.getSeconds();
        return hora * 3600 + minutos * 60 + segundos;
    }

    // Si es un número, asumir que ya son segundos
    if (typeof fecha === 'number') {
        return fecha;
    }


    console.error('Tipo de fecha no válido en Time2Seconds:', typeof fecha, fecha);
    return 0;
}

function Time2String(fecha) {
     // Copiada de DubAppScriptBreakdown.js
     if (!fecha && fecha !== 0) { return ""; } // Permitir 0 segundos

    // Si ya es un string con formato hh:mm:ss
    if (typeof fecha === 'string' && /^\d{2}:\d{2}:\d{2}$/.test(fecha)) {
        return fecha;
    }

    let totalSegundos = 0;
     if (typeof fecha === 'string') {
         totalSegundos = String2Seconds(fecha); // Intentar convertir string hh:mm:ss a segundos
     } else if (fecha instanceof Date) {
         totalSegundos = Time2Seconds(fecha); // Convertir Date a segundos
     } else if (typeof fecha === 'number') {
         totalSegundos = fecha; // Asumir que es un número de segundos
     } else {
         console.error('Formato de Timecode no válido en Time2String:', fecha);
         return "";
     }

    const horas = Math.floor(totalSegundos / 3600);
    const minutos = Math.floor((totalSegundos % 3600) / 60);
    const segundos = Math.floor(totalSegundos % 60); // Usar floor para segundos

    // Formateamos los componentes en una cadena hh:mm:ss
    // Asegurar que siempre tenga 2 dígitos con padStart
    const cadenaFormateada = `${String(horas).padStart(2, '0')}:${String(minutos).padStart(2, '0')}:${String(segundos).padStart(2, '0')}`;

    return cadenaFormateada;
}


function CharacterName(cadena) {
    // Copiada de DubAppScriptBreakdown.js
    if (!cadena) return "";
    var mapa = {
        'á': 'a', 'é': 'e', 'í': 'i', 'ó': 'o', 'ú': 'u',
        'Á': 'A', 'É': 'E', 'Í': 'I', 'Ó': 'O', 'Ú': 'Ú', // Corregido Ú
        'ü': 'u', 'Ü': 'U', 'ç': 'c', 'Ç': 'C', 'Â': 'A',
        'Ê': 'E', 'Ô': 'O', 'Ã': 'A', 'Õ': 'O', 'À': 'A', 'È': 'E', 'Ì': 'I', 'Ò': 'O', 'Ù': 'U' // Añadidos más acentos/tildes
    };

    var textoNormalizado = cadena.split('').map(function (char) {
        return mapa[char] || char;
    }).join('');

    // Quitar todo lo que está entre paréntesis y corchetes (no anidados)
    var sinParentesisCorchetes = textoNormalizado.replace(/\s*[\(\[].*?[\)\]]\s*/g, '');


    // Quitar los espacios en blanco al principio y al final
    var sinEspacios = sinParentesisCorchetes.trim();

    // Cualquier caracter extra - Lista ampliada de caracteres a eliminar
    var sincaracteres = sinEspacios.replace(/[\(\)\[\]\?\\\=\:;,\.'¨\*\+!@#\$%&\{\}_`"´'’]/g, ''); // Añadidos !@#$, etc.

    // Reemplazar múltiples espacios con un solo espacio
    sincaracteres = sincaracteres.replace(/\s+/g, ' ');

    // Pasar todo a mayúsculas
    var enMayusculas = sincaracteres.toUpperCase();

    return enMayusculas;
}

function LoopNumber(seconds2check) {
    // Copiada de DubAppScriptBreakdown.js
     if (typeof seconds2check !== 'number' || isNaN(seconds2check) || seconds2check < 0) {
         //console.warn(`Invalid input for LoopNumber: ${seconds2check}`);
         return 1; // O manejar como error, pero 1 es un fallback razonable
     }
    var auxLoop = Math.floor((seconds2check) / 15) + 1;
    return auxLoop;
}

function EstimatedTimecode() {
    // Copiada de DubAppScriptBreakdown.js
    // Modifica el array global `script`.
    if (!script || script.length === 0) {
        console.warn("No script data available for EstimatedTimecode.");
        return;
    }
    console.log(`Calculating EstimatedTimecode for ${script.length} lines.`);

    var fila; var auxLoopTCIn; var auxLoopTCOut; var checkAux;
    // Asegurarse de que rateSecondsWord está definido
    if (typeof rateSecondsWord === 'undefined') rateSecondsWord = 1.6;


    for (var k = 0; k < script.length; k++) {
        fila = script[k];
         // item structure: [timecodeInSeg, timecodeOutSeg, sourceName, cleanedDialogue, isTimecodeOutEmpty, statusMark, timecodeInStr, timecodeOutStrOriginal, estimatedTimecodeStr, checkMark, originalDialogue]
        if (fila[5] === "Dismissed") {
             script[k][8] = ""; // EstimatedTimecode (string)
             script[k][9] = ""; // Check status
            continue;
         }

        // Calcular segundos estimados a agregar (basado en diálogo limpio)
        var palabras = fila[3].split(" ").filter(word => word.length > 0).length; // Cuenta palabras no vacías
        var segundosAAgregar = parseInt(palabras / rateSecondsWord); // Usa rateSecondsWord global
        var totalSegundosEstimado = fila[0] + segundosAAgregar; // Calcula el total de segundos estimados desde TimecodeIn original

        var bypass = false; // Inicializar bypass
        checkAux = ""; // Inicializar checkAux

        // REGLA 1: Si tiene Timecode out original cargado (fila[1] no está null/empty), no es check y es bypass.
         // Asegurarse de que fila[1] es el TimecodeOut original en segundos
        if (fila[1] !== null && fila[1] !== "" && typeof fila[1] === 'number') {
            bypass = true;
            checkAux = "";
             totalSegundosEstimado = fila[1]; // Si hay OUT original, la estimación "efectiva" es el OUT original para cálculo de loop TCOut para Check
        } else {
            // Lógica original para bypass solo si NO tiene timecode out original
            if (k + 1 < script.length) {
                for (var m = k + 1; m < script.length; m++) {
                    if (script[m][5] === "Dismissed") { continue; }
                     // Usar el timecode IN original de la siguiente línea (script[m][0]) para la comparación
                    if (script[m][0] <= totalSegundosEstimado) {
                        totalSegundosEstimado = script[m][0]; bypass = true;
                    } else {
                        // Comparar con diferencia de menos de 5 segundos
                        if (script[m][0] - totalSegundosEstimado < 5) {
                            bypass = true;
                        }
                    }
                    break; // Solo comparar con la primera siguiente línea no Dismissed
                }
            }
        }

        auxLoopTCIn = LoopNumber(script[k][0]); // Loop del timecode IN original
        auxLoopTCOut = LoopNumber(totalSegundosEstimado); // Loop del timecode estimado (o OUT original si existía)

        // NUEVA REGLA 2: Si la diferencia entre auxLoopTCOut y auxLoopTCIn es mayor a 1, es "Check"
        // y no hay bypass (a menos que ya se haya establecido por la REGLA 1 usando OUT original).
         // La REGLA 1 ya manejó el caso de OUT original que fuerza bypass="" y check="".
         // Entonces, esta regla solo se aplica si no había OUT original (bypass=false inicialmente).
        if (Math.abs(auxLoopTCOut - auxLoopTCIn) > 1) {
            checkAux = "Check";
             // timecodeOutRevisionFlag global se usa para marcar que hay al menos un check por timecode
             // Para esta revisión, no necesitamos modificar la flag global persistente.
             // if (!fila[1]) { timecodeOutRevisionFlag = true; } // Esto no es relevante para este script de revisión
        } else {
            // Lógica original de check si no se aplicó la NUEVA REGLA 2
            // y si bypass no es true por la REGLA 1 o la lógica de proximidad.
            if (checkAux !== "Check" && !bypass) { // Solo aplicar si no hay bypass y no se marcó por la REGLA 2
                if (auxLoopTCIn != auxLoopTCOut) {
                    // Buscar la próxima intervención del MISMO personaje
                    var proximaIntervencionEncontrada = false;
                    for (var m = k + 1; m < script.length; m++) {
                        if (script[m][5] === "Dismissed") { continue; }
                        if (script[k][2] === script[m][2]) { // Mismo personaje (usando sourceName normalizado)
                            proximaIntervencionEncontrada = true;
                            // Comparar el loop de inicio de la próxima intervención (LoopNumber(script[m][0]))
                            // con el loop estimado de salida de la línea actual (auxLoopTCOut)
                            if (LoopNumber(script[m][0]) != auxLoopTCOut) {
                                checkAux = "Check";
                                // if (!fila[1]) { timecodeOutRevisionFlag = true; } // No relevante aquí
                            }
                            break; // Encontrada la próxima intervención del mismo personaje, salir
                        }
                    }
                     // Si el personaje no vuelve a hablar en el resto del script
                    if (!proximaIntervencionEncontrada && (m === script.length || k === script.length - 1)) {
                         if (auxLoopTCIn !== auxLoopTCOut) { // Solo marcar check si el loop estimado cambió y no hay más intervenciones
                             checkAux = "Check";
                            // if (!fila[1]) { timecodeOutRevisionFlag = true; } // No relevante aquí
                         }
                    }
                }
            }
        }

        // Calcular el timecode estimado en formato string hh:mm:ss
         // OJO: totalSegundosEstimado fue modificado por la lógica de bypass.
         // El timecode estimado de SALIDA REAL, si no hay OUT original,
         // es fila[0] + segundosAAgregar.
        let finalEstimatedTimecodeSec = fila[0] + segundosAAgregar;
         // Si había OUT original (REGLA 1), el timecode estimado de salida es el OUT original.
         if (fila[1] !== null && fila[1] !== "" && typeof fila[1] === 'number') {
             finalEstimatedTimecodeSec = fila[1];
         }


        var nuevoTimecodeString = Time2String(finalEstimatedTimecodeSec); // Usa Time2String global

        script[k][8] = nuevoTimecodeString; // Almacena el timecode estimado (o OUT original) en formato string
        script[k][9] = checkAux; // Almacena el estado "Check" o ""
    }
     console.log('EstimatedTimecode calculation complete.');
}

// Adaptación de AddLoops2Character para contar loops únicos localmente
function AddLoops2CharacterRevised(characterAux, loopsAux, addedAux, characterLoopCounts) {
    // characterLoopCounts es un objeto donde las claves son nombres de personaje
    // y los valores son Sets de strings de loop numbers.
    if (loopsAux === "") { return; }
    var sources = characterAux.split("/");
    var loops = loopsAux.split(",").map(loop => loop.trim()).filter(loop => loop !== ""); // Split by comma and trim/filter empty

    if (loops.length === 0) { return; } // Si no hay loops válidos, salir

    for (var k = 0; k < sources.length; k++) {
        var individualSource = sources[k].trim();
        if (individualSource != "" && inhibited.indexOf(individualSource.toUpperCase()) === -1) {
             // Usar el nombre normalizado
            const characterName = typeof CharacterName !== 'undefined' ? CharacterName(individualSource) : individualSource.toUpperCase();

             // Inicializar si no existe
            if (!characterLoopCounts[characterName]) {
                characterLoopCounts[characterName] = new Set(); // Usar Set para contar loops únicos
            }

            // Añadir cada loop al Set
            loops.forEach(loop => {
                 if (loop !== "") {
                    characterLoopCounts[characterName].add(loop); // Los loops ya están como strings por LoopNumber
                 }
            });
        }
    }
}

// Adaptación de ExtractCharacter para obtener solo los nombres únicos presentes en el script
// No necesita lógica de DWO_Character o DWO_FilesCharacter aquí.
function ExtractCharacterRevised(scriptArray) {
    let uniqueCharacterNames = new Set(); // Usar Set para almacenar nombres únicos

    if (!scriptArray || scriptArray.length === 0) {
         console.warn("No script data available for ExtractCharacterRevised.");
         return [];
    }

    for (var j = 0; j < scriptArray.length; j++) {
        var sourceValue = scriptArray[j][2]; // Source Name (normalizado)
        if (sourceValue) { // Asegurarse de que no es null/undefined/empty
             var sources = sourceValue.split("/");

            for (var k = 0; k < sources.length; k++) {
                var individualSource = sources[k].trim();
                if (individualSource != "" && inhibited.indexOf(individualSource.toUpperCase()) === -1) {
                     // El nombre ya debería estar normalizado si ExtractDialogLine llama a CharacterName
                     // Pero para seguridad, lo normalizamos aquí de nuevo si es necesario
                     const characterName = typeof CharacterName !== 'undefined' ? CharacterName(individualSource) : individualSource.toUpperCase();
                     if (characterName) { // Asegurarse de que el nombre normalizado no es vacío
                        uniqueCharacterNames.add(characterName);
                     }
                }
            }
        }
    }

    // Devuelve un array de nombres únicos
    return Array.from(uniqueCharacterNames);
}


function Rand(n) {
    // Copiada de DubAppScriptBreakdown.js - No estrictamente necesaria para este script
    var min = Math.pow(10, n - 1);
    var max = Math.pow(10, n) - 1;
    var numeroAleatorio = Math.floor(Math.random() * (max - min + 1)) + min;
    return numeroAleatorio;
}

// processControlChanges no se usa según el requisito 5.

// databaseID.getID() se usa en initializeGlobals() y se asume accesible globalmente.

// Funciones para obtener información de contacto tampoco se usan en este script de revisión.


// --- Función Principal de Revisión ---

/**
 * Revisa y compara los loops liquidados para un archivo de script específico.
 * Procesa el script, calcula los loops por personaje y compara con los datos existentes
 * en DWO_CharacterProduction, grabando las diferencias.
 *
 * @param {string} file_ID El ID del archivo en DWO_Files a reprocesar (Columna A).
 */
function RevisionScript(file_ID) {
    console.log(`Iniciando RevisionScript para File_ID: ${file_ID}`);
    try {
        // Asegurar la inicialización de variables globales como sheetID, folderId, etc.
         initializeGlobals(); // Asegura que ssActive y IDs están configurados

        isNow = Utilities.formatDate(new Date(), TIMEZONE, TIMESTAMP_FORMAT);

        // 1. Cargar datos del archivo DWO_Files
        // Buscar por File_ID en Columna A (indice 0)
        OpenSheet("DWO_Files", 1, file_ID, 0, ssActive);
        if (!auxFilteredValues || auxFilteredValues.length === 0) {
            console.error(`File_ID ${file_ID} no encontrado en DWO_Files.`);
            return; // Salir si el archivo no se encuentra
        }
        const fileDetails = auxFilteredValues[0];
        const fileNameInFolder = fileDetails[6]; // Columna G (6) - Nombre del archivo en "Uploaded/"
        const productionID = fileDetails[2]; // Columna C (2) - Production ID
        const projectID = fileDetails[15]; // Columna P (15) - Project ID
        const userID = fileDetails[12]; // Columna M (12) - User ID
        const auxVersion = fileDetails[14]; // Columna O (14) - Versión del archivo

        const fileNameClean = fileNameInFolder.replace("Uploaded/", "").trim();
        if (!fileNameClean) {
             console.error(`Nombre de archivo inválido en DWO_Files para File_ID ${file_ID}. Columna G: "${fileNameInFolder}"`);
             return;
        }

        console.log(`File Details - File_ID: ${file_ID}, File Name (Clean): ${fileNameClean}, ProductionID: ${productionID}, ProjectID: ${projectID}, UserID: ${userID}, Version: ${auxVersion}`);


        // 2. Obtener el archivo de script de Drive por nombre en la carpeta especificada
        let scriptFileBlob;
        let mimeType;
        let fileDriveIdActual = null; // Para almacenar el ID de Drive real si se encuentra
        try {
             // Asegurarse de que folderId está definido por initializeGlobals
             if (!folderId) {
                 throw new Error("Upload folder ID is not defined.");
             }
             const folder = DriveApp.getFolderById(folderId);
             const files = folder.getFilesByName(fileNameClean);

             if (!files.hasNext()) {
                 throw new Error(`Archivo "${fileNameClean}" no encontrado en la carpeta de Drive (ID: ${folderId}).`);
             }
             const file = files.next();
             scriptFileBlob = file.getBlob();
             mimeType = file.getMimeType();
             fileDriveIdActual = file.getId(); // Obtener el ID real del archivo en Drive
             console.log(`Archivo de script encontrado en Drive. Nombre: ${file.getName()}, ID: ${fileDriveIdActual}, MimeType: ${mimeType}`);

        } catch (e) {
             console.error(`Error al obtener archivo de Drive: ${e.message}`);
             return; // Salir en caso de error de archivo
        }


        // 3. Reprocesar el script para extraer diálogos y timecodes
        // Resetear el array global 'script'
        script = [];
        let extractResult = ""; // Variable para capturar el resultado de ExtractDialogLine si devuelve string

        // Adaptamos la lógica de ExtractDialogLine para usar el blob y devolver el resultado string si hay error
         try {
             // Crear un archivo temporal en Drive a partir del blob para usar DocumentApp/SpreadsheetApp
             let fileTemp;
             if (mimeType === "application/vnd.openxmlformats-officedocument.wordprocessingml.document" || mimeType === "application/msword") {
                  fileTemp = Drive.Files.insert({ title: "temp_doc_for_revision_" + userID, mimeType: MimeType.GOOGLE_DOCS }, scriptFileBlob, { convert: true });
                  // Procesar Google Doc temporal
                  const doc = DocumentApp.openById(fileTemp.id);
                  const body = doc.getBody();
                  const tables = body.getTables();
                  if (tables.length === 0) {
                      extractResult = "No se encontró ninguna tabla en el documento de Word.";
                      console.error(extractResult);
                  } else {
                      // Procesar la primera tabla
                      const dataTable = tables[0];
                      for (let i4 = 0; i4 < dataTable.getNumRows(); i4++) {
                         const rowAux = dataTable.getRow(i4);
                         const fila = [];
                         for (let j = 0; j < rowAux.getNumCells(); j++) {
                             fila.push(rowAux.getCell(j).getText());
                         }
                         if (fila[0] === null || fila[0].trim() === "") {
                             continue;
                         }
                         const timecodeInStr = fila[0].replace(/\n/g, '').trim().substring(0, 8);
                         const timecodeIn = String2Seconds(timecodeInStr); // Usa String2Seconds
                         const source = CharacterName(fila[1]); // Usa CharacterName
                         let dialogue = fila[2];
                         let loopAuxStatus = "";

                         if (inhibited.includes(source.toUpperCase()) || source === "") {
                             loopAuxStatus = "Dismissed";
                         } else {
                             if (auxVersion === "Final version: Script_upload_lite" && dialogue.includes("[+]")) {
                                 loopAuxStatus = "Added";
                             }
                         }

                         dialogue = dialogue.replace(/\n|\[[^\]]*\]/g, '').trim();

                          if (!dialogue && loopAuxStatus !== "Dismissed") {
                              loopAuxStatus = "Dismissed";
                          }

                         script.push([
                             timecodeIn,       // 0: TimecodeIn (segundos)
                             null,             // 1: TimecodeOut original (segundos)
                             source,           // 2: Source (normalizado)
                             dialogue,         // 3: Dialogue (limpio)
                             true,             // 4: Flag: isTimecodeOutEmpty (siempre true para docx/doc)
                             loopAuxStatus,    // 5: Status (Dismissed, Added, "")
                             timecodeInStr,    // 6: TimecodeIn (string original)
                             "",               // 7: TimecodeOut (string original)
                             "",               // 8: EstimatedTimecode (string) - se llenará después
                             "",               // 9: Check status - se llenará después
                             fila[2]           // 10: Original Dialogue
                         ]);
                      }
                  }
                  // Borrar archivo temporal
                  Drive.Files.remove(fileTemp.id);
                  console.log('Extraction from DOC/DOCX complete.');

             } else if (mimeType === 'application/vnd.ms-excel' ||
                        mimeType === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
                        mimeType === 'application/vnd.oasis.opendocument.spreadsheet') {

                  fileTemp = Drive.Files.insert({ title: "temp_sheet_for_revision_" + userID, mimeType: MimeType.GOOGLE_SHEETS }, scriptFileBlob, { convert: true });
                  // Procesar Google Sheet temporal
                  const sheetTemp = SpreadsheetApp.openById(fileTemp.id).getSheets()[0]; // Asume primera hoja
                  const dataValues = sheetTemp.getDataRange().getValues();

                  if (dataValues.length === 0 || dataValues[0][0] !== "IN-TIMECODE" || dataValues[0][1] !== "OUT-TIMECODE" || dataValues[0][2] !== "SOURCE" || dataValues[0][3] !== "TRANSCRIPTION") {
                       extractResult = "El archivo de hoja de cálculo no contiene los encabezados esperados.";
                       console.error(extractResult);
                  } else {
                      // Recorrer datos y cargar el array 'script'
                      for (var i4 = 1; i4 < dataValues.length; i4++) { // Empieza en 1 para saltar la fila de encabezado
                         const fila = dataValues[i4];
                         if (fila[0] === null || fila[0].toString().trim() === "") { // Check columna IN-TIMECODE
                             continue;
                         }
                         const timecodeInStr = fila[0].toString().trim().substring(0, 8);
                         const timecodeIn = String2Seconds(timecodeInStr); // Usa String2Seconds

                         const source = CharacterName(fila[2]); // Usa CharacterName
                         let dialogue = fila[3]; // Columna D (3)
                         let loopAuxStatus = "";

                         if (inhibited.includes(source.toUpperCase()) || source === "") {
                             loopAuxStatus = "Dismissed";
                         } else {
                             if (auxVersion === "Final version: Script_upload_lite" && dialogue.includes("[+]")) {
                                 loopAuxStatus = "Added";
                             }
                         }

                         dialogue = dialogue.replace(/\n|\[[^\]]*\]/g, '').trim();

                         if (!dialogue && loopAuxStatus !== "Dismissed") {
                              loopAuxStatus = "Dismissed";
                         }

                         let timecodeOutStr = fila[1] ? fila[1].toString().trim().substring(0, 8) : ""; // Columna B (1)
                         let timecodeOut = timecodeOutStr ? String2Seconds(timecodeOutStr) : null;
                         let flagStatus = (timecodeOut === null || timecodeOutStr === ""); // true if TimecodeOut is empty

                         script.push([
                             timecodeIn,       // 0: TimecodeIn (segundos)
                             timecodeOut,      // 1: TimecodeOut original (segundos)
                             source,           // 2: Source (normalizado)
                             dialogue,         // 3: Dialogue (limpio)
                             flagStatus,       // 4: Flag: isTimecodeOutEmpty
                             loopAuxStatus,    // 5: Status (Dismissed, Added, "")
                             timecodeInStr,    // 6: TimecodeIn (string original)
                             timecodeOutStr,   // 7: TimecodeOut (string original)
                             "",               // 8: EstimatedTimecode (string) - se llenará después
                             "",               // 9: Check status - se llenará después
                             fila[3]           // 10: Original Dialogue (Columna D en este formato)
                         ]);
                      }
                  }
                  // Borrar archivo temporal
                  Drive.Files.remove(fileTemp.id);
                  console.log('Extraction from Sheet complete.');

             } else {
                  extractResult = `Formato de archivo no reconocido para revisión (${mimeType}).`;
                  console.error(extractResult);
             }

         } catch (e) {
              console.error(`Error durante la extracción/procesamiento del archivo temporal: ${e.message}`);
              // Si un error ocurre ANTES de borrar fileTemp, intentar borrarlo.
              if (fileTemp && fileTemp.id) {
                  try {
                      Drive.Files.remove(fileTemp.id);
                      console.log(`Temporary file ${fileTemp.id} removed after error.`);
                  } catch(e2) {
                      console.error(`Error removing temporary file ${fileTemp.id}: ${e2.message}`);
                  }
              }
              extractResult = `Error crítico durante la extracción del script: ${e.message}`; // Capturar el error principal
         }

        if (extractResult !== "" || script.length === 0) {
             console.error(`No se pudo procesar el script. Motivo: ${extractResult}`);
             return; // Salir si hubo error en la extracción o no se extrajeron líneas
        }
        console.log(`Extracción de diálogo completada. Líneas procesadas: ${script.length}`);


        // 4. Calcular timecodes estimados y Check status
        EstimatedTimecode(); // Modifica el array global 'script'
        console.log('Timecodes estimados y Check status calculados.');


        // 5. Calcular Loops por Personaje (con lógica de ScriptBreakdown)
        const characterLoopCounts = {}; // { "NOMBRE_PERSONAJE": Set { "loopNum1", "loopNum2", ... }, ...}

        script.forEach(function (item) {
            // item structure: [timecodeInSeg, timecodeOutSeg, sourceName, cleanedDialogue, isTimecodeOutEmpty, statusMark, timecodeInStr, timecodeOutStrOriginal, estimatedTimecodeStr, checkMark, originalDialogue]
            if (item[5] !== "Dismissed") { // Solo procesar líneas no descartadas
                let loopAuxString = "";
                const auxIn = LoopNumber(item[0]); // Loop de inicio (basado en timecode IN original)

                let auxOutLoop;
                if (item[4]) { // Si Timecode OUT original está vacío (isTimecodeOutEmpty is true)
                     // Usar el timecode estimado de salida (item[8] es el string, convertir a segundos)
                     const estimatedTimecodeSec = item[8] ? Time2Seconds(item[8]) : item[0];
                     auxOutLoop = LoopNumber(estimatedTimecodeSec); // Loop de fin estimado

                } else { // Si Timecode OUT original existe
                    // Usar el timecode OUT original (item[1] ya está en segundos)
                    const originalTimecodeOutSec = item[1];
                    auxOutLoop = LoopNumber(originalTimecodeOutSec); // Loop de fin original
                }

                // Generar el string de loops cubiertos (similar a ScriptBreakdown para AddLoops2Character)
                 // Esto asegura que contamos los loops EXACTAMENTE como el script original los registraba
                if (auxIn === auxOutLoop) {
                     loopAuxString = auxIn.toString();
                 } else {
                      for (let i2 = auxIn; i2 <= auxOutLoop; i2++) {
                         loopAuxString = loopAuxString === "" ? i2.toString() : loopAuxString + ", " + i2.toString();
                      }
                 }


                // Acumular loops únicos para el personaje/s usando la versión Revised
                 AddLoops2CharacterRevised(item[2], loopAuxString, item[5], characterLoopCounts); // item[2] es el nombre del Source normalizado
            }
        });
         console.log('Cálculo de loops por personaje (conteo único) completado.');
         // characterLoopCounts ahora contiene { "CHARACTER_NAME": Set { "loop1", "loop2", ... }, ...}


        // 6. Cargar datos de DWO_Character para mapear Nombres a IDs
         // Buscar por Project ID en Columna B (indice 1)
        OpenSheet("DWO_Character", 1, projectID, 2, ssActive); // ndxCol=1 (ProjectID), key=projectID, ndxCol2=2 (CharacterName)
        const dwoCharacterData = auxFilteredValues; // Datos filtrados de DWO_Character
        const dwoCharacterMapByName = new Map(); // Mapa para buscar Character_ID por nombre normalizado
        if(dwoCharacterData && dwoCharacterData.length > 0){
             dwoCharacterData.forEach(row => {
                 const charID = row[0]; // Columna A (0) Character_ID
                 const charName = CharacterName(row[2]); // Columna C (2) Character Name (normalizado)
                 if (charName) {
                    dwoCharacterMapByName.set(charName, charID);
                 }
             });
             console.log(`Cached ${dwoCharacterMapByName.size} Character IDs from DWO_Character for ProjectID ${projectID}.`);
        } else {
             console.log(`No records found in DWO_Character for ProjectID ${projectID}.`);
        }


        // 7. Cargar datos existentes de DWO_CharacterProduction para esta producción
        // Usamos OpenSht para obtener el índice de fila original y buscar por Production ID en Columna C (indice 2)
        OpenSht("DWO_CharacterProduction", 0, 3, productionID, 0, ssActive); // keyCol=3 (ProductionID), keyValue=productionID
        const dwoCharacterProductionSheet = auxSheet; // La hoja completa para escritura
        const dwoCharacterProductionData = auxFilteredValues; // Datos filtrados CON índice original (última columna)
        // Construir un mapa para buscar rápidamente por CharacterID dentro de esta producción
        const charProdMapByCharId = new Map();
        if(dwoCharacterProductionData && dwoCharacterProductionData.length > 0){
             dwoCharacterProductionData.forEach(row => {
                 const characterID = row[1]; // Columna B (1) Character_ID
                 const rowIndex = row[row.length - 1]; // Última columna es el índice original (0-based in auxValues)
                 charProdMapByCharId.set(characterID, {
                     data: row.slice(0, row.length - 1), // Guardar datos sin el índice extra
                     rowIndex: rowIndex
                 });
             });
             console.log(`Cached ${charProdMapByCharId.size} CharacterProduction entries for ProductionID ${productionID}.`);
        } else {
             console.log(`No existing records found in DWO_CharacterProduction for ProductionID: ${productionID}.`);
        }


        // 8. Comparar loops calculados y grabar diferencias en DWO_CharacterProduction
        // Recorrer los personajes que encontramos en el script y calculamos loops para
        console.log("Comparing calculated loops with existing data and preparing updates.");
        let rowsToUpdate = []; // Array para acumular los datos de las filas a actualizar


        for (const charName in characterLoopCounts) {
            if (characterLoopCounts.hasOwnProperty(charName)) {
                const calculatedLoopCount = characterLoopCounts[charName].size; // Número de loops únicos calculados

                 // 8a. Buscar el Character_ID usando el nombre normalizado y ProjectID (usando el mapa de DWO_Character)
                const charID = dwoCharacterMapByName.get(charName);

                if (charID) {
                     // 8b. Buscar la entrada correspondiente en DWO_CharacterProduction usando el Character_ID y ProductionID
                     const foundCharProdEntry = charProdMapByCharId.get(charID);

                    if (foundCharProdEntry) {
                        const existingData = foundCharProdEntry.data; // Datos existentes de la fila en CP
                        const rowIndex = foundCharProdEntry.rowIndex; // Índice original 0-based en auxValues

                        const plannedLoops = existingData[5] ? parseInt(existingData[5]) : 0; // Columna F (5) Planned loops
                        const finalLoops = existingData[22] ? parseInt(existingData[22]) : 0; // Columna W (22) Final loops
                        let loopsAdded = existingData[24] || ""; // Columna Y (24) Loops Added
                        let finalLoopsAdded = existingData[25] || ""; // Columna Z (25) Loops Added (Revision)
                        let comments = existingData[8] || ""; // Columna I (8) Comments

                        let difference = 0;
                        let targetColumnIndex = -1; // Índice de columna a actualizar (24 para Y, 25 para Z)

                        if (auxVersion === "Final version: Script_upload_lite") {
                             // Comparar con Final loops (Columna W)
                            difference = calculatedLoopCount - finalLoops;
                            targetColumnIndex = 25; // Columna Z (índice 25)
                            finalLoopsAdded = difference.toString(); // Grabar la diferencia en Z
                            // loopsAdded = ""; // No es necesario limpiar la otra si solo se graba en una

                            comments += ` // Revisión Loops Finales: Calculados ${calculatedLoopCount}, Existentes ${finalLoops}, Diferencia ${difference} (${isNow})`;

                        } else { // Versión preliminar
                             // Comparar con Planned loops (Columna F)
                            difference = calculatedLoopCount - plannedLoops;
                            targetColumnIndex = 24; // Columna Y (índice 24)
                            loopsAdded = difference.toString(); // Grabar la diferencia en Y
                            // finalLoopsAdded = ""; // No es necesario limpiar la otra

                            comments += ` // Revisión Loops Planificados: Calculados ${calculatedLoopCount}, Existentes ${plannedLoops}, Diferencia ${difference} (${isNow})`;
                        }

                        // Preparar la fila actualizada - copiar los datos existentes y modificar las columnas relevantes
                        const updatedRow = existingData.slice(); // Copiar todos los datos existentes
                        updatedRow[24] = loopsAdded; // Columna Y
                        updatedRow[25] = finalLoopsAdded; // Columna Z
                        updatedRow[8] = comments; // Columna I

                        // Añadir la información de actualización a la lista
                        // range: dwoCharacterProductionSheet.getRange(rowIndex + 2, 1, 1, updatedRow.length),
                        // +2 porque rowIndex es 0-based index in auxValues (que empieza en la fila 2 de la hoja) + convertir a 1-based
                         if (dwoCharacterProductionSheet) {
                             rowsToUpdate.push({
                                range: dwoCharacterProductionSheet.getRange(rowIndex + 2, 1, 1, updatedRow.length),
                                values: [updatedRow]
                            });
                            console.log(`Prepared update for Character "${charName}" (ID: ${charID}). Calculated Loops: ${calculatedLoopCount}. Difference ${difference} recorded in Col ${targetColumnIndex === 24 ? 'Y' : 'Z'}.`);
                         } else {
                            console.error("DWO_CharacterProduction sheet object is null. Cannot prepare range.");
                         }


                    } else {
                        console.log(`Personaje "${charName}" (ID: ${charID}) encontrado en DWO_Character pero no registrado en DWO_CharacterProduction para ProductionID ${productionID}. No se grabará la diferencia.`);
                    }
                } else {
                    console.log(`Personaje "${charName}" encontrado en script pero no encontrado en DWO_Character para ProjectID ${projectID}. No se grabará la diferencia.`);
                }
            }
        }

        // 9. Escribir las actualizaciones acumuladas en DWO_CharacterProduction
        if (rowsToUpdate.length > 0) {
            console.log(`Iniciando escritura de ${rowsToUpdate.length} filas en DWO_CharacterProduction.`);
            let writeSuccess = false;
            let retryCount = 0;
            const maxRetries = 3;

            while (!writeSuccess && retryCount < maxRetries) {
                 try {
                     // Escribir cada fila individualmente ya que no son rangos contiguos
                     rowsToUpdate.forEach(update => {
                         update.range.setValues(update.values);
                     });
                     writeSuccess = true;
                     console.log('Escritura en DWO_CharacterProduction exitosa.');

                 } catch (error) {
                     retryCount++;
                     console.error(`Error en escritura de DWO_CharacterProduction (Intento ${retryCount}): ${error.message}`);
                     if (retryCount === maxRetries) {
                          // Lanza el error final si fallan todos los reintentos
                         throw new Error(`Fallo en escritura de DWO_CharacterProduction después de ${maxRetries} intentos: ${error.message}`);
                     }
                     Utilities.sleep(1000 * retryCount); // Espera exponencial antes de reintentar
                 }
            }

        } else {
            console.log('No hay filas para actualizar en DWO_CharacterProduction.');
        }

        console.log(`Revisión para File_ID: ${file_ID} completada.`);

    } catch (e) {
        console.error(`Error general durante la revisión para File_ID ${file_ID}: ${e.message}`);
        // Aquí podrías añadir lógica adicional para manejar el error,
        // como enviar una notificación, registrar en una hoja de log, etc.
        throw e; // Re-lanzar el error para que se muestre en los logs de Apps Script
    }
}

// Nota: La función databaseID.getID() debe estar definida y ser accesible
// en este proyecto de Google Apps Script para que initializeGlobals funcione.
// Si databaseID es una librería, asegúrate de haberla añadido al proyecto.

/**
 * Función de entrada para ejecutar el script de revisión de loops.
 * Especifica el File_ID del archivo DWO_Files a procesar.
 */
function runRevision() {
  // <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
  // ** MODIFICA ESTE VALOR **
  // Reemplaza "ID_DEL_ARCHIVO_AQUI" con el File_ID real del archivo en DWO_Files
  // que deseas que el script revise.
  const file_ID_a_revisar = "ID_DEL_ARCHIVO_AQUI";
  // <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

  if (file_ID_a_revisar === "ID_DEL_ARCHIVO_AQUI") {
    console.error("Por favor, reemplaza 'ID_DEL_ARCHIVO_AQUI' con el File_ID real antes de ejecutar.");
    SpreadsheetApp.getUi().alert("Error", "Por favor, reemplaza 'ID_DEL_ARCHIVO_AQUI' con el File_ID real en la función runRevision() antes de ejecutar.", SpreadsheetApp.getUi().ButtonSet.OK);
    return; // Detener la ejecución si el ID no ha sido cambiado
  }

  console.log(`Llamando a RevisionScript con File_ID: ${file_ID_a_revisar}`);
  RevisionScript(file_ID_a_revisar);
}
