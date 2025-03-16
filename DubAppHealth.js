/*
### Estructura General
DubAppHealth.js
├── checkActive()
│   └── checkSheet() [llamada múltiple para diferentes hojas]
└── Funciones Utilitarias
    ├── checkDuplicates()
    ├── checkVersions()
    ├── dateHandle()
    ├── isValidDate() 
    ├── String.prototype.replaceAll()
    └── daysBetweenToday()

### Explicación de Funciones

Funciones Principales:
1. checkActive()
   - Función principal que verifica la consistencia entre diferentes hojas
   - Revisa múltiples hojas de trabajo relacionadas con DWO (Dubbing Work Order)
   - Lanza un error si encuentra casos para reenviar o duplicados

2. checkSheet(sourceSS, sheet2check)
   - Compara una hoja específica entre dos spreadsheets
   - Verifica duplicados y diferencias en timestamps
   - Registra casos que necesitan ser actualizados en la hoja de control

Funciones Utilitarias:
3. dateHandle(d, timezone, timestamp_format)
   - Formatea fechas según zona horaria y formato especificado
   
4. isValidDate(d)
   - Valida si un objeto es una fecha válida

5. String.prototype.replaceAll(search, replacement)
   - Extensión del prototipo String para reemplazar todas las ocurrencias

6. daysBetweenToday(dateParam)
   - Calcula días entre una fecha dada y hoy

Variables Globales Importantes:
- Configuración de timezone y formatos de timestamp
- Arrays y objetos para tracking de diferentes hojas (Active, Total, Phantom, Log)
- Contadores y registros de casos duplicados
- Referencias a hojas de cálculo y sus valores

El sistema está diseñado para mantener la sincronización entre diferentes hojas de 
cálculo en un sistema de gestión de dubbing, con énfasis en el control de versiones 
y la detección de inconsistencias.
*/

//Global declaration
const timezone = "GMT-3";
const timestamp_format = "dd/MM/yyyy HH:mm:ss"; // Timestamp Format.
const timestamp_format2 = "yyyy/MM/dd HH:mm:ss"; // Timestamp alt Format.

const sheetCache = {
  initialized: false,
  sheets: new Map()
};

//Flexible loading
let ssActive, ssTotal, ssLogs, ssNoTrack;
let labelValues, userValues, labelNDX, userNDX;
var sourceSheet = null;
var sourceValues = null;
var sourceNDX = null;
var sourceRow = null;
var witnessSheet = null;  
var witnessValues = null;
var witnessNDX = null;
var witnessRow = null;
var logSheet = null;
var logValues = null;
var logNDX = null;
var containerSheet = null;
var containerValues = null;
var containerNDX = null;
var containerNDX2 = null;
var conTask = null;
var countDWO2Total = 0;

//Previous loading witness
var ssWitness = null;
var sheetNameWitness = null;

//Global ID
const allIDs = databaseID.getID();

//General variables
var resendCases = 0; var duplicateTotal = 0; var duplicateDetail = "";

function checkActive() {
  try {
    //Check consistency
    conTask = SpreadsheetApp.openById(allIDs['controlID']);
    if (!conTask) {
      throw new Error(`No se pudo abrir el documento de control con ID: ${allIDs['controlID']}`);
    }
    
    conTask = conTask.getSheetByName("CON-TaskCurrent");
    if (!conTask) {
      throw new Error("No se encontró la hoja CON-TaskCurrent en el documento de control");
    }

    // Obtener la estructura desde databaseID
    const estructura = databaseID.getStructure();
    const hojas = Object.keys(estructura);

    let errores = [];
    for (const hoja of hojas) {
      try {
        console.log(`Verificando hoja: ${hoja}`);
        checkSheet("DubAppActive01", hoja);
      } catch (error) {
        errores.push(`Error procesando ${hoja}: ${error.message}`);
        console.error(`Error detallado para ${hoja}:`, error);
      }
    }

    if (errores.length > 0) {
      throw new Error(`Errores múltiples:\n${errores.join('\n')}`);
    }

    if (resendCases > 0 || duplicateTotal > 0) {
      let errMsg = "";
      if (resendCases > 0) { errMsg = `DubApp Health check. Total cases: ${resendCases} // `; }
      if (duplicateTotal > 0) { errMsg = errMsg + duplicateDetail; }
      console.log(errMsg);
      throw new Error(errMsg);
    }
  } catch (error) {
    console.error("Error en checkActive:", error);
    throw error;
  }
}

function checkSheet(sourceSS, sheet2check) {
  try {
    // Verificar que la hoja está configurada
    if (!SHEET_CONFIG[sheet2check]) {
      console.error(`La hoja ${sheet2check} no está configurada en SHEET_CONFIG`);
      return false;
    }

    //Obtain confront SS
    witnessSS = "DubAppTotal01";

    console.log(`Intentando cargar hoja ${sheet2check} desde ${sourceSS}...`);
    
    //Load source
    try {
      LazyLoad(sourceSS, sheet2check);
      if (!containerSheet || !containerValues) {
        throw new Error(`LazyLoad no pudo cargar ${sheet2check} desde ${sourceSS}`);
      }
    } catch (error) {
      console.error(`Error cargando source: ${error.message}`);
      return false;
    }
    
    sourceSheet = containerSheet;
    sourceValues = containerValues;
    sourceNDX = containerNDX;
    
    console.log(`Intentando cargar hoja ${sheet2check} desde ${witnessSS}...`);
    
    //Load witness
    try {
      LazyLoad(witnessSS, sheet2check);
      if (!containerSheet || !containerValues) {
        throw new Error(`LazyLoad no pudo cargar ${sheet2check} desde ${witnessSS}`);
      }
    } catch (error) {
      console.error(`Error cargando witness: ${error.message}`);
      return false;
    }
    
    witnessSheet = containerSheet;
    witnessValues = containerValues;
    witnessNDX = containerNDX;

    // Solo procedemos con la comparación si ambas cargas fueron exitosas
    if (!sourceValues || !witnessValues) {
      console.log(`Saltando comparación de ${sheet2check} - datos incompletos`);
      return false;
    }

    console.log(`Procesando ${sheet2check}: ${sourceValues?.length || 0} filas en source, ${witnessValues?.length || 0} filas en witness`);

    //Clear filters
    let aux = sourceSheet.getFilter(); 
    if (aux != null) {
      try {
        aux.remove();
      } catch (error) {
        console.warn(`No se pudo eliminar el filtro en source: ${error.message}`);
      }
    }
    
    aux = witnessSheet.getFilter(); 
    if (aux != null) {
      try {
        aux.remove();
      } catch (error) {
        console.warn(`No se pudo eliminar el filtro en witness: ${error.message}`);
      }
    }
    
    // Obtain label data
    if (!labelNDX || !labelValues) {
      throw new Error('Datos de etiquetas no inicializados correctamente');
    }

    labelRow = labelNDX.indexOf(sheet2check, 0);
    if (labelRow === -1) {
      throw new Error(`No se encontraron etiquetas para ${sheet2check}`);
    }

    taskLabelNames = labelValues[labelRow];
    taskLabelActions = labelValues[labelRow + 1];
    
    if (!taskLabelNames || !taskLabelActions) {
      throw new Error(`Configuración de etiquetas incompleta para ${sheet2check}`);
    }

    taskColKey = taskLabelActions.indexOf("K") - 1;
    taskColUser = taskLabelNames.indexOf("Last user", 0) - 1; 
    taskColChange = taskLabelNames.indexOf("Last change", 0) - 1;

    if (taskColKey === -2 || taskColUser === -2 || taskColChange === -2) {
      throw new Error(`Columnas requeridas no encontradas en ${sheet2check}`);
    }

    //Loop
    for (sourceRow = 0; sourceRow < sourceValues.length; sourceRow++) {
      //Obtain current info
      sourceKeyAux = sourceValues[sourceRow][taskColKey] + "";
      if (sourceKeyAux == "") {
        continue;
      }

      if (sourceNDX.indexOf(sourceKeyAux, sourceRow + 1) != -1) {
        //Duplicate key
        duplicateTotal = duplicateTotal + 1;
        duplicateDetail = duplicateDetail + " " + sourceSS + "/" + sheet2check + ": " + sourceKeyAux;
      }

      //Find key source in witness
      witnessRow = witnessNDX.indexOf(sourceKeyAux);

      //Not present
      if (witnessRow == -1) {
        sourceUser = sourceValues[sourceRow][taskColUser];
        sourceChange = sourceValues[sourceRow][taskColChange];
        try {
          conTask.appendRow([sheet2check, sourceKeyAux, sourceChange, sourceSS, "INSERT_ROW", sourceUser, "01 Pending", "", "checkSheet"]);
          resendCases++;
        } catch (error) {
          console.error(`Error al agregar fila a conTask: ${error.message}`);
        }
        continue;
      } else if (witnessNDX.indexOf(sourceKeyAux, witnessRow + 1) != -1) {
        //Duplicate key
        duplicateTotal = duplicateTotal + 1;
        duplicateDetail = duplicateDetail + " " + witnessSS + "/" + sheet2check + ": " + sourceKeyAux;
      }

      //Which is newer
      try {
        sourceAux = dateHandle(sourceValues[sourceRow][taskColChange], timezone, timestamp_format2);
        witnessAux = dateHandle(witnessValues[witnessRow][taskColChange], timezone, timestamp_format2);
      } catch (error) {
        console.error(`Error procesando fechas para ${sourceKeyAux}: ${error.message}`);
        continue;
      }

      //Check if equal
      if (sourceAux == witnessAux) {
        continue;
      }

      resendCases++;
      if (witnessAux > sourceAux) {
        witnessUser = witnessValues[witnessRow][taskColUser];
        witnessChange = witnessValues[witnessRow][taskColChange];
        try {
          conTask.appendRow([sheet2check, sourceKeyAux, witnessChange, witnessSS, "EDIT", witnessUser, "01 Pending", "", "checkSheet"]);
        } catch (error) {
          console.error(`Error al agregar fila a conTask: ${error.message}`);
        }
      } else {
        sourceUser = sourceValues[sourceRow][taskColUser];
        sourceChange = sourceValues[sourceRow][taskColChange];
        try {
          conTask.appendRow([sheet2check, sourceKeyAux, sourceChange, sourceSS, "EDIT", sourceUser, "01 Pending", "", "checkSheet"]);
        } catch (error) {
          console.error(`Error al agregar fila a conTask: ${error.message}`);
        }
      }
    }

    return true;

  } catch (error) {
    console.error(`Error general en checkSheet para ${sheet2check}: ${error.message}`);
    return false;
  }
}

/*GENERAL UTILITIES*/

function dateHandle(d,timezone, timestamp_format) {
  if ( isValidDate(d) )
  {
    return Utilities.formatDate(d, timezone, timestamp_format);
  }
  else
  {
    return d;
  }
}

function isValidDate(d) {
  if ( Object.prototype.toString.call(d) !== "[object Date]" )
    return false;
  return !isNaN(d.getTime());
}

String.prototype.replaceAll = function(search, replacement) {
  var target = this;
  if(search == null || replacement == null) {
    return target;
  } else {
    return target.replace(new RegExp(search, 'g'), replacement);
  }
};

function daysBetweenToday(dateParam) {
  let auxNow = new Date();
  let auxDays = Math.floor((auxNow.getTime() - dateParam.getTime()) / (1000 * 60 * 60 * 24));
  return auxDays; 
}