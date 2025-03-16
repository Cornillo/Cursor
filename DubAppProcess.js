/*
### DubApp Process - Sistema de Procesamiento de Colas
Este script maneja el procesamiento de tareas en cola para el sistema DubApp.
Gestiona la sincronización de datos entre diferentes hojas de cálculo y
controla el flujo de trabajo para actualizar registros en tiempo real.

### Descripción General
El script procesa tareas pendientes en la hoja CON-TaskCurrent, comparando datos
entre hojas "source" y "witness" para detectar cambios y actualizar registros.
Utiliza un sistema de locks para evitar ejecuciones simultáneas y mantiene
un registro detallado de los cambios realizados.

### Árbol de Llamadas de Funciones
processQueue()
├── DataLoad(ssAux, sheetNameAux, keyAux)
│   ├── inhibitAppend(auxEnvironment, auxTable, auxKey)
│   └── LazyLoad(ssAux, sheetNameAux)
├── ObtainVariation(taskAux)
│   ├── dateHandle(d, timezone, timestamp_format)
│   └── uncodeChannelEventType(chainSource, chainWitness)
│       └── LazyLoad("DubAppActive01", "DWO-ChannelEventType")

### Funciones Principales:
- processQueue(): Función principal que procesa las tareas pendientes en la cola
- DataLoad(ssAux, sheetNameAux, keyAux): Carga datos de las hojas source y witness
- ObtainVariation(taskAux): Detecta variaciones entre registros source y witness

### Funciones de Gestión de Locks:
- clearAllLocks(): Limpia todos los locks que puedan haber quedado activos
- checkLockStatus(): Verifica el estado actual de los locks y propiedades relacionadas

### Funciones Utilitarias:
- dateHandle(d, timezone, timestamp_format): Maneja formatos de fecha
- isValidDate(d): Verifica si una fecha es válida
- daysBetweenToday(dateParam): Calcula días entre una fecha y hoy
- setCache(auxKey, valueToCache): Almacena valores en caché
- getCachedValue(auxKey): Recupera valores de la caché
- inhibitAppend(auxEnvironment, auxTable, auxKey): Controla la adición de registros
- uncodeChannelEventType(chainSource, chainWitness): Decodifica tipos de eventos

### Funciones de Control:
- Llamador(): Inicializa el entorno para ejecución manual

### Guía de Solución de Problemas:
1. Error "The starting column of the range is too small":
   - Este error ocurre cuando se intenta acceder a una columna con índice menor a 1
   - El script incluye validaciones para evitar este problema

2. Proceso bloqueado o locks persistentes:
   - Ejecutar la función clearAllLocks() para liberar todos los locks
   - Usar checkLockStatus() para diagnosticar el estado actual

3. Errores de acceso a datos:
   - Verificar que las hojas existan y tengan los nombres correctos
   - Comprobar que las columnas esperadas estén presentes

4. Problemas de rendimiento:
   - Revisar el tamaño de los datos procesados
   - Considerar procesar en lotes más pequeños

5. Errores en el procesamiento de fechas:
   - Verificar el formato de las fechas en las hojas de origen
   - Comprobar la configuración de zona horaria (timezone)
*/

//Global declaration
const allIDs = databaseID.getID();  // Llamamos a la función getID()
const timezone = allIDs["timezone"];
const timestamp_format = allIDs["timestamp_format"]; // Timestamp Format.
var verboseFlag = null;
var descreetFlag = null;

//Flexible loading
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
var ConControl = null; 
var ConTask = null;

//Previous loading witness
var ssPrevious = null;
var sheetNamePrevious = null;

//Global ID

//Count check
var counter = {
  dwo: 0,
  production: 0,
  synopsisproject: 0, 
  event: 0,
  synopsisproduction: 0,
  mixandedit: 0,
  observation: 0};

var inhibitedAppend=[];

//Global Label
var labelRow = null; 
var taskLabelNames = null;
var taskLabelActions = null; 
var taskColKey = null;
var taskColUser = null; 
var taskColChange = null; 
var taskColLog = null;

const sheetCache = {
  initialized: false,
  sheets: new Map()};

// Definir relacionesHojas al inicio del archivo
const relacionesHojas = {
  'DWO': {
    columna: 'B',
    dwoCol: 'B'
  },
  'DWO-Production': {
    columna: 'A',
    dwoCol: 'B'
  },
  'DWO-Event': {
    columna: 'A',
    dwoCol: 'BX'
  },
  'DWO_Character': {
    columna: 'A',
    dwoCol: 'B'
  },
  'DWO_CharacterProduction': {
    columna: 'A',
    dwoCol: 'X'
  },
  'DWO_Files': {
    columna: 'A',
    dwoCol: 'P'
  },
  'DWO-MixAndEdit': {
    columna: 'A',
    dwoCol: 'Q'
  },
  'DWO-Observation': {
    columna: 'A',
    dwoCol: 'Z'
  },
  'DWO_FilesLines': {
    columna: 'A',
    dwoCol: 'Q'
  },
  'DWO_FilesCharacter': {
    columna: 'A',
    dwoCol: 'N'
  },
  'DWO_Song': {
    columna: 'A',
    dwoCol: 'P'
  },
  'DWO_SongDetail': {
    columna: 'A',
    dwoCol: 'G'
  },
  'DWO-SynopsisProject': {
    columna: 'A',
    dwoCol: 'A'
  },
  'DWO-SynopsisProduction': {
    columna: 'A',
    dwoCol: 'S'
  }
};

//
function processQueue() {
 //OnTime trigger for tables DubAppControl01/CON-TaskCurrent
 //Version dev 18/3/24
 //Database architecture doc https://docs.google.com/document/d/1vtoE9m8mkgFHjs-C27YjS55kobLsNK01086UE4s0h38/edit

 // Añadir contador de reintentos para el acceso inicial
 const MAX_RETRIES = 3;
 let retryCount = 0;
 let ss;

 // Definir las constantes de tiempo al inicio de la función
 const START_TIME = Date.now();
 const MAX_EXECUTION_TIME = 8 * 60 * 1000; // 8 minutos

 while (retryCount < MAX_RETRIES) {
   try {
     //Lock script
     var lock = LockService.getScriptLock();
     if (!lock.tryLock(1)) {  // Intenta obtener el lock inmediatamente
       console.log('Another thread of processQueue already running (lock service)');
       if (lock) lock.releaseLock();
       return;    
     }

     //Check if service enabled 
     ss = SpreadsheetApp.openById(allIDs['controlID']);
     break; // Si tiene éxito, salir del bucle
     
   } catch (e) {
     retryCount++;
     if (retryCount === MAX_RETRIES) {
       console.error('Failed to access spreadsheet after ' + MAX_RETRIES + ' attempts: ' + e.toString());
       if (lock) lock.releaseLock();
       throw e;
     }
     // Esperar antes de reintentar (tiempo exponencial)
     Utilities.sleep(Math.pow(2, retryCount) * 1000);
   }
 }

 try {
   ConControl = ss.getSheetByName("CON-Control");
   var controlArray = ConControl.getRange('A2:M2').getValues();
   verboseFlag = controlArray[0][11];
   descreetFlag = controlArray[0][12];

   //Check if process is operational
   if(controlArray[0][0] == false){
     if(verboseFlag === true) {
       console.log('processQueue not operational.')
     };
     lock.releaseLock();
     return;
   }

   //Check if it is running time
   var now = Utilities.formatDate(new Date(), timezone, "yyyy-MM-dd HH:mm:ss");

   if(controlArray[0][9]>controlArray[0][10]) {
     //Check if its already 6 minutes blocked considering next run
     var auxTime = controlArray[0][1];
     auxTime.setMinutes(auxTime.getMinutes() + 6);
     var nextRunPlus = Utilities.formatDate(auxTime, timezone, "yyyy-MM-dd HH:mm:ss");
     //Its Ok
     if (nextRunPlus > now) {
       if(verboseFlag === true) {
         console.log('Another thread of processQueue already running');
       };
       lock.releaseLock();
       return;
     } else {
       if(verboseFlag === true) {
         console.log('Process next runtime reestablish');
       };
       //Block another thread
       var nextRun = new Date();
       auxTime= nextRun.getMinutes() + 6;
       nextRun.setMinutes(auxTime);
       ConControl.getRange(2,2).setValue(nextRun);
     }
   }

   var nextRun = Utilities.formatDate(controlArray[0][1], timezone, "yyyy-MM-dd HH:mm:ss");
   if( nextRun>now ) {
     if(verboseFlag === true) {
       console.log('processQueue on time. It will run '+controlArray[0][1])
     };
     lock.releaseLock();
     return;
   }

   //Set Last check begin
   ConControl.getRange(2,10).setValue(now);
   var scriptProperties = PropertiesService.getScriptProperties();
   SpreadsheetApp.flush();

   try {
     //Open CON-TaskCurrent
     ConTask = ss.getSheetByName("CON-TaskCurrent");
     var ConTaskEnd = ConTask.getLastRow();
     var ConTaskData = ConTask.getRange(2, 1, ConTaskEnd, ConTask.getLastColumn()); 
     var ConTaskValues = ConTaskData.getValues();
     var ConTaskNDX = ConTaskValues.map(function(r){ return r[6]; });
     //Load and save las case in Cache
     var lastRow = 0;
     var FirstPending = ConTaskNDX.indexOf("01 Pending", lastRow);
     var FirstRetry = ConTaskNDX.indexOf("04 Retry", lastRow);     

     //Something to process 
     if(FirstPending !== -1 || FirstRetry !== -1) {

       if((FirstPending < FirstRetry || FirstRetry === -1) && FirstPending != -1) {
         var ConTaskBegin = FirstPending; 
       } else {
         var ConTaskBegin = FirstRetry;
       }

       //Check if there is a DWO change to prioritize
       let onlyDWOProcess = false;
       //Begin loop
       for (var iTask = ConTaskBegin; iTask < ConTaskEnd - 1; iTask++) {
         // Check for pending / Continue if 02 Incorporated or 03 Inhibited
         if( ConTaskValues[iTask][6]!="01 Pending" && ConTaskValues[iTask][6]!="04 Retry" ) {continue;}
         // If repeated key, discarded
         if((iTask >0 && iTask - 1 >= ConTaskBegin) || ConTaskValues[iTask][1]=="") {
           if((ConTaskValues[iTask][0]==ConTaskValues[iTask - 1][0] && ConTaskValues[iTask][1]==ConTaskValues[iTask - 1][1] && ConTaskValues[iTask - 1][7]=="") || ConTaskValues[iTask][1]=="") {
             continue;
           }
         }
         if (ConTaskValues[iTask][0]==="DWO") {
           onlyDWOProcess = true;
           if(verboseFlag === true) {
             console.log('DWO prioritized run');
           };
           break;
         }
       }

       //Begin loop
       for (var iTask = ConTaskBegin; iTask < ConTaskEnd - 1; iTask++) {
         // Verificar tiempo de ejecución
         if (Date.now() - START_TIME > MAX_EXECUTION_TIME) {
           if(verboseFlag === true) {
             console.log('Maximum execution time reached, stopping process');
           }
           break;
         }
         
         // Check for pending / Continue if 02 Incorporated or 03 Inhibited
         if( ConTaskValues[iTask][6]!="01 Pending" && ConTaskValues[iTask][6]!="04 Retry") {continue;}

         // If only DWO process
         if (onlyDWOProcess === true && ConTaskValues[iTask][0]!="DWO") {
           continue;
         }
       
         // Saves track info
         ConTask.getRange(iTask + 2,8).setValue(now); 

         // If repeated key, discarded
         if((iTask >0 && iTask - 1 >= ConTaskBegin) || ConTaskValues[iTask][1]=="") {
           if((ConTaskValues[iTask][0]==ConTaskValues[iTask - 1][0] && ConTaskValues[iTask][1]==ConTaskValues[iTask - 1][1] && ConTaskValues[iTask - 1][7]=="") || ConTaskValues[iTask][1]=="") {
             ConTask.getRange(iTask + 2,7).setValue("05 Discarded");
             continue;
           }
         }

         //Current case
         var furtherAction = DataLoad(ConTaskValues[iTask][3],ConTaskValues[iTask][0], ConTaskValues[iTask][1]);

         if(furtherAction=="Discarded"){
           ConTask.getRange(iTask + 2,7).setValue("05 Discarded");
           continue;
         }

         // If current = 04 Retry, save in comment
         var retryNumber = scriptProperties.getProperty('retryNumber');
         if( ConTaskValues[iTask][6]==="04 Retry" || (retryNumber != null && retryNumber != "" )) {
           // Previous error
           retryNumber = parseInt(retryNumber);
           var errorMSG = scriptProperties.getProperty('errorMSG');
           if (isNaN(retryNumber)) {retryNumber = "1"; scriptProperties.setProperty('retryNumber', retryNumber);errorMSG="";};
           scriptProperties.setProperty('retryNumber', retryNumber);
           ConTask.getRange(iTask + 2,9).setValue("04 Retry ("+retryNumber+") "+errorMSG);
         } else {
           // Mark as 04 Retry just in halt case
           ConTask.getRange(iTask + 2,7).setValue("04 Retry");
         };

         if(verboseFlag === true) {
           console.log("Process: "+iTask+" // "+ConTaskValues[iTask][3]+" // "+ConTaskValues[iTask][0]+" // Key: "+ConTaskValues[iTask][1]+ "//"+ ConTaskValues[iTask][8] )
         }

         // PROCESS BEGIN
         var variationResult = ObtainVariation(ConTaskValues[iTask]);

         if( variationResult["variationStatus"] === "unchanged" ) {
           ConTask.getRange(iTask + 2,7).setValue("06 Unchanged");
           continue;
         } else if( variationResult["variationStatus"] === "source missed key" ) {
           ConTask.getRange(iTask + 2,7).setValue("07 Source missed key");
           continue;
         } else {
           //Recording process
           if(variationResult["logAddHTML"] && variationResult["logAddHTML"] !== "" && 
              (descreetFlag==false || ConTaskValues[iTask][5]!="appsheet@mediaaccesscompany.com")) {
             //Log
             logSheet.getRange(logSheet.getLastRow() + 1, 1, 1, 7).setValues([
               [
                 ConTaskValues[iTask][1],
                 ConTaskValues[iTask][2],
                 ConTaskValues[iTask][3]+" / "+ConTaskValues[iTask][0],  
                 ConTaskValues[iTask][5],
                 variationResult["logAddHTML"] || "",
                 variationResult["logAddPlain"] || "",
                 variationResult["variationCode"] || ""
               ]
             ]);

             //Source update Log
             if(variationResult["sourceData"] && variationResult["sourceData"][taskColLog] && 
                variationResult["sourceData"][taskColLog]!="") {
               let aux = variationResult["sourceData"][taskColLog];
               let columnIndex = Math.max(1, taskColLog+1);
               sourceSheet.getRange(sourceRow+2, columnIndex).setValue(aux);
             }
             // If comment
             if(variationResult["cleanComment"] && variationResult["cleanComment"] > -1) {
               let columnIndex = Math.max(1, variationResult["cleanComment"]+1);
               sourceSheet.getRange(sourceRow+2, columnIndex).setValue("");
             }
           }
           let aux = [variationResult["sourceData"] || []];
           
           if( variationResult["variationStatus"] === "append" ) {
             //New witness
             witnessSheet.appendRow(aux[0]);
           } else {
             //Witness overwrite with Source
             // Asegurar que aux[0].length sea al menos 1
             let columnCount = Math.max(1, aux[0].length);
             witnessSheet.getRange(witnessRow+2, 1, 1, columnCount).setValues(aux);
           //            witnessSheet.getRange(witnessRow+2,1,1, witnessSheet.getLastColumn()).setValues(aux);
           }

           //DWO Status changed
           ConTask.getRange(iTask + 2,7).setValue("02 Incorporated");
         }
         
         // Añadir pequeña pausa cada X iteraciones para evitar sobrecarga
         if (iTask % 20 === 0) {
           Utilities.sleep(1000);
         }
       }
     }
     // Reset flag of any chance of retry
     var nullValue = ""; scriptProperties.setProperty('retryNumber', nullValue); scriptProperties.setProperty('errorMSG', nullValue);
   // Error detection
   } catch (e) {
     // Mejorar el registro de errores
     console.error('Error in processQueue: ' + e.toString());
     console.error('Stack: ' + e.stack);
     
     // Asegurarse de liberar el lock en caso de error
     if (lock) lock.releaseLock();
     
     // Propagar el error
     throw e;
   } finally {
     SpreadsheetApp.flush();
     lock.releaseLock();
   }
   
   var currentTime = Utilities.formatDate(new Date, timezone, "HH:mm");
   var workdayTimeFrom = controlArray[0][2].toTimeString();
   var workdayTimeTo = controlArray[0][3].toTimeString();

   var weekday = new Date(); weekday = weekday.getDay(); 
   if (currentTime > workdayTimeFrom && currentTime< workdayTimeTo && weekday!=6 && weekday!=0 ){
     var nextTriggerMin = controlArray[0][4];
   } else {
     var nextTriggerMin = controlArray[0][5];
   }
   //Set Last check timestamp
   ConControl.getRange(2,2).setValue(nextTriggerMin);

   //Set Last check end
   nextRun = new Date();
   let aux2= nextRun.getMinutes() + nextTriggerMin;
   nextRun.setMinutes(aux2);
   ConControl.getRange(2,2).setValue(nextRun);
   ConControl.getRange(2,11).setValue(Utilities.formatDate(new Date(), timezone, timestamp_format));

 } catch (e) {
   // Mejorar el registro de errores
   console.error('Error in processQueue: ' + e.toString());
   console.error('Stack: ' + e.stack);
   
   // Asegurarse de liberar el lock en caso de error
   if (lock) lock.releaseLock();
   
   // Propagar el error
   throw e;
 }
}


//
/* BUSINESS UTILITIES*/
function DataLoad(ssAux, sheetNameAux, keyAux) {
  let actionReturn = ""; 
  let ssAux2 = ""; 
  let flag = true;
  let dwoKeyAux = "";

  if(inhibitAppend(ssAux, sheetNameAux, keyAux)===false) {
    actionReturn="Discarded";
    return actionReturn;
  }

  //Call for Source
  LazyLoad(ssAux, sheetNameAux);

  sourceSheet = containerSheet;
  sourceValues = containerValues;
  sourceNDX = containerNDX;

  // Obtain label data
  labelRow = labelNDX.indexOf(sheetNameAux,0);
  taskLabelNames = labelValues[labelRow];
  taskLabelActions = labelValues[labelRow + 1];
  taskColKey = taskLabelActions.indexOf("K") - 1;
  taskColUser = taskLabelNames.indexOf("Last user", 0) - 1; 
  taskColChange = taskLabelNames.indexOf("Last change", 0) - 1;
  taskColLog = taskLabelNames.indexOf("LogTrack", 0) - 1;
  var auxLabel = labelValues[labelRow + 2][0];

  //Check if previous call was equal 
  if(ssPrevious === ssAux && sheetNamePrevious === sheetNameAux) {return actionReturn;}
  if( ssAux === "DubAppActive01") {
    ssAux2 = "DubAppTotal01";
  } else {
    if(auxLabel == "[Split]") {
      let auxRow = sourceNDX.indexOf(keyAux);
      // Buscar la configuración de la hoja en relacionesHojas
      let dwoColLetter = relacionesHojas[sheetNameAux]?.dwoCol || 'A';
      // Convertir letra(s) de columna a número
      let dwoCol = dwoColLetter.split('')
        .reduce((acc, c) => (acc * 26) + (c.charCodeAt(0) - 'A'.charCodeAt(0) + 1), 0) - 1;
      let dwoKeyAux = sourceValues[auxRow][dwoCol];

      LazyLoad("DubAppActive01", "DWO");
      let activeDWONDX = containerNDX;

      LazyLoad("DubAppTotal01", sheetNameAux);
      let totalDWOValues = containerValues;
      let totalDWONDX = containerNDX;

      if(activeDWONDX.indexOf(dwoKeyAux) !== -1 || 
         (sheetNameAux === 'DWO' && (() => {
           let dwoRow = totalDWONDX.indexOf(keyAux);
           // Verificar que dwoRow sea válido y que totalDWOValues[dwoRow] tenga al menos 59 columnas
           return dwoRow !== -1 && 
                  dwoRow < totalDWOValues.length && 
                  totalDWOValues[dwoRow].length > 58 && 
                  totalDWOValues[dwoRow][58] === "(01) On track: DWO";
         })())
      ) {
        ssAux2 = "DubAppActive01";
      } else {
        ssAux2 = "DubAppLogs01";
      }
    } else {
      ssAux2 = "DubAppActive01";
    }
  };


  if (ssAux2!="") {
    LazyLoad(ssAux2, sheetNameAux);
    witnessSheet = containerSheet; 
    witnessValues = containerValues;
    witnessNDX = containerNDX;
  }

  //Call for log
  
  if(sheetNamePrevious !== sheetNameAux) {
    sheetNameAux2 = sheetNameAux+"Log";
    LazyLoad("DubAppLogs01", sheetNameAux2);
  
    logSheet = containerSheet;
    logValues = containerValues;
    logNDX = containerNDX;
  }

  //New witness
  ssPrevious = ssAux;
  sheetNamePrevious = sheetNameAux;
  
  return actionReturn;
}

function ObtainVariation(taskAux) {

 // Initialize return
 var variationAux = {
   variationStatus: "unchanged",
   variationCode: null,
   logAddHTML: "",
   logAddPlain: "",
   cleanComment: -1,
   statusSplit: false,
   sourceData: false,
   sourceRow: false,
   witnessData: false,
   witnessRow: false,
   dwoChangeStatus: "no"
 };
 var sourceAux = null; var witnessAux = null; var sourceLog = null; var witnessLog = null; 

 // Obtain source & witness
 var aux3 = taskAux[1]+"";
 sourceRow = sourceNDX.indexOf(aux3); 
 if(sourceRow == -1) {variationAux["variationStatus"]="source missed key"; return variationAux;};
 variationAux["sourceRow"]=sourceRow; 
 
 // Verificar que sourceRow sea válido antes de acceder a sourceValues
 if (sourceRow < 0 || sourceRow >= sourceValues.length) {
   variationAux["variationStatus"]="source missed key"; 
   return variationAux;
 }
 
 var sourceData = sourceValues[sourceRow]; 
 witnessRow = witnessNDX.indexOf(aux3);
 if(witnessRow == -1) {
   // Not exist in witness
   aux3=taskAux[0]+aux3;
   //Check if already appended in the same batch
   if(inhibitAppend(taskAux[3], taskAux[0], aux3)) {
     //Ok to append
     variationAux["variationStatus"]="append"; variationAux["variationCode"]=0;
   } else {
     //Already appended
     variationAux ["variationStatus"] = "unchanged";
     return variationAux;
   }
 } else {
   // Verificar que witnessRow sea válido antes de acceder a witnessValues
   if (witnessRow < 0 || witnessRow >= witnessValues.length) {
     variationAux["variationStatus"]="witness missed key"; 
     return variationAux;
   }
   
   // Exists in witness
   var witnessData = witnessValues[witnessRow]; 
   variationAux["witnessData"] = witnessValues[witnessRow]; 
   variationAux["witnessRow"]=witnessRow;

   //Halt if equal
   if(sourceData == witnessData) {return variationAux;};

   // Loop for variations
   for(var caseAux = 0; caseAux < sourceData.length; ++caseAux){
     witnessLog = ""; sourceLog = "";

     if(caseAux == taskColChange) {
       caseAction="DT";
     } else {
       // Asegurar que caseAux+1 sea un índice válido para taskLabelActions
       var caseAction = (caseAux+1 < taskLabelActions.length) ? taskLabelActions[caseAux+1] : "X";
     } 
     // Skip if X or Key
     if(caseAction=="K" || caseAction=="X") {continue};

     //Check if value changed
     // If Date or Datetime format
     if(caseAction.substring(0,1) == 'D') {
       if(caseAction == "DT" || caseAux == taskColChange) {
         // Datetime
         sourceAux = dateHandle(sourceData[caseAux], timezone, timestamp_format);
         // Verificar que caseAux sea un índice válido para witnessData
         witnessAux = (witnessData && caseAux < witnessData.length) ? 
                      dateHandle(witnessData[caseAux], timezone, timestamp_format) : "";
       } else {
         // Date
         sourceAux = dateHandle(sourceData[caseAux], timezone, "dd/MM/yyyy");
         witnessAux = dateHandle(witnessData[caseAux], timezone, "dd/MM/yyyy");
       };
     } else {sourceAux = sourceData[caseAux]; witnessAux = witnessData[caseAux]};

     // Check if different
     if(sourceAux === witnessAux) {continue};

     // If update
     if (variationAux["variationStatus"]=="unchanged") {
       variationAux["variationStatus"]="update"; variationAux["variationCode"]=1;
     }

     //Special interpretation
     if(caseAction == "" || caseAction.substring(0,1) == 'D') {
       // for most of cases and date
       witnessLog=witnessAux; sourceLog=sourceAux;
     // ! = Channel Event Type
     } else if(caseAction == "!") {
       let response = uncodeChannelEventType(sourceAux, witnessAux);
       sourceLog=response["chainIn"]; witnessLog=response["chainOut"];
     } // @ = User
     else if(caseAction == "@") {
       let aux = userNDX.indexOf(witnessAux); if (aux != -1) {witnessLog = userValues[aux][1];} 
       aux = userNDX.indexOf(sourceAux); if (aux != -1) {sourceLog = userValues[aux][1];} 
     } // F = File
     else if(caseAction == "F") {
       if(sourceAux!="" && witnessAux!="") {sourceLog = "File changed"} else
       if(sourceAux!="") {sourceLog = "File uploaded"} else {sourceLog = "File deleted"};
       witnessAux="";
     // C = Comment
     } else if(caseAction == "C") {
       variationAux["cleanComment"]=caseAux;
       witnessAux=""; sourceLog="'"+sourceAux+"'";
       variationAux["variationStatus"]="Comment included"; variationAux["variationCode"]=3;
     } // : = Extract suffix
     else if(caseAction.substring(0,1) == ':') {
       if (Object.prototype.toString.call(witnessAux) === '[object String]') {
         witnessLog = witnessAux.replaceAll(caseAction,"");
       }
       if (Object.prototype.toString.call(sourceAux) === '[object String]') {
         sourceLog = sourceAux.replaceAll(caseAction,"");
       }
     // Otherwise
     } else {
       witnessLog="'"+witnessAux+"'"; sourceLog="'"+sourceAux+"'";
     }
     if(caseAux != taskColChange) {
       if(witnessAux==="" || witnessAux === null) {
         variationAux["logAddHTML"]=variationAux["logAddHTML"] + "<br><b>["+taskLabelNames[caseAux + 1]+"]:</b> "+sourceLog+"</li>";
         variationAux["logAddPlain"]=variationAux["logAddPlain"] + "["+taskLabelNames[caseAux + 1]+"]: "+sourceLog+" deleted // " ;
       } else if (sourceAux==="" || sourceAux === null) {
         variationAux["logAddHTML"]=variationAux["logAddHTML"] + "<br><b>["+taskLabelNames[caseAux + 1]+"]:</b> <del>"+witnessLog+"</del></li>";
         variationAux["logAddPlain"]=variationAux["logAddPlain"] + "["+taskLabelNames[caseAux + 1]+"]: "+witnessLog+" deleted // " ;
       } else {
         variationAux["logAddHTML"]=variationAux["logAddHTML"] + "<br><b>["+taskLabelNames[caseAux + 1]+"]:</b> <del>"+witnessLog+"</del> ► "+sourceLog+"</li>";
         variationAux["logAddPlain"]=variationAux["logAddPlain"] + "["+taskLabelNames[caseAux + 1]+"]:"+witnessLog+" ► "+sourceLog+" // " ;
       }
     }
   }
 }
 //Log completeness
 if(variationAux ["variationStatus"] !== "unchanged") {
   let aux = userNDX.indexOf(sourceData[taskColUser]); var auxHTML=""; var auxPlain="";
   // Different User
   if(variationAux ["variationStatus"] != "append" && sourceData[taskColUser] != witnessData[taskColUser]) {
     if (aux != -1) {
       auxHTML = "<hr><b><small>Changed by </small><mark>"+userValues[aux][1]+"</mark> ";
       auxPlain = "Changed by "+userValues[aux][1]+" ";
     };
   } else if (variationAux ["variationStatus"] == "append" && aux != -1) {
       auxHTML = "<b><small>Created by </small><mark>"+userValues[aux][1]+"</mark> ";
       auxPlain = "Created by "+userValues[aux][1]+" ";
   } else {
     auxHTML="<br>";
   }
   // Different Day
   if(variationAux ["variationStatus"] != "append") {
       // Verificar que taskColChange sea un índice válido para sourceData y witnessData
       if (taskColChange >= 0 && taskColChange < sourceData.length && 
           witnessData && taskColChange < witnessData.length) {
         sourceAux = dateHandle(sourceData[taskColChange], timezone, "yyyy-MM-dd");
         witnessAux = dateHandle(witnessData[taskColChange], timezone, "yyyy-MM-dd");
         if (sourceAux != witnessAux) {      
           sourceAux = dateHandle(sourceData[taskColChange], timezone, timestamp_format);
         } else {
           sourceAux = dateHandle(sourceData[taskColChange], timezone, "HH:mm:ss");
         }
       } else {
         // Si no son válidos, usar un valor por defecto
         sourceAux = "Fecha desconocida";
       }
       variationAux["logAddHTML"]= auxHTML +"<small>"+ sourceAux+ "</small></b>" + variationAux["logAddHTML"];
       variationAux["logAddPlain"]= auxPlain + sourceAux +" // " + variationAux["logAddPlain"] + "\n";
   } else {
     // Verificar que taskColChange sea un índice válido para sourceData
     if (taskColChange >= 0 && taskColChange < sourceData.length) {
       sourceAux = dateHandle(sourceData[taskColChange], timezone, timestamp_format);
     } else {
       // Si no es válido, usar un valor por defecto
       sourceAux = "Fecha desconocida";
     }
     variationAux["logAddHTML"]= auxHTML +"<small>"+ sourceAux + "</small></b><li>";
     variationAux["logAddPlain"]= auxPlain + sourceAux + "\n";
   }

   //Reset comment
   if( variationAux["cleanComment"] > -1 ) {sourceData[variationAux["cleanComment"]]=""};
   //Add to previous log
   if (taskColLog >= 0 && taskColLog < sourceData.length) {
     // Si no existe, inicializarlo como cadena vacía
     if (sourceData[taskColLog] === undefined || sourceData[taskColLog] === null) {
       sourceData[taskColLog] = "";
     }
     sourceData[taskColLog] = sourceData[taskColLog] + variationAux["logAddHTML"];
   }
 }
 variationAux["sourceData"]=sourceData;
 
 return variationAux
}

function Llamador() { 
 //Open control
 var ss = SpreadsheetApp.openById(allIDs['controlID']);
 ConControl = ss.getSheetByName("CON-Control");
 //Check if process is operational
 var controlArray = ConControl.getRange('A2:M2').getValues();
 if(controlArray[0][0] == false ) {return;}
 verboseFlag = controlArray[0][11];
 descreetFlag = controlArray[0][12];
 ConTask = ss.getSheetByName("CON-TaskCurrent");
 };

/**
 * Limpia todos los locks que puedan haber quedado activos por interrupciones inesperadas.
 * Esta función debe ejecutarse manualmente cuando se detecten problemas con locks persistentes.
 * 
 * Realiza las siguientes acciones:
 * 1. Libera el lock principal del script
 * 2. Intenta forzar la liberación de cualquier lock persistente
 * 3. Limpia propiedades de script relacionadas con locks o estados de procesamiento
 * 4. Restablece el estado de control en la hoja CON-Control
 * 5. Limpia entradas específicas de la caché del script
 * 
 * @example
 * // Ejecutar manualmente desde el editor de scripts cuando el proceso quede bloqueado
 * function manualUnlock() {
 *   clearAllLocks();
 * }
 * 
 * @return {boolean} - true si la operación fue exitosa, false en caso de error
 */
function clearAllLocks() {
  try {
    console.log('Iniciando liberación de todos los locks...');
    
    // Liberar el lock principal del script
    const scriptLock = LockService.getScriptLock();
    if (scriptLock.hasLock()) {
      scriptLock.releaseLock();
      console.log('Lock principal liberado');
    }
    
    // Intentar forzar la liberación de cualquier lock persistente
    try {
      // Crear un nuevo lock y liberarlo inmediatamente para asegurar que no queden locks activos
      const forceLock = LockService.getScriptLock();
      forceLock.tryLock(10000); // Intentar obtener un lock con 10 segundos de timeout
      if (forceLock.hasLock()) {
        forceLock.releaseLock();
        console.log('Lock forzado liberado correctamente');
      }
    } catch (error) {
      console.warn(`Error al forzar liberación de lock: ${error.message}`);
    }
    
    // Limpiar propiedades de script que puedan estar relacionadas con locks o estados de procesamiento
    try {
      const scriptProperties = PropertiesService.getScriptProperties();
      
      // Limpiar propiedades específicas relacionadas con el procesamiento
      const propertiesToClear = ['retryNumber', 'errorMSG', 'processingStatus'];
      
      propertiesToClear.forEach(propKey => {
        if (scriptProperties.getProperty(propKey)) {
          scriptProperties.deleteProperty(propKey);
          console.log(`Propiedad eliminada: ${propKey}`);
        }
      });
      
      // Opcionalmente, buscar y eliminar otras propiedades relacionadas con locks
      const allKeys = scriptProperties.getKeys();
      const lockRelatedKeys = allKeys.filter(key => 
        key.toLowerCase().includes('lock') || 
        key.toLowerCase().includes('process') || 
        key.toLowerCase().includes('retry')
      );
      
      lockRelatedKeys.forEach(key => {
        scriptProperties.deleteProperty(key);
        console.log(`Propiedad relacionada con locks eliminada: ${key}`);
      });
    } catch (error) {
      console.warn(`Error al limpiar propiedades de script: ${error.message}`);
    }
    
    // Restablecer el estado de control en la hoja CON-Control
    try {
      const ss = SpreadsheetApp.openById(allIDs['controlID']);
      const ConControl = ss.getSheetByName("CON-Control");
      
      // Restablecer el contador de ejecuciones
      ConControl.getRange(2, 10).setValue(0); // Restablecer Last check begin
      ConControl.getRange(2, 11).setValue(0); // Restablecer Last check end
      
      // Establecer la próxima ejecución para dentro de 1 minuto
      var nextRun = new Date();
      nextRun.setMinutes(nextRun.getMinutes() + 1);
      ConControl.getRange(2, 2).setValue(nextRun);
      
      console.log('Estado de control restablecido correctamente');
    } catch (error) {
      console.warn(`Error al restablecer estado de control: ${error.message}`);
    }
    
    // Limpiar la caché del script
    try {
      const cache = CacheService.getScriptCache();
      // No podemos limpiar toda la caché directamente, pero podemos invalidar entradas específicas
      // relacionadas con el procesamiento si las conocemos
      const cacheKeysToInvalidate = ['processStatus', 'lastRun', 'queueStatus'];
      
      cacheKeysToInvalidate.forEach(key => {
        cache.remove(key);
        console.log(`Entrada de caché invalidada: ${key}`);
      });
    } catch (error) {
      console.warn(`Error al limpiar caché: ${error.message}`);
    }
    
    console.log('Proceso de liberación de locks completado exitosamente');
    return true;
  } catch (error) {
    console.error('Error al liberar locks:', error.message);
    return false;
  }
}

function uncodeChannelEventType(chainSource, chainWitness) {
  // Cargamos la hoja DWO-ChannelEventType usando LazyLoad
  LazyLoad("DubAppActive01", "DWO-ChannelEventType");
  
  // Ahora podemos usar las variables globales que LazyLoad ha inicializado:
  // containerValues y containerNDX

  var response = {
    chainIn: "",
    chainOut: ""
  };

  // Si alguna de las cadenas es nula o vacía, retornamos la respuesta vacía
  if (!chainSource || !chainWitness) {
    return response;
  }

  let auxChainSource = chainSource.split(" , ");
  let auxChainWitness = chainWitness.split(" , ");
  let aux = null;
  let eventTypeRow = null;

  // Loop Source
  let comma = "";
  for(var loopT = 0; loopT < auxChainSource.length; loopT++) {
    aux = auxChainWitness.indexOf(auxChainSource[loopT]);
    if(aux == -1) {
      eventTypeRow = containerNDX.indexOf(auxChainSource[loopT]);
      if(eventTypeRow != -1) {
        response["chainIn"] = response["chainIn"] + comma + containerValues[eventTypeRow][13];
        comma = " , ";
      }
    }
  }

  // Loop Witness
  comma = "";
  for(var loopT = 0; loopT < auxChainWitness.length; loopT++) {
    aux = auxChainSource.indexOf(auxChainWitness[loopT]);
    if(aux == -1) {
      eventTypeRow = containerNDX.indexOf(auxChainWitness[loopT]);
      if(eventTypeRow != -1) {
        response["chainOut"] = response["chainOut"] + comma + containerValues[eventTypeRow][13];
        comma = " , ";
      }
    }
  }
  return response;
}

//
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

 function inhibitAppend(auxEnvironment, auxTable, auxKey) {
 //Check if case already was appended
   let key2find = auxEnvironment+" / "+auxTable+" / "+auxKey;

   if(inhibitedAppend.indexOf(key2find)==-1) {
     inhibitedAppend.push(key2find);
     return true;
   }
   return false;
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

 function setCache(auxKey, valueToCache) {

   var cache = CacheService.getScriptCache(); // Obtener la caché del script
   var valueType = typeof valueToCache; // Obtener el tipo de la variable

   // Crear un objeto JSON que contenga el valor y su tipo
   var cacheObject = {
     value: valueToCache,
     type: valueType
   };

   // Convertir el objeto JSON a cadena y guardar en la caché
   cache.put(auxKey, JSON.stringify(cacheObject), 21600); // 6 horas
 }


 function getCachedValue(auxKey) {
   var cache = CacheService.getScriptCache(); // Obtener la caché del script
   var cachedValue = cache.get(auxKey); // Recuperar el valor del caché

   if (cachedValue) {
     // Convertir la cadena JSON de vuelta a un objeto
     var cacheObject = JSON.parse(cachedValue);
     
     // Recuperar el valor y su tipo
     var value = cacheObject.value;
     var type = cacheObject.type;

     // Convertir el valor a su tipo original
     switch (type) {
       case 'number':
         value = Number(value);
         break;
       case 'string':
         value = String(value);
         break;
       case 'boolean':
         value = Boolean(value);
         break;
       // Agregar otros tipos según sea necesario
     }

     Logger.log('Valor en caché: ' + value + ' (Tipo: ' + typeof value + ')');
     return value;
   } else {
     Logger.log('No se encontró el valor en la caché.');
     return -1;
   }
 }

/**
 * Verifica el estado actual de los locks y propiedades relacionadas con el procesamiento.
 * Útil para diagnosticar problemas de locks persistentes.
 * 
 * Recopila y devuelve información sobre:
 * 1. Si hay un lock activo en el script
 * 2. Propiedades de script relacionadas con el procesamiento (retryNumber, errorMSG, etc.)
 * 3. Estado actual de la hoja de control (CON-Control)
 * 4. Entradas relevantes en la caché del script
 * 
 * @example
 * // Ejecutar manualmente desde el editor de scripts para diagnosticar problemas
 * function diagnoseIssues() {
 *   const status = checkLockStatus();
 *   console.log(JSON.stringify(status, null, 2));
 * }
 * 
 * @return {Object} - Objeto con información detallada sobre el estado de los locks y propiedades
 *   - hasScriptLock: {boolean} Indica si hay un lock activo
 *   - scriptProperties: {Object} Propiedades de script relacionadas con el procesamiento
 *   - controlSheetStatus: {Object} Estado actual de la hoja de control
 *   - cacheEntries: {Object} Entradas relevantes en la caché
 */
function checkLockStatus() {
  try {
    console.log('Verificando estado de locks y propiedades...');
    
    const result = {
      hasScriptLock: false,
      scriptProperties: {},
      controlSheetStatus: {},
      cacheEntries: {}
    };
    
    // Verificar si hay un lock activo
    try {
      const scriptLock = LockService.getScriptLock();
      result.hasScriptLock = scriptLock.hasLock();
      console.log(`Lock activo: ${result.hasScriptLock}`);
    } catch (error) {
      console.warn(`Error al verificar lock: ${error.message}`);
      result.lockError = error.message;
    }
    
    // Verificar propiedades de script relacionadas con el procesamiento
    try {
      const scriptProperties = PropertiesService.getScriptProperties();
      const propertiesToCheck = ['retryNumber', 'errorMSG', 'processingStatus'];
      
      propertiesToCheck.forEach(propKey => {
        const propValue = scriptProperties.getProperty(propKey);
        if (propValue) {
          result.scriptProperties[propKey] = propValue;
          console.log(`Propiedad ${propKey}: ${propValue}`);
        }
      });
      
      // Buscar otras propiedades relacionadas con locks
      const allKeys = scriptProperties.getKeys();
      const lockRelatedKeys = allKeys.filter(key => 
        key.toLowerCase().includes('lock') || 
        key.toLowerCase().includes('process') || 
        key.toLowerCase().includes('retry')
      );
      
      lockRelatedKeys.forEach(key => {
        if (!result.scriptProperties[key]) {
          result.scriptProperties[key] = scriptProperties.getProperty(key);
          console.log(`Propiedad relacionada con locks ${key}: ${result.scriptProperties[key]}`);
        }
      });
    } catch (error) {
      console.warn(`Error al verificar propiedades de script: ${error.message}`);
      result.propertiesError = error.message;
    }
    
    // Verificar estado de la hoja de control
    try {
      const ss = SpreadsheetApp.openById(allIDs['controlID']);
      const ConControl = ss.getSheetByName("CON-Control");
      const controlValues = ConControl.getRange('A2:M2').getValues()[0];
      
      result.controlSheetStatus = {
        operational: controlValues[0],
        nextRun: controlValues[1],
        lastCheckBegin: controlValues[9],
        lastCheckEnd: controlValues[10],
        verboseFlag: controlValues[11],
        descreetFlag: controlValues[12]
      };
      
      console.log('Estado de hoja de control:', JSON.stringify(result.controlSheetStatus));
    } catch (error) {
      console.warn(`Error al verificar hoja de control: ${error.message}`);
      result.controlSheetError = error.message;
    }
    
    // Verificar entradas de caché
    try {
      const cache = CacheService.getScriptCache();
      const cacheKeysToCheck = ['processStatus', 'lastRun', 'queueStatus'];
      
      cacheKeysToCheck.forEach(key => {
        const cacheValue = cache.get(key);
        if (cacheValue) {
          result.cacheEntries[key] = cacheValue;
          console.log(`Entrada de caché ${key}: ${cacheValue}`);
        }
      });
    } catch (error) {
      console.warn(`Error al verificar caché: ${error.message}`);
      result.cacheError = error.message;
    }
    
    console.log('Verificación de estado de locks completada');
    return result;
  } catch (error) {
    console.error('Error al verificar estado de locks:', error.message);
    return { error: error.message };
  }
}
