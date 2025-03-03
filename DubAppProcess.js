/*
### Árbol de Llamadas de Funciones
processQueue()
├── DataLoad(ssAux, sheetNameAux, keyAux)
│   ├── inhibitAppend(auxEnvironment, auxTable, auxKey)
│   └── LazyLoad(ssAux, sheetNameAux)
├── ObtainVariation(taskAux)
│   ├── dateHandle(d, timezone, timestamp_format)
│   └── uncodeChannelEventType(chainSource, chainWitness)
│       └── LazyLoad("DubAppActive01", "DWO-ChannelEventType")

Funciones Utilitarias:
- dateHandle(d, timezone, timestamp_format)
- isValidDate(d)
- daysBetweenToday(dateParam)
- setCache(auxKey, valueToCache)
- getCachedValue(auxKey)
- inhibitAppend(auxEnvironment, auxTable, auxKey)

Funciones de Control:
- Llamador()
- reset()
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

 let lock = null;
 try {
   // Reducir el número máximo de reintentos y ajustar el tiempo máximo de ejecución
   const MAX_RETRIES = 2;
   const MAX_EXECUTION_TIME = 6 * 60 * 1000; // 6 minutos para dar más margen
   var BATCH_SIZE = 10; // Reducir el tamaño del lote para procesar menos registros por iteración
   const START_TIME = Date.now();

   let retryCount = 0;
   let ss;

   while (retryCount < MAX_RETRIES) {
     try {
       lock = LockService.getScriptLock();
       // Aumentar el timeout del lock a 5 segundos
       if (!lock.tryLock(5000)) {
         console.log('Otro proceso está ejecutándose');
         return;    
       }

       //Check if service enabled 
       ss = SpreadsheetApp.openById(allIDs['controlID']);
       break; // Si tiene éxito, salir del bucle
       
     } catch (e) {
       retryCount++;
       if (retryCount === MAX_RETRIES) {
         console.error('Failed to access spreadsheet after ' + MAX_RETRIES + ' attempts: ' + e.toString());
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
       // Asegurarse de que ConTaskEnd sea al menos 2 para evitar arrays vacíos
       if (ConTaskEnd < 2) {
         if(verboseFlag === true) {
           console.log('No tasks to process');
         }
         // Actualizar la columna K para indicar que el proceso ha terminado
         ConControl.getRange(2,11).setValue(Utilities.formatDate(new Date(), timezone, timestamp_format));
         return;
       }
       
       var ConTaskData = ConTask.getRange(2, 1, ConTaskEnd - 1, ConTask.getLastColumn()); 
       var ConTaskValues = ConTaskData.getValues();
       var ConTaskNDX = ConTaskValues.map(function(r){ return r[6]; });
       //Load and save las case in Cache
       var lastRow = 0;
       var FirstPending = ConTaskNDX.indexOf("01 Pending", lastRow);
       var FirstRetry = ConTaskNDX.indexOf("04 Retry", lastRow);     
       var onlyDWOProcess=false;
       //Something to process 
       if(FirstPending !== -1 || FirstRetry !== -1) {
         //Check if there is a DWO to process
         
         // Buscar el primer DWO pendiente
         for (var i = FirstPending; i < ConTaskEnd - 1; i++) {
           if ((ConTaskValues[i][6] == "01 Pending" || ConTaskValues[i][6] == "04 Retry") && 
                ConTaskValues[i][0] === "DWO") {
              FirstPending = i;
              onlyDWOProcess=true;
              BATCH_SIZE=999;
             break;
           }
         }

         //Procesar en lotes más pequeños
         for (var iTask = FirstPending; iTask < ConTaskEnd - 1; iTask += BATCH_SIZE) {
           // Verificar tiempo límite con más margen
           if (Date.now() - START_TIME > MAX_EXECUTION_TIME) {
             if(verboseFlag === true) {
               console.log('Tiempo máximo de ejecución alcanzado - deteniendo proceso');
             }
             // Actualizar el estado antes de salir
             var nextRun = new Date();
             nextRun.setMinutes(nextRun.getMinutes() + 2);
             ConControl.getRange(2,2).setValue(nextRun);
             ConControl.getRange(2,11).setValue(Utilities.formatDate(new Date(), timezone, timestamp_format));
             return;
           }
           
           // Obtener el rango del lote actual y filtrarlo si es necesario
           let currentBatch = ConTaskValues.slice(iTask, iTask + BATCH_SIZE);
           if (onlyDWOProcess) {
             // Filtrar para incluir solo casos DWO
             currentBatch = currentBatch.filter(task => task[0] === "DWO");
           }

           // Procesar el lote ordenado
           for (const task of currentBatch) {
             // Verificar tiempo de ejecución
             if (Date.now() - START_TIME > MAX_EXECUTION_TIME) {
               if(verboseFlag === true) {
                 console.log('Maximum execution time reached, stopping process');
               }
               break;
             }
             
             const j = ConTaskValues.indexOf(task);
             
             // Check for pending / Continue if 02 Incorporated or 03 Inhibited
             if( ConTaskValues[j][6]!="01 Pending" && ConTaskValues[j][6]!="04 Retry") {continue;}
         
             // Saves track info
             ConTask.getRange(j + 2,8).setValue(now); 

             // If repeated key, discarded
             if((j >0 && j - 1 >= FirstPending) || ConTaskValues[j][1]=="") {
               if((ConTaskValues[j][0]==ConTaskValues[j - 1][0] && ConTaskValues[j][1]==ConTaskValues[j - 1][1] && ConTaskValues[j - 1][7]=="") || ConTaskValues[j][1]=="") {
                 ConTask.getRange(j + 2,7).setValue("05 Discarded");
                 continue;
               }
             }

             var forcedStringKey = ConTaskValues[j][1].toString();

             //Current case
             var furtherAction = DataLoad(ConTaskValues[j][3],ConTaskValues[j][0], forcedStringKey);

             if(furtherAction=="Discarded"){
               ConTask.getRange(j + 2,7).setValue("05 Discarded");
               continue;
             }

             // If current = 04 Retry, save in comment
             var retryNumber = scriptProperties.getProperty('retryNumber');
             if( ConTaskValues[j][6]==="04 Retry" || (retryNumber != null && retryNumber != "" )) {
               // Previous error
               retryNumber = parseInt(retryNumber);
               var errorMSG = scriptProperties.getProperty('errorMSG');
               if (isNaN(retryNumber)) {retryNumber = "1"; scriptProperties.setProperty('retryNumber', retryNumber);errorMSG="";};
               scriptProperties.setProperty('retryNumber', retryNumber);
               ConTask.getRange(j + 2,9).setValue("04 Retry ("+retryNumber+") "+errorMSG);
             } else {
               // Mark as 04 Retry just in halt case
               ConTask.getRange(j + 2,7).setValue("04 Retry");
             };

             if(verboseFlag === true) {
               console.log("Process: "+j+" // "+ConTaskValues[j][3]+" // "+ConTaskValues[j][0]+" // Key: "+ConTaskValues[j][1]+ "//"+ ConTaskValues[j][8] )
             }

             // PROCESS BEGIN
             var variationResult = ObtainVariation(ConTaskValues[j]);

             if( variationResult["variationStatus"] === "unchanged" ) {
               ConTask.getRange(j + 2,7).setValue("06 Unchanged");
               continue;
             } else if( variationResult["variationStatus"] === "source missed key" ) {
               ConTask.getRange(j + 2,7).setValue("07 Source missed key");
               continue;
             } else {
               //Recording process
               if(variationResult["logAddHTML"]!="" && (descreetFlag==false || ConTaskValues[j][5]!="appsheet@mediaaccesscompany.com")) {
                 //Log
                 logSheet.getRange(logSheet.getLastRow() + 1, 1, 1, 7).setValues([
                   [
                     ConTaskValues[j][1],
                     ConTaskValues[j][2],
                     ConTaskValues[j][3]+" / "+ConTaskValues[j][0],  
                     ConTaskValues[j][5],
                     variationResult["logAddHTML"],
                     variationResult["logAddPlain"],
                     variationResult["variationCode"]
                   ]
                 ]);

                 //Source update Log
                 if(variationResult["sourceData"][taskColLog]!="" && taskColLog >= 0) {
                   let aux = variationResult["sourceData"][taskColLog];
                   sourceSheet.getRange(sourceRow+2,taskColLog+1).setValue(aux);
                 }
                 // If comment
                 if(variationResult["cleanComment"] > -1) {sourceSheet.getRange(sourceRow+2,variationResult["cleanComment"]+1).setValue("");};
               }
               let aux = [variationResult["sourceData"]];
               
               if( variationResult["variationStatus"] === "append" ) {
                 //New witness
                 witnessSheet.appendRow(aux[0]);
               } else {
                 //Witness overwrite with Source
                 witnessSheet.getRange(witnessRow+2,1,1, aux[0].length).setValues(aux);
               //            witnessSheet.getRange(witnessRow+2,1,1, witnessSheet.getLastColumn()).setValues(aux);
               }

               //DWO Status changed
               ConTask.getRange(j + 2,7).setValue("02 Incorporated");
             }
             
           }
           
           // Añadir pausa más corta entre lotes para optimizar el tiempo
           Utilities.sleep(500);
           
           // Forzar flush más frecuente
           if (iTask % (BATCH_SIZE * 2) === 0) {
             SpreadsheetApp.flush();
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
     
     // Asegurarse de actualizar el estado en CON-Control para indicar que el proceso no está bloqueado
     try {
       if (ConControl) {
         ConControl.getRange(2,11).setValue(Utilities.formatDate(new Date(), timezone, timestamp_format));
       }
     } catch (updateError) {
       console.error('Error updating control status: ' + updateError.toString());
     }
     
     // Propagar el error
     throw e;
   } finally {
     // Asegurarse de que el lock siempre se libere, incluso en caso de error
     try {
       if (lock && lock.hasLock()) {
         SpreadsheetApp.flush();
         lock.releaseLock();
         console.log('Lock released successfully');
       }
     } catch (lockError) {
       console.error('Error releasing lock: ' + lockError.toString());
     }
   }
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
  try {
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

    // Validar que labelNDX y labelValues existan y estén inicializados
    if (!labelNDX || !labelValues) {
      console.error(`Labels no inicializados para ${sheetNameAux}`);
      actionReturn = "Discarded";
      return actionReturn;
    }

    // Obtain label data con validación
    labelRow = labelNDX.indexOf(sheetNameAux, 0);
    
    if (labelRow === -1) {
      if(verboseFlag === true) {
        console.log(`No se encontró configuración de labels para ${sheetNameAux}`);
      }
      // Usar configuración por defecto para hojas sin labels específicos
      taskLabelNames = [];
      taskLabelActions = [];
      taskColKey = -1;
      taskColUser = -1;
      taskColChange = -1;
      taskColLog = -1;
    } else {
      // Asegurarse de que existan los datos necesarios
      if (labelValues[labelRow] && labelValues[labelRow + 1]) {
        taskLabelNames = labelValues[labelRow];
        taskLabelActions = labelValues[labelRow + 1];
        
        taskColKey = taskLabelActions.indexOf("K") - 1;
        taskColUser = taskLabelNames.indexOf("Last user", 0) - 1; 
        taskColChange = taskLabelNames.indexOf("Last change", 0) - 1;
        taskColLog = taskLabelNames.indexOf("LogTrack", 0) - 1;
      } else {
        console.error(`Datos de label incompletos para ${sheetNameAux}`);
        // Usar valores por defecto
        taskLabelNames = [];
        taskLabelActions = [];
        taskColKey = -1;
        taskColUser = -1;
        taskColChange = -1;
        taskColLog = -1;
      }
    }

    if(taskColLog<0){
      var taskColLogMsg="No column log present";
    } else {
      var taskColLogMsg=taskColChange.toString();
    }
    
  
    if(verboseFlag === true) {
      console.log("Sheet: " + sheetNameAux + " | taskColLog: " + taskColLogMsg + " | labelRow: " + labelRow);
    }
    
    var auxLabel = labelValues[labelRow + 2][0];

    //Check if previous call was equal 
    if(ssPrevious === ssAux && sheetNamePrevious === sheetNameAux) {return actionReturn;}
    if( ssAux === "DubAppActive01") {
      ssAux2 = "DubAppTotal01";
    } else {
      if(auxLabel == "[Split]") {
        let auxRow = sourceNDX.indexOf(keyAux);
        // Buscar la configuración de la hoja en relacionesHojas
        let dwoColLetter = relacionesHojas[sheetNameAux].dwoCol;
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
             return dwoRow !== -1 && totalDWOValues[dwoRow][58] === "(01) On track: DWO";
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
    
  } catch (error) {
    console.error(`Error en DataLoad para ${sheetNameAux}: ${error.toString()}`);
    console.error('Stack: ' + error.stack);
    return "Discarded";
  }
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
   // Exists in witness
   var witnessData = witnessValues[witnessRow]; variationAux["witnessData"] = witnessValues[witnessRow]; variationAux["witnessRow"]=witnessRow;

   //Halt if equal
   if(sourceData == witnessData) {return variationAux;};

   // Loop for variations
   for(var caseAux = 0; caseAux < sourceData.length; ++caseAux){
     witnessLog = ""; sourceLog = "";

     if(caseAux == taskColChange) {
       caseAction="DT";
     } else {
       var caseAction = taskLabelActions[caseAux+1];
     } 
     // Skip if X or Key
     if(caseAction=="K" || caseAction=="X") {continue};

     //Check if value changed
     // If Date or Datetime format
     if(caseAction.substring(0,1) == 'D') {
       if(caseAction == "DT" || caseAux == taskColChange) {
         // Datetime
         sourceAux = dateHandle(sourceData[caseAux], timezone, timestamp_format);
         witnessAux = dateHandle(witnessData[caseAux], timezone, timestamp_format);
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
       sourceAux = dateHandle(sourceData[taskColChange], timezone, "yyyy-MM-dd");
       witnessAux = dateHandle(witnessData[taskColChange], timezone, "yyyy-MM-dd");
       if (sourceAux != witnessAux) {      
         sourceAux = dateHandle(sourceData[taskColChange], timezone, timestamp_format);
       } else {
         sourceAux = dateHandle(sourceData[taskColChange], timezone, "HH:mm:ss");
       }
       variationAux["logAddHTML"]= auxHTML +"<small>"+ sourceAux+ "</small></b>" + variationAux["logAddHTML"];
       variationAux["logAddPlain"]= auxPlain + sourceAux +" // " + variationAux["logAddPlain"] + "\n";
   } else {
     sourceAux = dateHandle(sourceData[taskColChange], timezone, timestamp_format);
     variationAux["logAddHTML"]= auxHTML +"<small>"+ sourceAux + "</small></b><li>";
     variationAux["logAddPlain"]= auxPlain + sourceAux + "\n";
   }

//console.log(variationAux["logAddHTML"]); /*Borrar*/

   //Reset comment
   if( variationAux["cleanComment"] > -1 ) {sourceData[variationAux["cleanComment"]]=""};
   //Add to previous log
   if(taskColLog>0){
    sourceData[taskColLog]=sourceData[taskColLog]+variationAux["logAddHTML"];
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
