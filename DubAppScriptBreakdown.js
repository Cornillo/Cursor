/**
 * Árbol de llamadas de funciones y descripciones:
 * 
 * Main Execution Flow:
 * 1. ScriptUpload() - Función principal de procesamiento
 *    ├── 2. initializeGlobals() - Inicializa variables globales
 *    ├── 3. OpenSheet() - Abre y carga hojas de cálculo
 *    ├── 4. ScriptBreakdown() - Procesa el script subido
 *    │    ├── 5. ExtractDialogLine() - Extrae diálogos del script
 *    │    └── 6. ExtractCharacter() - Extrae personajes del script
 *    └── 7. schedulePDFWitness() - Programa la generación del PDF
 *         └── 8. asyncPDFWitness() - Ejecuta la generación del PDF
 *              └── 9. PDFWitness() - Genera el PDF con el desglose
 * 
 * Helper Functions:
 * - OpenSht() - Abre y filtra hojas de cálculo
 * - String2Seconds() - Convierte timecode a segundos
 * - Time2Seconds() - Convierte tiempo a segundos
 * - Time2String() - Convierte tiempo a string formateado
 * - Seconds2String() - Convierte segundos a timecode
 * - CharacterName() - Normaliza nombres de personajes
 * - LoopNumber() - Calcula número de loop según timecode
 * - ClearRow() - Limpia filas en hojas de cálculo
 * - processControlChanges() - Procesa cambios pendientes
 * - cleanupResources() - Limpia recursos temporales
 * - logOperationStatus() - Registra estado de operaciones
 * - sendErrorNotification() - Envía notificaciones de error
 * 
 * API Endpoints:
 * - doPost(e) - Maneja solicitudes POST
 * - doGet(e) - Maneja solicitudes GET y consulta de estado
 * - getUserStatus() - Obtiene estado de procesamiento
 */

//Global variablesdata*
var sheet = "";
var data = "";
var script = [];
var characters = [];
var charactersNDX = [];
var dwoCharacter;
var dwoCharacterData;
var dwoCharacterNDX;
var dwoCharacterProduction;
var dwoCharacterProductionData;
var dwoCharacterProductionNDX;
var dwoCharacterProductionNDX2;
var dwoCharacterProductionFiltered;
var loopCharacter = [];
var recordedCharacter = [];
var inhibited = ["MAIN TITLE", "GRAPHICS INSERTS", "PRINCIPAL PHOTOGRAPHY", "NONE", "GRAFICA", "BURNEDIN SUBS", "EPISODIC TITLE"];
var auxSheet;
var auxValues;
var auxFilteredValues;
var auxNDX;
var auxNDX2;
var auxRow;
var auxSheetDWO_Files;
var auxCaseDWO_Files;
var auxRowDWO_Files;
var auxLabel;
var auxProject;
var characterRevisionFlag = false;
var timecodeOutRevisionFlag = false;
var firstEventSeconds = 99999999;
var rateSecondsWord = 1.6;
var ssActive;
var eventID;
var appRepository = "Uploaded/";
var fasttrack = false;
var verboseFlag = false;
// GDoc template file ID in G drive
const templateFileID = '1YdfzaVjOunoCkvNf1Y8Agp6XDgF1o_Nm9Glu0T91msg';
const TEMPLATE_MAIL_ID = "1V6j-cQCpSHBNGfEkn1LrGyv9ccrOejb9v2QMhI7oVTI";
const BCC_EMAIL = "appsheet@mediaaccesscompany.com";
var isNow;
const allIDs = databaseID.getID();
const sheetID = allIDs["activeID"];
const folderId = allIDs["uploaded"];
const controlID = allIDs["controlID"];
const folderWitnessId = allIDs["loopsWitness"];
var auxVersion;
var channelID;
var productionID;
var file_ID;
var controlSheet; // Nueva variable global para CON-TaskCurrent
var control2add = [];

// Obtener configuración de timezone y formato de timestamp desde databaseID
const TIMEZONE = allIDs.timezone;
const TIMESTAMP_FORMAT = allIDs.timestamp_format;

const sheetCache = {
	initialized: false,
	sheets: new Map()
};

// Mover estas variables al objeto de configuración global
const CONFIG = {
	verboseFlag: false,
	fasttrack: false,
	appRepository: "Uploaded/",
	templateFileID: '1YdfzaVjOunoCkvNf1Y8Agp6XDgF1o_Nm9Glu0T91msg',
	TEMPLATE_MAIL_ID: "1V6j-cQCpSHBNGfEkn1LrGyv9ccrOejb9v2QMhI7oVTI",
	BCC_EMAIL: "appsheet@mediaaccesscompany.com"
};

/**
 * Función para logging condicional basado en verboseFlag
 * Solo imprime mensajes si verboseFlag es true o si isImportant es true
 */
function conditionalLog(message, isImportant = false) {
    // Siempre mostrar mensajes marcados como importantes, independientemente de verboseFlag
    if (isImportant || verboseFlag) {
        console.log(message);
    }
}

function initializeGlobals() {
	// Variables globales que necesitan ser accesibles en todo el script
	if (typeof verboseFlag === 'undefined') verboseFlag = CONFIG.verboseFlag;
	if (typeof fasttrack === 'undefined') fasttrack = CONFIG.fasttrack;
	if (typeof appRepository === 'undefined') appRepository = CONFIG.appRepository;
}

function call() {
	ScriptUpload(
		"DCVI",
		"5A1F389B-E479-49A1-9F4C-87FECF8E3156",
		"pBVMnTTl",
		"appsheet@mediaaccesscompany.com",
		"F613c3753",
		"DCVI - Big City Greens - S4 - Episode F101 - Dialogue Translation (only finals)"
	);
	}

function call3() {
	PDFWitness("Netflix - Cheat  Unfinished Business - S01 - Episode 102 - ARS conform", "8d03b6f1");
	}

function call2() {
ScriptUpload(
	"Netflix",
	"B1B99441-06B2-4CEF-A0C9-7ED2CB3E825B",
	"bMO8YHUz",
	"appsheet@mediaaccesscompany.com",
	"ac515e22",
	"Netflix - Another Test Script"
);
}

function call4() {
	ScriptUpload(
		"Netflix",
		"0CDFC5B7-A182-4A3F-9347-C3651DA26DE9",
		"c618b02e",
		"appsheet@mediaaccesscompany.com",
		"2fa5f372",
		"Dabba Cartel - S01 Trailer - Teaser (Series) Teaser 2"
	);
	}

/**
* Función principal de carga de scripts
*/
function ScriptUpload(channelID, projectID, productionID, userID, file_ID, pdfname) {
    console.log(`Iniciando proceso inmediato: ${channelID} / ${projectID} / ${productionID} / ${userID} / ${file_ID}`);
    
    try {
        // Abrir la hoja de cálculo y actualizar el estado a "En proceso"
        ssActive = SpreadsheetApp.openById(sheetID);
        isNow = Utilities.formatDate(new Date(), TIMEZONE, TIMESTAMP_FORMAT);
        
        // Load DWO_FIles para actualizar el estado
        OpenSheet("DWO_Files", 1, file_ID, 0, ssActive);
        var auxSheetDWO_Files_temp = auxSheet;
        var auxRowDWO_Files_temp = auxRow + 1;
        auxCaseDWO_Files = auxFilteredValues;  
        
        // Log estado actual del archivo
        conditionalLog(`TRAZABILIDAD - Estado actual antes del procesamiento: ${auxCaseDWO_Files[0][11]}, Version: ${auxCaseDWO_Files[0][14]}, EventID: ${auxCaseDWO_Files[0][3]}`);
        
        // NO actualizar el estado al principio, solo leer el estado existente
        // El estado se actualizará al final del procesamiento
        
        // Crear una ejecución separada para iniciar el proceso principal
        // Usar PropertiesService en lugar de CacheService para mayor confiabilidad
        const props = PropertiesService.getScriptProperties();
        const jobKey = "ScriptUpload_" + file_ID;
        
        // Almacenar los parámetros en PropertiesService
        props.setProperty(jobKey, JSON.stringify({
            channelID: channelID,
            projectID: projectID,
            productionID: productionID,
            userID: userID,
            file_ID: file_ID,
            pdfname: pdfname,
            startTime: isNow
        }));
        
        // Configurar un trigger que se ejecute inmediatamente
        ScriptApp.newTrigger('processScriptUploadFromCache')
            .timeBased()
            .after(1000) // 1 segundo de retraso mínimo
            .create();
        
        conditionalLog('Proceso iniciado en background, script retornando inmediatamente', true);
        
        // Retornar inmediatamente para que AppSheet no espere
        return auxCaseDWO_Files[0][11]; // Retornar el estado actual sin modificarlo
        
    } catch (error) {
        console.error('Error al iniciar el proceso:', error);
        sendErrorNotification(error, 'ScriptUpload', {
            channelID: channelID,
            projectID: projectID,
            productionID: productionID,
            userID: userID,
            file_ID: file_ID
        });
        return "(09) Failed: DWOFiles - Error: " + error.message;
    }
}

/**
* Función que recupera datos de las propiedades del script y ejecuta el procesamiento
*/
function processScriptUploadFromCache(e) {
    // Limpiar el trigger que ejecutó esta función
    try {
        // Obtener todos los triggers y buscar el que coincide con el ID actual
        const allTriggers = ScriptApp.getProjectTriggers();
        for (let i = 0; i < allTriggers.length; i++) {
            // Verificar si es el trigger actual
            if (allTriggers[i].getUniqueId() === e.triggerUid) {
                // Eliminar el trigger usando el objeto correcto, no el ID
                ScriptApp.deleteTrigger(allTriggers[i]);
                break;
            }
        }
    } catch (error) {
        console.error("Error al eliminar trigger:", error);
    }
    
    // Buscar todos los trabajos pendientes en las propiedades
    const props = PropertiesService.getScriptProperties();
    const allProps = props.getProperties();
    
    try {
        // Obtener todas las claves que comienzan con ScriptUpload_
        const pendingJobs = Object.keys(allProps)
            .filter(key => key.startsWith("ScriptUpload_"));
            
        if (pendingJobs.length === 0) {
            conditionalLog("No hay trabajos pendientes en las propiedades");
            return;
        }
        
        conditionalLog(`Se encontraron ${pendingJobs.length} trabajos pendientes`);
        
        // Procesar el primer trabajo pendiente
        const firstJobKey = pendingJobs[0];
        try {
            const params = JSON.parse(allProps[firstJobKey]);
            
            // Eliminar este trabajo de las propiedades para que no se procese nuevamente
            props.deleteProperty(firstJobKey);
            
            conditionalLog(`Procesando trabajo: ${firstJobKey} - Parámetros: ${JSON.stringify(params)}`);
            
            // Procesar este trabajo
            processScriptUpload(params);
            
            conditionalLog("Trabajo procesado exitosamente: " + firstJobKey);
            
        } catch (error) {
            console.error('Error al procesar trabajo desde propiedades:', error);
            props.deleteProperty(firstJobKey); // Eliminar trabajo con error
        }
        
        // Verificar de nuevo todas las claves tras el procesamiento
        const remainingJobsKeys = Object.keys(props.getProperties())
            .filter(key => key.startsWith("ScriptUpload_"));
        
        // Si hay más trabajos pendientes, programar otra ejecución
        if (remainingJobsKeys.length > 0) {
            conditionalLog(`Quedan ${remainingJobsKeys.length} trabajos pendientes, programando próxima ejecución`);
            ScriptApp.newTrigger('processScriptUploadFromCache')
                .timeBased()
                .after(10000) // 10 segundos entre procesamiento de trabajos
                .create();
        } else {
            conditionalLog("No quedan más trabajos pendientes");
        }
    } catch (propError) {
        console.error("Error al manejar las propiedades:", propError);
        // Asegurar que siempre se programe el siguiente trigger si hay un error
        ScriptApp.newTrigger('processScriptUploadFromCache')
            .timeBased()
            .after(30000) // 30 segundos si hubo un error
            .create();
    }
}

/**
* Función que procesa la carga de scripts
*/
function processScriptUpload(params) {
    try {
        const channelID = params.channelID;
        const projectID = params.projectID;
        const productionID = params.productionID;
        const userID = params.userID;
        const file_ID = params.file_ID;
        const pdfname = params.pdfname;
        
        if (!channelID || !projectID || !productionID || !userID || !file_ID) {
            throw new Error('Parámetros incompletos para processScriptUpload');
        }

        conditionalLog(`Ejecutando proceso: ${channelID} / ${projectID} / ${productionID} / ${userID} / ${file_ID}`);

        // Open sheet (de nuevo, porque esta es una ejecución separada)
        ssActive = SpreadsheetApp.openById(sheetID);
        isNow = Utilities.formatDate(new Date(), TIMEZONE, TIMESTAMP_FORMAT);

        // Inicializar variables globales
        initializeGlobals();

        // Load DWO_FIles
        OpenSheet("DWO_Files", 1, file_ID, 0, ssActive);
        auxSheetDWO_Files = auxSheet;
        auxCaseDWO_Files = auxFilteredValues;
        auxRowDWO_Files = auxRow + 1;
        eventID = auxCaseDWO_Files[0][3];
        auxVersion = auxCaseDWO_Files[0][14];
        var auxStatus = auxCaseDWO_Files[0][11];
        var resultAux = "";
        var statusAux = "";

        conditionalLog(`TRAZABILIDAD - Iniciando processScriptUpload con auxStatus: ${auxStatus}, auxVersion: ${auxVersion}, eventID: ${eventID}`);

        //First round
        if (auxStatus === "(01) Pending: DWOFiles") {
            conditionalLog(`TRAZABILIDAD - Entrando en flujo "(01) Pending"`);
            var fileName = auxCaseDWO_Files[0][6].replace(appRepository, "");

            // Obtain script file
            var folder = DriveApp.getFolderById(folderId);
            var files = folder.getFilesByName(fileName);

            if (files.hasNext()) {
                var file = files.next();
                var fileID = file.getId();
                conditionalLog(`TRAZABILIDAD - Iniciando ScriptBreakdown para archivo: ${fileName}, ID: ${fileID}`);
                resultAux = ScriptBreakdown(channelID, projectID, productionID, userID, fileID, file_ID);
                // Determine status
                if (resultAux === "" && (fasttrack || (!characterRevisionFlag && !timecodeOutRevisionFlag))) {
                    statusAux = "(99) Completed: DWOFiles";
                } else {
                    statusAux = characterRevisionFlag && timecodeOutRevisionFlag
                        ? "(04) Check pending / New characters & Timecode out: DWOFiles"
                        : characterRevisionFlag
                            ? "(02) Check pending / New characters: DWOFiles"
                            : timecodeOutRevisionFlag
                                ? "(03) Check pending / Timecode out: DWOFiles"
                                : "(09) Failed: DWOFiles";
                }
                conditionalLog(`TRAZABILIDAD - ScriptBreakdown completado, characterRevisionFlag: ${characterRevisionFlag}, timecodeOutRevisionFlag: ${timecodeOutRevisionFlag}, statusAux: ${statusAux}`);
            } else {
                resultAux = "File not found: " + fileName;
                conditionalLog(resultAux);
                statusAux = "(09) Failed: DWOFiles";
            }

        } else if (auxStatus === "(05) Complete Upload: DWOFiles") {
            conditionalLog(`TRAZABILIDAD - Entrando en flujo "(05) Complete Upload", fasttrack será establecido a true`);
            ReloadDialogLine(file_ID);
            fasttrack = true;
            conditionalLog(`TRAZABILIDAD - Iniciando ScriptBreakdown para Complete Upload con fasttrack: ${fasttrack}`);
            resultAux = ScriptBreakdown(channelID, projectID, productionID, userID, fileID, file_ID);
            statusAux = "(99) Completed: DWOFiles";
            conditionalLog(`TRAZABILIDAD - ScriptBreakdown completado para Complete Upload, fasttrack: ${fasttrack}, characterRevisionFlag: ${characterRevisionFlag}, timecodeOutRevisionFlag: ${timecodeOutRevisionFlag}, statusAux forzado a: ${statusAux}`);
        } else {
            // Si ya está en algún otro estado, no hacer nada
            conditionalLog(`TRAZABILIDAD - Saltando procesamiento porque el estado actual es: ${auxStatus}`);
            return;
        }

        //Reapertura forzada
        //dwoCharacterProduction = ssActive.getSheetByName("DWO_CharacterProduction");

        if (statusAux === "(99) Completed: DWOFiles") {
            conditionalLog(`TRAZABILIDAD - Procesando Characters para estado ${statusAux}, total caracteres: ${characters.length}`);
            //Graba Characters ***
            var currentValue; var updatedValue;
            let rowsToWrite = []; // Array para acumular filas a escribir
            let characterRowsToWrite = []; // Array para acumular filas de DWO_Character
            
            logOperationStatus('InicioProcesamiento', {
                totalCharacters: characters.length,
                version: auxVersion,
                productionID: productionID
            });

            for (let i = 0; i < characters.length; i++) {
                const fila = characters[i]; 
                if (fila[7] === "0") { continue; }
                
                try {
                    if (fila[12] === "Character not present in project: Breakdown_mark") {
                        // Preparar fila para DWO_Character
                        const characterRow = [fila[0], projectID, fila[3], null, null, null, null, null, null, null, fila[8], "(01) Enabled: Generic", fila[10], fila[11], null, null, null, null, null, null, null];
                        characterRowsToWrite.push(characterRow);
                        
                        logOperationStatus('CharacterPrepared', {
                            character: fila[3],
                            characterID: fila[0]
                        });

                        control2add.push({
                            sheet: "DWO_Character",
                            key: fila[0],
                            action: "INSERT_ROW",
                            user: fila[10]
                        });

                        // Preparar fila para DWO_CharacterProduction
                        var auxSpecialAttribute = (auxVersion === "Final version: Script_upload_lite") ? "Final loops added: CharacterProduction_Attributes" : "";
                        // Usar formato consistente para el ID - siempre con guión
                        const productionRow = [productionID + "-" + fila[0], fila[0], productionID, null, null, fila[7], null, null, fila[8], "(01) Recording pending: DWOCharacterProduction", fila[10], fila[11], null, null, null, null, auxSpecialAttribute, null, null, null, null, null, null, projectID];
                        
                        rowsToWrite.push(productionRow);
                        
                        logOperationStatus('ProductionRowPrepared', {
                            character: fila[3],
                            productionID: productionID,
                            rowIndex: rowsToWrite.length
                        });

                        control2add.push({
                            sheet: "DWO_CharacterProduction",
                            key: productionID + "-" + fila[0], // Usar el mismo formato con guión
                            action: "INSERT_ROW",
                            user: fila[10]
                        });
                    } else {
                        // Verifica si ya está creado en DWO_CharacterProduction
                        let aux1 = dwoCharacterProductionNDX.indexOf(fila[2]);
                        // Verificar que caractherID existe y que pertenece a la producción actual
                        while (aux1 !== -1 && dwoCharacterProductionData[aux1][2] !== productionID) {
                            aux1 = dwoCharacterProductionNDX.indexOf(fila[2], aux1 + 1);
                        }
                        
                        if (aux1 === -1) {
                            //No está creado
                            var auxSpecialAttribute = (auxVersion === "Final version: Script_upload_lite") ? "Final loops added: CharacterProduction_Attributes" : "";
                            const newRow = [productionID + "-" + fila[2], fila[2], productionID, null, null, fila[7], null, null, fila[8], "(01) Recording pending: DWOCharacterProduction", fila[10], fila[11], null, null, null, null, auxSpecialAttribute, null, null, null, null, null, null, projectID];
                            rowsToWrite.push(newRow);
                            
                            logOperationStatus('NewCharacterPrepared', {
                                character: fila[3],
                                productionID: productionID,
                                rowIndex: rowsToWrite.length,
                                characterID: fila[2]
                            });

                            control2add.push({
                                sheet: "DWO_CharacterProduction",
                                key: productionID + "-" + fila[2],
                                action: "INSERT_ROW",
                                user: fila[10]
                            });
                        } else {
                            //Está creado en proyecto ***
                            var charProd = dwoCharacterProductionData[aux1];
                            //Versión FINAL 
                            if (auxVersion === "Final version: Script_upload_lite") {
                                //No se grabó la primera vez
                                if(dwoCharacterProductionData[aux1][7]===""){
                                    // Sobre escribe loops
                                    charProd[5] = fila[7]; // Graba en Planned loops
                                    charProd[22] = ""; // Limpia loops Finales
                                    charProd[3] = ""; // Limpia additional
                                    charProd[6] = ""; // Limpia extra citation
                                    charProd[20] = ""; // Limpia additional loops
                                    charProd[21] = ""; // Limpia additional loops
                                } else {
                                    //
                                    if (!charProd[16].includes("Final loops added: CharacterProduction_Attributes")) {
                                        if(charProd[6] === 1) {
                                            // Hay retoma creada
                                            if(charProd[17] === "") {
                                                //No se grabó todavía : sobre escribe
                                                charProd[8] = charProd[8] + "// Final loops (previous: "+ charProd[3] +") overwritted";
                                                charProd[22] = fila[7]; // sobreescribe loops Finales
                                                charProd[3] = fila[7]; // sobreescribe additional loops
                                                charProd[20] = ""; // Limpia additional loops 2
                                                charProd[21] = ""; // Limpia additional loops 3
                                            } else {
                                                //Ya se grabó / Error
                                                charProd[8] = charProd[8] + "// Error: Retake already completed";
                                            }
                                        } else if(charProd[6] > 1) {
                                            // Hay más de una retoma creada / Error
                                            charProd[8] = charProd[8] + "// Error: More than one retake created";
                                        } else {
                                            // No hay retoma creada
                                            charProd[22] = fila[7]; // escribe loops Finales
                                            charProd[3] = fila[7]; // escribe additional loops
                                            charProd[6] = 1; // Crea extra citation
                                            charProd[20] = ""; // Limpia additional loops 2
                                            charProd[21] = ""; // Limpia additional loops 3
                                        }
                                    } else {
                                    //Es reprocesamiento
                                        // Variaron los loops
                                        if(charProd[5] !== fila[7]) {
                                            //No se grabó todavía
                                            if(charProd[17] === "") {
                                                //Sobre escribe loops
                                                charProd[8] = charProd[8] + "// Final loops (previous: "+ charProd[3] +") overwritted";
                                                charProd[3] = fila[7]; // sobreescribe loops Finales
                                                charProd[22] = fila[7]; // sobreescribe loops Finales
                                                charProd[20] = ""; // Limpia additional loops 2
                                                charProd[21] = ""; // Limpia additional loops 3
                                            } else {
                                                //Ya se grabó
                                                charProd[8] = charProd[8] + "// Error: Retake already completed";
                                            }
                                        } else {
                                            continue;
                                        }
                                    }
                                }
                                // Preparar todos los valores a actualizar
                                let currentValue = charProd[16];
                                let attributesToAdd = [
                                    "Check pending: CharacterProduction_Attributes",
                                    "Final loops added: CharacterProduction_Attributes"
                                ];
                                
                                // Construir el string de atributos
                                let attributes = currentValue || "";
                                attributesToAdd.forEach(attr => {
                                    if (!attributes.includes(attr)) {
                                        attributes = attributes ? `${attributes} , ${attr}` : attr;
                                    }
                                });

                                // Actualizar status si es necesario
                                if (charProd[9] === "(04) Dismissed: DWOCharacterProduction") {
                                    charProd[9] = "(01) Recording pending: DWOCharacterProduction";
                                    charProd[8] = charProd[8] + "// Dissmissed case reopened";
                                    const checkPendingAttr = "Check pending: CharacterProduction_Attributes";
                                    
                                    if (!attributes.includes(checkPendingAttr)) {
                                        attributes = attributes ? `${attributes} , ${checkPendingAttr}` : checkPendingAttr;
                                    }
                                }
                                charProd[16]=attributes;
                            } else {
                                //Versión preliminar reprocesada
                                // Variaron los loops
                                if(charProd[5] !== fila[7]) {
                                    //Todavía no se grabó
                                    if(charProd[7] === "") {
                                        //Sobre escribe loops
                                        charProd[5] = fila[7]; // Graba en Planned loops
                                        charProd[22] = ""; // Limpia loops Finales
                                        charProd[3] = ""; // Limpia additional
                                        charProd[6] = ""; // Limpia extra citation
                                        charProd[20] = ""; // Limpia additional loops
                                        charProd[21] = ""; // Limpia additional loops
                                    } else {
                                        //Ya se grabó
                                        charProd[8] = charProd[8] + "// Error: Loops changed ("+fila[7]+ ") and recording already completed";
                                        charProd[16]="Check pending: CharacterProduction_Attributes , "+charProd[16];
                                    }
                                } else {
                                    continue;
                                }
                            }
                        
                            // Agregar a rowsToWrite en lugar de escribir directamente
                            rowsToWrite.push(charProd);
                        }
                    }
                } catch (error) {
                    logOperationStatus('ErrorProcesamiento', {
                        character: fila[3],
                        error: error.toString(),
                        stack: error.stack
                    });
                    throw error;
                }
            }
            
            // Escribir todas las filas acumuladas en DWO_Character
            if (characterRowsToWrite.length > 0) {
                let retryCount = 0;
                const maxRetries = 3;
                let writeSuccess = false;
                
                while (!writeSuccess && retryCount < maxRetries) {
                    try {
                        const startRow = dwoCharacter.getLastRow() + 1;
                        dwoCharacter.getRange(startRow, 1, characterRowsToWrite.length, characterRowsToWrite[0].length)
                            .setValues(characterRowsToWrite);
                        writeSuccess = true;
                        
                        logOperationStatus('EscrituraCharacterExitosa', {
                            filaInicio: startRow,
                            totalFilas: characterRowsToWrite.length,
                            intento: retryCount + 1
                        });
                    } catch (error) {
                        retryCount++;
                        logOperationStatus('ErrorEscrituraCharacter', {
                            intento: retryCount,
                            error: error.toString()
                        });
                        
                        if (retryCount === maxRetries) {
                            throw new Error(`Fallo en escritura de DWO_Character después de ${maxRetries} intentos: ${error}`);
                        }
                        Utilities.sleep(1000 * retryCount);
                    }
                }
            }
            
            // Escribir todas las filas acumuladas en DWO_CharacterProduction
            if (rowsToWrite.length > 0) {
                let retryCount = 0;
                const maxRetries = 3;
                let writeSuccess = false;
                
                while (!writeSuccess && retryCount < maxRetries) {
                    try {
                        const startRow = dwoCharacterProduction.getLastRow() + 1;
                        dwoCharacterProduction.getRange(startRow, 1, rowsToWrite.length, rowsToWrite[0].length)
                            .setValues(rowsToWrite);
                        writeSuccess = true;
                        
                        logOperationStatus('EscrituraProductionExitosa', {
                            filaInicio: startRow,
                            totalFilas: rowsToWrite.length,
                            intento: retryCount + 1
                        });
                    } catch (error) {
                        retryCount++;
                        logOperationStatus('ErrorEscrituraProduction', {
                            intento: retryCount,
                            error: error.toString()
                        });
                        
                        if (retryCount === maxRetries) {
                            throw new Error(`Fallo en escritura de DWO_CharacterProduction después de ${maxRetries} intentos: ${error}`);
                        }
                        Utilities.sleep(1000 * retryCount);
                    }
                }
            }

            // Configurar el trigger asíncrono para PDFWitness
            console.log('TRAZABILIDAD - Intentando programar PDFWitness para:', pdfname, file_ID);
            const scheduled = schedulePDFWitness(pdfname, file_ID);
            
            if (!scheduled) {
                console.error('TRAZABILIDAD - Falló la programación de PDFWitness');
                statusAux = "(09) Failed: DWOFiles";
            } else {
                console.log('TRAZABILIDAD - PDFWitness programado exitosamente');
            }
        }

        //Save output
        //Status y Message en una sola operación
        // Obtener los valores actuales de las columnas O y P desde auxCaseDWO_Files
        // Las columnas O y P corresponden a los índices 14 y 15 en el array (0-based)
        var columnO = auxCaseDWO_Files[0][14]; // Columna O
        var columnP = auxCaseDWO_Files[0][15]; // Columna P

        conditionalLog(`TRAZABILIDAD - Actualizando estado final a: ${statusAux}, mensaje: ${resultAux}`);

        var updateRange = auxSheetDWO_Files.getRange(auxRowDWO_Files + 1, 10, 1, 8);
        var updateValues = [[
            resultAux, // columna 10 (J)
            "", // columna 11 (K)
            statusAux, // columna 12 (L)
            userID, // columna 13 (M)
            isNow, // columna 14 (N)
            columnO, columnP, // columnas 15-16 (O-P) - Preservar valores existentes
            channelID + " / " + projectID + " / " + productionID + " / " + userID + " / " + file_ID + " / " + pdfname// columna 17 (Q)
        ]];
        updateRange.setValues(updateValues);

        // Procesar los cambios acumulados
        processControlChanges();

        // En lugar de retornar statusAux, grabarlo en DWO-Event
        // Buscar en DWO-Event la fila donde columna A tiene el valor de eventID
        OpenSheet("DWO-Event", 0, "", 0, ssActive);
        var dwoEventSheet = auxSheet;
        var dwoEventData = auxValues;
        var filaEncontrada = -1;

        // Buscar la fila con el eventID en la columna A
        for (var i = 0; i < dwoEventData.length; i++) {
            if (dwoEventData[i][0] === eventID) {
                // i+2 porque: +1 por el índice base 0 y +1 por la fila de encabezado en la hoja
                filaEncontrada = i + 2;
                break;
            }
        }

        if (filaEncontrada > 0) {
            // Columna BC es la número 55 (considerando que A es 1, BC sería 55)
            dwoEventSheet.getRange(filaEncontrada, 55).setValue(statusAux);
            conditionalLog(`TRAZABILIDAD - Estado actualizado en DWO-Event para eventID: ${eventID} con valor: ${statusAux} en fila ${filaEncontrada}`);
        } else {
            console.error(`TRAZABILIDAD - No se encontró el eventID ${eventID} en DWO-Event`);
        }
        
        conditionalLog('TRAZABILIDAD - Proceso de ScriptUpload completado exitosamente');
        
    } catch (error) {
        console.error('TRAZABILIDAD - Error en processScriptUpload:', error);
        
        // Intentar actualizar el estado en caso de error, para evitar que quede en "En procesamiento"
        try {
            if (auxSheetDWO_Files && auxRowDWO_Files) {
                var updateRange = auxSheetDWO_Files.getRange(auxRowDWO_Files + 1, 12, 1, 1);
                updateRange.setValue("(09) Failed: DWOFiles - Error: " + error.message.substring(0, 100));
                
                // También actualizar la columna de mensajes
                var msgRange = auxSheetDWO_Files.getRange(auxRowDWO_Files + 1, 10, 1, 1);
                msgRange.setValue("Error: " + error.message.substring(0, 255));
                
                conditionalLog(`TRAZABILIDAD - Actualizado estado a error: (09) Failed: DWOFiles - Error: ${error.message.substring(0, 100)}`);
                // NO actualizar DWO-Event en caso de error, como en la versión original
            }
        } catch (err2) {
            console.error('TRAZABILIDAD - Error al actualizar estado de error:', err2);
        }
        
        sendErrorNotification(error, 'processScriptUpload', params);
    }
}

function ScriptBreakdown(channelID, projectID, productionID, userID, fileID, file_ID, pdfname) {

// Load label
OpenSht("DWO", 0, 2, projectID, 0, ssActive);
if (auxFilteredValues[0][6] === "") {
	auxLabel = auxFilteredValues[0][7];
} else {
	auxLabel = auxFilteredValues[0][0] + " - " + auxFilteredValues[0][6];
}

OpenSht("DWO-Production", 0, 1, productionID, 0, ssActive);
auxLabel = auxLabel + " - " + auxFilteredValues[0][3]
auxLabel = auxLabel.replace(/:/g, " -") + " - loops";

// Open production characters
OpenSht("DWO_CharacterProduction", 2, 3, productionID, 2, ssActive);
dwoCharacterProduction = auxSheet;
dwoCharacterProductionData = auxValues;
dwoCharacterProductionFiltered = auxFilteredValues;
dwoCharacterProductionNDX = auxNDX;
dwoCharacterProductionNDX2 = auxNDX2;

// Read according to channel and leave it ready
if (script.length === 0) {
	var resultAux = ExtractDialogLine(channelID, fileID, userID);
} else {
	var resultAux = "";
}

if (resultAux === "") {
  var scriptBuffer = [];
  // Extract characters
  ExtractCharacter(projectID, file_ID, userID);

  var loopAux;
  // First pass of script
  script.forEach(function (item) {
	var auxIn = LoopNumber(item[0]);
	var auxOut;

	if (item[5] != "Dismissed") {
		if (item[4]) {
			auxOut = "";
			loopAux = auxIn.toString();
		} else {
			auxOut = LoopNumber(item[1]); loopAux = "";
			// Count timecodes between In and Out
			for (var i2 = auxIn; i2 <= auxOut; i2++) {
				loopAux = loopAux === "" ? i2.toString() : loopAux + ", " + i2.toString();
			}
		}

		AddLoops2Character(item[2], loopAux, item[5]);
	}

	//Logger.log('Boya 1 / Project ID: ' + projectID);

	//    scriptAux.push([item[6].substring(0, 8), item[7].substring(0, 8), item[2], item[3], item[5]]);
	scriptBuffer.push([file_ID + "-S" + Rand(8),
	file_ID,
	item[2],
	item[6].substring(0, 8),
	item[7].substring(0, 8),
	item[3].toString(),
		"",
	item[9],
	loopAux,
	item[8],
		"",
		"(01) Enabled: Generic",
	userID,
	isNow,
	item[5],
	item[10],
	projectID])

  });

  // Characters persistence
  OpenSheet("DWO_FilesCharacter", 0, "", 0, ssActive);
  var sheetDWO_FilesCharacter = auxSheet;

  // Clear previous values
  if (auxRow != -1) {
	  ClearRow(sheetDWO_FilesCharacter, 1, file_ID);
  }

  sheetDWO_FilesCharacter.getRange(sheetDWO_FilesCharacter.getLastRow() + 1, 1, characters.length, characters[0].length).setValues(characters);

  // Script persistence
  OpenSheet("DWO_FilesLines", 0, "", 0, ssActive);
  sheetDWO_FilesLines = auxSheet;

  // Clear previous values
  if (auxRow != -1) {
	  ClearRow(sheetDWO_FilesLines, 1, file_ID);
  }

  sheetDWO_FilesLines.getRange(sheetDWO_FilesLines.getLastRow() + 1, 1, scriptBuffer.length, scriptBuffer[0].length).setValues(scriptBuffer);

}
return resultAux;
}

function ReloadDialogLine(file_ID) {
//Second round // Load values
//Characters
OpenSheet("DWO_FilesCharacter", 2, file_ID, 0, ssActive);
var dataCharacter = auxFilteredValues;
var dataCharNDX = dataCharacter.map(function (r) { return r[3].toString(); });
// Script 
OpenSheet("DWO_FilesLines", 2, file_ID, 0, ssActive);
var dataLines = auxFilteredValues;
var aux3; var loopAux;
//Loop
for (var i = 0; i < dataLines.length; i++) {
	//Checks if character is replaced
	aux3 = dataCharNDX.indexOf(dataLines[i][2]);
	if (aux3 != -1) {
		if (dataCharacter[aux3][4] != "") {
			for (var j = 0; j < dataCharacter.length; j++) {
				if (dataCharacter[j][0] == dataCharacter[aux3][4]) {
					dataLines[i][2] = dataCharacter[j][3];
				}
			}
		}
	}

	script.push([Time2Seconds(dataLines[i][3]), Time2Seconds(dataLines[i][4]), dataLines[i][2], dataLines[i][5], true, dataLines[i][14], Time2String(dataLines[i][3]), Time2String(dataLines[i][4]), "", "", dataLines[i][15]]);
}
}


function ExtractDialogLine(channelID, fileID, userID) {

var timecodeIn; var timecodeOut; var flagStatus; var fila; var loopAux; var rowAux; var auxCell; var source; var dialogue; var comment;

const fileAux = DriveApp.getFileById(fileID);
var mimeType = fileAux.getMimeType();
//-----------------------------------------------------------//
if (mimeType === "application/vnd.openxmlformats-officedocument.wordprocessingml.document" || mimeType === "application/msword") {

	var blob = fileAux.getBlob();
	var fileTemp = Drive.Files.insert({}, blob, { convert: true });
	var id = fileTemp["id"];
	var doc = DocumentApp.openById(id);
	var body = doc.getBody()

	// Encontrar la primera tabla en el documento
	var tables = body.getTables();
	if (tables.length === 0) {
		throw new Error("No se encontró ninguna tabla en el documento.");
	}

	data = tables[0];
	timecodeOut = "";
	flagStatus = true; 
	// Recorre los data y carga la script
	for (var i4 = 0; i4 < data.getNumRows(); i4++) {
		rowAux = data.getRow(i4);
		fila = [];
		for (var j = 0; j < rowAux.getNumCells(); j++) {
			auxCell = rowAux.getCell(j);
			fila.push(auxCell.getText());
		}
		if (fila[0] === null || fila[0] === "") { 
			continue;
		}

		fila[0] = fila[0].replace(/\n/g, '').trim();
		timecodeIn = String2Seconds(fila[0].substring(0, 8));

		//comment = (fila[2].match(/\[(.*?)\]/g) ?? []).join(' ');

		var dialogue = fila[2];
		if (inhibited.includes(fila[1]?.toUpperCase()) || fila[1] === "") {
			loopAux = "Dismissed";
		} else {
			if (auxVersion === "Final version: Script_upload_lite" && dialogue.includes("[+]")) {
				loopAux = "Added"
			} else {
				loopAux = "";
			}
		}
		var source = CharacterName(fila[1]);
		dialogue = dialogue.replace(/\n|\[[^\]]*\]/g, '').trim();

		// Si después de quitar los corchetes y su contenido no queda texto, marcar como Dismissed
		if (!dialogue) {
			loopAux = "Dismissed";
		}

		script.push([timecodeIn, timecodeOut, source, dialogue, flagStatus, loopAux, fila[0], timecodeOut, "", "", fila[2]]);
	}

	EstimatedTimecode();

	// Borrar el archivo de Google Docs creado
	Drive.Files.remove(fileTemp.id);
	//-----------------------------------------------------------//
} else if (mimeType === 'application/vnd.ms-excel' || 
           mimeType === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
           mimeType === 'application/vnd.oasis.opendocument.spreadsheet') {
    fasttrack = true;

    //Convert XLSX/ODS a GSheet
    var xlsxBlob = fileAux.getBlob();
    let auxTitle = "archivo temporal DubApp borrar por favor _" + userID;

    // Configura los parámetros para la conversión
    var resource = {
        title: auxTitle,
        mimeType: MimeType.GOOGLE_SHEETS
    };

    // Crea el archivo Google Sheets a partir del blob del archivo .xlsx
    var fileTemp = Drive.Files.insert(resource, xlsxBlob, { convert: true });

    // Abre la sheet de cálculo y selecciona la sheet activa
    var sheet = SpreadsheetApp.openById(fileTemp.id);

    // Obtiene todos los data de la sheet
    var data = sheet.getDataRange().getValues();

    if (data[0][0] !== "IN-TIMECODE" || data[0][1] !== "OUT-TIMECODE" || data[0][2] !== "SOURCE" || data[0][3] !== "TRANSCRIPTION") {
        return "The uploaded file does not contain the correct headers. Please retry"
    }

    // Recorre los data y carga la script
    for (var i4 = 1; i4 < data.length; i4++) { // Empieza en 1 para saltar la fila de encabezado
        fila = data[i4];
        if (fila[0] === null || fila[0] === "") { 
            continue;
        }
        timecodeIn = String2Seconds(fila[0].substring(0, 8));
        var dialogue = fila[4];
        if (inhibited.includes(fila[2]?.toUpperCase()) || fila[2] === "") {
            loopAux = "Dismissed";
        } else {
            if (auxVersion === "Final version: Script_upload_lite" && dialogue.includes("[+]")) {
                loopAux = "Added"
            } else {
                loopAux = "";
            }
        }
        var source = CharacterName(fila[2]);
        dialogue = dialogue.replace(/\n|\[[^\]]*\]/g, '').trim();

        // Si después de quitar los corchetes y su contenido no queda texto, marcar como Dismissed
        if (!dialogue) {
            loopAux = "Dismissed";
        }

        if (loopAux != "Dismissed") { 
			if (fila[1] === null || fila[1] === "") {
				timecodeOut = "";
				flagStatus = true; 
			} else {
				timecodeOut = String2Seconds(fila[1].substring(0, 8));
				flagStatus = false;
			}
		}

        script.push([timecodeIn, timecodeOut, source, dialogue, flagStatus, loopAux, fila[0], fila[1], "", "", fila[4]]);
    }

    EstimatedTimecode();

    Drive.Files.remove(fileTemp.id);
} else {
	return "The uploaded file is in an unrecognized format (" + mimeType + "). Please retry";
}
return "";
}

function EstimatedTimecode() {

var fila; var auxLoopTCIn; var auxLoopTCOut; var checkAux;
for (var k = 0; k < script.length; k++) {
	fila = script[k];
	if (fila[5] === "Dismissed") { continue; }

	var palabras = fila[3].split(" ").length; // Cuenta las palabras en la cuarta columna
	var segundosAAgregar = parseInt(palabras / 1.6); // Calcula los segundos a agregar      
	var totalSegundos = fila[0] + segundosAAgregar; // Calcula el total de segundos
	var bypass = false;

	//Compara con siguiente entrada de diálogo
	if (k + 1 < script.length) {
		for (var m = k + 1; m < script.length; m++) {
			if (script[m][5] === "Dismissed") { continue; }
			if (script[m][0] <= totalSegundos) {
				totalSegundos = script[m][0]; bypass = true;
			} else {
				if (script[m][0] - totalSegundos < 5) {
					bypass = true;
				}
			}
			break;
		}
	}

	//Chequea condiciones para supervisión
	auxLoopTCIn = LoopNumber(script[k][0]);
	auxLoopTCOut = LoopNumber(totalSegundos);

	//Condiciones de check
	checkAux = "";

	if (auxLoopTCIn != auxLoopTCOut) {
		//Busca próxima intervención del personaje
		for (var m = k + 1; m < script.length; m++) {
			if (script[k][2] === script[m][2]) {
				if (LoopNumber(script[m][0]) != auxLoopTCOut && bypass === false) {
					checkAux = "Check";
					if (!fila[1]) {timecodeOutRevisionFlag = true};
				}
				break
			}
		}
		if (m === script.length && bypass === false) {
			checkAux = "Check";
			if (!fila[1]) {timecodeOutRevisionFlag = true};
		}
	}

	var nuevasHoras = Math.floor(totalSegundos / 3600);
	totalSegundos %= 3600;
	var nuevosMinutos = Math.floor(totalSegundos / 60);
	var nuevosSegundos = Math.floor(totalSegundos % 60);

	var nuevoTimecode =
		(nuevasHoras < 10 ? "0" : "") + nuevasHoras + ":" +
		(nuevosMinutos < 10 ? "0" : "") + nuevosMinutos + ":" +
		(nuevosSegundos < 10 ? "0" : "") + nuevosSegundos;

	script[k][8] = nuevoTimecode; // Almacena el nuevo timecode en la novena columna
	script[k][9] = checkAux;
}

}

function AddLoops2Character(characterAux, loopsAux, addedAux) {
//If empty, return
if (loopsAux === "") { return }
//Split character
var sources = characterAux.split("/"); // Divide el contenido por "/"
var loops = loopsAux.split(", ");
var loopsCharacterAux; var aux2;

for (var k = 0; k < sources.length; k++) {
	var individualSource = sources[k].trim(); // Elimina espacios en blanco alrededor
	let characterRow = charactersNDX.indexOf(individualSource);
	if (characterRow !== -1 && individualSource != "") {
		for (var m = 0; m < loops.length; m++) {
			//Reject cause: recorded but not line added/changed
			if (characters[characterRow][12] === "Character already recorded: Breakdown_mark" && addedAux != "Added") { continue }
			loopsCharacterAux = characters[characterRow][6].split(", ");
			if (loopsCharacterAux.indexOf(loops[m]) === -1) {
				//Set
				if (loopsCharacterAux[0] === "") {
					aux2 = 1;
				} else {
					aux2 = loopsCharacterAux.length + 1;
				}
				characters[characterRow][7] = aux2.toString();
				if (characters[characterRow][6] === "") {
					characters[characterRow][6] = loops[m].toString()
				} else {
					characters[characterRow][6] = characters[characterRow][6] + ", " + loops[m]
				}
			}
		}
	}
}
return;
}


function ExtractCharacter(projectID, file_ID, userID) {

var auxCharacterRow;
//Load project character
OpenSheet("DWO_Character", 1, projectID, 2, ssActive);
dwoCharacter = auxSheet;
dwoCharacterData = auxValues;
dwoCharacterNDX = auxNDX;

// Recorre script y extrae los valores únicos de Source

for (var j = 0; j < script.length; j++) {
	var sourceValue = script[j][2];
	var sources = sourceValue.split("/"); // Divide los personajes con "/"

	//Recorre los personajes involucrados en una sola linea
	for (var k = 0; k < sources.length; k++) {
		var individualSource = sources[k].trim(); // Elimina espacios en blanco alrededor
		if (charactersNDX.indexOf(individualSource) === -1 && individualSource != "" && inhibited.indexOf(individualSource.toUpperCase()) === -1) {
			charactersNDX.push(individualSource);
			auxCharacterRow = -1;
			for (var m = 0; m < auxFilteredValues.length; m++) {
				if (auxFilteredValues[m][2].toUpperCase() === individualSource) {
					auxCharacterRow = m;
					break;
				}
			}
			//Personaje ya presente
			if (auxCharacterRow !== -1) {
				//Si ya se grabó y es el final de un preliminar, se marca
				var auxFlag = "";
				if (auxVersion === "Final version: Script_upload_lite") {
					//Busca estado de la grabación

					var aux5 = dwoCharacterProductionNDX2.indexOf(auxFilteredValues[auxCharacterRow][0]) 
					if (aux5 !== -1 && dwoCharacterProductionFiltered[aux5][7] !== "") {
						auxFlag = "Character already recorded: Breakdown_mark";
					}
				}

				// [0] File_Line_ID	[1] File_ID	[2] Character_ID	[3] Character_Name	[4] Character_Equivalence [5] Attributes	[6] Loops	[7] Loops count	[8] Comments		[9] Status [10] Lasts	user [11]  Last change
				characters.push([file_ID + "-S" + Rand(8), file_ID, auxFilteredValues[auxCharacterRow][0], individualSource, "", auxFilteredValues[auxCharacterRow][4], "", "0", "", "(01) Pending: DWOFiles", userID, isNow, auxFlag, projectID]);
			} else {
				//Personaje nuevo
				characters.push([file_ID + "-S" + Rand(8), file_ID, "", individualSource, "", "", "", "0", "", "(02) Check pending / New characters: DWOFiles", userID, isNow, "Character not present in project: Breakdown_mark", projectID]);
				characterRevisionFlag = true;
			}
		}
	}
}
// Ordenar la matriz
characters.sort(function (a, b) {
	return a[3].localeCompare(b[3]); // Ordenar alfabéticamente por la cuarta columna 
});

// ReCrear índice en matriz unidimensional (charactersNDX)
charactersNDX = characters.map(function (character) {
	return character[3]; // Obtener el valor de la cuarta columna
});

return;
}


//FUNCTIONS

/*function SendEmail(destinatario, nombreDestinatario, idDocumento, idPDF, asunto) {
// Convertir el documento de Google Docs a HTML
var documento = DocumentApp.openById(idDocumento);
var urlDocumento = 'https://docs.google.com/feeds/download/documents/export/Export?id=' + idDocumento + '&exportFormat=html';
var respuesta = UrlFetchApp.fetch(urlDocumento, {
	headers: {
		Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
	}
});
var contenidoHTML = respuesta.getContentText();

// Personalizar el contenido con el nombre del destinatario
var contenidoPersonalizado = contenidoHTML.replace("{{nombre}}", nombreDestinatario);

// Obtener el archivo PDF de Google Drive
var archivoPDF = DriveApp.getFileById(idPDF);

// Enviar el correo con el contenido HTML personalizado y el archivo PDF adjunto
MailApp.SendEmail({
	to: destinatario,
	subject: asunto,
	htmlBody: contenidoPersonalizado,
	attachments: [archivoPDF.getAs("application/pdf")]
});
}*/

function OpenSheet(sheetNameAux, ndxCol, key, ndxCol2, ss) {
//OpenSheet("Sheet-to-load", ndx-col, key-value-to-filter, col-to-filter, sheet)
// 
// auxSheet*
// auxValues (complete load)*
// auxNDX if ndx-col > 0*
// auxNDX2 if key-value-to-filter = "" and col-to-filter > 0
// auxFilteredValues if key-value-to-filter <> "" and col-to-filter > 0
// auxRow if key-value-to-filter <> "" and col-to-filter > 0 and result = 1

LazyLoad("DubAppActive01", sheetNameAux);
auxSheet = containerSheet;
var lastRow = auxSheet.getLastRow();
//If empty
if (lastRow === 1) { auxRow = -1; auxValues = []; auxNDX = []; auxNDX2 = []; return }
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
	}
}
}

function OpenSht(sheetNameAux, ndxColValues, keyCol, keyValue, ndxColFiltered, ss) {
// auxSheet / For append
// auxValues (complete load)
// ndxColValues > 0 -> auxNDX
// auxFilteredValues if keyCol & keyValue >""
// ndxColFiltered if keyCol & keyValue >"" & ndxColFiltered > 0 -> auxNDX2

// auxRow if key-value-to-filter <> "" and col-to-filter > 0 and result = 1
LazyLoad("DubAppActive01", sheetNameAux);
auxSheet = containerSheet;
var lastRow = auxSheet.getLastRow();

if (lastRow === 1) {
	auxRow = -1; auxValues = []; auxNDX = []; auxNDX2 = []; auxFilteredValues = [];
	return;
}
auxValues = containerValues;

if (ndxColValues > 0) {
	auxNDX = auxValues.map(r => r[ndxColValues - 1].toString());
}

if (keyCol !== 0 && keyValue != "") {
	// Filtrar valores de la columna keyCol por el valor keyValue
	var filteredValues = auxValues
		.map((row, index) => ({ row: row, index: index })) // Añadir el índice (número de fila) a cada fila
		.filter(item => item.row[keyCol - 1].toString() === keyValue); // Filtrar por keyValue

	// Agregar una columna con el número de fila original en auxValues
	auxFilteredValues = filteredValues.map(item => {
		var newRow = item.row.slice(); // Copiar la fila original
		newRow.push(item.index); // Agregar el número de fila (sumamos 2 para compensar el índice base 0 y el encabezado)
		return newRow;
	});

	if (ndxColFiltered > 0) {
		auxNDX2 = auxFilteredValues.map(r => r[ndxColFiltered - 1].toString());
	}
}
}



function String2Seconds(cadenaDuracion) {

// Divide la cadena en partes (horas, minutos, segundos)
var partes = cadenaDuracion.split(":");
var horas = parseInt(partes[0], 10);
var minutos = parseInt(partes[1], 10);
var segundos = parseInt(partes[2], 10);

var duracion = (horas * 3600 + minutos * 60 + segundos);
return duracion;
}

function Time2Seconds(fecha) {
if (!fecha) return fecha;

// Si es un string, convertirlo a componentes de tiempo
if (typeof fecha === 'string') {
	const partes = fecha.split(':');
	if (partes.length === 3) {
		const hora = parseInt(partes[0], 10);
		const minutos = parseInt(partes[1], 10);
		const segundos = parseInt(partes[2], 10);
		return hora * 3600 + minutos * 60 + segundos;
	}
	return 0; // o manejar el error como prefieras
}

// Si es un objeto Date
if (fecha instanceof Date) {
	const hora = fecha.getHours();
	const minutos = fecha.getMinutes();
	const segundos = fecha.getSeconds();
	return hora * 3600 + minutos * 60 + segundos;
}

// Si no es ninguno de los tipos esperados
console.error('Tipo de fecha no válido:', typeof fecha);
return 0; // o manejar el error como prefieras
}

function Time2String(fecha) {
    if (!fecha) { return ""; }
    
    // Si es un string en formato HH:MM:SS
    if (typeof fecha === 'string') {
        // Verificar si ya está en formato correcto
        if (/^\d{2}:\d{2}:\d{2}$/.test(fecha)) {
            return fecha;
        }
        // Intentar parsear el string
        const partes = fecha.split(':');
        if (partes.length === 3) {
            const hora = partes[0].padStart(2, '0');
            const minutos = partes[1].padStart(2, '0');
            const segundos = partes[2].padStart(2, '0');
            return `${hora}:${minutos}:${segundos}`;
        }
    }
    
    // Si es un objeto Date
    if (fecha instanceof Date) {
        const hora = fecha.getHours().toString().padStart(2, '0');
        const minutos = fecha.getMinutes().toString().padStart(2, '0');
        const segundos = fecha.getSeconds().toString().padStart(2, '0');
        return `${hora}:${minutos}:${segundos}`;
    }
    
    // Si es un número (segundos totales)
    if (typeof fecha === 'number') {
        const horas = Math.floor(fecha / 3600).toString().padStart(2, '0');
        const minutos = Math.floor((fecha % 3600) / 60).toString().padStart(2, '0');
        const segundos = Math.floor(fecha % 60).toString().padStart(2, '0');
        return `${horas}:${minutos}:${segundos}`;
    }
    
    console.error('Formato de Timecode no válido:', fecha);
    return "";
}

function Seconds2String(duracion) {
const horas = Math.floor(duracion / 3600);
const minutos = Math.floor((duracion % 3600) / 60);
const segundos = duracion % 60;

// Formateamos los componentes en una cadena hh:mm:ss
const cadenaFormateada = `${horas.toString().padStart(2, '0')}:${minutos.toString().padStart(2, '0')}:${segundos.toString().padStart(2, '0')}`;

return cadenaFormateada;
}

function CharacterName(cadena) {
var mapa = {
	'á': 'a', 'é': 'e', 'í': 'i', 'ó': 'o', 'ú': 'u',
	'Á': 'A', 'É': 'E', 'Í': 'I', 'Ó': 'O', 'Ú': 'U',
	'ü': 'u', 'Ü': 'U', 'ç': 'c', 'Ç': 'C', 'Â': 'A',
	'Ê': 'E', 'Ô': 'O', 'Ã': 'A', 'Õ': 'O', 'À': 'A'

};

var textoNormalizado = cadena.split('').map(function (char) {
	return mapa[char] || char;
}).join('');

// Quitar todo lo que está entre paréntesis
var sinParentesis = textoNormalizado.replace(/\(.*?\)/g, '');

// Quitar todo lo que está entre corchetes
var sinCorchetes = sinParentesis.replace(/\[.*?\]/g, '');

// Quitar los espacios en blanco al principio y al final
var sinEspacios = sinCorchetes.trim();

// Cualquier caracter extra
var sincaracteres = sinEspacios.replace(/[\(\)\[\]\?\\\=\:;,\.'¨\*\+]/g, '');

// Pasar todo a mayúsculas
var enMayusculas = sincaracteres.toUpperCase();

return enMayusculas;
}

function LoopNumber(seconds2check) {

var auxLoop = Math.floor((seconds2check) / 15) + 1;
return auxLoop;
}


function PDFWitness(pdfname, file_ID) {

try {
	console.log("LoopPDFWitness " + pdfname + " / " + file_ID);

	pdfname = pdfname.replace(/[\/\\:*?"<>|()\[\]]/g, ' ');

	// Open sheet
	ssActive = SpreadsheetApp.openById(sheetID);
	auxLabel = pdfname;

	//Load project character
	OpenSheet("DWO_FilesCharacter", 2, file_ID, 0, ssActive);
	let characterAux = [["Character", "Loops related", "Loops count"]];
	for (let i = 0; i < auxFilteredValues.length; i++) {
		let newRow = [];
		newRow.push(auxFilteredValues[i][3].toString()); // Columna 3
		newRow.push(auxFilteredValues[i][6].toString()); // Columna 6
		newRow.push(auxFilteredValues[i][7].toString()); // Columna 9
		// Agregar la nueva fila a la nueva matriz
		characterAux.push(newRow);
	}

	//Load script
	OpenSheet("DWO_FilesLines", 2, file_ID, 0, ssActive);
	let scriptAux = [["Timecode In", "Timecode Out", "Character", "Dialogue", "Loop related"]];
	for (let i = 0; i < auxFilteredValues.length; i++) {
		let newRow = [];
		newRow.push(Time2String(auxFilteredValues[i][3])); // Columna 4
		newRow.push(Time2String(auxFilteredValues[i][4])); // Columna 5
		newRow.push(auxFilteredValues[i][2].toString()); // Columna 3
		if(auxFilteredValues[i][5] === null){
			newRow.push(auxFilteredValues[i][5].toString()); // Columna 6 without signs
		} else {
			newRow.push(auxFilteredValues[i][15].toString()); // Columna 16 Original
		}
		newRow.push(auxFilteredValues[i][8].toString()); // Columna 9
		// Agregar la nueva fila a la nueva matriz
		scriptAux.push(newRow);
	}

	// Crear la tabla HTML
	let html = '<html><head><style>' +
		'body { font-family: "Arial Narrow", Arial, sans-serif; font-size: 10pt; margin: 20px; }' +
		'table { border: 1px solid black; border-collapse: collapse; width: 100%; table-layout: fixed; }' +
		'th, td { border: 1px solid black; padding: 4px 6px; vertical-align: top; font-size: 10pt; }' +
		'td:last-child { word-wrap: break-word; white-space: normal; }' +
		'td:not(:last-child) { white-space: nowrap; }' +
		'.dialogue-cell { word-wrap: break-word; white-space: pre-wrap; overflow-wrap: break-word; }' +
		'.header-cell { font-weight: bold; background-color: #f8f8f8; word-break: break-word; white-space: normal; }' +
		'.inhibited-row td { color: #999999; font-style: italic; word-break: break-word; white-space: pre-wrap; }' +
		'.added-row td { color: #FF0000; font-weight: bold; word-break: break-word; white-space: pre-wrap; }' +
		'h1 { font-family: "Arial Narrow", Arial, sans-serif; font-size: 14pt; margin: 15px 0; }' +
		'h2 { font-family: "Arial Narrow", Arial, sans-serif; font-size: 12pt; margin: 15px 0; }' +
		'.page-break { page-break-before: always; }' +
		'@page { margin: 0.5in; }' +
		'</style></head><body>' +
		`<h1>${pdfname}</h1>` +
		'<h2>Loops count</h2>' +
		'<table border="1">';
	characterAux.forEach((row, rowIndex) => {
		html += '<tr>';
		row.forEach((cell, index) => {
			if (rowIndex === 0) { // Encabezados
				if (index === 0) { // Character
					html += `<td style="width: 25%" class="header-cell">${cell}</td>`;
				} else if (index === 1) { // Loops related
					html += `<td style="width: 60%" class="header-cell">${cell}</td>`;
				} else { // Count
					html += `<td style="width: 15%" class="header-cell">${cell}</td>`;
				}
			} else {
				if (index === 0) { // Character
					html += `<td style="width: 25%">${cell}</td>`;
				} else if (index === 1) { // Loops related
					let processedCell = '';
					if (cell) {
						processedCell = cell.toString()
							.replace(/&/g, '&amp;')
							.replace(/</g, '&lt;')
							.replace(/>/g, '&gt;')
							.replace(/"/g, '&quot;')
							.replace(/'/g, '&#039;')
							.replace(/\s*,\s*/g, ', '); // Normalizar espacios alrededor de comas
					}
					html += `<td style="width: 60%; word-break: break-word; white-space: pre-wrap;" class="loops-cell">${processedCell}</td>`;
				} else { // Count
					html += `<td style="width: 15%; text-align: center;">${cell}</td>`;
				}
			}
		});
		html += '</tr>';
	});
	html += '</table>';

	// Agregar salto de página y tabla Script breakdown
	html += '<div class="page-break">' +
		'<h2>Script breakdown</h2>' +
		'<table border="1">' +
		'<tr>';

	// Encabezados con ancho fijo y estilo específico
	const headers = ['IN-TIMECODE', 'OUT-TIMECODE', 'CHARACTER', 'DIALOGUE', 'LOOP'];
	headers.forEach((header, index) => {
		let width, style;
		if (index === 0 || index === 1) {
			width = '12%';
		} else if (index === 2) {
			width = '18%';
		} else if (index === 3) {
			width = '45%';
		} else {
			width = '13%';
		}
		style = `width: ${width}; font-weight: bold; background-color: #f8f8f8; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;`;
		html += `<td style="${style}">${header}</td>`;
	});
	html += '</tr>';

	// Contenido de la tabla (sin generar nuevos encabezados)
	scriptAux.slice(1).forEach(row => {  // Usar slice(1) para omitir la fila de encabezados
		const isInhibited = inhibited.includes(row[2]);
		const isAdded = row[3] && row[3].includes("[+]");
		
		html += `<tr${isInhibited ? ' class="inhibited-row"' : (isAdded ? ' class="added-row"' : '')}>`;
		row.forEach((cell, index) => {
			let cellStyle = '';
			if (index === 0 || index === 1) { // Timecode columns
				cellStyle = 'width: 12%; white-space: nowrap;';
			} else if (index === 2) { // Character column
				cellStyle = 'width: 18%; word-break: break-word; white-space: pre-wrap; overflow-wrap: break-word;';
			} else if (index === 3) { // Dialogue column
				cellStyle = 'width: 45%; word-break: break-word; white-space: pre-wrap; overflow-wrap: break-word;';
			} else { // Loop column
				cellStyle = 'width: 13%; white-space: nowrap;';
			}
			
			let processedCell = '';
			if (cell) {
				processedCell = cell.toString()
					.replace(/&/g, '&amp;')
					.replace(/</g, '&lt;')
					.replace(/>/g, '&gt;')
					.replace(/"/g, '&quot;')
					.replace(/'/g, '&#039;')
					.replace(/\n/g, '<br>');
			}
			
			html += `<td style="${cellStyle}">${processedCell}</td>`;
		});
		html += '</tr>';
	});
	html += '</table></div>';

	// Convertir el HTML a PDF
	const blob = Utilities.newBlob(html, 'text/html', 'table.html');
	const pdf = blob.getAs('application/pdf').setName(pdfname);

	// Guardar el PDF en la carpeta especificada
	const folder = DriveApp.getFolderById(folderWitnessId);
	const file = folder.createFile(pdf);
	const docID = file.getId();

	// Buscar información en DWO_Files para enviar el email
	OpenSheet("DWO_Files", 1, file_ID, 0, ssActive);
	if (auxFilteredValues && auxFilteredValues.length > 0) {
		const dwoFilesData = auxFilteredValues[0];
		
		// Extraer los datos necesarios para el email
		const destinatario = dwoFilesData[4]; // Columna E
		
		// Solo enviar el email si hay un destinatario válido
		if (destinatario && destinatario.trim() !== "") {
			const remitente = dwoFilesData[12]; // Columna M
			const nombreDestinatario = "";
			const idDocumento = "1A32AoqCjHrfSaW35ZapKtolw9JqMR0jnDhCd3c8FOnY";
			const idPDF = "";
			const asunto = "DubApp: Script uploaded for " + pdfname;
			
			// Preparar el parámetro de versión - quitar el sufijo ": Script_upload_lite"
			let versionTexto = dwoFilesData[14] || ""; // Columna O
			versionTexto = versionTexto.replace(/\s*:\s*Script_upload_lite$/, "");
			
			const parametros = "Title::DubApp: Script uploaded by translator" + 
				"||Header::A new script was uploaded for your attention: " + 
				"||Detail::" + pdfname + 
				"||Footer::" + versionTexto;
			
			const cc = dwoFilesData[5]; // Columna F
			const bcc = "appsheet@mediaaccesscompany.com";
			
			// Enviar el email utilizando la librería SendEmail
			try {
				if(verboseFlag) {
					console.log("Enviando email a: " + destinatario + " sobre: " + pdfname);
				}
				
				// Llamar a la función AppSendEmailX de la librería SendEmail
				SendEmail.AppSendEmailX(
					destinatario,
					remitente,
					nombreDestinatario,
					idDocumento,
					idPDF,
					asunto,
					parametros,
					cc,
					bcc
				);
				
				if(verboseFlag) {
					console.log("Email enviado exitosamente");
				}
			} catch (emailError) {
				console.error("Error al enviar email: " + emailError.message);
				// No lanzamos el error para que continúe la ejecución
			}
		} else if(verboseFlag) {
			console.log("No se envió email porque el destinatario está vacío");
		}
	}

} catch (e) {
	Logger.log('Error: ' + e.message);
}
}

function Rand(n) {
// Calcula el valor mínimo y máximo para n cifras
var min = Math.pow(10, n - 1);
var max = Math.pow(10, n) - 1;

// Genera un número aleatorio entre min y max
var numeroAleatorio = Math.floor(Math.random() * (max - min + 1)) + min;

return numeroAleatorio;
}

function ClearRow(ssAux, keyCol, keyAux) {

// Define el rango de datos
var datos = ssAux.getDataRange().getValues();

// Define las filas que cumplen la condición (por ejemplo, filas donde la columna A está vacía)
var filasAEliminar = [];
for (var i = 0; i < datos.length; i++) {
	if (datos[i][keyCol] === keyAux) { // Cambia esta condición según tus necesidades
		filasAEliminar.push(i + 1);
	}
}

// Agrupa las filas en rangos continuos
var rangosContinuos = [];
var inicio = null;
for (var j = 0; j < filasAEliminar.length; j++) {
	if (inicio === null) {
		inicio = filasAEliminar[j];
	}
	if (j === filasAEliminar.length - 1 || filasAEliminar[j] + 1 !== filasAEliminar[j + 1]) {
		rangosContinuos.push([inicio, filasAEliminar[j]]);
		inicio = null;
	}
}

// Limpia los contenidos de los rangos continuos
for (var k = 0; k < rangosContinuos.length; k++) {
	var rango = rangosContinuos[k];
	ssAux.getRange(rango[0], 1, rango[1] - rango[0] + 1, ssAux.getLastColumn()).clearContent();
}
}

function sendErrorNotification(error, functionName, data) {
	const adminEmail = BCC_EMAIL;
	const subject = `Error en ${functionName} - DubApp`;
	const VERSION = "v1.0.1 - 27/01/2024"; // Agregamos versión para tracking
	
	const body = `VERSIÓN DEL SCRIPT: ${VERSION}
Se produjo un error en la función ${functionName}
Timestamp: ${new Date().toISOString()}
Error: ${error.toString()}
Stack: ${error.stack || 'No stack trace disponible'}
Data: ${JSON.stringify(data, null, 2)}
Config actual:
- verboseFlag: ${verboseFlag}
- fasttrack: ${fasttrack}
- appRepository: ${appRepository}
- Timezone: ${TIMEZONE}
`;

	MailApp.sendEmail(adminEmail, subject, body);
	
	// También registrar en el log de Apps Script
	console.error(`[${VERSION}] Error en ${functionName}:`, {
		timestamp: new Date().toISOString(),
		error: error.toString(),
		stack: error.stack,
		data: data
	});
}

/**
* Procesa los cambios acumulados en control2add
*/
function processControlChanges() {
if (!control2add.length) return;

controlSheet = SpreadsheetApp.openById(controlID).getSheetByName("CON-TaskCurrent"); // Inicializar controlSheet

const now = Utilities.formatDate(new Date(), TIMEZONE, TIMESTAMP_FORMAT);
const rows = control2add.map(change => [
	change.sheet,                  // Nombre de la hoja
	change.key,                    // Clave del registro
	now,
	"DubAppActive01",
	change.action,                 // Tipo de acción (INSERT_ROW, EDIT, etc)
	change.user,                   // Usuario
	"01 Pending"
]);

// Agregar todas las filas de una vez
controlSheet.getRange(controlSheet.getLastRow() + 1, 1, rows.length, 7)
	.setValues(rows);

// Limpiar el array después de procesar
control2add = [];
}

function schedulePDFWitness(pdfname, file_ID) {
    try {
        if (!pdfname || !file_ID) {
            throw new Error('Parámetros inválidos para schedulePDFWitness');
        }

        // Limpiar triggers existentes
        const triggers = ScriptApp.getProjectTriggers();
        triggers.forEach(trigger => {
            if (trigger.getHandlerFunction() === 'asyncPDFWitness') {
                ScriptApp.deleteTrigger(trigger);
            }
        });

        // Guardar parámetros
        const props = PropertiesService.getScriptProperties();
        const params = {
            'pdfname': pdfname,
            'file_ID': file_ID,
            'scheduledTime': Utilities.formatDate(new Date(), TIMEZONE, TIMESTAMP_FORMAT)
        };
        
        Object.keys(params).forEach(key => {
            props.setProperty(key, params[key]);
        });
        
        // Crear trigger para ejecutar en 1 segundo
        ScriptApp.newTrigger('asyncPDFWitness')
            .timeBased()
            .after(1000) // 1 segundo
            .create();
            
        conditionalLog('PDFWitness programado:', params);
        return true;

    } catch (error) {
        console.error('Error en schedulePDFWitness:', error);
        sendErrorNotification(error, 'schedulePDFWitness', {
            pdfname: pdfname,
            file_ID: file_ID
        });
        return false;
    }
}

function asyncPDFWitness() {
    conditionalLog('Iniciando asyncPDFWitness');
    try {
        // Obtener parámetros guardados
        const props = PropertiesService.getScriptProperties();
        const pdfname = props.getProperty('pdfname');
        const file_ID = props.getProperty('file_ID');
        
        if (!pdfname || !file_ID) {
            throw new Error('Parámetros no encontrados en propiedades');
        }

        // Ejecutar PDFWitness
        PDFWitness(pdfname, file_ID);
        
        // Limpiar propiedades después de la ejecución
        props.deleteAllProperties();
        
        conditionalLog('PDFWitness completado exitosamente');
        
    } catch (error) {
        console.error('Error en asyncPDFWitness:', error);
        sendErrorNotification(error, 'asyncPDFWitness', {
            error: error.toString(),
            stack: error.stack
        });
    }
}

function cleanupResources() {
    try {
        const props = PropertiesService.getScriptProperties();
        props.deleteAllProperties();
        
        const triggers = ScriptApp.getProjectTriggers();
        triggers.forEach(trigger => {
            if (trigger.getHandlerFunction() === 'asyncPDFWitness') {
                ScriptApp.deleteTrigger(trigger);
            }
        });
        
        console.log('Recursos limpiados exitosamente');
    } catch (error) {
        console.error('Error al limpiar recursos:', error);
    }
}

// Agregar función helper para logging al inicio del archivo
function logOperationStatus(operation, details) {
    if (verboseFlag) {
        const timestamp = new Date().toISOString();
        console.log(`${timestamp} - ${operation}: ${JSON.stringify(details)}`);
    }
}