//Global ID
const allIDs = databaseID.getID();

function DWOObs(obsAuthor, currentFileName, eventID, mixandeditID){

 //Params
/*function DWOObs(){
  var currentFileName = "98ef0604.FileImport.135127.txt"
  var folderID = "1pkvWYHN_uX4t47zhL8k1ooni8zX0A606";
  var eventID="3NJcoWhg-297YmDIrqBL8LhS4crYKcQIB";
  var mixandeditID = "98ef0604";
  var obsAuthor = "appsheet@mediaaccesscompany.com";
  var ID = "1FLv7GXmuTOIb0kMK1zuG4_SyPCtU_wjFWph9Ip_IYVU";
*/
  //VAR
  var timezone = allIDs['timezone'];
  var timestamp_format = allIDs['timestamp_format']; // Timestamp Format. 
  var eventTime = Utilities.formatDate(new Date(), timezone, timestamp_format);
  var initialStatus ="(00) Draft: DWOObservation";
  var currentFileID = "";

  //Data
  var ssActive = SpreadsheetApp.openById(allIDs['activeID']);
  var DWOObs = ssActive.getSheetByName('DWO-Observation');

  var DWOObslastRow =DWOObs.getLastRow()-1;
  var data =[];

  // BEGINNING OF NEW CODE TO PREVENT DUPLICATE OBSERVATIONS
  var dwoEventSheet = ssActive.getSheetByName('DWO-Event'); // Assuming DWO-Event sheet is in the active spreadsheet

  if (!dwoEventSheet) {
    Logger.log("Error: La hoja 'DWO-Event' no se encontró. No se puede verificar duplicados basados en Production ID.");
    // Considerar si se debe detener la ejecución o continuar sin la verificación de duplicados.
    // Por ahora, continuará, pero podría generar duplicados si la hoja falta.
  }

  var allDwoEventRows = dwoEventSheet ? dwoEventSheet.getDataRange().getValues() : [];
  var allDwoObservationRows = DWOObs.getDataRange().getValues();

  // 1. Find the Production ID for the current eventID
  var targetProductionID = null;
  if (dwoEventSheet) {
    // Start from 1 to skip header row
    for (var i = 1; i < allDwoEventRows.length; i++) {
      if (allDwoEventRows[i][0] && allDwoEventRows[i][0].toString() === eventID) { // Column 'Event ID' is 0
        targetProductionID = allDwoEventRows[i][1]; // Column 'Production ID' is 1
        break;
      }
    }
  }

  var relatedEventIDs = [];
  if (targetProductionID && dwoEventSheet) {
    // Start from 1 to skip header row
    for (var i = 1; i < allDwoEventRows.length; i++) {
      if (allDwoEventRows[i][1] && allDwoEventRows[i][1].toString() === targetProductionID) { // Column 'Production ID' is 1
        relatedEventIDs.push(allDwoEventRows[i][0].toString()); // Column 'Event ID' is 0
      }
    }
  } else if (dwoEventSheet) {
    // If no targetProductionID found for the eventID, or if eventID itself wasn't found,
    // consider only the current eventID for checking duplicates to be safe, or log an error.
    // For now, if targetProductionID is null, relatedEventIDs will be empty, and no cross-event duplicates will be checked.
    // If eventID was found but had no Production ID, relatedEventIDs would also be empty.
    // If the eventID itself is what we want to check against, add it:
    // if (!targetProductionID && eventID) relatedEventIDs.push(eventID); 
    // However, the request was to check based on common Production ID, so an empty relatedEventIDs if no ProdID is found seems correct.
  }

  // 3. Collect existing observations (type, timecode) for the relatedEventIDs
  var existingObservationsMap = {}; // Using an object/map for efficient lookup: "type|timecode" -> true
  // Start from 1 to skip header row
  for (var i = 1; i < allDwoObservationRows.length; i++) {
    var obsEventID = allDwoObservationRows[i][1] ? allDwoObservationRows[i][1].toString() : null; // Column 'Event ID' is 1
    if (obsEventID && relatedEventIDs.indexOf(obsEventID) !== -1) {
      var obsType = allDwoObservationRows[i][2] ? allDwoObservationRows[i][2].toString() : ""; // Column 'Observation type' is 2
      var obsTimecode = allDwoObservationRows[i][3] ? allDwoObservationRows[i][3].toString() : ""; // Column 'Timecode' is 3
      // Normalize timecode format if necessary, assuming it's already HH:MM:SS
      existingObservationsMap[obsType + "|" + obsTimecode] = true;
    }
  }
  // END OF NEW CODE TO PREVENT DUPLICATE OBSERVATIONS

  //Objects
  var ssNotrack = SpreadsheetApp.openById(allIDs['noTrackID']);
  var APPObjects = ssNotrack.getSheetByName('App-Objects');
  var lastRow = APPObjects.getLastRow();
  var lastCol = APPObjects.getLastColumn()
  var APPObjectsData = APPObjects.getRange(2,1,lastRow - 1, lastCol); 
  var APPObjectsValues = APPObjectsData.getValues(); 
  var APPObjectsNDX = APPObjectsValues.map(function(r){ return r[1]; });

  var folder = DriveApp.getFolderById(allIDs['proToolID']);
  var files = folder.getFiles();
  while (files.hasNext()){
    file = files.next();
    if (currentFileName === file.getName()) {
      currentFileID = file.getId();
      break;
    } 
  }

  //Obtain content
  var content = DriveApp.getFileById(currentFileID).getBlob().getDataAsString();

  var contentData = content.split(/\r?\n/);
  //Look for "#"
  var beginImport = false;
  for (var rowRead = 11; rowRead < contentData.length; rowRead++)
  {
    var aux = "\t";
    var importLine =  contentData[rowRead].split(/\t/);
    //If empty then loop
    if(importLine[1]=="") {continue;}

    if(beginImport && importLine.length>2) {
      //Importing
      var auxObsID = newKey(8);
      //TIMECODE
      var auxTimecode=importLine[1].trim();
      var parts = auxTimecode.split(':');
      auxTimecode = parts.slice(0, 3).join(':');
      
      //MARKER NAME
      var auxMarker=importLine[4].trim();

      //COMMENT
      var aux3 = auxMarker.split("_");
      if (aux3.length > 1) {
        var auxComment = aux3[1];
      } else {
        var auxComment ="";
      }

      //OSB TYPE
      var auxObsType="";
      //Extract Type
      if(auxMarker.length>0){
        var aux1 = auxMarker.split(" ");
        var aux2 = aux1[0].trim();

        //Check if exist
        var obsRow = APPObjectsNDX.indexOf("Observation",0); 

        while (obsRow != -1) {
          if ( aux2== APPObjectsValues[obsRow][2]) {
            auxObsType = APPObjectsValues[obsRow][0]+": Observation";
            break;
          }
          var obsRow = APPObjectsNDX.indexOf("Observation",obsRow+1); 
        }
      }

      // IF TYPE OUT OF CATALOG
      if(obsRow == -1 && auxMarker.length>0) {
        auxComment=aux2 + " " + auxComment;
      }

      // CHARACTER
      if(aux3.length>0) {
        var auxCharacter= aux3[0].substring(aux2.length + 1).trim();
      } else {
        var auxCharacter="";
      }

      // NEW CHECK FOR DUPLICATES
      var mapKey = auxObsType + "|" + auxTimecode;
      if (existingObservationsMap[mapKey]) {
        Logger.log("Observación duplicada omitida: Tipo=" + auxObsType + ", Timecode=" + auxTimecode + " (relacionada por Production ID: " + (targetProductionID || 'N/A') + ")");
        continue; // Skip this observation as it already exists
      }
      // END OF NEW CHECK

      // SAVE OBS
      data.push([auxObsID, eventID, auxObsType, auxTimecode, "", auxCharacter, auxComment, obsAuthor, eventTime, "", "", "", mixandeditID, "", "", "", "", "", "", initialStatus, obsAuthor, eventTime]);
    } else {
      if(importLine[0].trim()=="#") {
        beginImport = true;
      }
    }
  }
  if(data.length>0) {
    DWOObs.getRange(DWOObslastRow + 2,1,data.length, data[0].length).setValues(data);
  }
}

function newKey(len) {
  const possible = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ123456789";
  let id = "";
  if(!Number(len)) {
    throw new Error("The length must be an integer.")
  }
  for(var i=0; i<len; i++) {
    id += possible.charAt(Math.floor(Math.random() * possible.length));
  }
  return id;
}