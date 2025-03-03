const CONSTANTS = {
  DURATION_FORMAT: "HH:mm:ss",
  TEXT_MARKERS: {
    DRAFT_TEXT: "D  R  A  F  T     C  O  P  Y",
    NO_COMMENTS: "\n                                THERE ARE NO COMMENTS ON THIS EPISODE"
  }
};

const allIDs = databaseID.getID();
const timezone = allIDs.timezone;
const timestamp_format = allIDs.timestamp_format;
const destinationFolderID = allIDs.qcCreativeFolderID;
const paramID = allIDs.activeID;

//Constants
const eventTime = Utilities.formatDate(new Date(), timezone, timestamp_format); 

function call() {
  var paramEventID = "noyqwiAF-SKiZvbg6sUP5pkng4gWwX9Vs";
  var paramFileName = "Bob's Burgers F284 LAS QC " + eventTime;
  var paramProduction = "Bob's Burgers F284";
  var paramPM = "Denise Surce";
  var paramQCerName = "Paula Nuñez";
  var paramVersion = "Final";
  var paramFrom = "Episode QC: Sandra Brizuela";
  var paramPrevious = "1O_ter7gxL_Bw3UnY4vsLxhO71AuWkiyL";
  var paramGrant = "gustavo.cornillon@mediaaccesscompany.com";

  var fileID=QCCreativeReport(paramEventID, paramProduction, paramPM, paramQCerName, paramFileName, paramVersion, paramGrant, paramPrevious, paramFrom, "1dX8YDAbfasnf7au-nDNcott-gxvrVM8A7T7bBFUkp_8");
  Logger.log(fileID);
}

/**
 * Genera un reporte de control de calidad creativo en formato PDF
 * @param {string} paramEventID - ID del evento
 * @param {string} paramProduction - Nombre de la producción
 * @param {string} paramPM - Nombre del Project Manager
 * @param {string} paramQCerName - Nombre del QCer
 * @param {string} paramFileName - Nombre del archivo
 * @param {string} paramID - ID del documento
 * @param {string} paramVersion - Versión del documento ('Draft' o 'Final')
 * @param {string} paramGrant - Email para compartir el documento
 * @param {string} paramPrevious - ID del documento anterior
 * @param {string} paramFrom - Origen del reporte
 * @param {string} templateFileID - ID de la plantilla
 * @returns {string} ID del archivo PDF generado
 * @throws {Error} Si faltan parámetros requeridos o hay errores en el proceso
 */
function QCCreativeReport(paramEventID, paramProduction, paramPM, paramQCerName, paramFileName, paramVersion, paramGrant, paramPrevious, paramFrom, templateFileID) {
  try {
    const eventTime = Utilities.formatDate(new Date(), timezone, timestamp_format);
    
    Logger.log({paramEventID: paramEventID, paramProduction: paramProduction, paramPM:  paramPM, paramQCerName: paramQCerName, paramFileName: paramFileName, paramVersion: paramVersion, paramGrant: paramGrant, paramPrevious: paramPrevious, paramFrom: paramFrom});
//function QCCreativeReport() {

    const duration_format = "HH:mm:ss";
    
    var aux1="";

    //Paramaters replacement
    /*const 
    const paramID = "1gdCgdj7bBGUxFMBfsBB1yupNyxvNnbp9b1ZffZfS3wA";*/


    //Open tables
    //Observation
    const TableDWOObs =  SpreadsheetApp.openById(paramID).getSheetByName('DWO-Observation');
    const lastRow = TableDWOObs.getLastRow();
    const lastCol = TableDWOObs.getLastColumn()
    const TableDWOObsData = TableDWOObs.getRange(2,1,lastRow - 1, lastCol); 
    const TableDWOObsValues = TableDWOObsData.getValues();
    const TableDWOObsNDX = TableDWOObsValues.map(function(r){ return r[1]; });
    
    const templateFile = DriveApp.getFileById(templateFileID);
    const destinationFolder = DriveApp.getFolderById(destinationFolderID);

    if(paramVersion == "Draft") {
      paramFileName=paramFileName+" DRAFT";
    }

    if(paramPrevious!="") {
    //Delete Previous PDF
      try {
        DriveApp.getFileById(paramPrevious).setTrashed(true);
      } catch (e) {
        paramPrevious="";
      }
    }
    
    // Create an intermediate copy of the template GDoc file called "intermediateFile"
    // This copy will also be a GDoc file
    // We create copy of the template file because original template will be continuously reused.
    // Assign a unique file name to the new intermediate file because in multiuser system the file name needs
    // to be unique. So assign a date time stamp to the intermediate file name or some better unique qualifier such as unique ID.

    const intermediateFile = templateFile.makeCopy(paramFileName, destinationFolder);
    const auxID = intermediateFile.getId();
    const intermediateFileToEdit = DocumentApp.openById(auxID);

    //Populate the data fetched from the AppSheet record by the GAS as an argument to the GAS function.
      
    const doc= intermediateFileToEdit.getBody();
    
    doc.replaceText("<<Date>>", eventTime);
    doc.replaceText("<<Production>>", paramProduction);
    doc.replaceText("<<PM>>", paramPM);
    doc.replaceText("<<QCer Name>>", paramQCerName);
    doc.replaceText("<<From>>", paramFrom);
    if(paramVersion != "Draft") {
      doc.replaceText("D  R  A  F  T     C  O  P  Y", "");
    }

    const auxTable = doc.getTables()[0];
    var cell="";

    // Observation loop
    var auxObsRow = TableDWOObsNDX.indexOf(paramEventID,0); 
    if(auxObsRow == -1) {
      //No observations
      // Delete table
      auxTable.removeFromParent();
      // Replace text
      doc.replaceText("Hi, please consider the following changes for this episode:", 
      "\n                                THERE ARE NO COMMENTS ON THIS EPISODE");
    } else {
      //With Observations
      var auxTableLineCount = 1;
      // Extract current Obs
      var currentObs = [];
      while (auxObsRow != -1) {
        // Si la observación tiene uno de los estados válidos
        if(['(24) Only notification: DWOObservation', 
            '(01) Resolution pending: DWOObservation',
            '(00) Draft: DWOObservation'].includes(TableDWOObsValues[auxObsRow][19])) {
          currentObs.push(TableDWOObsValues[auxObsRow]);
        }
        var auxObsRow = TableDWOObsNDX.indexOf(paramEventID,auxObsRow + 1); 
      }
      var totalObs = currentObs.length;
      if(totalObs != 0 ) {
        //Sort extract
        currentObs.sort(function(a, b) {
          return a[3] - b[3];
        });
        for(i = 0; i < totalObs; i++ ) {
          //Reject options
          if(!currentObs[i][6] || currentObs[i][6]==="Cambiar por " || currentObs[i][19]==="(99) Canceled: DWOObservation" ) {
            continue;
          }

          //Add a file
          var tableRow = auxTable.appendTableRow();
          // Ahora puedes agregar celdas (TableCells) a la fila
          var tableCell = tableRow.appendTableCell();
          tableCell = tableRow.appendTableCell();
          tableCell = tableRow.appendTableCell();
          auxTable.setBorderColor('#000000')
          auxTable.setBorderWidth(1);

          //Timecode
          cell = auxTable.getCell(auxTableLineCount, 0);
          if(currentObs[i][3]!="") {
            aux1 = formatTimecode(currentObs[i][3]);
            cell.setText(aux1);
          }

          //Character
          cell = auxTable.getCell(auxTableLineCount, 1);
          cell.setText(currentObs[i][5]);

          //Comment
          cell = auxTable.getCell(auxTableLineCount, 2);
          if(currentObs[i][6]==''){
            //Type
            var auxType = " *"+currentObs[i][2]+"* ";
            auxType = auxType.replace(': Observation','');
            auxType = auxType.toUpperCase();
            auxType = currentObs[i][4]+auxType+currentObs[i][9];
            cell.setText(auxType);
          } else {
            cell.setText(currentObs[i][6]);
          }
          
          Logger.log({timecode: aux1, character: currentObs[i][5]});
          // Observation type highligting 
          //        aux1 = auxType.length;
          //        cell.editAsText().setBold(0, aux1, true);
          //        cell.editAsText().setForegroundColor(0, aux1, '#FF0000');

          auxTableLineCount ++;
        }
      }
      formatTextInCell(auxTable)
    }
    // Save the intermdiate G doc file after changes have been made to it by populating the data fetched by the 
    // AppScript from the AppSheet record.
      
      intermediateFileToEdit.saveAndClose();
    // Do necessary housekeeping to convert the intermediate G doc file to final PDF file used as report
      
      const folder = DriveApp.getFolderById(destinationFolderID);
      const fileID = intermediateFileToEdit.getId();
      const docFile = DriveApp.getFileById(fileID);

    //Give a name to the PDF file and save it in the same destination folder
      var file = folder.createFile(docFile.getBlob()).setName(paramFileName+'.pdf');
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW)

    //Delete the inermediate Gdoc file because a PDF has been created now.
      DriveApp.getFileById(fileID).setTrashed(true);

      // Configurar permisos sin notificaciones usando la API REST
      const url = "https://www.googleapis.com/drive/v3/files/" + file.getId() + "/permissions";
      const options = {
        method: "post",
        contentType: "application/json",
        headers: {
          Authorization: "Bearer " + ScriptApp.getOAuthToken()
        },
        payload: JSON.stringify({
          role: "reader",
          type: "user",
          emailAddress: paramGrant,
          sendNotificationEmail: false
        }),
        muteHttpExceptions: true
      };
      
      UrlFetchApp.fetch(url, options);

      return file.getId();
  } catch (error) {
    Logger.log('Error en QCCreativeReport2: ' + error.toString());
    throw new Error('Error al procesar el reporte: ' + error.message);
  }
}


function formatTextInCell(tabla) {
  var numRows = tabla.getNumRows();
  for (var i = 0; i < numRows; i++) {
    var celda = tabla.getRow(i).getCell(2);
    var texto = celda.editAsText();

    //Check if empty
    if (texto.getText() === "") {
      continue; // Salta a la siguiente iteración del bucle
    }

    var regexUnderline = /\~(.*?)\~/g;
    var matchUnderline;
    while (matchUnderline = regexUnderline.exec(texto.getText())) {
      var inicio = matchUnderline.index;
      var fin = inicio + matchUnderline[0].length; 
      texto.replaceText("~" + matchUnderline[1] + "~", matchUnderline[1]);
      texto.setUnderline(inicio, fin - 3, true);
    }

    var regexBold = /\*(.*?)\*/g;
    var matchBold;
    while (matchBold = regexBold.exec(texto.getText())) {
      var inicio = matchBold.index;
      var fin = inicio + matchBold[0].length;
      texto.replaceText("\\*" + matchBold[1] + "\\*", matchBold[1]);
      if(inicio <= fin - 3){
        texto.setBold(inicio, fin - 3, true);
      }
    }

    var regexItalic = /\_(.*?)\_/g;
    var matchItalic;
    while (matchItalic = regexItalic.exec(texto.getText())) {
      var inicio = matchItalic.index;
      var fin = inicio + matchItalic[0].length;
      texto.replaceText("_" + matchItalic[1] + "_", matchItalic[1]);
      if(inicio <= fin - 3){
        texto.setItalic(inicio, fin - 3, true);
      }
    }

    var regexDeleted = /\%(.*?)\%/g;
    var matchDeleted;
    while (matchDeleted = regexDeleted.exec(texto.getText())) {
      var inicio = matchDeleted.index;
      var fin = inicio + matchDeleted[0].length;
      texto.replaceText("%" + matchDeleted[1] + "%", matchDeleted[1]);
      if(inicio <= fin - 3){
        texto.setStrikethrough(inicio, fin - 3, true);
      }
    }
  }
}


function formatTimecode(dateTime) {
  var formattedDateTime = Utilities.formatDate(dateTime, Session.getScriptTimeZone(), "HH:mm:ss");
  return formattedDateTime;
}