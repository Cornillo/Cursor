//Global
const allIDs = databaseID.getID();
const sheetID= allIDs["activeID"];
var auxSheet;
var auxValues;
var auxFilteredValues;
var auxNDX;
var auxNDX2;
var auxRow;
// GDoc template file ID in G drive
const templateFileID = '1YdfzaVjOunoCkvNf1Y8Agp6XDgF1o_Nm9Glu0T91msg';
const templateMailID = "1U1hnS2_wfrbElVWjOE0OFxIQNItyPaPliZRWV6e4BzQ";
const folderId= allIDs["uploaded"];
const folderWitnessId= allIDs["loopsWitness"];
var auxVersion;

//URL Web App https://script.google.com/macros/s/AKfycbybS_HbLZfAoC_btdyxwTaGC70NJO5GQNMvlgiMt6WTlbJJjyGY7Jk2rzoS5D0sGehnxg/exec

function call(){

PDFWitness("Big City Greens - S4 - Episode F098 - Dialogue Translation (only finals)","15a308ce", "","");
}

function doPost(e) {
var params = JSON.parse(e.postData.contents);
var pdfname = params.pdfname;
var file_ID = params.file_ID;
var auxProduction = params.auxProduction;
var auxUserName = params.auxUserName;

// Llamar a la función PDFWitness con los parámetros proporcionados
PDFWitness(pdfname, file_ID, auxProduction, auxUserName);

return ContentService.createTextOutput("Success");
}


function PDFWitness(pdfname, file_ID, auxProduction, auxUserName){

try {
  console.log("LoopPDFWitness "+pdfname+" / "+file_ID);

  pdfname=pdfname.replace(/[\/\\:*?"<>|()\[\]]/g, ' ');

  // Open sheet
  ssActive = SpreadsheetApp.openById(sheetID);
  auxLabel=pdfname;

  //Load project character
  OpenSheet("DWO_FilesCharacter", 2, file_ID, 0, ssActive);
  let characterAux=[["Character","Loops related", "Loops count"]];
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
  let scriptAux=[["Timecode In","Timecode Out", "Character","Dialogue","Loop related"]];
  for (let i = 0; i < auxFilteredValues.length; i++) {
    let newRow = [];
    newRow.push(Time2String(auxFilteredValues[i][3])); // Columna 4
    newRow.push(Time2String(auxFilteredValues[i][4])); // Columna 5
    newRow.push(auxFilteredValues[i][2].toString()); // Columna 3
    newRow.push(auxFilteredValues[i][5].toString()); // Columna 6
    newRow.push(auxFilteredValues[i][8].toString()); // Columna 9
    // Agregar la nueva fila a la nueva matriz
    scriptAux.push(newRow);
  }

  // Crear la tabla HTML
  let html = '<html><head><style>table, th, td {border: 1px solid black;border-collapse: collapse;</style></head><body><table border="1"><h2>Script breakdown</h2>';
  scriptAux.forEach(row => {
    html += '<tr>';
    row.forEach(cell => {
      html += `<td>${cell}</td>`;
    });
    html += '</tr>';
  });
  html += '</table></body></html>';

  // Crear la segunda tabla HTML
  html += '<html><head><style>table, th, td {border: 1px solid black;border-collapse: collapse;</style></head><body><table border="1"><h2>Loops by characters</h2>';
  characterAux.forEach(row => {
    html += '<tr>';
    row.forEach(cell => {
      html += `<td>${cell}</td>`;
    });
    html += '</tr>';
  });
  html += '</table></body></html>';

  // Convertir el HTML a PDF
  const blob = Utilities.newBlob(html, 'text/html', 'table.html');
  const pdf = blob.getAs('application/pdf').setName(pdfname);

  // Guardar el PDF en la carpeta especificada
  const folder = DriveApp.getFolderById(folderWitnessId);
  const file = folder.createFile(pdf);
  const docID = file.getId();

  // Load DWO_FIles
  auxCaseDWO_Files=auxFilteredValues;
  auxVersion = auxCaseDWO_Files[0][14];
  var auxSentTo = auxCaseDWO_Files[0][12];
  if(auxVersion==="Add loops prelim/final difference: Script_upload_lite") {
    SendEmail.AppSendEmail(auxSentTo, auxUserName, templateMailID,  docID, "DubApp: Script Breakdown "+auxProduction);
  }

} catch (e) {
  Logger.log('Error: ' + e.message);
}
}

function OpenSheet(sheetNameAux, ndxCol, key, ndxCol2, ss) {
//OpenSheet("Sheet-to-load", ndx-col, key-value-to-filter, col-to-filter, sheet)
// 
// auxSheet*
// auxValues (complete load)*
// auxNDX if ndx-col > 0*
// auxNDX2 if key-value-to-filter = "" and col-to-filter > 0
// auxFilteredValues if key-value-to-filter <> "" and col-to-filter > 0
// auxRow if key-value-to-filter <> "" and col-to-filter > 0 and result = 1

auxSheet = ss.getSheetByName(sheetNameAux);
var lastRow = auxSheet.getLastRow();
//If empty
if(lastRow === 1) {auxRow=-1; auxValues=[]; auxNDX=[]; auxNDX2=[]; return}
var lastCol = auxSheet.getLastColumn();
var auxData = auxSheet.getRange(2,1,lastRow - 1, lastCol); 
auxValues = auxData.getValues();
if(ndxCol2 !=0) {
  auxNDX2 = auxValues.map(function(r){ return r[ndxCol2-1].toString(); });
}
if(ndxCol>0) {
  auxNDX = auxValues.map(function(r){ return r[ndxCol-1].toString(); });
  if(key!=""){
    if(ndxCol2 !=0) {
      var auxcase = auxNDX2.indexOf(key);
    } else {
      var auxcase = auxNDX.indexOf(key);
    }
    auxFilteredValues = [];
    while (auxcase !== -1){
      auxRow = auxcase;
      auxFilteredValues.push(auxValues[auxcase]);

      if(ndxCol2 !=0) {
        auxcase = auxNDX2.indexOf(key, auxcase + 1);
      } else {
        auxcase = auxNDX.indexOf(key, auxcase + 1);
      }
    }
  } 
}
}

function Time2String(fecha) {
  if(!fecha){return ""}
  const hora = fecha.getHours();
  const minutos = fecha.getMinutes();
  const segundos = fecha.getSeconds();

  // Formateamos los componentes en una cadena hh:mm:ss
  const cadenaFormateada = `${hora.toString().padStart(2, '0')}:${minutos.toString().padStart(2, '0')}:${segundos.toString().padStart(2, '0')}`;

  return cadenaFormateada;
}
