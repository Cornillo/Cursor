//GLOBAL
const allIDs = databaseID.getID();
const activeID= allIDs["activeID"];
const noTrackID= allIDs["noTrackID"];

const ccResumenDiario = "paula.nunez@mediaaccesscompany.com , irina.nugoli@mediaaccesscompany.com, mylai.hernandez@mediaaccesscompany.com , rocio.panizo@mediaaccesscompany.com , martina.chiarullo@mediaaccesscompany.com , denise.labat@mediaaccesscompany.com";
const ccoEnvioActor = "paula.nunez@mediaaccesscompany.com , appsheet@mediaaccesscompany.com";

// Abre la hoja de cálculo activa
var spreadsheet;
var settlementSheetID;
var ssActive;
var ssNoTrack;
var ssSettlementSheet;
var auxSheet;
var auxValues;
var auxFilteredValues;
var auxNDX;
var auxNDX2;
var auxRow;
const timezone = allIDs["timezone"];
var actorValues;
var actorNDX;
var rateValues;
var rateNDX;
var rateItemValues;
var rateItemNDX;
var charNDX;
var charValues;
var charProdValues;
var charProdNDX;
var productionValues;
var productionValuesNDX;
var projectValues;
var projectValuesNDX;
var songValues;
var songNDX;
var userValues;
var userNDX;
var actorProdNDX=[];
var actorProd=[];
var dataAdult = [];
var dataRetired = [];
var dataMex = [];
var data1617 = [];
var dataMinor = [];
var dataDub = [];
var dataUndetermined = [];
var dataDirDialog = [];
var dataDirMusic = [];
var songIDWitness = [];
var disabledStatus = ["(04) Dismissed: DWOCharacterProduction", "(05) Covered by neutral: DWOCharacterProduction", "(06) Covered as ditty: DWOCharacterProduction", "(09) Unnamed broken down: DWOCharacterProduction", "(07) Covered by previous recording: DWOCharacterProduction"];

//FUNCTIONS
/*function call2(){
var fechaInicial = new Date("2025-05-06T00:00:00"); // Año-Mes-Día
var fechaFinal = new Date("2025-05-06T00:00:00");
actorsSettlement2(fechaInicial,fechaFinal, "Mailing");
}*/

function call(){
var hojaActiva = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("RESUMEN"); // Obtener la hoja RESUMEN

var fechaInicial = hojaActiva.getRange("C5").getValue();
var fechaFinal =   hojaActiva.getRange("C7").getValue();
actorsSettlement(fechaInicial,fechaFinal);

}


function call3(){
var fechaInicial = new Date(Date.UTC(2025, 1, 10, 3, 0, 0)); // 10/2/25 00:00 GMT-3
var fechaFinal = new Date(Date.UTC(2025, 1, 11, 3, 0, 0));   // 11/2/25 00:00 GMT-3
actorsSettlement2(fechaInicial,fechaFinal, "Mailing resume");

}

function actorsSettlement(auxFromDate, auxToDate){
actorsSettlement2(auxFromDate, auxToDate, "Settlement")
}

function modalMailing(){
  // Crear fecha inicial (hoy a las 00:00)
  var fechaInicial = new Date();
  fechaInicial.setHours(0, 0, 0, 0);
  
  Logger.log('Ejecución mailing período: ' + fechaInicial);
  actorsSettlement2(fechaInicial, fechaInicial, "Mailing");
}

function modalMailingDebug(){
  // Establecer fecha inicial como hoy a las 00:00
  var fechaInicial = new Date("2025-04-08T00:00:00");
  var fechaFinal = new Date("2025-04-08T00:00:00"); // Fecha y hora actual
  actorsSettlement2(fechaInicial,fechaFinal, "Mailing");
  }

function actorsSettlement2(auxFromDate, auxToDate, modal){
var charProdSheet;
auxToDate.setDate(auxToDate.getDate() + 1);

ssActive = SpreadsheetApp.openById(activeID);
ssNoTrack = SpreadsheetApp.openById(noTrackID);

//Filtrado Character
OpenSht("DWO_Character",1,0,"",0, ssActive);
charNDX=auxNDX;
charValues=auxValues;

//Filtrado CharacterProduction
OpenSht("DWO_CharacterProduction",1,0,"",0, ssActive);
charProdValues=auxValues;
charProdNDX=auxNDX;

//Inicia con Agregados
if(modal==="Settlement"){
  spreadsheet= SpreadsheetApp.getActiveSpreadsheet();
  settlementSheetID=spreadsheet.getId();
  processAggregatedData()
}

//Dialogues 
charProdSheet = auxValues.filter(function(fila) {
  var fechaEnFila = new Date(fila[7]);
  var status = fila[9];
  return fechaEnFila >= auxFromDate && fechaEnFila < auxToDate 
  && status!="(04) Dismissed: DWOCharacterProduction" 
  && status!="(05) Covered by neutral: DWOCharacterProduction" 
  && status!="(09) Unnamed broken down: DWOCharacterProduction" 
  && status!="(07) Covered by previous recording: DWOCharacterProduction" 
  && status!="(06) Covered as ditty: DWOCharacterProduction";
});

if(charProdSheet.length>0) {
  ExtractWorkCompleted(charProdSheet, "Diálogo");
}

const retakeLevels = [1, 2, 3];
const retakeColumns = [17, 18, 19];

retakeLevels.forEach((level, index) => {
  charProdSheet = charProdValues.filter(function(fila) {
    var fechaEnFila = new Date(fila[retakeColumns[index]]); 
    var extraCitation = fila[6];
    return extraCitation > (level - 1) && fechaEnFila >= auxFromDate && fechaEnFila < auxToDate ;
  });

  if (charProdSheet.length > 0) {
    ExtractWorkCompleted(charProdSheet, `Retoma ${level}`);
  }
});

//Songs
//Filtrado DWO_SongDetail
OpenSht("DWO_SongDetail",0,0,"",0, ssActive);
charProdValues = auxValues;
OpenSht("DWO_Song",1,0,"",0, ssActive);
songValues = auxValues;
songNDX = auxNDX;

charProdSheet = charProdValues.filter(function(fila) {
  var fechaEnFila = new Date(fila[18]); 
  return fechaEnFila >= auxFromDate && fechaEnFila < auxToDate;
});

if(charProdSheet.length>0) {
  ExtractWorkCompleted(charProdSheet, "Songs");
}

//Musical directors
if(modal==="Settlement"){
  charProdSheet = songValues.filter(function(fila) {
    var fechaEnFila = new Date(fila[9]); 
    return fechaEnFila >= auxFromDate && 
           fechaEnFila < auxToDate && 
           fila[11] !== "(99) Dismissed: DWOSong";
  });

  if(charProdSheet.length>0) {
    ExtractSongsCompleted(charProdSheet);
  }

  //Dialog Directors
  ExtractEvents(auxFromDate, auxToDate);
}

var auxHTML="";

//Order by name
dataAdult=CheckUnified(dataAdult);
if(modal==="Settlement"){
  SaveSheet(dataAdult, "ADULTOS");
} else {
  if (dataAdult.length > 0) {
    auxHTML = AgregarTablaHTML("ADULTOS", auxHTML, dataAdult, modal);
  }
}
dataRetired=CheckUnified(dataRetired);
if(modal==="Settlement"){
  SaveSheet(dataRetired, "JUBILADOS");
} else {
  if (dataRetired.length > 0) {
    auxHTML=AgregarTablaHTML("JUBILADOS", auxHTML, dataRetired, modal);
  }
}
data1617=CheckUnified(data1617);
if(modal==="Settlement"){
  SaveSheet(data1617, "ENTRE 16 Y 17 ");
} else {
  if (data1617.length > 0) {
    auxHTML=AgregarTablaHTML("ENTRE 16 Y 17", auxHTML, data1617, modal);
  }
}
dataMinor=CheckUnified(dataMinor);
if(modal==="Settlement"){
  SaveSheet(dataMinor, "MENORES");
} else {
  if (dataMinor.length > 0) {
    auxHTML=AgregarTablaHTML("MENORES", auxHTML, dataMinor, modal);
  }
}

if(modal==="Settlement"){
  dataMex=CheckUnified(dataMex);
  SaveSheet(dataMex, "MEXICO");
  SaveSheet(dataUndetermined, "CHECK PENDING");
  SaveSheet(dataDub, "DUB");
  SaveSheet(dataDirDialog, "DIRECTORES DIALOGOS");
  SaveSheet(dataDirMusic, "DIRECTORES CANCIONES");
} else {
  if(dataDub.length >0){
    auxHTML=AgregarTablaHTML("OTROS", auxHTML, dataDub, modal);
  }
  if(auxHTML!=""){
    //SendEmail
    SendEmail.AppSendEmailX("appsheet@mediaaccesscompany.com","ar.info@mediaaccesscompany.com", "", "1SVblIt8tS5pLSoB6lSvgfGDdIzZRSaUBLyDXOYHPIvY", "", "DubApp: Comunicación diaria de grabaciones a actores", "Detalle::"+auxHTML+"||Fecha::"+Utilities.formatDate(auxFromDate, Session.getScriptTimeZone(), 'dd/MM/yyyy'), ccResumenDiario, "");
  }
}
}



function reintentarConEspera(funcion, maxIntentos = 5) {
for (let i = 0; i < maxIntentos; i++) {
  try {
    return funcion();
  } catch (e) {
    if (e.toString().includes('429') && i < maxIntentos - 1) {
      Utilities.sleep(Math.pow(2, i) * 1000); // Espera exponencial
      continue;
    }
    throw e;
  }
}
}
function AgregarTablaHTML(titulo, htmlBase, datos, modal) {
if (!datos || datos.length === 0) {
  return htmlBase;
}

let nuevoHTML = htmlBase;

nuevoHTML += `<h3>${titulo}</h3>`;

// Primera tabla (completa)
nuevoHTML += '<table border="1" style="border-collapse: collapse; width: 100%; font-size: 11px;">';

// Encabezados de la primera tabla
nuevoHTML += '<tr>';
const encabezados = [
  'Nombre', 'DNI', 'Producción', 'Formato', 'Tipo', 'Personaje', 
  'Intervención', 'Loops', 'Monto', 'Email'
];
encabezados.forEach(encabezado => {
  nuevoHTML += `<th style="background-color: #d9d9d9; padding: 8px; font-size: 11px;">${encabezado}</th>`;
});
nuevoHTML += '</tr>';

// Datos de la primera tabla
datos.forEach(fila => {
  nuevoHTML += '<tr>';
  [0, 1, 3, 4, 5, 6, 7, 8].forEach(colIndex => {
    nuevoHTML += `<td style="padding: 8px;">${fila[colIndex]}</td>`;
  });
  // Formatear el monto
  nuevoHTML += `<td style="padding: 8px;">${formatCurrency(fila[9])}</td>`;
  // Email
  nuevoHTML += `<td style="padding: 8px;">${fila[10]}</td>`;
  nuevoHTML += '</tr>';
});
nuevoHTML += '</table><br><br>';

// Segunda tabla para el email - solo enviar si es modo "Mailing"
if (modal === "Mailing") {
  let tablaEmail = '';
  let nombreAnterior = '';
  let emailActual = '';

  datos.forEach((fila, index) => {
    // Si cambia el nombre, enviar email con la tabla acumulada
    if (index > 0 && nombreAnterior !== '' && nombreAnterior !== fila[0] && tablaEmail !== '') {
      const primeraPalabra = nombreAnterior.split(' ')[0];
      const fechaGrabacion = datos[index - 1][2];
      SendEmail.AppSendEmailX(emailActual,"ar.info@mediaaccesscompany.com", nombreAnterior, "1kzuhDW8pmxJVUoYty1AzWtSXUsoXom01rwJhoAirpMg", "", "Media Access Company: Detalle de grabaciones", "Detalle::"+tablaEmail+"||Name::"+primeraPalabra+"||Fecha::"+fechaGrabacion, "", ccoEnvioActor);
      tablaEmail = ''; // Reiniciar tabla
    }

    // Actualizar nombre y email actuales
    nombreAnterior = fila[0];
    emailActual = fila[10];

    // Si es la primera fila o cambió el nombre, iniciar nueva tabla
    if (tablaEmail === '') {
      tablaEmail = '<table border="1" style="border-collapse: collapse; width: 100%; font-size: 11px;">';
      // Encabezados
      tablaEmail += '<tr>';
      const encabezadosEmail = [
        'Producción', 'Formato', 'Tipo', 'Personaje', 
        'Intervención', 'Loops', 'Monto'
      ];
      encabezadosEmail.forEach(encabezado => {
        tablaEmail += `<th style="background-color: #f8f8f8; padding: 8px; font-size: 11px;">${encabezado}</th>`;
      });
      tablaEmail += '</tr>';
    }

    // Agregar fila con las columnas específicas
    tablaEmail += '<tr>';
    [3, 4, 5, 6, 7, 8].forEach(colIndex => {
      tablaEmail += `<td style="padding: 8px;">${fila[colIndex]}</td>`;
    });
    // Formatear el monto
    tablaEmail += `<td style="padding: 8px;">${formatCurrency(fila[9])}</td>`;
    tablaEmail += '</tr>';
  });

  // Enviar el último email si hay datos pendientes
  if (tablaEmail !== '' && nombreAnterior !== '') {
    const primeraPalabra = nombreAnterior.split(' ')[0];
    const fechaGrabacion = datos[datos.length - 1][2];
    SendEmail.AppSendEmailX(emailActual,"ar.info@mediaaccesscompany.com", nombreAnterior, "1kzuhDW8pmxJVUoYty1AzWtSXUsoXom01rwJhoAirpMg", "", "Media Access Company: Detalle de grabaciones", "Detalle::"+tablaEmail+"||Name::"+primeraPalabra+"||Fecha::"+fechaGrabacion, "", ccoEnvioActor);
  }
}

return nuevoHTML;
}


function CheckUnified(matrixParam) {
// Ordenar la matriz por las primeras 6 columnas 
matrixParam.sort(function(a, b) { 
  for (var i = 0; i < 6; i++) { 
    if (a[i] < b[i]) return -1;
    if (a[i] > b[i]) return 1; 
  } 
  return 0;
});

var result = [];
var currentGroup = null;

for (var i = 0; i < matrixParam.length; i++) {
  var key = matrixParam[i].slice(0, 6).join("|");

  // Crear un nuevo grupo si es la primera iteración o la clave ha cambiado
  if (currentGroup === null || key !== currentGroup.key || (!matrixParam[i][7].includes("Diálogo") && !matrixParam[i][7].includes("Retoma"))) {
    // Si existe un grupo actual, agregarlo a los resultados
    if (currentGroup !== null) {
      if (currentGroup.colX > 1) {
        // Llamar a la función Amount

        var aux1 =  convertirStringAFecha(currentGroup.key.split("|")[2]);
        currentGroup.col9 = Amount(currentGroup.col8, convertirStringAFecha(currentGroup.key.split("|")[2]), currentGroup.key.split("|")[5], "", "");
        currentGroup.col6 = currentGroup.col6bis;
      }
      result.push([
        ...currentGroup.key.split("|"),
        currentGroup.col6.join("\n"),
        currentGroup.col7.join("\n"),
        currentGroup.col8,
        currentGroup.col9,
        ...currentGroup.extraColumns
      ]);
    }

    // Iniciar un nuevo grupo
    currentGroup = {
      key: key,
      col6: [matrixParam[i][6]],
      col6bis: [matrixParam[i][6] + " (" + matrixParam[i][8] + ")"],
      col7: [matrixParam[i][7]],
      col8: parseFloat(matrixParam[i][8].toString().replace(',', '.')) || 0,
      col9: parseFloat(matrixParam[i][9].toString().replace(',', '.')) || 0,
      colX: 1,
      extraColumns: matrixParam[i].slice(10) // Capturar columnas adicionales
    };
  } else {
    // Agregar datos al grupo actual
    currentGroup.col6.push(matrixParam[i][6]);
    currentGroup.col7.push(matrixParam[i][7]);
    currentGroup.col6bis.push(matrixParam[i][6] + " (" + matrixParam[i][8] + ")");
    currentGroup.col8 += parseFloat(matrixParam[i][8].toString().replace(',', '.')) || 0;
    currentGroup.col9 += parseFloat(matrixParam[i][9].toString().replace(',', '.')) || 0;
    currentGroup.colX++;
  }
}

// Añadir el último grupo si existe
if (currentGroup !== null) {
  if (currentGroup.colX > 1) {
    // Llamar a la función Amount
    currentGroup.col9 = Amount(currentGroup.col8, currentGroup.key.split("|")[2], currentGroup.key.split("|")[5], "", "");
    currentGroup.col6 = currentGroup.col6bis;
  }
  result.push([
    ...currentGroup.key.split("|"),
    currentGroup.col6.join("\n"),
    currentGroup.col7.join("\n"),
    currentGroup.col8,
    currentGroup.col9,
    ...currentGroup.extraColumns
  ]);
}

// Asegurarse de que cada fila tenga el mismo número de columnas
var maxLength = result.reduce((max, row) => Math.max(max, row.length), 0);
result = result.map(row => {
  while (row.length < maxLength) {
    row.push('');
  }
  return row;
});

return result;
}


function processAggregatedData() {
  var ss = SpreadsheetApp.openById(settlementSheetID);
  var sheetAggregates = ss.getSheetByName('AGREGADOS');
  
  // Verificar si la hoja existe
  if (!sheetAggregates) {
    Logger.log('La hoja AGREGADOS no existe - saltando el procesamiento de agregados');
    return;
  }

  // Verificar si hay datos en la hoja
  var lastRow = sheetAggregates.getLastRow();
  if (lastRow <= 2) {
    Logger.log('No hay datos para procesar en la hoja AGREGADOS');
    return;
  }

  var dataAggregates = sheetAggregates.getRange(3, 1, lastRow - 2, sheetAggregates.getLastColumn()).getValues();

  // Recorrer los datos y copiar a las matrices correspondientes
  for (var i = 0; i < dataAggregates.length; i++) {
    var category = dataAggregates[i][0].toString().trim();

    switch (category) {
      case 'ADULTOS':
        dataAdult.push(dataAggregates[i].slice(1));
        break;
      case 'JUBILADOS':
        dataRetired.push(dataAggregates[i].slice(1));
        break;
      case 'MEXICO':
        dataMex.push(dataAggregates[i].slice(1));
        break;
      case 'ENTRE 16 Y 17':
        data1617.push(dataAggregates[i].slice(1));
        break;
      case 'MENORES':
        dataMinor.push(dataAggregates[i].slice(1));
        break;
      case 'DUB':
        dataDub.push(dataAggregates[i].slice(1));
        break;
      default:
        Logger.log('Categoría no reconocida: ' + category);
    }
  }
}


function ExtractEvents(auxFromDate, auxToDate) { 
const SEARCH_TYPES = {
  RECORDING: "Recording: EventPhase",
  RECORDING2: "Recording #2: EventPhase"
};

const EVENT_STATUS = {
  COMPLETED: "(120) Phase completed: DWOEvent",
  IN_PRODUCTION: "(118) Partial delivery: DWOEvent"
};

// Inicializar arrays usando Set para evitar duplicados
const dialogChannelEventTypNDX = new Set();
const finalChannelEventTypNDX = new Set();
const prelimChannelEventTypNDX = new Set();
const withoutPreliminar = new Set();
const dwoWithoutPreliminar = new Set();

// Obtener producciones sin preliminar
OpenSht("DWO", 0, 0, "", 0, ssActive);
auxValues.forEach(row => {
  if (row[67]==="false") {
    dwoWithoutPreliminar.add(row[1]);
  }
});

// Obtener proyectos sin preliminar
OpenSht("DWO-Production", 0, 0, "", 0, ssActive);
auxValues.forEach(row => {
  if (row[42]?.includes("Excludes preliminar (recording): ProductionAltAttributes")) {
    withoutPreliminar.add(row[0]);
  }
});

// Procesar DWO-ChannelEventType
OpenSht("DWO-ChannelEventType", 0, 0, "", 0, ssActive);

auxValues.forEach(row => {
  if ([SEARCH_TYPES.RECORDING, SEARCH_TYPES.RECORDING2].includes(row[3]) && 
      row[28] === "(01) Enabled: Generic") {
    
    dialogChannelEventTypNDX.add(row[1]);
    if (row[3] === SEARCH_TYPES.RECORDING2) {
      finalChannelEventTypNDX.add(row[1]);
    } else {
      prelimChannelEventTypNDX.add(row[1]);
    }
  }
});

// Procesar DWO-Event
OpenSht("DWO-Event", 0, 59, "(120) Phase completed: DWOEvent", 0, ssActive);

// Filtrar y mapear los valores para obtener solo la columna [1] de los casos que cumplen la condición
const recordingEvents = auxFilteredValues
  .filter(row => prelimChannelEventTypNDX.has(row[3]))
  .map(row => row[1]);

//*** */
var auxfecha; var auxtipoevento; var auxestado; var auxdescartar;
const dialogRecordingEvents = auxFilteredValues.filter(row => {
  const fechaEnFila = new Date(row[15]);

  auxfecha = fechaEnFila >= auxFromDate && fechaEnFila < auxToDate;
  if(auxfecha) {
    auxtipoevento = dialogChannelEventTypNDX.has(row[3]);
    if(auxtipoevento) {
      auxestado = [EVENT_STATUS.COMPLETED, EVENT_STATUS.IN_PRODUCTION].includes(row[58]);
      if(auxestado) {
        auxdescartar = finalChannelEventTypNDX.has(row[3]) && recordingEvents.includes(row[1]);
      }
    }
  }

  // Debug específico
  /*if(row[1]==="23de6b3f5-GyzUb" && dialogChannelEventTypNDX.has(row[3])) {
    Logger.log('Debug - Fecha: ' + auxfecha);
    Logger.log('Debug - Tipo Evento: ' + auxtipoevento);
    Logger.log('Debug - Estado: ' + auxestado);
    Logger.log('Debug - Descartar: ' + auxdescartar);
  }*/

  return auxfecha && auxtipoevento && auxestado && !auxdescartar;
});

// Procesar eventos filtrados
dialogRecordingEvents.forEach(row => {
  const auxProject = labelProject(row[1], row[15]);
  dataDirDialog.push([
    auxProject.dirdialog,
    auxProject.title,
    auxProject.formatType,
    auxProject.contentType,
    auxProject.projectType,
    auxProject.duration,
    auxProject.service,
    row[15],
    auxProject.mainCharMinor,
    auxProject.assistant
  ]);
});
} 

function ExtractSongsCompleted(charProdSheet){
var auxProject;
//Armado de data
for (var i = 0; i < charProdSheet.length; i++) {
  auxCharProd = charProdSheet[i];
  //Director´s song settlement
  var auxSeconds= Time2Seconds(auxCharProd[6]) - Time2Seconds(auxCharProd[5]) - auxCharProd[7];
  auxSeconds = auxSeconds.toString();
  auxProject=labelProject(auxCharProd[1], "");
  dataDirMusic.push([auxProject["dirmusic"], auxProject["title"], auxCharProd[4].replace(": Song_Type",""), auxSeconds, auxCharProd[9]]);
}
}


function ExtractWorkCompleted(charProdSheet, groupRequest){

//Filtrado Character
if(groupRequest==="Songs"){
  var paramCharacter= 3;
} else {
  var paramCharacter= 1;
}

var j; var charProd=[]; var charProdNDX=[]; var caseAux; var caseAux2; var j; var k;
//Recorre línea por línea obteniendo los personajes unicos de la tanda
for (var i = 0; i < charProdSheet.length; i++) {
  if(groupRequest!=="Songs" && charProdSheet[i][13]!==""){
    caseAux = charProdSheet[i][paramCharacter]+"//"+charProdSheet[i][15];
    caseAux2 = charProdSheet[i][paramCharacter];
  } else {
    caseAux = charProdSheet[i][paramCharacter];
    caseAux2=caseAux;
  }
  if(charProdNDX.indexOf(caseAux)!==-1) {
    continue;
  }
  j = charNDX.indexOf(caseAux2);
  if (j != -1) { 
    charProd.push(charValues[j]); 
    charProdNDX.push(caseAux2); 
    // Si hay alt_character
    if(charValues[j][8]!=""){
      if(charProdNDX.indexOf(charValues[j][8])!==-1) {
        continue;
      }
      k = charNDX.indexOf(charValues[j][8]);
      charProd.push(charValues[k]); 
      charProdNDX.push(charValues[j][8]); 
    }
  }
}

var auxCharacter; var auxCharacterName; var auxActor_ID; var auxLoops; var auxAmount; var auxCharacter; var aux1; var auxSong;
var auxCharProd; var auxActorObject; var auxGrupoEtario; var auxProject; var auxDate; var auxGroupRequest;
  
//Armado de data
for (var i = 0; i < charProdSheet.length; i++) {
  auxCharProd = charProdSheet[i];
  
  //Busca char
  if(auxCharProd[paramCharacter]!="") {
    auxCharacter = charProd[charProdNDX.indexOf(auxCharProd[paramCharacter])] ;
    if(groupRequest!=="Songs" && auxCharProd[13]!==""){
      //With parent
      auxCharacterName=auxCharacter[2]+" / "+auxCharProd[15];
    } else {
      //Alt character
      if(auxCharacter[8]!=""){
        auxCharacter = charProd[charProdNDX.indexOf(auxCharacter[8])] ;
      }
      //Armado
      auxCharacterName=auxCharacter[2];
    }
  }

  if(groupRequest==="Songs"){
    if(auxCharProd[paramCharacter]==""){
      auxCharacterName="CHORUS"
    }

    aux1=songNDX.indexOf(auxCharProd[1]);
    if(aux1!=-1){
      auxSong=songValues[aux1];
      if(auxSong[11]==="(99) Dismissed: DWOSong") {continue};
      // Verificar si la canción está cubierta por diálogo
      if(auxSong[8] && auxSong[8].toString().includes("Cover by dialog recoding: Song_ExtraAttrib")) {continue};
      auxGroupRequest = auxCharProd[10].replace(": Songs recording: RateTeam","")+ " / " + auxSong[2];
    //        auxGroupRequest = auxSong[4].replace(": Song_Type","");
    } else {
      auxGroupRequest = "Undetermined";
    }

    if(auxCharProd[5]!=""){
      auxActor_ID=auxCharProd[5];
    } else if (auxCharacter[4].includes("Alt actor (sing): Character_Attributes") && auxCharacter[7]!="") {
      auxActor_ID=auxCharacter[7];
    } else if (auxCharacter[6]!=""){
      auxActor_ID=auxCharacter[6];
    } else {
      // Buscar en DWO_CharacterProduction el actor_id correspondiente
      var characterProductionSheet = SpreadsheetApp.openById(databaseID.getID().activeID).getSheetByName("DWO_CharacterProduction");
      var characterProductionData = characterProductionSheet.getDataRange().getValues();
      var foundActor = false;
      
      // Verificar si ya está cargado en memoria, si no, buscar en la hoja
      for (var k = 1; k < characterProductionData.length; k++) {
        if (characterProductionData[k][1] == auxCharProd[paramCharacter] && characterProductionData[k][2] == auxSong[1]) {
          auxActor_ID = characterProductionData[k][4]; // Columna E contiene el actor_id
          foundActor = true;
          break;
        }
      }
      
      // Si no se encontró, usar un valor predeterminado o dejarlo vacío
      if (!foundActor) {
        auxActor_ID = "";
      }
    }
    auxProject=labelProject(auxSong[1], "");

  } else {
    if(!auxGroupRequest) {
      auxGroupRequest = groupRequest;
    }

    //Actor
    if(auxCharProd[4]!="") {
      auxActor_ID=auxCharProd[4];      
    } else {
      auxActor_ID=auxCharacter[6];
    }
    auxProject=labelProject(auxCharProd[2], "");
  }
  auxActorObject=actorData(auxActor_ID);

  if(groupRequest==="Diálogo") {
    //Date
    auxDate = auxCharProd[7] ;
    //Loops
    if(auxCharProd[13]!==""){
      //With parent
      auxLoops=auxCharProd[3];
    } else {
      //Without parent
      if(auxCharProd[6]<1 && auxCharProd[3]>0){
        auxLoops=auxCharProd[5]+auxCharProd[3];
      } else {
        auxLoops=auxCharProd[5];
      }
    }
    //Amount
    if(auxProject["service"].includes("Audio Description")) {
      var aux2 = "Audio Description";
    } else {
      var aux2 = "LipSync";
    }
    auxAmount=Amount(auxLoops, auxDate, aux2, auxCharProd[16], auxProject["duration"]);

  } else if(groupRequest.includes("Retoma")) {
    //Date
    var auxRetake = parseInt(groupRequest.match(/\d+$/)[0], 10);

    if(auxRetake===1) {
      auxDate=auxCharProd[17];
      auxLoops=auxCharProd[3];
    } else if(auxRetake===2) {
      auxDate=auxCharProd[18];
      auxLoops=auxCharProd[20];
    } else if(auxRetake===3) {
      auxDate=auxCharProd[19];
      auxLoops=auxCharProd[21];
    } 

    //Loops
    auxAmount=Amount(auxLoops, auxDate , "LipSync", "", "");

  } else if(groupRequest==="Songs") {
    auxDate=auxCharProd[18];
    auxAmount = calculateSongAmount(auxCharProd[10], auxDate);
    auxLoops="-";
  }

  //Grupo etario
  if(auxActorObject["retired"]){
    auxGrupoEtario="JUBILADOS";
  } else {
    auxGrupoEtario=grupoEtario(auxActorObject["birth_date"],auxDate); 
  }

  if(auxActorObject["settlement"] === "Dub: Settlement_actor_attributes"){
    dataDub.push([auxActorObject["name"], auxActorObject["ID"], formatDate(auxDate) , auxProject["title"], auxProject["projectType"], auxProject["service"], auxCharacterName, auxGroupRequest, auxLoops, auxAmount, auxActorObject["email"]]);
  } else if(auxActorObject["country"] === "México: Country"){
    dataMex.push([auxActorObject["name"], auxActorObject["ID"], formatDate(auxDate), auxProject["title"], auxProject["projectType"], auxProject["service"], auxCharacterName, auxGroupRequest, auxLoops, auxAmount, auxActorObject["email"]]);
  } else if(auxActorObject["name"] === "Undetermined" || auxAmount===0 || auxActorObject["ID"]===""){
    dataUndetermined.push([auxActorObject["name"], auxActorObject["ID"], formatDate(auxDate), auxProject["title"], auxProject["projectType"], auxProject["service"], auxCharacterName, auxGroupRequest, auxLoops, auxAmount, auxProject["assistant"], auxActorObject["email"]]);
  } else if(auxGrupoEtario === "ADULTOS"){
    dataAdult.push([auxActorObject["name"], auxActorObject["ID"], formatDate(auxDate) , auxProject["title"], auxProject["projectType"], auxProject["service"], auxCharacterName, auxGroupRequest, auxLoops, auxAmount, auxActorObject["email"]]);
  } else if(auxGrupoEtario === "JUBILADOS"){
    dataRetired.push([auxActorObject["name"], auxActorObject["ID"], formatDate(auxDate) , auxProject["title"], auxProject["projectType"], auxProject["service"], auxCharacterName, auxGroupRequest, auxLoops, auxAmount, auxActorObject["email"]]);
  } else if(auxGrupoEtario === "ENTRE 16 Y 17"){
    data1617.push([auxActorObject["name"], auxActorObject["ID"], formatDate(auxDate) , auxProject["title"], auxProject["projectType"], auxProject["service"], auxCharacterName, auxGroupRequest, auxLoops, auxAmount, auxActorObject["email"]]);
  } else {
    dataMinor.push([auxActorObject["name"], auxActorObject["ID"], formatDate(auxDate), auxProject["title"], auxProject["projectType"], auxProject["service"], auxCharacterName, auxGroupRequest, auxLoops, auxAmount, auxActorObject["email"]]);
  }

}
}

function SaveSheet(auxMatrix, auxSheetName) {

//Empareja matriz
auxMatrix=equalizeRows(auxMatrix);

//Access
if(!ssSettlementSheet) {
  ssSettlementSheet = SpreadsheetApp.openById(settlementSheetID);
  CleanFilters(ssSettlementSheet);
}

var sheet = ssSettlementSheet.getSheetByName(auxSheetName); 

if (sheet != null) { // Check if the sheet exists
  var range = sheet.getRange(3, 1, sheet.getLastRow() - 1, sheet.getLastColumn()); 
  range.clearContent(); 
} else {
  Logger.log("Sheet not found: " + sheetName);
}

if(auxMatrix.length===0){return};

//Order data
auxMatrix=orderMatrixByFirstColumn(auxMatrix);

//Save
  sheet.getRange(3, 1, auxMatrix.length, auxMatrix[0].length).setValues(auxMatrix); // Write the matrix to the sheet

}

function equalizeRows(matrix) {
// Encontrar la longitud de la fila más larga
var maxLength = 0;
matrix.forEach(function(row) {
  if (row.length > maxLength) {
    maxLength = row.length;
  }
});

// Agregar columnas vacías a las demás filas para que tengan la misma longitud
matrix = matrix.map(function(row) {
  while (row.length < maxLength) {
    row.push(''); // Puedes cambiar '' por otro valor si lo prefieres
  }
  return row;
});

return matrix;
}


function CleanFilters(ss) {

var currentSheets = ss.getSheets();

for (var i = 0; i < currentSheets.length; i++) {
  var filterAux = currentSheets[i].getFilter();
  if (filterAux !== null) {
    filterAux.remove();
  }
}
}

function Seconds2String(duracion) {
const horas = Math.floor(duracion / 3600);
const minutos = Math.floor((duracion % 3600) / 60);
const segundos = duracion % 60;

// Formateamos los componentes en una cadena hh:mm:ss
const cadenaFormateada = `${horas.toString().padStart(2, '0')}:${minutos.toString().padStart(2, '0')}:${segundos.toString().padStart(2, '0')}`;

return cadenaFormateada;
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

function Time2Seconds(fecha) {
if(!fecha){return fecha}
const hora = fecha.getHours();
const minutos = fecha.getMinutes();
const segundos = fecha.getSeconds();

const totalSegundos = hora * 3600 + minutos * 60 + segundos;
return totalSegundos;
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

function orderMatrixByFirstColumn(matrix) {
// Sort the matrix using a custom comparison function
matrix.sort(function(a, b) {
  // Compare the first elements of each row (a[0] and b[0])
  if (a[0] < b[0]) {
    return -1; // a should come before b
  } else if (a[0] > b[0]) {
    return 1; // b should come before a
  } else {
    return 0; // a and b are equal
  }
});

return matrix; // Return the sorted matrix
}

function calculateSongAmount(songTask, songCompletedDate) {
  // Asegurarse de que los datos de tarifas estén cargados
  // Usamos índice 1 (Col A) para DWO-Rate porque ahí está el Rate ID (PK) que relaciona con RateItem
  if (!rateItemValues || !rateValues || !rateNDX) {
    openSheet("DWO-Rate", 1, "", 0, ssNoTrack);
    rateValues = auxValues;
    rateNDX = auxNDX; // Indexado por Col A de DWO-Rate
    openSheet("DWO-RateItem", 1, "", 0, ssNoTrack); // Este carga RateItem, el índice Col A no es relevante aquí directamente
    rateItemValues = auxValues;
    // No necesitamos rateItemNDX para este cálculo, iteraremos.
  }

  // Convertir fecha completado a Date, usar hoy si es nula/inválida
  let effectiveDate = null;
  try {
      effectiveDate = convertToDate(songCompletedDate);
      // Verificar si la conversión resultó en una fecha válida
      if (!(effectiveDate instanceof Date && !isNaN(effectiveDate.getTime()))) {
          //Logger.log('Fecha de completado inválida o vacía: ${songCompletedDate}. Usando fecha actual.');
          Logger.log('Fecha de completado inválida o vacía: ' + songCompletedDate + '. Usando fecha actual.');
          effectiveDate = new Date();
          effectiveDate.setHours(0, 0, 0, 0); // Estandarizar a medianoche
      }
  } catch (e) {
      //Logger.log('Error al convertir fecha: ${songCompletedDate}. Usando fecha actual. Error: ${e}');
      Logger.log('Error al convertir fecha: ' + songCompletedDate + '. Usando fecha actual. Error: ' + e);
      effectiveDate = new Date();
      effectiveDate.setHours(0, 0, 0, 0);
  }

  // Iterar sobre RateItems
  for (let i = 0; i < rateItemValues.length; i++) {
    const rateItem = rateItemValues[i];
    const itemRateID = rateItem[1];     // Col B: Rate ID (FK a DWO-Rate.RateID)
    const paymentMethod = rateItem[2]; // Col C: Current payment method
    const itemTask = rateItem[3];     // Col D: Task (Asumido, por favor verifica)
    const itemStatus = rateItem[5];   // Col F: Status

    // Verificar Task y Status
    if (itemTask === songTask && itemStatus === "(01) Enabled: Generic") {
      // Encontrar el Rate correspondiente usando el Rate ID de RateItem
      const rateIndex = rateNDX.indexOf(itemRateID.toString()); // Buscamos el FK en el índice del PK de DWO-Rate

      if (rateIndex !== -1) {
        const rate = rateValues[rateIndex];
        const validFromStr = rate[2]; // Col C: Valid from en DWO-Rate
        const validUntilStr = rate[3]; // Col D: Valid until en DWO-Rate

        try {
            const validFrom = convertToDate(validFromStr);
            let validUntil = null;
            if (validUntilStr) {
                validUntil = convertToDate(validUntilStr);
            }

            // Verificar que las fechas sean válidas antes de comparar
            const isValidFrom = validFrom instanceof Date && !isNaN(validFrom.getTime());
            const isValidUntil = !validUntilStr || (validUntil instanceof Date && !isNaN(validUntil.getTime())); // Válido si está vacío o es fecha válida

            if (isValidFrom && isValidUntil) {
                // Comprobar rango de fechas (ValidFrom <= effectiveDate AND (ValidUntil >= effectiveDate OR ValidUntil is blank))
                if (validFrom <= effectiveDate && (!validUntil || validUntil >= effectiveDate)) {
                    // Tarifa válida encontrada!
                    const amount = parseFloat(paymentMethod.toString().replace(',', '.')) || 0;
                    //Logger.log('Tarifa encontrada para Tarea: ${songTask}, Fecha: ${effectiveDate.toISOString()}. Monto: ${amount}');
                    Logger.log('Tarifa encontrada para Tarea: ' + songTask + ', Fecha: ' + effectiveDate.toISOString() + '. Monto: ' + amount);
                    return amount; // Devolver el monto
                }
            } else {
                 //Logger.log('Fechas inválidas en DWO-Rate para Rate ID ${itemRateID}. From: ${validFromStr}, Until: ${validUntilStr}');
                 Logger.log('Fechas inválidas en DWO-Rate para Rate ID ' + itemRateID + '. From: ' + validFromStr + ', Until: ' + validUntilStr);
            }
        } catch (e) {
            //Logger.log('Error al procesar fechas para Rate ID ${itemRateID} (RateItem Col B) y RateItem ID ${rateItem[0]} (RateItem Col A): ${e}');
            Logger.log('Error al procesar fechas para Rate ID ' + itemRateID + ' (RateItem Col B) y RateItem ID ' + rateItem[0] + ' (RateItem Col A): ' + e);
            // Continuar con el siguiente item si hay error de fecha
        }
      } else {
          // Logger.log('Rate ID ${itemRateID} de RateItem no encontrado en DWO-Rate.'); // Puede ser muy verboso
      }
    }
  }

  // Si no se encontró ninguna tarifa válida
  //Logger.log('No se encontró tarifa válida para Tarea: ${songTask}, Fecha: ${effectiveDate.toISOString()}');
  Logger.log('No se encontró tarifa válida para Tarea: ' + songTask + ', Fecha: ' + effectiveDate.toISOString());
  return 0; // Devolver 0 por defecto
}

function labelProject(auxProduction_ID, recordingDate) {
if(!productionValues) {
  openSheet("DWO-Production", 1,"",0, ssActive);
  productionValues = auxValues;
  productionValuesNDX = auxNDX;
}
if(!projectValues) {
  openSheet("DWO", 2,"",0, ssActive);
  projectValues = auxValues;
  projectValuesNDX = auxNDX;
}
if(!userValues) {
  openSheet("App-User", 1,"",0, ssNoTrack);
  userValues = auxValues;
  userNDX = auxNDX;
}

//Busca Production
var auxProductionRow = productionValuesNDX.indexOf(auxProduction_ID);
var auxReference=productionValues[auxProductionRow][3];
var auxProjectID=productionValues[auxProductionRow][1];
if(productionValues[auxProductionRow][57]!="") {
  var auxDuration=productionValues[auxProductionRow][57];
  auxDuration=Time2String(auxDuration);
} else if(productionValues[auxProductionRow][56]!="") {
  var auxDuration=productionValues[auxProductionRow][56];
  auxDuration=Time2String(auxDuration);
} else {
  var auxDuration=productionValues[auxProductionRow][23]*60;
  auxDuration=Seconds2String(auxDuration);
}

var auxCompleted = productionValues[auxProductionRow][13];

//Busca
var auxProjectRow = projectValuesNDX.indexOf(auxProjectID);
var auxFormatType = projectValues[auxProjectRow][3].replace(": FormatType","");
var auxContentType = projectValues[auxProjectRow][4].replace(": ContentType","");
var auxProjectType = projectValues[auxProjectRow][5].replace(": ProjectType","");
var auxServiceType = projectValues[auxProjectRow][16].replace(": Services","");
var auxIncludesPreliminar= projectValues[auxProjectRow][67];

if(auxIncludesPreliminar && productionValues[auxProductionRow][42].includes("Excludes preliminar (recording): ProductionAltAttributes")){
  auxIncludesPreliminar=false;
}

//Dialog dir
if(productionValues[auxProductionRow][44]!="") {
  var auxDirDialog = productionValues[auxProductionRow][44];
} else {
  var auxDirDialog = projectValues[auxProjectRow][28];
}
var auxDir = userNDX.indexOf(auxDirDialog);
if(auxDir!=-1) {
  auxDirDialog=userValues[auxDir][1];
  var auxAssistant = userValues[auxDir][11];
  auxAssistant = userValues[auxDir][1];
}


//Musical dir
/*  if(productionValues[auxProductionRow][7]!="") {
  var auxDirMusic = productionValues[auxProductionRow][7];
} else {*/
  var auxDirMusic = projectValues[auxProjectRow][29];
//}
if(auxDirMusic!="") {
  auxDirMusic = userValues[userNDX.indexOf(auxDirMusic)][1];
}

if(auxFormatType==="Series"){
  var auxTitle=projectValues[auxProjectRow][7].replace(projectValues[auxProjectRow][0]+": ","");
  auxTitle=auxTitle+" / "+projectValues[auxProjectRow][8]+" / Ep. "+auxReference;
} else {
  var auxTitle = projectValues[auxProjectRow][6] + (auxReference === "Main production" ? "" : " / " + auxReference);
}

//If minority check requested
var auxMainCharMinor= false;
if(recordingDate!="") {
  //Check all Characters
  var aux1=charNDX.indexOf(auxProjectID);
  while (aux1!=-1){
    //If Main check actor minority
    if(charValues[aux1][4].includes("Main: Character_Attributes")) {
      var auxActor=actorData(charValues[aux1][6]);
      if(grupoEtario(auxActor["birth_date"], recordingDate)==="MENORES"){
        var auxMainCharMinor= true;
      }        
    }
    var aux1=charNDX.indexOf(auxProjectID, aux1 + 1);
  }
}

return {
    title: auxTitle,
    projectType: auxFormatType+" / "+auxContentType,
    formatType: auxFormatType,
    contentType: auxContentType,
    projectType: auxProjectType, 
    service: auxServiceType,
    dirdialog: auxDirDialog,
    dirmusic: auxDirMusic,
    assistant: auxAssistant,
    duration: auxDuration,
    completed: auxCompleted,
    mainCharMinor: auxMainCharMinor,
    preliminar: auxIncludesPreliminar
    };

}

function grupoEtario(fechaInicio, fechaFin) {

if(!fechaInicio || !fechaFin) {return "ADULTOS"};

// Asegurarse de que las fechas son objetos Date válidos
fechaInicio = new Date(fechaInicio);
fechaFin = new Date(fechaFin);

// Calcular la diferencia en milisegundos
var diferenciaMilisegundos = fechaFin - fechaInicio;

// Convertir a años (aproximado, considerando años bisiestos)
var milisegundosPorAnio = 1000 * 60 * 60 * 24 * 365.25; 
var aniosTranscurridos = Math.floor(diferenciaMilisegundos / milisegundosPorAnio);

if(aniosTranscurridos < 16) {
  return "MENORES";
} else if(aniosTranscurridos < 18) {
  return "ENTRE 16 Y 17";
} else {
  return "ADULTOS";
}
}

function Amount(auxLoops, auxDate, auxService, auxAttribute, auxDuration) {

if (!rateItemValues) {
  openSheet("DWO-Rate", 2, "", 0, ssNoTrack);
  rateValues = auxValues;
  rateNDX = auxNDX;
  openSheet("DWO-RateItem", 1, "", 0, ssNoTrack);
  rateItemValues = auxValues;
  rateItemNDX = auxNDX;
}

auxDate=convertToDate(auxDate);

if (auxService === "Voice Over" || auxService === "LipSync" || auxService === "Lip sync" || auxService === "Audio Description") {
  var auxRow = rateNDX.indexOf("Loops recording: RateTeam");
  while (auxRow != -1) {
    var aux22 = rateValues[auxRow][2];
    var aux23 = rateValues[auxRow][3] ? rateValues[auxRow][3] : null;
    if (aux22 <= auxDate && (!aux23 || aux23 >= auxDate)) {
      var auxRate_ID = rateValues[auxRow][0];
      break;
    }
    auxRow = rateNDX.indexOf("Loops recording: RateTeam", auxRow + 1);
  }
  if (auxService === "Audio Description") {
    auxDuration = String2Seconds(auxDuration) / 60;
    var auxRateAmount;
  }
  auxRow = rateItemNDX.indexOf(auxRate_ID);
  var auxCitation;
  var auxLoopValue;
  while (auxRow != -1) {
    if (rateItemValues[auxRow][0] != auxRate_ID) { continue; }
    if (rateItemValues[auxRow][1] == "Valor Loop: Loops recording: RateTeam") { auxLoopValue = rateItemValues[auxRow][2]; }
    if (rateItemValues[auxRow][1] == "Citación mínima: Loops recording: RateTeam") { auxCitation = rateItemValues[auxRow][2]; }
    if (auxService === "Audio Description") {
      if (auxDuration >= 90 && rateItemValues[auxRow][1] == "Audio descripción (90´-104´): Loops recording: RateTeam") {
        auxRateAmount = rateItemValues[auxRow][2];
      } else if (auxDuration >= 75 && auxDuration < 90 && rateItemValues[auxRow][1] == "Audio descripción (75'-89´): Loops recording: RateTeam") {
        auxRateAmount = rateItemValues[auxRow][2];
      } else if (auxDuration >= 60 && auxDuration < 75 && rateItemValues[auxRow][1] == "Audio descripción (60'-74´): Loops recording: RateTeam") {
        auxRateAmount = rateItemValues[auxRow][2];
      } else if (auxDuration >= 45 && auxDuration < 60 && rateItemValues[auxRow][1] == "Audio descripción (45'-59´): Loops recording: RateTeam") {
        auxRateAmount = rateItemValues[auxRow][2];
      } else if (auxDuration >= 30 && auxDuration < 45 && rateItemValues[auxRow][1] == "Audio descripción (30'-44´): Loops recording: RateTeam") {
        auxRateAmount = rateItemValues[auxRow][2];
      } else if (auxDuration >= 15 && auxDuration < 30 && rateItemValues[auxRow][1] == "Audio descripción (15´-29´): Loops recording: RateTeam") {
        auxRateAmount = rateItemValues[auxRow][2];
      } else {
        auxRateAmount = 0;
      }
    }
    auxRow = rateItemNDX.indexOf(auxRate_ID, auxRow + 1);
  }

  //AD Rate comparison
  if (auxService === "Audio Description") {
    if (auxLoops < 6 ? auxCitation : auxCitation + ((auxLoops - 6) * auxLoopValue) > auxRateAmount) {
      return (auxLoops < 6 ? auxCitation : auxCitation + ((auxLoops - 6) * auxLoopValue));
    } else {
      return auxRateAmount;
    }
  } else if (auxLoops < 7) {
    // Check if minimum is required
    ////var auxActorProd=actorProdNDX.indexOf(auxActorProdKey);
    ////auxActorProd=actorProd[auxActorProd][2];
    if (auxAttribute.includes("No minimum citation loops: CharacterProduction_Attributes")) {
      return (auxLoops * auxLoopValue);
    } else {
      return auxCitation;
    }
  } else {
    return auxCitation + ((auxLoops - 6) * auxLoopValue);
  }
}
}

function convertirStringAFecha(dateString) {
var partes = dateString.split('/');
var dia = parseInt(partes[0], 10);
var mes = parseInt(partes[1], 10) - 1; // Los meses en JavaScript van de 0 a 11
var año = parseInt(partes[2], 10) + 2000; // Convertir el año a formato completo

return new Date(año, mes, dia);
}


function actorData(auxActorID) {

auxActorID=auxActorID.toString();

if(!actorValues) {
  openSheet("DWO_Actor", 1,"",0, ssNoTrack);
  actorValues = auxValues;
  actorNDX = auxNDX;
}
var auxActorRow = actorNDX.indexOf(auxActorID);
if(auxActorRow != -1 ){
  
  return {
    name: actorValues[auxActorRow][1]+" "+actorValues[auxActorRow][2],
    birth_date: actorValues[auxActorRow][16],
    country: actorValues[auxActorRow][11],
    ID: actorValues[auxActorRow][13],
    settlement: actorValues[auxActorRow][31],
    email:actorValues[auxActorRow][6],
    retired: actorValues[auxActorRow][17].includes("Jubilado: Actor_attributes")
    };
} else {
  return {
    name: "Undetermined",
    surname: "",
    birth_date: "",
    country: "",
    ID: "",
    settlement: "",
    email:"",
    retired: false
    };
}

}

function openSheet(sheetNameAux, ndxCol, key, ndxCol2, ss) {
//openSheet("Sheet-to-load", ndx-col, key-value-to-filter, col-to-filter, sheet)
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

function OpenSht(sheetNameAux, ndxColValues, keyCol, keyValue, ndxColFiltered, ss) {
// auxSheet / For append
// auxValues (complete load)
// ndxColValues > 0 -> auxNDX
// auxFilteredValues if keyCol & keyValue >""
// ndxColFiltered if keyCol & keyValue >"" & ndxColFiltered > 0 -> auxNDX2

// auxRow if key-value-to-filter <> "" and col-to-filter > 0 and result = 1
  auxSheet = ss.getSheetByName(sheetNameAux);
  var lastRow = auxSheet.getLastRow();

  if (lastRow === 1) {
    auxRow = -1; auxValues = []; auxNDX = []; auxNDX2 = []; auxFilteredValues = [];
    return;
  }

  var lastCol = auxSheet.getLastColumn();
  auxValues = auxSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  if (ndxColValues > 0) {
    auxNDX = auxValues.map(r => r[ndxColValues - 1].toString());
  }

  if (keyCol !== 0 && keyValue!="") {
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


function processAdultCases(){
  sendMarkedEmails("ADULTOS");
}

function processMinorCases(){
  sendMarkedEmails("MENORES");
}

function processRetiredCases(){
  sendMarkedEmails("JUBILADOS");
}

function process1617Cases(){
  sendMarkedEmails("ENTRE 16 Y 17 ");
}

function sendMarkedEmails(sheetName) {
  var sheetNameResumen = 'RESUMEN';

  // Obtener la hoja activa
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetResumen = ss.getSheetByName(sheetNameResumen);
  var sheet2Work = ss.getSheetByName(sheetName);

  // Validar que las hojas existan
  if (!sheetResumen || !sheet2Work) {
    throw new Error('No se encontró la hoja ' + (!sheetResumen ? sheetNameResumen : sheetName));
  }

  var data = sheet2Work.getDataRange().getValues();
  
  // Verificar si hay casos marcados
  var hayCasosMarcados = data.slice(2).some(row => row[11] === true && row[10]);
  
  if (!hayCasosMarcados) {
    throw new Error('No hay casos seleccionados');
  }

  // Abrir la hoja de cálculo y obtener las fechas de la sheet RESUMEN
  var fechaDesde = formatDate(sheetResumen.getRange('C5').getValue());
  var fechaHasta = formatDate(sheetResumen.getRange('C7').getValue());
  var fechasAux = fechaDesde + " - " + fechaHasta;

  // Inicializar variables
  var groups = {};

  // Mantener un registro de las filas que necesitan actualización
  var rowsToUpdate = [];
  
  // Recorrer los datos (empezando desde la fila 3)
  for (var i = 2; i < data.length; i++) {
    if (data[i][11] === true && data[i][10]) {
      // Guardar el número de fila para actualizar después
      rowsToUpdate.push(i + 1); // +1 porque las filas en Sheets empiezan en 1
      
      var key = data[i][0] + "|" + data[i][1];
      var nombre = data[i][0].split(' ')[0]; // Primera palabra del string Actor
      
      if (!groups[key]) {
        groups[key] = {
          actor: data[i][0],
          nombre: nombre,
          dni: formatInteger(data[i][1]), // Formatear como número entero
          email: data[i][10],
          sumColumnJ: 0,
          count: 0,
          detailAux: ''
        };
      }
      
      // Agregar los valores a detailAux como una fila de tabla HTML
      groups[key].detailAux += '<tr>';
      // Formatear la fecha de la columna C (índice 2) y omitir las columnas E y F (índices 4 y 5)
      groups[key].detailAux += '<td style="border: 1px solid black; padding: 5px; text-align: right; font-size: 12px;">' + formatDate(data[i][2]) + '</td>'; // Fecha grabación
      groups[key].detailAux += '<td style="border: 1px solid black; padding: 5px; font-size: 12px;">' + data[i][3] + '</td>'; // Proyecto
      groups[key].detailAux += '<td style="border: 1px solid black; padding: 5px; font-size: 12px;">' + replaceNewlineWithBreak(data[i][6]) + '</td>'; // Personaje
      groups[key].detailAux += '<td style="border: 1px solid black; padding: 5px; font-size: 12px;">' + replaceNewlineWithBreak(data[i][7]) + '</td>'; // Tipo intervención
      groups[key].detailAux += '<td style="border: 1px solid black; padding: 5px; text-align: right; font-size: 12px;">' + data[i][8] + '</td>'; // Loops
      groups[key].detailAux += '<td style="border: 1px solid black; padding: 5px; text-align: right; font-size: 12px;">' + formatCurrency(data[i][9]) + '</td>'; // Monto
      groups[key].detailAux += '</tr>';
      
      // Sumar el valor de la columna J (índice 9) a sumColumnJ
      groups[key].sumColumnJ += data[i][9];
      groups[key].count++;
    }
  }

  // Crear la tabla HTML en detalleAux para cada grupo
  for (var key in groups) {
    groups[key].sumColumnJ = formatCurrency(groups[key].sumColumnJ); // Formatear como moneda
    groups[key].detailAux = '<table style="border-collapse: collapse; border: 1px solid black; font-size: 12px;"><tr><th style="border: 1px solid black; padding: 5px; font-size: 12px;">Fecha grabación</th><th style="border: 1px solid black; padding: 5px; font-size: 12px;">Proyecto</th><th style="border: 1px solid black; padding: 5px; font-size: 12px;">Personaje</th><th style="border: 1px solid black; padding: 5px; font-size: 12px;">Tipo intervención</th><th style="border: 1px solid black; padding: 5px; font-size: 12px;">Loops</th><th style="border: 1px solid black; padding: 5px; font-size: 12px;">Monto</th></tr>' + 
                            groups[key].detailAux + '</table>';
    // Llamar a la función Dummy con los datos agrupados
    var paramaAux = "Name::" + groups[key].nombre + "||Detalle liquidacion::" + groups[key].detailAux + "||Monto::" + groups[key].sumColumnJ + "||NyA::" + groups[key].actor + "||desdeHasta::" + fechasAux + "||DNI::" + groups[key].dni;
    //SendEmail.AppSendEmailX("appsheet@mediaaccesscompany.com", "ar.info@mediaaccesscompany.com", "appsheet@mediaaccesscompany.com", "1KbeZbRH5qBBvZEzsK1qNa87CdydgqsAjTvmzm1aUrWY", "", "Media Access Company - Detalle liquidación noviembre 2024", paramaAux, "appsheet@mediaaccesscompany.com",""  );
    SendEmail.AppSendEmailX(groups[key].email, "ar.info@mediaaccesscompany.com", groups[key].actor, "1KbeZbRH5qBBvZEzsK1qNa87CdydgqsAjTvmzm1aUrWY", "", "Media Access Company - Detalle liquidación", paramaAux, "", "appsheet@mediaaccesscompany.com , paula.nunez@mediaaccesscompany.com");
    //SendEmail.AppSendEmailX("appsheet@mediaaccesscompany.com", "", groups[key].actor, "1KbeZbRH5qBBvZEzsK1qNa87CdydgqsAjTvmzm1aUrWY", "", "Media Access Company - Detalle liquidación noviembre 2024", paramaAux, "appsheet@mediaaccesscompany.com",""  );
  }

  // Actualizar solo la columna L de las filas procesadas
  if (rowsToUpdate.length > 0) {
    rowsToUpdate.forEach(rowNum => {
      sheet2Work.getRange(rowNum, 12).setValue(false); // Columna L es 12
    });
  }
}


function replaceNewlineWithBreak(text) 
{ return text.replace(/\n/g, '<br>'); }

function formatDate(date) { 
var d = new Date(date); var day = String(d.getDate()).padStart(2, '0'); var month = String(d.getMonth() + 1).padStart(2, '0'); // Los meses son de 0 a 11 
var year = String(d.getFullYear()).slice(-2); 
return day + '/' + month + '/' + year; 
} 

function formatCurrency(value) 
{ return value.toLocaleString('es-AR', { style: 'currency', currency: 'ARS', minimumFractionDigits: 2 }); } 

function formatInteger(value) 
{ return Math.round(value).toString(); }

function convertToDate(auxDate) {
if (typeof auxDate === 'string') {
  var parts = auxDate.split('/');
  var day = parseInt(parts[0], 10);
  var month = parseInt(parts[1], 10) - 1;
  var year = parseInt(parts[2], 10);
  
  // Verificar que el año no sea ya >= 2025
  if (year < 100) {
    year += 2000;
  }
  
  return new Date(year, month, day);
}
return auxDate;
}