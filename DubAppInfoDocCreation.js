// Cargar configuración desde databaseID.js
const config = databaseID.getID();
const FolderDestiny = DriveApp.getFolderById(config.infoDocFolder);
const googleDocTemplate = DriveApp.getFileById(config.infoDocTemplate);
const databaseID = config.activeID;

// Abre DWO
const dwo = SpreadsheetApp.openById(databaseID).getSheetByName("DWO");
const dwoRange = dwo.getRange(1,1,dwo.getLastRow(),dwo.getLastColumn());
const dataDWO = dwoRange.getValues();
const dataDWONDX = dataDWO.map(r => r[1]);

// Abre DWO-Production
const dwoProd = SpreadsheetApp.openById(databaseID).getSheetByName("DWO-Production");
const dwoProdRange = dwoProd.getRange(1,1,dwoProd.getLastRow(),dwoProd.getLastColumn());
const dataDwoProd = dwoProdRange.getValues();
const dataDwoProdNDX = dataDwoProd.map(function(r){ return r[1]; });

// Abre User
const appUser = SpreadsheetApp.openById(databaseID).getSheetByName("App-User");
const appUserRange = appUser.getRange(1,1,appUser.getLastRow(),appUser.getLastColumn());
const dataappUser = appUserRange.getValues();
const dataappUserNDX = dataappUser.map(function(r){ return r[0]; });

var editors = ["sandra.brizuela@mediaaccesscompany.com"];
var viewers = [];

function Llamador() {
  InfoDoc("8275370E-EA55-4EE7-81D3-16DCAD6E80C0");
}

/*function CreateDoc(projectName, pmName, editors, viewers ){*/
function InfoDoc(projectID) {
  if (!projectID) {
    console.error('Project ID es requerido');
    return null;
  }

  console.log(`projectID: ${projectID}`);
  const rowNum = dataDWONDX.indexOf(projectID, 0);
  
  if (rowNum === -1) {
    Logger.log(`Invalid project ID: ${projectID}`);
    return null;
  }

  // Leer datos del proyecto
  const projectData = dataDWO[rowNum];
  const projectName = `${projectData[6]}${projectData[7]}`;
  const services = projectData[16].replace(': Services', '');
  
  // Procesar usuarios y permisos
  const pmMail = userCorpoEmail(projectData[35], viewers);
  const pmName = pmMail !== -1 ? dataappUser[pmMail][1] : 'PM No encontrado';

  // Agregar editores
  [projectData[33], projectData[34], projectData[54]].forEach(email => {
    userCorpoEmail(email, editors);
  });

  // Agregar viewers
  [projectData[28], projectData[29]].forEach(email => {
    const aux = userCorpoEmail(email, viewers);
    if (aux !== -1) {
      userCorpoEmail(dataappUser[aux][11], viewers);
    }
  });

  // Optimizar el bucle de producciones alternativas
  const processAlternative = (row, type, email, list) => {
    if (row[42].includes(type)) {
      userCorpoEmail(email, list);
    }
  };

  let currentIndex = 0;
  while ((currentIndex = dataDwoProdNDX.indexOf(projectID, currentIndex)) !== -1) {
    const prodRow = dataDwoProd[currentIndex];
    
    if (prodRow[42]) {
      processAlternative(prodRow, 'Alt AD Scriptwriter', prodRow[32], editors);
      processAlternative(prodRow, 'Alt Translator', prodRow[29], editors);
      // ... procesar otros alternativos ...
    }
    currentIndex++;
  }

  var docID = projectData[70];
  if(docID =="" || docID=="PENDING"){
    //Copy template
    var doc = googleDocTemplate.makeCopy(`TempFileCopy`, FolderDestiny);
    //Customize copy
    var docID = doc.getId();

    dwo.getRange(rowNum + 1, 71).setValue(docID);
    doc.setName("DubApp - INFO - "+projectName)
    const docEdit = DocumentApp.openById(docID);
    const body = docEdit.getBody();
    body.replaceText('PROJNAME', projectName);
    body.replaceText('PMNAME', pmName);
    body.replaceText('REQSERV', services); 
  } else {
    //Only Grant
    var doc = DocumentApp.openById(docID);
  }
  doc.addEditors(editors);
  doc.addViewers(viewers);
  return docID;
}

String.prototype.replaceAll = function(search, replacement) {
  return this.replace(new RegExp(search, 'g'), replacement);
};

function userCorpoEmail(userEmail, grantList) {
  if (!userEmail) return -1;
  
  const auxUserF = dataappUserNDX.indexOf(userEmail, 0);
  if (auxUserF === -1) return -1;

  const userCorpoEmail = dataappUser[auxUserF][2];
  if (userCorpoEmail && !grantList.includes(userCorpoEmail)) {
    grantList.push(userCorpoEmail);
  }
  return auxUserF;
}

