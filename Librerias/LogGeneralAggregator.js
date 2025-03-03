/*
DESCRIPCIÓN GENERAL:
Este script maneja la sincronización y registro de cambios entre el Aggregator y DubApp, 
específicamente para programas, episodios y trabajos.

FUNCIONES PRINCIPALES:

StatusLog(e)
- Punto de entrada principal que maneja cambios en las hojas AggProgram, AggEpisode y AggWork
- Detecta cambios en columnas específicas y registra los cambios

appendProgram/Episode/Work(sht, appLogSht, activeRow) 
- Registran cambios de estado en los logs correspondientes

dubAppCheck(rowCase)
- Verifica y sincroniza datos entre Aggregator y DubApp
- Maneja la creación/actualización de proyectos y producciones

refreshDWO/Production(data, row)
- Actualizan datos en DubApp basados en cambios del Aggregator

ÁRBOL DE LLAMADAS:
StatusLog
├── appendProgram
├── appendEpisode  
├── appendWork
└── dubAppCheck
    ├── loadProgramValues
    ├── loadEpisodeValues
    ├── dubAppCheckProject
    │   ├── dubAppProject
    │   └── refreshDWO
    │       ├── glossary
    │       ├── dubAppSeries
    │       └── addContask
    └── dubAppCheckProduction
        ├── dubAppProduction
        └── refreshProduction
            └── addContask

FUNCIONES DE SOPORTE:
- glossary: Maneja equivalencias entre Aggregator y DubApp
- openSheet: Abre y carga datos de hojas
- newKey: Genera IDs únicos
*/

//Global declaration
  //Flexible loading
  var ssActiveDubApp = null;
  var dubAppTotalID="1d258eRpy3NovzQ1wKX_WVRCfXhxc2u2pwpYUuTDyfIw";
  var dubAppIDControl = "1P-kQdID7dwG4UUezuF0QuJq7G3swDaAwaznBNwnNvCk";
  var aggregatorID ="1sNOJ0f0yYDMqxKo97C7tx1npwxuyBP0-BmnXrFIGEzA";

  var Contask = null;

  var programSheet = null;
  var programValues = null;
  var programNDX = null;
  var programRow = null;

  var episodeSheet = null;
  var episodeValues = null;
  var episodeNDX = null;
  var episodeRow = null;

  var productionSheet = null;
  var productionValues = null;
  var productionNDX = null;
  var productionNDX2 = null;
  var productionNDX3 = null;
  var productionRow = null;

  var seriesSheet = null;
  var seriesValues = null;
  var seriesNDX = null;

  var glossarySheet = null;
  var glossaryValues = null;
  var glossaryNDXAgg = null;
  var glossaryNDXDubApp = null;

  var dwoSheet = null;
  var dwoValues = null;
  var dwoNDX = null;
  var dwoNDX2 = null;
  var dwoRow = null;
  var dwoProjectID = null;
  var dwoFormat =  null;
  var dwoEstimatedTRT = null;

  var sSht = null;
  var sht = null;
  var shtName = null;

  var timezone = "GMT-3";
  var timestamp_format = "dd/MM/yyyy HH:mm:ss";
  var tmpTimeStamp = Utilities.formatDate(new Date(), timezone, timestamp_format);

  var lastUser = null;

  var debug=false;

  var verbose=true;



function StatusLog(e) {
  if(debug) {
    shtName = "AggProgram";
    sSht = SpreadsheetApp.openById(aggregatorID);
    sht = sSht.getSheetByName(shtName);
    var rngW = 1;
    var activeCol = 26;
    var activeRow = 1006;
    var rngStartIndex = 0;
    var rngEndIndex = 0;
  } else {
    sSht = e.source;
    sht = sSht.getActiveSheet();
    shtName = sht.getName();
    var activeRng = sht.getActiveRange();
    var rngW = activeRng.getWidth();
    var activeCol = activeRng.getColumn();
    var activeRow = activeRng.getRow();
    var rngStartIndex = activeCol;
    var rngEndIndex = activeCol + rngW - 1;
  }
  var appLogSht;

  if (verbose) {console.log(shtName+" // "+activeRow)};

  if(activeRow==1) {return};

  if (shtName === "AggProgram") {
    appLogSht = sSht.getSheetByName("AggLogProgram");
    if ((rngW == 1 && activeCol == 26) || (rngStartIndex <= 26 && rngEndIndex >= 26)) {
      appendProgram(sht, appLogSht, activeRow);
    }
    var dubAppRef = sht.getRange(activeRow,48).getValue();
    if(dubAppRef!="") {
      if (verbose) {console.log("dubAppCheck(activeRow) "+activeRow)};
      dubAppCheck(activeRow);
    }

  } else if (shtName === "AggEpisode") {
    appLogSht = sSht.getSheetByName("AggLogEpisode");
    if ((rngW == 1 && activeCol == 31) || (rngStartIndex <= 31 && rngEndIndex >= 31)) {
      appendEpisode(sht, appLogSht, activeRow);
    }
    dubAppCheck(activeRow);
    
  } else if (shtName === "AggWork") {
    appLogSht = sSht.getSheetByName("AggLogWork");
    if ((rngW == 1 && activeCol == 41) || (rngStartIndex <= 41 && rngEndIndex >= 41)) {
      appendWork(sht, appLogSht, activeRow);
    }
  }
}

function appendProgram(sht, appLogSht, activeRow) {
  var tmpStatus = sht.getRange(activeRow, 26).getValue();
  var tmpStatusLog = sht.getRange(activeRow, 33).getValue();
  if (tmpStatus != tmpStatusLog) {
    var tmpProgramID = "'" + sht.getRange(activeRow, 1).getValue().toString();
    var tmpUserMail = sht.getRange(activeRow, 28).getValue();
    var tmpStatus = sht.getRange(activeRow, 26).getValue();

    var tmpTypeChange = "Status";
    var tmpComment = sht.getRange(activeRow, 24).getValue();
    var tmpNextstamp;
    appLogSht.appendRow([tmpTimeStamp, tmpProgramID, tmpUserMail, tmpStatus, tmpTypeChange, tmpComment, tmpNextstamp]);
    if (verbose) {console.log("Append Program log "+tmpProgramID)};
    sht.getRange(activeRow, 33).setValue(tmpStatus);
  }
}

function appendEpisode(sht, appLogSht, activeRow) {
  var tmpStatus = sht.getRange(activeRow, 31).getValue();
  var tmpStatusLog = sht.getRange(activeRow, 42).getValue();
  if (tmpStatus != tmpStatusLog) {
    var tmpProdNumber = sht.getRange(activeRow, 2).getValue();
    var tmpUserMail = sht.getRange(activeRow, 33).getValue();

    var tmpTypeChange = "Status";
    var tmpComment = sht.getRange(activeRow, 30).getValue();
    var tmpNextstamp;
    appLogSht.appendRow([tmpTimeStamp, tmpProdNumber, tmpUserMail, tmpStatus, tmpTypeChange, tmpComment, tmpNextstamp]);
    if (verbose) {console.log("Append Episode log "+tmpProdNumber)};
    sht.getRange(activeRow, 42).setValue(tmpStatus);
    sht.getRange(activeRow, 30).setValue("");
  }
}

function appendWork(sht, appLogSht, activeRow) {
  var tmpStatus = sht.getRange(activeRow, 41).getValue();
  var tmpStatusLog = sht.getRange(activeRow, 53).getValue();
  if (tmpStatus != tmpStatusLog) {
    var tmpLocalizacion = sht.getRange(activeRow, 1).getValue() + ": " + sht.getRange(activeRow, 2).getValue();
    var tmpUserMail = sht.getRange(activeRow, 43).getValue();

    var tmpTypeChange = "Status";
    var tmpComment = sht.getRange(activeRow, 40).getValue();
    var tmpNextstamp;
    appLogSht.appendRow([tmpTimeStamp, tmpLocalizacion, tmpUserMail, tmpStatus, tmpTypeChange, tmpComment, tmpNextstamp]);
    if (verbose) {console.log("Append Work log "+tmpLocalizacion)};
    sht.getRange(activeRow, 53).setValue(tmpStatus);
    sht.getRange(activeRow, 40).setValue("");
  }
}

function dubAppCheck(rowCase) {
  if (verbose) {console.log("dubAppCheck(rowCase) "+rowCase)};
  if (programValues === null) {
    loadProgramValues();
  }

  if (shtName !== "AggProgram") {
    if (episodeValues === null) {
      loadEpisodeValues();
    }
    var auxRowCase = rowCase;
    var auxKey = episodeValues[rowCase - 2][0]+"";
    rowCase = programNDX.indexOf(auxKey) + 2;
  }

  if (shtName !== "AggProgram" || ( shtName === "AggProgram" && programValues[rowCase - 2][47] !== "")) {
    var auxCase = programValues[rowCase - 2];
    if(auxCase[25]==="2) DWO recibido") {
      if (dubAppCheckProject(rowCase, auxCase) && shtName !== "AggProgram") {
        var auxCase = episodeValues[auxRowCase - 2];
        dubAppCheckProduction(auxRowCase, auxCase);
      }
    }
  }
}

function loadProgramValues() {
  programSheet = sSht.getSheetByName("AggProgram");
  var lastRow = programSheet.getLastRow();
  var lastCol = programSheet.getLastColumn();
  var auxData = programSheet.getRange(2, 1, lastRow - 1, lastCol);
  programValues = auxData.getValues();
  programNDX = programValues.map(function (r) { return r[0].toString(); });
}

function loadEpisodeValues() {
  episodeSheet = sSht.getSheetByName("AggEpisode");
  var lastRow = episodeSheet.getLastRow();
  var lastCol = episodeSheet.getLastColumn();
  var auxData = episodeSheet.getRange(2, 1, lastRow - 1, lastCol);
  episodeValues = auxData.getValues();
  episodeNDX = episodeValues.map(function (r) { return r[1].toString(); });
}

function dubAppCheckProject(rowCase, dataCase) {
  var isLinkable = false;
  var servicesCase = dataCase[37];
  servicesCase = servicesCase.split(" , ");

  if (dataCase[21] == "Non Stop" && (servicesCase.indexOf("Dub") != -1 || servicesCase.indexOf("Voice Over") != -1 || servicesCase.indexOf("Audio description") != -1 || servicesCase.indexOf("Metadata") != -1)) {
    let dubProjectAuxRow = dubAppProject(dataCase[0]);
    let dubProjectAuxKey = refreshDWO(dataCase, dubProjectAuxRow);
    if (dataCase[48] === "") {
      programSheet.getRange(rowCase, 48).setValue("Active");
      programSheet.getRange(rowCase, 49).setValue(dubProjectAuxKey);
    }
    isLinkable=true;
  }
  if (verbose) {console.log("Project elegible "+isLinkable)};
  return isLinkable;
}

function dubAppCheckProduction(rowCase, dataCase) {
  dubProductionAuxRow = dubAppProduction(dataCase[1])
  dubProductionAuxKey = refreshProduction(dataCase, dubProductionAuxRow);
  if (dataCase[64] === "") {
    episodeSheet.getRange(rowCase, 64).setValue("Active");
    episodeSheet.getRange(rowCase, 65).setValue(dubProductionAuxKey);
  }
}

function refreshProduction(aggEpisodeCase, dubProductionRow) {
  var auxChangeKey = null;
  lastUser = aggEpisodeCase[32];
  var auxProductionNumber = aggEpisodeCase[1] + "";
  var auxType = dwoFormat === "Series: FormatType" ? "Episode: Series: FormatType" : "Movie: Standalone: FormatType";
  var auxEstimatedTRT = dwoEstimatedTRT;
  var auxReference = "F"+Utilities.formatString("%003d", aggEpisodeCase[3]); 
  var auxEpisodeTitle = aggEpisodeCase[2];
  var auxNomenclature = aggEpisodeCase[65];
  var auxTitleSpa = aggEpisodeCase[5];

  if (dubProductionRow === -1) {
    dubAppProduction("");
    var auxProductionID = auxProductionNumber + "-" + newKey(5);
    var buffer = [auxProductionID, dwoProjectID, auxType, auxReference, auxEpisodeTitle, null, null, null, null, auxNomenclature, null, null, null, null, null, null, null, null, null, auxTitleSpa, null, null, null, auxEstimatedTRT, "Unblock", true, null, null, null, null, false, null, null, null, "(01) Waiting for assets: DWOProduct", lastUser, tmpTimeStamp, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, auxProductionNumber, null, null, null, null];
    productionSheet.appendRow(buffer);
    addContask("DWO-Production", auxProductionID, "INSERT_ROW");
    if (verbose) {console.log("Append DubApp Production "+auxProductionID)};
    auxChangeKey = auxProductionID;
  } else {
    var dubappProductionCase = productionValues[dubProductionRow];
    var updateNeeded = false;
    auxChangeKey = dubappProductionCase[0];
    
    if (auxEstimatedTRT !== dubappProductionCase[23]) {
      productionSheet.getRange(dubProductionRow + 2, 24).setValue(auxEstimatedTRT);
      updateNeeded = true;
    }
    if (auxReference !== dubappProductionCase[3]) {
      productionSheet.getRange(dubProductionRow + 2, 4).setValue(aggEpisodeCase[3]);
      updateNeeded = true;
    }
    if (auxEpisodeTitle !== dubappProductionCase[4]) {
      productionSheet.getRange(dubProductionRow + 2, 5).setValue(aggEpisodeCase[2]);
      updateNeeded = true;
    }
    if (auxNomenclature !== dubappProductionCase[9]) {
      productionSheet.getRange(dubProductionRow + 2, 10).setValue(aggEpisodeCase[65]);
      updateNeeded = true;
    }
    if (auxTitleSpa !== dubappProductionCase[19]) {
      productionSheet.getRange(dubProductionRow + 2, 20).setValue(aggEpisodeCase[5]);
      updateNeeded = true;
    } 

    if (updateNeeded) {
      addContask("DWO-Production", auxChangeKey, "EDIT");
      if (verbose) {console.log("Update DubApp Production "+auxChangeKey)};
    }
  }

  return auxChangeKey;
}

function refreshDWO(aggProgramCase, dubProjectRow) {

  lastUser = aggProgramCase[27];
  var auxChannel = glossary("Aggregator", "Channel" + aggProgramCase[30]);
  var auxProgramID = aggProgramCase[0] + "";
  dwoFormat = glossary("Aggregator", "Format" + aggProgramCase[4]);
  var auxSerie = "";
  var auxSeason = "";
  var auxTitle = "";

  if (dwoFormat == "Series: FormatType") {
    var arraySerie = aggProgramCase[1].split(" // ");
    auxSerie = arraySerie[0];
    auxSeason = arraySerie[1];
    var aux = dubAppSeries(auxChannel, auxSerie, aggProgramCase[2]);
  } else {
    auxTitle = aggProgramCase[1];
  }

  var auxTitleTranslated = aggProgramCase[2];
  var auxQCCreative = aggProgramCase[28];
  var auxContent = glossary("Aggregator", "Content" + aggProgramCase[39]);
  dwoEstimatedTRT = glossary("Aggregator", "Duration" + aggProgramCase[7]);
  var auxMailDataOut = glossary("Aggregator", "MailDataOut");
  var auxDuration = !aggProgramCase[7] ? "" : aggProgramCase[7]+": Duration option";
  var auxEpisodes = aggProgramCase[17];
  var arrayServices = aggProgramCase[37].split(" , ");
  var auxServices = arrayServices.map(function(service) { return glossary("Aggregator", "Services" + service); }).join(" / ");
  var auxOriginalLanguage = glossary("Aggregator", "OrigLang" + aggProgramCase[5]);

  if (dubProjectRow === -1) {
    dwoProjectID = auxProgramID + "-" + newKey(30);
    dubAppProject("");
    var buffer = [auxChannel, dwoProjectID, "Argentina", dwoFormat, auxContent, "Live action: ProjectType", auxTitle, auxChannel+": "+auxSerie, auxSeason, null, tmpTimeStamp, 3, "Aggregator", dwoEstimatedTRT, auxDuration, auxEpisodes, auxServices, "Realization: Contract Currency , Lyricist: Contract Currency", "2.0: MixSoundRequest", false, auxOriginalLanguage, "LAS: RequestedLanguage", false, null, null, null, null, null, null, null, null, null, false, null, null, "denise.surce@mediaaccesscompany.com", null, null, null, null, null, null, null, "Nat Geo: Sandra Brizuela", auxQCCreative, null, null, null, "Unblock", null, null, "F", null, null, null, null, auxMailDataOut, null, "(01) On track: DWO", lastUser, tmpTimeStamp, null, null, null, null, "", null, null, null, auxProgramID, null, null, auxTitleTranslated, null, null, null, null, null, null, null, null, null, null];
    dwoSheet.appendRow(buffer);
    addContask("DWO", dwoProjectID, "INSERT_ROW");
    if (verbose) {console.log("Append DubApp DWO "+dwoProjectID)};
    auxChangeKey = dwoProjectID;
  } else {
    var dubappDWOCase = dwoValues[dubProjectRow];
    dwoProjectID = dubappDWOCase[1];
    var updateNeeded = false;
    if (dwoFormat === "Series: FormatType") {
      if (dubappDWOCase[8] !== auxSeason) {
        dwoSheet.getRange(dubProjectRow + 2, 9).setValue(auxSeason);
        updateNeeded = true;
      }
      if (dubappDWOCase[15] !== auxEpisodes) {
        dwoSheet.getRange(dubProjectRow + 2, 16).setValue(auxEpisodes);
        updateNeeded = true;
      }
    } else {
      if (dubappDWOCase[6] !== auxTitle) {
        dwoSheet.getRange(dubProjectRow + 2, 7).setValue(auxTitle);
        updateNeeded = true;
      }
    }
    if (dubappDWOCase[72] !== auxTitleTranslated) {dwoSheet.getRange(dubProjectRow + 2, 74).setValue(auxTitleTranslated); updateNeeded = true;}
    if (dubappDWOCase[4] !== auxContent) {dwoSheet.getRange(dubProjectRow + 2, 5).setValue(auxContent); updateNeeded = true;}
    if (dubappDWOCase[44] !== auxQCCreative) {dwoSheet.getRange(dubProjectRow + 2, 45).setValue(auxQCCreative); updateNeeded = true;}
    if (dubappDWOCase[14] !== auxDuration) {dwoSheet.getRange(dubProjectRow + 2, 15).setValue(auxDuration); updateNeeded = true;}
    if (dubappDWOCase[13] !== dwoEstimatedTRT) {dwoSheet.getRange(dubProjectRow + 2, 14).setValue(dwoEstimatedTRT); updateNeeded = true;}
    if (dubappDWOCase[16] !== auxServices) {dwoSheet.getRange(dubProjectRow + 2, 17).setValue(auxServices); updateNeeded = true;}
    if (updateNeeded) {
      addContask("DWO", dwoProjectID, "EDIT");
      if (verbose) {console.log("Update DubApp DWO "+dwoProjectID)};
    }

    return dwoProjectID;

  }

  return dwoProjectID;
}

function openEpisodes() {
  if (!episodeSheet) {
    episodeSheet = sSht.getSheetByName("AggEpisode");
    var lastRow = episodeSheet.getLastRow();
    var lastCol = episodeSheet.getLastColumn();
    var auxData = episodeSheet.getRange(2, 1, lastRow - 1, lastCol);
    episodeValues = auxData.getValues();
    episodeNDX = episodeValues.map(function (r) { return r[0].toString(); }); //Program ID
    episodeNDX2 = episodeValues.map(function (r) { return r[1].toString(); }); //Production Number
    episodeNDX3 = episodeValues.map(function (r) { return r[64].toString(); }); //Production ID
  }
}

function glossary(glossSource, glossKey) {
  if (!glossarySheet) {
    glossarySheet = sSht.getSheetByName("AggDubAppGlossary");
    var lastRow = glossarySheet.getLastRow();
    var lastCol = glossarySheet.getLastColumn();
    var auxData = glossarySheet.getRange(2, 1, lastRow - 1, lastCol);
    glossaryValues = auxData.getValues();
    glossaryNDXAgg = glossaryValues.map(function (r) { return r[1].toString(); }); //Agg key
    glossaryNDXDubApp = glossaryValues.map(function (r) { return r[2].toString(); }); //DubApp key
  }

  var equivalence = glossSource === "Aggregator" ? glossaryNDXAgg.indexOf(glossKey) : glossaryNDXDubApp.indexOf(glossKey);
  return equivalence !== -1 ? glossaryValues[equivalence][glossSource === "Aggregator" ? 4 : 3] : "";
}

function openSheet(sheetName) {
  if (!ssActiveDubApp) {
    ssActiveDubApp = SpreadsheetApp.openById(dubAppTotalID);
  }
  var sheet = ssActiveDubApp.getSheetByName(sheetName);
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  var auxData = sheet.getRange(2, 1, lastRow - 1, lastCol);
  return {
    values: auxData.getValues(),
    NDX: auxData.getValues().map(function (r) { return r[1].toString(); }),
    NDX2: auxData.getValues().map(function (r) { return r[69].toString(); })
  };
}

function dubAppProject(programID) {
  if (!ssActiveDubApp) {
    ssActiveDubApp = SpreadsheetApp.openById(dubAppTotalID);
  }
  if (!dwoSheet) {
    dwoSheet = ssActiveDubApp.getSheetByName("DWO");
    var lastRow = dwoSheet.getLastRow();
    var lastCol = dwoSheet.getLastColumn();
    var auxData = dwoSheet.getRange(2, 1, lastRow - 1, lastCol);
    dwoValues = auxData.getValues();
    dwoNDX = dwoValues.map(function (r) { return r[1].toString(); });
    dwoNDX2 = dwoValues.map(function (r) { return r[69].toString(); });
  }
  programID = programID + "";
  if (programID!="") {
    dwoRow = dwoNDX2.indexOf(programID);
    return dwoRow;
  } else {return -1}
}

function dubAppProduction(productionNumber) {
  if (!ssActiveDubApp) {
    ssActiveDubApp = SpreadsheetApp.openById(dubAppTotalID);
  }
  if (!productionSheet) {
    productionSheet = ssActiveDubApp.getSheetByName("DWO-Production");
    var lastRow = productionSheet.getLastRow();
    var lastCol = productionSheet.getLastColumn();
    var auxData = productionSheet.getRange(2, 1, lastRow - 1, lastCol);
    productionValues = auxData.getValues();
    productionNDX = productionValues.map(function (r) { return r[0].toString(); }); // Production ID
    productionNDX2 = productionValues.map(function (r) { return r[1].toString(); }); // Project ID
    productionNDX3 = productionValues.map(function (r) { return r[58].toString(); }); // Production Number
  }
  
  productionNumber = productionNumber + "";
  if (productionNumber!="") {
    productionRow = productionNDX3.indexOf(productionNumber);
    return productionRow;
  } else {return -1}
}

function dubAppSeries(seriesChannel, seriesTitle, seriesTitleSpa) {
  if (!ssActiveDubApp) {
    ssActiveDubApp = SpreadsheetApp.openById(dubAppTotalID);
  }
  if (!seriesSheet) {
    seriesSheet = ssActiveDubApp.getSheetByName("DWO-Series");
    var lastRow = seriesSheet.getLastRow();
    var lastCol = seriesSheet.getLastColumn();
    var auxData = seriesSheet.getRange(2, 1, lastRow - 1, lastCol);
    seriesValues = auxData.getValues();
    seriesNDX = seriesValues.map(function (r) { return r[1].toString(); }); // series title
  }
  let seriesRow = seriesNDX.indexOf(seriesTitle, 0);
  while (seriesRow != -1 && seriesValues[seriesRow][0] !== seriesChannel) {
    seriesRow = seriesNDX.indexOf(seriesTitle, seriesRow + 1);
  }
  if (seriesRow === -1) {
    var auxSeries = [seriesChannel, seriesTitle, seriesTitleSpa, "", "", "", "(01) Enabled: Generic", lastUser, tmpTimeStamp, "", "", "", ""];
    seriesSheet.appendRow(auxSeries);
    if (verbose) {console.log("Append DubApp Series "+seriesTitle)};
    addContask("DWO-Series", seriesTitle, "INSERT_ROW");
    seriesRow = seriesNDX.length;
    seriesNDX.push(seriesTitle);
    seriesValues.push(auxSeries);
  }
  return seriesRow;
}


function addContask(addTable, addKey, addAction) {
  if (!Contask) {
    Contask = SpreadsheetApp.openById(dubAppIDControl);
    Contask = Contask.getSheetByName("CON-TaskCurrent");
  }
  addKey = addKey + "";
  var newControl = [addTable, addKey, tmpTimeStamp, "DubAppActive01", addAction, lastUser, "01 Pending", "", "Aggregator"];
  Contask.appendRow(newControl);
  if (verbose) {console.log("Control "+addTable+" / "+addAction+" / "+addKey)};
}

function newKey(len) {
  const possible = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ123456789";
  let id = "";
  if (!Number(len)) {
    throw new Error("The length must be an integer.")
  }
  for (var i = 0; i < len; i++) {
    id += possible.charAt(Math.floor(Math.random() * possible.length));
  }
  return id;
}
