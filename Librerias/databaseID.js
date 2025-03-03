function getID() {
    var allIDs = {
    //Data bases
    activeID: "1d258eRpy3NovzQ1wKX_WVRCfXhxc2u2pwpYUuTDyfIw",
    totalID: "1rynSNh2wO6Izuty-DO1Nw5tVu0T-9saDGwleKYETIDM",
    logsID: "1Wattm0vFkINof0EzQRZaWzUlrJvs8cjbc9gq3q5wdhQ",
    noTrackID: "1WE7VAs-3jVu1Z7D7xtJa4LjtYd2e_HdfUd7sIiT3dRg",
    controlID: "1P-kQdID7dwG4UUezuF0QuJq7G3swDaAwaznBNwnNvCk",
    backupOrphansID: "1AN7RFQx_O2Exz0tmLxSY4afTf4RE2sF_4FK1",
    //Folders
    backupID: "1wbWjoEbYSRWRc3zE_opzLrqo4uUP9Rja",
    proToolID: "1pkvWYHN_uX4t47zhL8k1ooni8zX0A606",
    uploaded:"1ND3FGiWrrb5fDW02guJSlAO5CB_AXnVc",
    loopsWitness:"1h2q8drsnDjrRWX8iT-7BYrRhP__jSGWi",
    userTempID: "16g8NUHinLJ-hyKx_ZBS3ZKZP7BfIj0XM",
    infoDocFolder: "1l_glnJtUfU7iFdzm87s0T-s8vMIzjxvS",
    QCCreativeFolderID: '1zC3DwDaZgj6GXT1XsY3TXe0qeEt5huZu',
    //Templates
    infoDocTemplate: "1F5ta0KaphdRmVtHdm2zeU9g_bwLDsybaxgNPHC1DNss",
    //General config
    instalation: "01",
    timezone:"GMT-3",
    timestamp_format:"dd/MM/yyyy HH:mm:ss"
  };
    return allIDs;
}



function getStructure(){
  return {
    // Primero tablas dependientes
    'DWO-Production': { columna: 'A', dwoCol: 'B' },
    'DWO-Event': { columna: 'A', dwoCol: 'BX' },
    'DWO_Character': { columna: 'A', dwoCol: 'B' },
    'DWO_CharacterProduction': { columna: 'A', dwoCol: 'X' },
    'DWO_Files': { columna: 'A', dwoCol: 'P' },
    'DWO-MixAndEdit': { columna: 'A', dwoCol: 'Q' },
    'DWO-Observation': { columna: 'A', dwoCol: 'Z' },
    'DWO_FilesLines': { columna: 'A', dwoCol: 'Q' },
    'DWO_FilesCharacter': { columna: 'A', dwoCol: 'N' },
    'DWO_Song': { columna: 'A', dwoCol: 'P' },
    'DWO_SongDetail': { columna: 'A', dwoCol: 'G' },
    'DWO-SynopsisProject': { columna: 'A', dwoCol: 'A' },
    'DWO-SynopsisProduction': { columna: 'A', dwoCol: 'T' },
    // DWO al final
    'DWO': { columna: 'B' }
  };
}

//Clear all filters
function clearFilter(ss) {
  var ssId = ss.getId();
  var sheetIds = ss.getSheets();
  for (var i in sheetIds) {
    var requests = [{
      "clearBasicFilter": {
        "sheetId": sheetIds[i].getSheetId()
      }
    }];
    if (requests.length > 0) {
      Sheets.Spreadsheets.batchUpdate({'requests': requests}, ssId);
    }
  }
} 

//Ejemplo uso
/*
  // DB ID Loaded as global
  const allIDs = databaseID.getID();
  var ss = SpreadsheetApp.openById(allIDs['controlID']);
*/