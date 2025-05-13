function processCharacterProduction() {
    // IDs de las sheets
    var sheetIdDWO_CharacterProduction = '1d258eRpy3NovzQ1wKX_WVRCfXhxc2u2pwpYUuTDyfIw';
    var sheetIdDWO_Character = '1d258eRpy3NovzQ1wKX_WVRCfXhxc2u2pwpYUuTDyfIw';
    var sheetIdDWO = '1d258eRpy3NovzQ1wKX_WVRCfXhxc2u2pwpYUuTDyfIw';
    var sheetIdDWO_Actor = '1WE7VAs-3jVu1Z7D7xtJa4LjtYd2e_HdfUd7sIiT3dRg';
    var sheetIdCatalogo = '1V9_rRDAM_HuUto5iuCcyrinpU2_LvT53AuntXagm7eY';
  
    // Abrir las sheets
    var ss = SpreadsheetApp.openById(sheetIdDWO_CharacterProduction);
    var sheetDWO_CharacterProduction = ss.getSheetByName('DWO_CharacterProduction');
    var dataDWO_CharacterProduction = sheetDWO_CharacterProduction.getDataRange().getValues();
  
    ss = SpreadsheetApp.openById(sheetIdDWO_Character);
    var sheetDWO_Character = ss.getSheetByName('DWO_Character');
    var dataDWO_Character = sheetDWO_Character.getDataRange().getValues();
  
    ss = SpreadsheetApp.openById(sheetIdDWO);
    var sheetDWO = ss.getSheetByName('DWO');
    var dataDWO = sheetDWO.getDataRange().getValues();
  
    ss = SpreadsheetApp.openById(sheetIdDWO_Actor);
    var sheetDWO_Actor = ss.getSheetByName('DWO_Actor');
    var dataDWO_Actor = sheetDWO_Actor.getDataRange().getValues();
  
    // Abrir la hoja DWO_SongDetail para verificar participación en canciones
    ss = SpreadsheetApp.openById(sheetIdDWO_Character);
    var sheetDWO_SongDetail = ss.getSheetByName('DWO_SongDetail');
    var dataDWO_SongDetail = sheetDWO_SongDetail.getDataRange().getValues();
  
    ss = SpreadsheetApp.openById(sheetIdCatalogo);
    var sheetCatalogo = ss.getSheetByName('Catalogo');
  
    // Limpiar el contenido de la hoja Catalogo
    sheetCatalogo.clearContents();
  
    // Anular las selecciones de los filtros actuales
    var filter = sheetCatalogo.getFilter();
    if (filter) {
      var filterRange = filter.getRange();
      for (var col = 1; col <= filterRange.getNumColumns(); col++) {
        filter.removeColumnFilterCriteria(col);
      }
    }
  
    // Crear matrices globales por la clave
    var characterNDX = dataDWO_Character.map(function(r) { return r[0].toString(); });
    var projectNDX = dataDWO.map(function(r) { return r[1].toString(); });
    var actorNDX = dataDWO_Actor.map(function(r) { return r[0].toString(); });
  
    // Crear la matriz de resultados
    var result = [['Proyecto', 'Nombre_Actor', 'Personaje', 'Tipo_de_intervención']];
  
    // Recorrer DWO_CharacterProduction
    for (var i = 1; i < dataDWO_CharacterProduction.length; i++) {
      if (['(01) Recording pending: DWOCharacterProduction', '(02) Booked: DWOCharacterProduction', '(03) Completed: DWOCharacterProduction'].includes(dataDWO_CharacterProduction[i][9])) {
        var characterId = dataDWO_CharacterProduction[i][1].toString().trim();
        var characterIndex = characterNDX.indexOf(characterId);
  
        if (characterIndex !== -1) {
          var characterData = dataDWO_Character[characterIndex];
          var projectId = characterData[1].toString().trim();

          var projectIndex = projectNDX.indexOf(projectId);
          var project = projectIndex !== -1 ? (dataDWO[projectIndex][7] || dataDWO[projectIndex][6]) : '';
  
          var actorId = dataDWO_CharacterProduction[i][4].toString().trim() || characterData[6].toString().trim();

          var actorIndex = actorNDX.indexOf(actorId.toString().trim());
          var actorName = actorIndex !== -1 ? dataDWO_Actor[actorIndex][1] + ' ' + dataDWO_Actor[actorIndex][2] : '';
  
          if (actorName) { // Descartar casos donde Nombre_Actor está vacío
            var interventionType = '';
            if (characterData[4].includes('Main: Character_Attributes')) {
                interventionType += 'Principal';
            }
            if (characterData[4].includes('Recurring: Character_Attributes')) {
                if (interventionType) {
                    interventionType += ' y Recurrente';
                } else {
                    interventionType += 'Recurrente';
                }
            }
  
            // Verificar participación en canciones
            var hasSongParticipation = false;
            // Verificar si el proyecto tiene canciones (columna T en DWO)
            if (projectIndex !== -1 && dataDWO[projectIndex][19] === "true") {
              

              // Buscar coincidencias en DWO_SongDetail
              for (var j = 1; j < dataDWO_SongDetail.length; j++) {
                var songProjectId = dataDWO_SongDetail[j][6].toString().trim();
                var songCharacterId = dataDWO_SongDetail[j][3] ? dataDWO_SongDetail[j][3].toString().trim() : "";
                var songActorId = dataDWO_SongDetail[j][5] ? dataDWO_SongDetail[j][5].toString().trim() : "";
                var songStatus = dataDWO_SongDetail[j][14] || "";
                
                if (songProjectId === projectId && 
                    (songCharacterId === characterId || songActorId === actorId.toString().trim()) && 
                    songStatus !== "(99) Dismissed: DWOSong") {
                  hasSongParticipation = true;
                  break;
                }
              }
            }
            
            // Agregar información de participación en canciones
            if (hasSongParticipation) {
              if (interventionType) {
                interventionType += ", Incluye ditty/canción";
              } else {
                interventionType = "Incluye ditty/canción";
              }
            }
  
            var character = characterData[4].includes('Unnamed role: Character_Attributes') ? 
                            characterData[2] + " / " + dataDWO_CharacterProduction[i][15] : characterData[2];
  
            result.push([project, actorName, character, interventionType]);
          }
        } else {
          Logger.log('Character ID not found: ' + characterId);
        }
      }
    }
    
    // Añadir entradas de DWO_SongDetail donde la columna D está vacía y hay un ActorID en columna F
    // Crear un objeto para evitar duplicados por [proyecto, actor]
    var chorusEntries = {};
    
    for (var i = 1; i < dataDWO_SongDetail.length; i++) {
      var songStatus = dataDWO_SongDetail[i][14] || "";
      var songCharacterId = dataDWO_SongDetail[i][3] ? dataDWO_SongDetail[i][3].toString().trim() : "";
      var songActorId = dataDWO_SongDetail[i][5] ? dataDWO_SongDetail[i][5].toString().trim() : "";
      var songProjectId = dataDWO_SongDetail[i][6] ? dataDWO_SongDetail[i][6].toString().trim() : "";
      
      // Si no hay characterId pero hay actorId y no está descartado
      if (songActorId && !songCharacterId && songStatus !== "(99) Dismissed: DWOSong") {
        var projectIndex = projectNDX.indexOf(songProjectId);
        var actorIndex = actorNDX.indexOf(songActorId);
        
        if (projectIndex !== -1 && actorIndex !== -1) {
          var project = dataDWO[projectIndex][7] || dataDWO[projectIndex][6] || '';
          var actorName = dataDWO_Actor[actorIndex][1] + ' ' + dataDWO_Actor[actorIndex][2];
          
          // Usar proyecto+actor como clave para evitar duplicados
          var key = project + '|' + actorName;
          if (!chorusEntries[key]) {
            chorusEntries[key] = [project, actorName, "CHORUS", "Chorus"];
          }
        }
      }
    }
    
    // Añadir entradas de chorus al resultado
    for (var key in chorusEntries) {
      result.push(chorusEntries[key]);
    }
  
    // Unificar filas duplicadas
    var unifiedResult = {};
    for (var i = 1; i < result.length; i++) {
      var key = result[i].join('|'); // Crear una clave unificada para la fila
      if (!unifiedResult[key]) {
        unifiedResult[key] = result[i];
      }
    }
  
    var finalResult = [result[0]].concat(Object.values(unifiedResult));
  
    // Ordenar la matriz
    var headers = finalResult.shift(); // Quitar la cabecera para ordenar
    finalResult.sort(function(a, b) {
      if (a[0] !== b[0]) return a[0].localeCompare(b[0]); // Ordenar Proyecto ascendente
      if (a[1] !== b[1]) return a[1].localeCompare(b[1]); // Ordenar Nombre_Actor ascendente
      return a[2].localeCompare(b[2]); // Ordenar Personaje ascendente
    });
    finalResult.unshift(headers); // Añadir la cabecera nuevamente
  
    // Escribir los resultados en la sheet Catalogo
    sheetCatalogo.getRange(1, 1, finalResult.length, finalResult[0].length).setValues(finalResult);
  }
  