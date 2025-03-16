/**
 * Script de prueba para generar archivos STL compatibles con editores de subtítulos
 * Este archivo contiene funciones para probar la generación de archivos STL 
 * utilizando el formato que es compatible con editores comunes
 */

/**
 * Función principal para probar la generación de STL compatible
 */
function testCompatibleSTL() {
  Logger.log("Iniciando prueba de generación STL compatible...");
  
  // Crear datos de muestra con etiquetas de color como se ve en la imagen de ejemplo
  var subtitleData = createSampleStyledSubtitles();
  
  // Crear metadatos de ejemplo
  var metadata = {
    englishTitle: "SAMPLE STL TEST",
    spanishTitle: "PRUEBA STL COMPATIBLE",
    episodeNumber: "01",
    language: "es"
  };
  
  try {
    // Generar STL usando la función compatible
    var stlBlob = generateLegacySTL(subtitleData, metadata);
    
    // Verificar que se haya generado correctamente
    if (stlBlob && stlBlob.getBytes().length > 1024) {
      Logger.log("¡Éxito! Archivo STL compatible generado correctamente");
      Logger.log("Tamaño: " + stlBlob.getBytes().length + " bytes");
      
      // Guardar en Drive para pruebas
      var tempFolder = getOrCreateTempFolder();
      var stlFile = tempFolder.createFile(stlBlob);
      stlFile.setName("TEST_COMPATIBLE_STL.stl");
      
      // Hacer archivo accesible por URL
      stlFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      
      Logger.log("Archivo guardado. Puedes descargarlo desde: " + stlFile.getUrl());
      return {
        success: true,
        message: "STL compatible generado correctamente",
        fileName: "TEST_COMPATIBLE_STL.stl",
        fileUrl: stlFile.getUrl(),
        downloadUrl: stlFile.getDownloadUrl()
      };
    } else {
      throw new Error("El archivo STL no se generó correctamente");
    }
  } catch (error) {
    Logger.log("Error en la prueba: " + error.toString());
    return {
      success: false,
      message: "Error: " + error.message
    };
  }
}

/**
 * Crea un conjunto de subtítulos de muestra con etiquetas de formato
 * Simula el formato visto en la imagen de ejemplo con etiquetas de color
 * @return {Array} Array de subtítulos con formato
 */
function createSampleStyledSubtitles() {
  return [
    // Subtítulo con etiquetas de color rojo y amarillo
    ["00:00:06:00", "<font color=\"Red\"></font> <font color=\"Yellow\"></font>\\an8", "00:00:10:00"],
    
    // Subtítulo con etiquetas de color rojo y negro
    ["00:00:10:00", "<font color=\"Red\"></font> <font color=\"Black\"></font> <font color=\"Red\"></font>", "00:00:15:00"],
    
    // Subtítulo con etiquetas de color negro y magenta
    ["04:15:00:00", "\\an5<font color=\"Black\"></font> <font color=\"Magenta\"></font>", "08:04:15:00"],
    
    // Subtítulo con etiquetas de color negro y verde
    ["08:04:15:00", "\\an8<font color=\"Black\"></font> <font color=\"Green\"></font>", "09:01:00:00"],
    
    // Subtítulo con etiquetas de color cyan y negro
    ["09:01:00:00", "\\an8<font color=\"Cyan\"></font> <font color=\"Black\"></font>", "00:00:10:00"],
    
    // Subtítulo con varias etiquetas combinadas
    ["00:00:10:00", "<font color=\"Black\"></font> <font color=\"Green\"></font> <font color=\"Cyan\"></font>", "00:00:15:00"],
  ];
}

/**
 * Función para ejecutar la prueba de STL compatible desde el editor
 */
function runCompatibleSTLTest() {
  var result = testCompatibleSTL();
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

/**
 * Obtiene o crea una carpeta temporal para almacenar archivos STL
 * @return {Folder} Carpeta de Google Drive para uso temporal
 */
// Esta función ha sido movida a Code.gs para evitar duplicidad
// function getOrCreateTempFolder() {
//   try {
//     // Crear una carpeta temporal si no existe una reciente
//     var userProperties = PropertiesService.getUserProperties();
//     var lastTempFolderId = userProperties.getProperty('lastTempFolderId');
//     var tempFolder;
//     
//     if (!lastTempFolderId) {
//       // Crear una nueva carpeta temporal
//       var userEmail = Session.getActiveUser().getEmail();
//       var userName = userEmail.split('@')[0];
//       var timestamp = new Date().getTime();
//       var folderName = "STL_Temp_" + userName + "_" + timestamp;
//       tempFolder = DriveApp.createFolder(folderName);
//       
//       // Guardar el ID de la carpeta temporal
//       userProperties.setProperty('lastTempFolderName', folderName);
//       userProperties.setProperty('lastTempFolderId', tempFolder.getId());
//     } else {
//       // Usar la carpeta temporal existente
//       try {
//         tempFolder = DriveApp.getFolderById(lastTempFolderId);
//       } catch (e) {
//         // Si la carpeta no existe, crear una nueva
//         var userEmail = Session.getActiveUser().getEmail();
//         var userName = userEmail.split('@')[0];
//         var timestamp = new Date().getTime();
//         var folderName = "STL_Temp_" + userName + "_" + timestamp;
//         tempFolder = DriveApp.createFolder(folderName);
//         
//         // Guardar el ID de la carpeta temporal
//         userProperties.setProperty('lastTempFolderName', folderName);
//         userProperties.setProperty('lastTempFolderId', tempFolder.getId());
//       }
//     }
//     
//     return tempFolder;
//   } catch (e) {
//     Logger.log('Error al crear carpeta temporal: ' + e.toString());
//     return DriveApp.createFolder("STL_Temp_Error_" + new Date().getTime());
//   }
// } 