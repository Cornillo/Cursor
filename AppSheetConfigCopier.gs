/**
 * Configuración global del script
 */
const CONFIG = {
  // Credenciales OAuth
  CLIENT_ID: process.env.GOOGLE_CLIENT_ID || PropertiesService.getScriptProperties().getProperty('GOOGLE_CLIENT_ID'),
  CLIENT_SECRET: process.env.GOOGLE_CLIENT_SECRET || PropertiesService.getScriptProperties().getProperty('GOOGLE_CLIENT_SECRET'),
  OAUTH_SCOPE: 'https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/drive',
  
  // IDs y Access Keys de las aplicaciones
  SOURCE_APP_ID: 'DubAppProject233-665487746',      // Nombre exacto de la app origen
  SOURCE_ACCESS_KEY: 'V2-j7XGz-8dfnx-dUbdN-HrdUL-xkCV7-enN7X-LdG7s-Jvawn',
  
  TARGET_APP_ID: 'DubAppPlan-665487746',      // Nombre exacto de la app destino
  TARGET_ACCESS_KEY: 'V2-7nEnd-aeT9r-swFLx-u64ZS-XTiVw-gy522-rQkN8-T9bGM'   // Access Key de destino
};

function getAppConfig(appId, token) {
  const endpoint = `https://api.appsheet.com/api/v2/apps/${appId}/config`;
  const options = {
    method: 'GET',
    headers: {
      'Authorization': `Bearer ${token}`,
      'Accept': 'application/json'
    },
    muteHttpExceptions: true  // Para ver respuestas de error completas
  };
  
  const response = UrlFetchApp.fetch(endpoint, options);
  const responseCode = response.getResponseCode();
  
  if (responseCode !== 200) {
    Logger.log(`Error al acceder a la app ${appId}. Código: ${responseCode}`);
    Logger.log(`Respuesta: ${response.getContentText()}`);
    return null;
  }
  
  return JSON.parse(response.getContentText());
}

/**
 * Función para copiar una action específica por nombre
 * @param {string} actionName - Nombre de la action a copiar
 * @param {boolean} testMode - Si es true, solo muestra la configuración sin aplicarla
 */
function copySpecificAction(actionName, testMode = true) {
  const sourceApplicationId = CONFIG.SOURCE_APP_ID;
  const targetApplicationId = CONFIG.TARGET_APP_ID;
  const accessToken = getOAuthToken();
  
  try {
    // Obtener configuración de la action específica
    const actionConfig = getSpecificActionConfig(sourceApplicationId, actionName, accessToken);
    
    if (!actionConfig) {
      Logger.log(`No se encontró la action "${actionName}" en la app origen`);
      return;
    }

    // Mostrar la configuración encontrada
    Logger.log('\n=== CONFIGURACIÓN DE ACTION ENCONTRADA ===');
    Logger.log(`Nombre: ${actionName}`);
    Logger.log('Configuración:');
    Logger.log(JSON.stringify(actionConfig, null, 2));

    if (testMode) {
      Logger.log('\n=== MODO PRUEBA ===');
      Logger.log('La configuración NO se aplicó a la app destino');
      Logger.log('Para aplicar la configuración, ejecute con testMode = false');
      return;
    }

    // Si no es modo prueba, aplicar la configuración
    updateSpecificAction(targetApplicationId, actionName, actionConfig, accessToken);
    Logger.log('\nAction copiada exitosamente a la app destino');
    
  } catch (error) {
    Logger.log('Error al copiar la action: ' + error);
  }
}

function getSpecificActionConfig(appId, actionName, token) {
  const endpoint = `https://api.appsheet.com/api/v2/apps/${appId}/config`;
  const options = {
    method: 'GET',
    headers: {
      'Authorization': `Bearer ${token}`,
      'Accept': 'application/json'
    }
  };
  
  const response = UrlFetchApp.fetch(endpoint, options);
  const fullConfig = JSON.parse(response.getContentText());
  
  // Buscar la action específica
  if (fullConfig.actions) {
    const action = Object.entries(fullConfig.actions).find(([key, value]) => 
      value.name === actionName
    );
    
    if (action) {
      return {
        [action[0]]: action[1]  // Mantener la estructura key-value original
      };
    }
  }
  
  return null;
}

function updateSpecificAction(appId, actionName, actionConfig, token) {
  // Obtener configuración actual
  const endpoint = `https://api.appsheet.com/api/v2/apps/${appId}/config`;
  const getCurrentConfig = {
    method: 'GET',
    headers: {
      'Authorization': `Bearer ${token}`,
      'Accept': 'application/json'
    }
  };
  
  const currentConfig = JSON.parse(UrlFetchApp.fetch(endpoint, getCurrentConfig).getContentText());
  
  // Crear backup de la configuración actual
  const backupTimestamp = backupCurrentConfig(appId, token);
  Logger.log(`Backup creado con timestamp: ${backupTimestamp}`);
  
  // Actualizar solo la action específica
  const updatedActions = {
    ...currentConfig.actions,
    ...actionConfig
  };
  
  const updatedConfig = {
    ...currentConfig,
    actions: updatedActions
  };
  
  // Aplicar la actualización
  const updateOptions = {
    method: 'PUT',
    headers: {
      'Authorization': `Bearer ${token}`,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(updatedConfig)
  };
  
  return UrlFetchApp.fetch(endpoint, updateOptions);
}

// Función para listar todas las actions disponibles
function listAvailableActions() {
  try {
    // Loguear información de debug
    Logger.log('=== DEBUG INFO ===');
    
    // Usar el nombre exacto de la app
    const appName = 'DubAppProject233-665487746';
    const endpoint = `https://api.appsheet.com/api/v2/apps/${appName}/tables`;
    Logger.log(`Endpoint: ${endpoint}`);
    
    // Remover V2- y todos los guiones del Access Key
    const accessKey = CONFIG.SOURCE_ACCESS_KEY
      .replace('V2-', '')
      .replace(/-/g, '');
    Logger.log(`Access Key (sin guiones): ${accessKey}`);
    
    const options = {
      method: 'GET',
      headers: {
        'ApplicationAccessKey': accessKey,
        'Content-Type': 'application/json',
        'Accept': 'application/json'
      },
      muteHttpExceptions: true
    };
    
    Logger.log('Realizando petición...');
    Logger.log('Headers:');
    Logger.log(JSON.stringify(options.headers, null, 2));
    
    const response = UrlFetchApp.fetch(endpoint, options);
    const responseCode = response.getResponseCode();
    
    Logger.log(`Código de respuesta: ${responseCode}`);
    if (responseCode !== 200) {
      Logger.log('Headers de respuesta:');
      Logger.log(JSON.stringify(response.getAllHeaders(), null, 2));
      Logger.log('Contenido de respuesta:');
      Logger.log(response.getContentText().substring(0, 1000));
      
      // Intentar con otros endpoints
      const endpoints = [
        `https://api.appsheet.com/api/v2/apps/${appName}`,
        `https://api.appsheet.com/api/v2/apps/${appName}/actions`,
        `https://api.appsheet.com/api/v2/apps/${appName}/definition`
      ];
      
      Logger.log('\nIntentando con endpoints alternativos...');
      for (const altEndpoint of endpoints) {
        Logger.log(`\nProbando: ${altEndpoint}`);
        const altResponse = UrlFetchApp.fetch(altEndpoint, options);
        Logger.log(`Código de respuesta: ${altResponse.getResponseCode()}`);
      }
    } else {
      const config = JSON.parse(response.getContentText());
      Logger.log('\n=== CONFIGURACIÓN ENCONTRADA ===');
      Logger.log(JSON.stringify(config, null, 2));
    }
    
  } catch (error) {
    Logger.log('Error al listar actions: ' + error);
    Logger.log('Stack trace: ' + error.stack);
  }
}

/**
 * Función principal para ejecutar el copiado de actions
 * @param {boolean} testMode - Si es true, ejecuta en modo prueba
 */
function call() {
  // Configuración
  const actionName = 'Massive change  - DWO - 3'; // Reemplazar con el nombre de la action a copiar
  const testMode = true; // Cambiar a false para aplicar los cambios
  
  Logger.log('=== INICIANDO PROCESO DE COPIA DE ACTION ===');
  Logger.log(`Modo: ${testMode ? 'PRUEBA' : 'PRODUCCIÓN'}`);
  
  try {
    // Primero listar las actions disponibles
    Logger.log('\nListando actions disponibles...');
    listAvailableActions();
    
    // Luego copiar la action específica
    Logger.log('\nIniciando copia de action específica...');
    copySpecificAction(actionName, testMode);
    
  } catch (error) {
    Logger.log('Error en la ejecución: ' + error);
  }
  
  Logger.log('\n=== PROCESO FINALIZADO ===');
}

/**
 * Función para ejecutar en modo producción (aplicando cambios)
 */
function callProduction() {
  const actionName = 'Massive change  - DWO - 3'; // Reemplazar con el nombre de la action a copiar
  copySpecificAction(actionName, false);
}

/**
 * Función para ejecutar en modo prueba (solo mostrar configuración)
 */
function callTest() {
  const actionName = 'Massive change  - DWO - 3'; // Reemplazar con el nombre de la action a copiar
  copySpecificAction(actionName, true);
}

/**
 * Función para obtener el token OAuth
 * @returns {string} Token de acceso
 */
function getOAuthToken() {
  try {
    const properties = PropertiesService.getScriptProperties();
    let token = properties.getProperty('OAUTH_TOKEN');
    
    if (!token) {
      const service = OAuth2.createService('appsheet')
        .setAuthorizationBaseUrl('https://accounts.google.com/o/oauth2/auth')
        .setTokenUrl('https://accounts.google.com/o/oauth2/token')
        .setClientId(CONFIG.CLIENT_ID)
        .setClientSecret(CONFIG.CLIENT_SECRET)
        .setScope(CONFIG.OAUTH_SCOPE)
        .setPropertyStore(PropertiesService.getScriptProperties());
      
      if (!service.hasAccess()) {
        Logger.log('Autorización requerida. Por favor, ejecute la función authorize()');
        return null;
      }
      
      token = service.getAccessToken();
      properties.setProperty('OAUTH_TOKEN', token);
    }
    
    return token;
    
  } catch (error) {
    Logger.log('Error al obtener token OAuth: ' + error);
    return null;
  }
}

/**
 * Función para autorizar el script
 */
function authorize() {
  const service = OAuth2.createService('appsheet')
    .setAuthorizationBaseUrl('https://accounts.google.com/o/oauth2/auth')
    .setTokenUrl('https://accounts.google.com/o/oauth2/token')
    .setClientId(CONFIG.CLIENT_ID)
    .setClientSecret(CONFIG.CLIENT_SECRET)
    .setScope(CONFIG.OAUTH_SCOPE)
    .setPropertyStore(PropertiesService.getScriptProperties())
    .setCallbackFunction('authCallback');
    
  const authorizationUrl = service.getAuthorizationUrl();
  Logger.log('Abra esta URL para autorizar el script: ' + authorizationUrl);
}

/**
 * Callback para el proceso de autorización
 */
function authCallback(request) {
  const service = OAuth2.createService('appsheet')
    .setAuthorizationBaseUrl('https://accounts.google.com/o/oauth2/auth')
    .setTokenUrl('https://accounts.google.com/o/oauth2/token')
    .setClientId(CONFIG.CLIENT_ID)
    .setClientSecret(CONFIG.CLIENT_SECRET)
    .setScope(CONFIG.OAUTH_SCOPE)
    .setPropertyStore(PropertiesService.getScriptProperties())
    .setCallbackFunction('authCallback');
    
  const authorized = service.handleCallback(request);
  if (authorized) {
    return HtmlService.createHtmlOutput('Autorización completada. Puede cerrar esta ventana.');
  } else {
    return HtmlService.createHtmlOutput('Error en la autorización. Por favor, intente nuevamente.');
  }
}

/**
 * Función para limpiar las credenciales almacenadas
 */
function clearCredentials() {
  const properties = PropertiesService.getScriptProperties();
  properties.deleteProperty('OAUTH_TOKEN');
  Logger.log('Credenciales eliminadas');
} 