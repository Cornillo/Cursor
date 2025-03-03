/**
 * Árbol de llamadas y explicación del script:
 * 
 * call()
 *   └─> actualizarEmailUsuario(emailViejo, emailNuevo, oldEmailCorporate)
 *        ├─> Obtiene IDs de todas las spreadsheets mediante databaseID.getID()
 *        └─> Para cada spreadsheet (Active, Total, Logs, NoTrack, Control):
 *            ├─> Abre la spreadsheet
 *            └─> Para cada hoja en la spreadsheet:
 *                ├─> Excluye hojas 'DWO-LogLabels' e 'Index'
 *                ├─> Obtiene todos los datos de la hoja
 *                ├─> Procesa cada celda (saltando encabezados)
 *                │   └─> Reemplaza el email viejo por el nuevo
 *                └─> Caso especial para hoja 'App-User':
 *                    └─> Actualiza email en columna A y opcionalmente en C
 * 
 * Propósito: Actualizar el email de un usuario en todas las spreadsheets del sistema,
 * incluyendo menciones en cualquier celda y registros específicos en la hoja App-User.
 */

function call(){
    actualizarEmailUsuario("ignacio.recabeitia@nonstoptv.tv", "ignacio.recabeitia@mediaaccesscompany.com", true);
  }  

function actualizarEmailUsuario(emailUsuarioCambio, emailUsuarioNuevo, oldEmailCorporate) {
  // Obtener IDs de las spreadsheets
  const allIDs = databaseID.getID();
  const spreadsheets = [
    {id: allIDs.activeID, name: 'DubAppActive01'},
    {id: allIDs.totalID, name: 'DubAppTotal01'}, 
    {id: allIDs.logsID, name: 'DubAppLogs01'},
    {id: allIDs.noTrackID, name: 'DubAppNoTrack01'},
    {id: allIDs.controlID, name: 'DubAppControl01'}
  ];

  // Procesar cada spreadsheet
  spreadsheets.forEach(ss => {
    console.log(`Procesando spreadsheet: ${ss.name}`);
    const spreadsheet = SpreadsheetApp.openById(ss.id);
    const sheets = spreadsheet.getSheets();

    // Procesar cada hoja
    sheets.forEach(sheet => {
      const sheetName = sheet.getName();
      
      // Excluir hojas específicas
      if (sheetName === 'DWO-LogLabels' || sheetName === 'Index') {
        console.log(`Saltando hoja excluida: ${sheetName}`);
        return;
      }
      
      console.log(`Procesando hoja: ${sheetName}`);
      const lastRow = sheet.getLastRow();
      const lastCol = sheet.getLastColumn();
      
      if (lastRow <= 1) return; // Saltar hojas vacías

      // Obtener todos los datos
      const range = sheet.getRange(1, 1, lastRow, lastCol);
      const values = range.getValues();

      // Bandera para detectar cambios
      let hasChanges = false;

      // Procesar cada celda, empezando desde la fila 1 (después del encabezado)
      for (let i = 1; i < values.length; i++) {
        for (let j = 0; j < values[i].length; j++) {
          if (typeof values[i][j] === 'string') {
            const oldValue = values[i][j];
            // Reemplazar el email en el contenido
            const newValue = oldValue.replace(new RegExp(emailUsuarioCambio, 'g'), emailUsuarioNuevo);
            
            if (oldValue !== newValue) {
              values[i][j] = newValue;
              hasChanges = true;
            }
          }
        }
      }

      // Caso especial para App-User
      if (sheetName === 'App-User') {
        for (let i = 0; i < values.length; i++) {
          if (values[i][0] === emailUsuarioCambio) {
            values[i][0] = emailUsuarioNuevo;
            // Columna C (índice 2) - mantener valor previo o actualizar según oldEmailCorporate
            if (!oldEmailCorporate) {
              values[i][2] = emailUsuarioNuevo;
            }
            hasChanges = true;
          }
        }
      }

      // Actualizar la hoja si hubo cambios
      if (hasChanges) {
        range.setValues(values);
      }
    });
  });
}
