function AppSendEmailX(destinatario, remitente, nombreDestinatario, idDocumento, idPDF, asunto, parametros, cc, bcc) {
  // Obtener y convertir el documento a HTML con reintento
  const contenidoHTML = reintentarConEspera(() => {
    const urlDocumento = `https://docs.google.com/feeds/download/documents/export/Export?id=${idDocumento}&exportFormat=html`;
    const respuesta = UrlFetchApp.fetch(urlDocumento, {
      headers: {
        Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
      }
    });
    return respuesta.getContentText();
  });

  // Personalizar el contenido
  let cuerpo = contenidoHTML
    .replace("&lt;&lt;Nombre&gt;&gt;", nombreDestinatario)
    .replace("&lt;&lt;Logo&gt;&gt;", '<img src="cid:logoEmpresa" alt="Logo Media Access Company" style="width: 200px; margin-bottom: 20px;">');
  
  // Procesar parámetros adicionales
  if (parametros) {
    const datos = parametros.split("||");
    datos.forEach(par => {
      const [clave, valor] = par.split("::");
      cuerpo = cuerpo.replace(new RegExp("&lt;&lt;\\" + clave + "&gt;&gt;", 'g'), valor);
    });
  }

  // Obtener la imagen del logo
  const logoBlob = DriveApp.getFileById('122l7hTIpDHuAzTF_Iyf6GHLRrkdiWRJz').getBlob();
  
  // Preparar opciones del correo
  const opcionesCorreo = {
    to: destinatario,
    subject: asunto,
    htmlBody: cuerpo,
    from: 'Media Access Company <ar.info@mediaaccesscompany.com>',
    replyTo: remitente,
    inlineImages: {
      'logoEmpresa': logoBlob
    },
    ...(cc && { cc }),
    ...(bcc && { bcc })
  };

  // Adjuntar PDF si existe
  if (idPDF) {
    try {
      const archivoPDF = DriveApp.getFileById(idPDF);
      opcionesCorreo.attachments = [archivoPDF.getAs("application/pdf")];
    } catch (e) {
      console.error('Error al adjuntar PDF:', e.message);
    }
  }

  // Enviar el correo con reintento
  return reintentarConEspera(() => MailApp.sendEmail(opcionesCorreo));
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