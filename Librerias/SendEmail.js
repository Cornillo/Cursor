function AppSendEmail(destinatario, remitente, nombreDestinatario, idDocumento, idPDF, asunto, parametros) {
  AppSendEmailX(destinatario, remitente, nombreDestinatario, idDocumento, idPDF, asunto, parametros, "", "");
}

function AppSendEmailX(destinatario, remitente, nombreDestinatario, idDocumento, idPDF, asunto, parametros, cc, bcc) {
  // Validar destinatario
  let destinatarioFinal = destinatario;
  let asuntoFinal = asunto;
  
  if (!destinatario || !destinatario.includes('@') || destinatario.trim() === '') {
    destinatarioFinal = 'appsheet@mediaaccesscompany.com';
    asuntoFinal = 'SIN DESTINATARIO VALIDO - ' + asunto;
  }

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
      const patrones = [
        new RegExp(`&lt;&lt;[^>]*>${clave}[^>]*>&gt;&gt;`, 'g'),
        new RegExp(`&lt;&lt;${clave}&gt;&gt;`, 'g'),
        new RegExp(`&lt;&lt;[^&]*${clave}[^&]*&gt;&gt;`, 'g'),
        new RegExp(`&lt;&lt;\\s*${clave}\\s*&gt;&gt;`, 'g')
      ];
      
      patrones.forEach(patron => {
        const placeholderCompleto = cuerpo.match(patron);
        if (placeholderCompleto) {
          cuerpo = cuerpo.replace(placeholderCompleto[0], valor);
        }
      });
    });
  }

  // Obtener la imagen del logo
  const logoBlob = DriveApp.getFileById('122l7hTIpDHuAzTF_Iyf6GHLRrkdiWRJz').getBlob();
  
  // Actualizar opciones del correo con el destinatario y asunto validados
  const opcionesCorreo = {
    to: destinatarioFinal,
    subject: asuntoFinal,
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
      throw new Error('Error al adjuntar PDF: ' + e.message);
    }
  }

  // Enviar el correo con reintento
  return reintentarConEspera(() => MailApp.sendEmail(opcionesCorreo));
}

function reintentarConEspera(funcion, maxIntentos = 8) {
  for (let i = 0; i < maxIntentos; i++) {
    try {
      return funcion();
    } catch (e) {
      if (e.toString().includes('429') && i < maxIntentos - 1) {
        Utilities.sleep(Math.pow(2, i) * 1000);
        continue;
      }
      throw e;
    }
  }
}

function call() {
  AppSendEmail("appsheet@mediaaccesscompany.com", "prueba@prueba.com", "Tavo", "1pxau9jhDtg4NiOpqFYsnUgqY576laIoiXkr7dvBJ1ao", "", "Prueba de envío", 
    "Fecha::2/2/2024||Proyecto/Episodio::Prueba de Proyecto y Episodio||Personaje::Miliki||Loops::24||TipoCitacion::Retake 1||Mail asistente::prueba@prueba.com", 
    "cc@example.com", "bcc@example.com");
}
