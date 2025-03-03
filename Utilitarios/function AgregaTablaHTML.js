function AgregarTablaHTML(titulo, htmlBase, datos) {
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

  // Segunda tabla para el email
  let tablaEmail = '';
  let nombreAnterior = '';
  let emailActual = '';
  
  datos.forEach((fila, index) => {
    // Si cambia el nombre, enviar email con la tabla acumulada
    if (index > 0 && nombreAnterior !== '' && nombreAnterior !== fila[0] && tablaEmail !== '') {
      const primeraPalabra = nombreAnterior.split(' ')[0];
      const fechaGrabacion = datos[index - 1][2];
      AppSendEmailX("appsheet@mediaaccesscompany.com","ar.info@mediaaccesscompany.com", nombreAnterior, "1kzuhDW8pmxJVUoYty1AzWtSXUsoXom01rwJhoAirpMg", "", "Media Access Company: Detalle de grabaciones", "Detalle::"+tablaEmail+"||Name::"+primeraPalabra+"||Fecha::"+fechaGrabacion, "", "");
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
    AppSendEmailX("appsheet@mediaaccesscompany.com","ar.info@mediaaccesscompany.com", nombreAnterior, "1kzuhDW8pmxJVUoYty1AzWtSXUsoXom01rwJhoAirpMg", "", "Media Access Company: Detalle de grabaciones", "Detalle::"+tablaEmail+"||Name::"+primeraPalabra+"||Fecha::"+fechaGrabacion, "", "");
  }
  
  return nuevoHTML;
}