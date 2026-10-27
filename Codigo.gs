function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Portal de Búsqueda y Actualización de Moldería')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function formatearFechaInput(fecha) {
  if (!fecha || !(fecha instanceof Date)) return '';
  try {
    const d = new Date(fecha);
    if (isNaN(d.getTime())) return '';
    const anio = d.getFullYear();
    const mes = String(d.getMonth() + 1).padStart(2, '0');
    const dia = String(d.getDate()).padStart(2, '0');
    return `${anio}-${mes}-${dia}`;
  } catch (e) {
    return '';
  }
}

function formatearFecha(fecha) {
  if (!fecha) return '';
  if (fecha instanceof Date) {
    try {
      return Utilities.formatDate(fecha, Session.getScriptTimeZone(), 'dd/MM/yyyy');
    } catch (e) {
      return fecha.toString();
    }
  }
  return fecha.toString();
}

function buscarReferencia(referencia) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('MOLDERIA') || ss.getSheets()[0];
    const data = sheet.getDataRange().getValues();
    
    const refBuscada = String(referencia).trim();
    
    if (!refBuscada) {
      return { exito: false, mensaje: 'Por favor ingresa una referencia válida.' };
    }

    // Buscar en la columna F (índice 5)
    for (let i = 1; i < data.length; i++) {
      let refFila = String(data[i][5]).trim();
      
      if (refFila === refBuscada) {
        return {
          exito: true,
          numFila: i + 1,
          referencia: data[i][5],       // Col F
          descripcion: data[i][6],      // Col G
          grupo: data[i][7],            // Col H
          fEntregaIdeal: formatearFecha(data[i][10]), // Col K
          
          fEnvioCliente: formatearFechaInput(data[i][13]),         // Col N (14)
          fAprobadaCliente: formatearFechaInput(data[i][14]),      // Col O (15)
          fRoperoFinal: formatearFechaInput(data[i][15]),          // Col P (16)
          estadoMuestras: data[i][16] || 'APROBADA',             // Col Q (17)
          patronista: data[i][17] || '',                           // Col R (18)
          fInicioMolderia: formatearFechaInput(data[i][18]),       // Col S (19)
          llevaPlantillasConfeccion: data[i][19] === true || String(data[i][19]).toUpperCase() === 'VERDADERO' || String(data[i][19]).toUpperCase() === 'TRUE', // Col T (20)
          llevaPlantillasProcesos: data[i][20] === true || String(data[i][20]).toUpperCase() === 'VERDADERO' || String(data[i][20]).toUpperCase() === 'TRUE',   // Col U (21)
          fFinMolderia: formatearFechaInput(data[i][21]),          // Col V (22)
          observaciones: data[i][22] || ''                         // Col W (23)
        };
      }
    }
    
    return { exito: false, mensaje: 'La referencia ' + refBuscada + ' no fue encontrada.' };
    
  } catch (error) {
    return { exito: false, mensaje: 'Error al buscar: ' + error.toString() };
  }
}

function actualizarReferencia(datos) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('MOLDERIA') || ss.getSheets()[0];
    const fila = parseInt(datos.numFila);
    
    if (!fila || fila < 2) {
      return { exito: false, mensaje: 'Fila no válida para actualizar.' };
    }

    if (datos.fEnvioCliente) sheet.getRange(fila, 14).setValue(datos.fEnvioCliente); else sheet.getRange(fila, 14).clearContent();
    if (datos.fAprobadaCliente) sheet.getRange(fila, 15).setValue(datos.fAprobadaCliente); else sheet.getRange(fila, 15).clearContent();
    if (datos.fRoperoFinal) sheet.getRange(fila, 16).setValue(datos.fRoperoFinal); else sheet.getRange(fila, 16).clearContent();
    
    if (datos.estadoMuestras) {
      sheet.getRange(fila, 17).setValue(datos.estadoMuestras);
    } else {
      sheet.getRange(fila, 17).clearContent();
    }

    if (datos.patronista && datos.patronista.trim() !== "") {
      sheet.getRange(fila, 18).setValue(datos.patronista);
    } else {
      sheet.getRange(fila, 18).clearContent();
    }

    if (datos.fInicioMolderia) sheet.getRange(fila, 19).setValue(datos.fInicioMolderia); else sheet.getRange(fila, 19).clearContent();
    
    sheet.getRange(fila, 20).setValue(datos.llevaPlantillasConfeccion);   // Col T (20)
    sheet.getRange(fila, 21).setValue(datos.llevaPlantillasProcesos);     // Col U (21)
    
    if (datos.fFinMolderia) sheet.getRange(fila, 22).setValue(datos.fFinMolderia); else sheet.getRange(fila, 22).clearContent();
    if (datos.observaciones) sheet.getRange(fila, 23).setValue(datos.observaciones); else sheet.getRange(fila, 23).clearContent();

    registrarHistorial(ss, datos.referencia, datos.patronista);

    return { exito: true, mensaje: '¡Cambio guardado exitosamente!' };
    
  } catch (error) {
    return { exito: false, mensaje: 'Error al actualizar: ' + error.toString() };
  }
}

function registrarHistorial(ss, referencia, patronista) {
  try {
    let histSheet = ss.getSheetByName('Historial');
    if (!histSheet) {
      histSheet = ss.insertSheet('Historial');
      histSheet.appendRow(['FECHA Y HORA', 'REFERENCIA', 'PATRONISTA ASIGNADA']);
    }
    const fechaHora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    histSheet.appendRow([fechaHora, referencia, patronista || 'No asignada']);
  } catch (e) {
    console.error('Error al registrar en historial: ' + e.toString());
  }
}