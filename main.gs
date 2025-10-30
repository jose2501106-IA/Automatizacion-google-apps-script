// ===============================================================
// ==== REGISTRO PEREGRINACIÓN NACIONAL JUVENIL CUBILETE 2026 ====
// ===============================================================

/**
 * Se ejecuta automáticamente cada vez que abres la hoja de cálculo.
 * Crea un menú personalizado llamado "Herramientas de Pases" en la interfaz.
 */
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Herramientas de Pases')
      .addItem('Generar Pase(s) para Fila(s) Seleccionada(s)', 'procesarFilasManualmente')
      .addToUi();
}

/**
 * Lee las variables de configuración desde la hoja "Configuracion".
 */
function getConfig() {
  const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Configuracion');
  if (!configSheet) {
    throw new Error('No se pudo encontrar la hoja "Configuracion".');
  }
  return {
    slideTemplateId: configSheet.getRange('B1').getValue(),
    pdfFolderId: configSheet.getRange('B2').getValue(),
    adminEmail: configSheet.getRange('B3').getValue()
  };
}

/**
 * Esta es la función central que hace todo el trabajo.
 * @param {object} datosFila - Un objeto con {nombreCompleto, nombrePreferido, email, fila}.
 */
function generarYEnviarPase(datosFila) {
  const CONFIG = getConfig();
  const hojaRespuestas = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Respuestas de formulario 2'); 
  const { nombreCompleto, nombrePreferido, email, fila } = datosFila;
  
  const COLUMNA_ESTADO = 29;
  const COLUMNA_ENLACE_PDF = 30;

  try {
    if (!nombreCompleto || !email) {
      throw new Error("El nombre completo o el correo están vacíos en la fila " + fila);
    }
    
    const nombreParaSaludo = nombrePreferido || nombreCompleto.split(' ')[0];
    
    Logger.log(`Procesando Fila ${fila}: ${nombreCompleto} (${email})`);

    const qrApiUrl = `https://api.qrserver.com/v1/create-qr-code/?size=150x150&data=${encodeURIComponent(nombreCompleto)}`;
    const qrBlob = UrlFetchApp.fetch(qrApiUrl).getBlob().setName('codigoQR.png');

    const plantilla = DriveApp.getFileById(CONFIG.slideTemplateId);
    const nombreArchivo = `Pase - ${nombreCompleto}`;
    const nuevaCopiaId = plantilla.makeCopy(nombreArchivo).getId();
    
    const presentacion = SlidesApp.openById(nuevaCopiaId);
    const diapositiva = presentacion.getSlides()[0];
    diapositiva.replaceAllText('{{NOMBRE}}', nombreCompleto);
    const formas = diapositiva.getShapes();
    if (formas.length > 0) formas[0].replaceWithImage(qrBlob);
    else Logger.log('ADVERTENCIA: No se encontró forma para el QR.');
    presentacion.saveAndClose();

    const archivoSlide = DriveApp.getFileById(nuevaCopiaId);
    const carpetaDestino = DriveApp.getFolderById(CONFIG.pdfFolderId);
    const archivoPDF = archivoSlide.getAs('application/pdf').setName(nombreArchivo + '.pdf');
    const pdfGuardado = carpetaDestino.createFile(archivoPDF);
    const pdfUrl = pdfGuardado.getUrl();
    
    // ✅ AJUSTE FINAL: Usando el emoji de estrella (⭐) por su máxima compatibilidad.
    const asunto = '✅ ¡Tu lugar está confirmado! La aventura a Cristo Rey ⛰️ comienza hoy ⭐';
    
    const cuerpoHtml = `
      <!DOCTYPE html>
      <html lang="es"><head> <meta charset="UTF-8"> <meta name="viewport" content="width=device-width, initial-scale=1.0"> <title>¡Tu lugar está confirmado! La aventura a Cristo Rey comienza hoy.</title> </head><body style="margin: 0; padding: 0; background-color: #f4f4f4;"> <center style="width: 100%; table-layout: fixed; background-color: #f4f4f4; padding: 20px 0;"> <table style="width: 100%; max-width: 600px; margin: 0 auto; background-color: #ffffff; border-spacing: 0; font-family: Arial, Helvetica, sans-serif; color: #333333;"> <tr> <td style="padding: 40px 30px 30px 30px; text-align: left;"> <h1 style="font-size: 28px; font-weight: bold; color: #2c3e50; margin-top: 0; margin-bottom: 20px;"> ¡Hola, ${nombreParaSaludo}! </h1> <p style="font-size: 18px; line-height: 1.6; margin: 0 0 20px 0;"> Hay una llamada que se siente en lo profundo, una inquietud por buscar <u>algo más grande</u>. Y tú, con <strong>valentía</strong>, la has respondido. </p> <p style="font-size: 18px; line-height: 1.6; margin: 0 0 20px 0;"> ¡<strong>Oficialmente</strong>, eres parte de la <strong>Peregrinación Nacional Juvenil a Cristo Rey 2026</strong>! Tu "sí" ha resonado y estamos increíblemente emocionados de que te sumes a esta generación de jóvenes que no tiene miedo de ponerse en camino. </p> <p style="font-size: 18px; line-height: 1.6; margin: 0 0 30px 0;"> Este viaje es mucho más que solo llegar a la cima de una montaña. Es la oportunidad de <strong>encontrarte a ti mismo</strong>, de forjar amistades que duran toda la vida y de conectar con Dios de una manera que <u>te cambiará para siempre</u>. </p> <table style="width: 100%; border-spacing: 0;"> <tr> <td style="padding-top: 15px; border-top: 2px solid #eeeeee;"> <h2 style="font-size: 24px; color: #2c3e50; margin: 0 0 15px 0;"> Tu Llave a la Experiencia </h2> <p style="font-size: 18px; line-height: 1.6; margin: 0 0 20px 0;"> Adjunto a este correo encontrarás tu boleto personalizado. Pero queremos que lo veas como lo que realmente es: <strong>la primera pieza de un mapa</strong> que te llevará a vivir algo <u>extraordinario</u>. </p> <p style="font-size: 16px; line-height: 1.6; margin: 0 0 25px 0;"> Este documento contiene:<br> • Tu nombre de peregrino.<br> • Tu código QR único para un acceso ágil y rápido. </p> <table role="presentation" style="width: 100%; border-spacing: 0;"> <tr> <td> <p style="font-size: 16px; font-weight: bold; margin: 0 0 10px 0; text-align: center;"> &#128071; EL PRIMER PASO DE TU CAMINO &#128071; </p> <center> <a href="${pdfUrl}" target="_blank" style="background-color: #e67e22; color: #ffffff; font-size: 18px; font-weight: bold; text-decoration: none; padding: 15px 30px; border-radius: 8px; display: inline-block;"> DESCARGAR MI BOLETO AHORA </a> </center> <p style="font-size: 14px; text-align: center; color: #7f8c8d; margin: 15px 0 0 0;"> <em>Te recomendamos guardarlo en tu celular y asegurarte de que tus datos estén correctos. ¡Es tu pasaporte a la aventura!</em> </p> </td> </tr> </table> </td> </tr> </table> <table role="presentation" style="width: 100%; border-spacing: 0; margin-top:30px;"> <tr> <td style="padding-top: 30px; border-top: 2px solid #eeeeee;"> <h2 style="font-size: 24px; color: #2c3e50; margin: 0 0 15px 0;"> La Comunidad ya te Espera </h2> <p style="font-size: 18px; line-height: 1.6; margin: 0 0 25px 0;"> La peregrinación no empieza en el Cubilete, <strong>¡empieza ahora!</strong> Conéctate con otros peregrinos, comparte tu emoción y prepárate en comunidad. </p> <table role="presentation" style="width: 100%; border-spacing: 0;"> <tr> <td style="padding-bottom: 15px;"> <center> <a href="https://chat.whatsapp.com/D4KN9Fawze4JxaGd0lGMJV?mode=wwt" target="_blank" style="background-color: #25D366; color: #ffffff; font-size: 18px; font-weight: bold; text-decoration: none; padding: 15px 30px; border-radius: 8px; display: block; max-width: 300px;"> Unirse al Grupo de WhatsApp </a> </center> </td> </tr> </table> <table role="presentation" style="width: 100%; border-spacing: 0; margin-top:10px;"> <tr> <td> <center> <a href="https://www.instagram.com/jokmah_sf/" target="_blank" style="background-color: #3498db; color: #ffffff; font-size: 18px; font-weight: bold; text-decoration: none; padding: 15px 30px; border-radius: 8px; display: block; max-width: 300px;"> Seguir en Instagram </a> </center> </td> </tr> </table> </td> </tr> </table> <p style="font-size: 18px; line-height: 1.6; margin: 40px 0 10px 0;"> En las últimas semanas, te compartiremos guías de preparación, historias inspiradoras y todos los detalles para que vivas esta experiencia al máximo. </p> <p style="font-size: 18px; line-height: 1.6; font-weight: bold; margin: 0 0 20px 0;"> Tu lugar en la montaña ya tiene tu nombre. ¡Qué emoción saber que nos veremos pronto! </p> <p style="font-size: 18px; line-height: 1.6; margin: 30px 0 0 0;"> En el camino, </p> <p style="font-size: 18px; line-height: 1.6; font-weight: bold; margin: 0;"> Coordinación Jokmah<br> Peregrinación Nacional Juvenil 2026 </p> </td> </tr> </table> </center> </body></html>
    `;

    GmailApp.sendEmail(email, asunto, "Tu boleto ha llegado.", {
      htmlBody: cuerpoHtml,
      attachments: [pdfGuardado.getBlob()],
      name: 'Coordinación Jokmah 2026'
    });
    Logger.log(`Correo enviado a ${email}.`);

    archivoSlide.setTrashed(true);
    hojaRespuestas.getRange(fila, COLUMNA_ESTADO).setValue('Pase Enviado').setBackground('#d9ead3');
    hojaRespuestas.getRange(fila, COLUMNA_ENLACE_PDF).setValue(pdfUrl);
  } catch (error) {
    Logger.log(`Error en fila ${fila}: ${error.stack}`);
    hojaRespuestas.getRange(fila, COLUMNA_ESTADO).setValue(`ERROR: ${error.message}`).setBackground('#f4cccc');
    const asuntoError = `Error al generar pase PNJ para ${nombreCompleto}`;
    const cuerpoError = `Error procesando la fila ${fila}:\n\nNombre: ${nombreCompleto}\nEmail: ${email}\n\nDetalle: ${error.stack}`;
    GmailApp.sendEmail(CONFIG.adminEmail, asuntoError, cuerpoError);
  }
}

/**
 * Se activa AUTOMÁTICAMENTE con el envío del formulario.
 */
function activadorOnFormSubmit(e) {
  try {
    const valores = e.values;
    const datos = {
      email: valores[1],
      nombreCompleto: valores[3], 
      nombrePreferido: valores[4],
      fila: e.range.getRow()
    };
    generarYEnviarPase(datos);
  } catch (error) {
    const CONFIG = getConfig();
    const asuntoError = "Error CRÍTICO en la automatización de Pases";
    const cuerpoError = "El script falló al intentar leer los datos iniciales. Revisa la estructura de columnas.\n\nError: " + error.stack;
    GmailApp.sendEmail(CONFIG.adminEmail, asuntoError, cuerpoError);
    Logger.log("Error CRÍTICO: " + error.toString());
  }
}

/**
 * Se activa MANUALMENTE desde el menú personalizado.
 */
function procesarFilasManualmente() {
    const hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const rangoSeleccionado = hoja.getActiveRange();

    if (rangoSeleccionado.getHeight() > 50) {
        SpreadsheetApp.getUi().alert("Para evitar exceder los límites, por favor procesa menos de 50 filas a la vez.");
        return;
    }

    const filas = rangoSeleccionado.getValues();
    const encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];

    const indiceEmail = encabezados.indexOf("Dirección de correo electrónico");
    const indiceNombreCompleto = encabezados.indexOf("Nombre completo:");
    const indiceNombrePreferido = encabezados.indexOf("¿Cómo te gusta que te llamen?");
    
    filas.forEach((datosFila, index) => {
        const numeroFila = rangoSeleccionado.getRow() + index;
        const datos = {
            email: datosFila[indiceEmail],
            nombreCompleto: datosFila[indiceNombreCompleto],
            nombrePreferido: datosFila[indiceNombrePreferido],
            fila: numeroFila
        };
        if (datos.nombreCompleto && datos.email) {
            generarYEnviarPase(datos);
        }
    });
}
