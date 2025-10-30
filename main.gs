/**
 * ===============================================================
 * ==== REGISTRO PEREGRINACIÓN NACIONAL JUVENIL CUBILETE 2026 ====
 * ===============================================================
 * Script para automatizar la generación y envío de pases personalizados
 * a partir de las respuestas de un formulario de Google.
 */

/**
 * Se ejecuta al abrir la hoja de cálculo.
 * Crea el menú personalizado "Herramientas de Pases" en la interfaz.
 */
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Herramientas de Pases')
      .addItem('Generar Pase(s) para Fila(s) Seleccionada(s)', 'procesarFilasManualmente')
      .addToUi();
}

/**
 * Obtiene las variables de configuración desde la hoja "Configuracion".
 * Centraliza los IDs y correos para un mantenimiento fácil.
 * @returns {object} Un objeto con los IDs de la plantilla, la carpeta y el email del admin.
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
 * Función principal. Procesa los datos de una fila para generar un pase,
 * guardarlo en Drive y enviarlo por correo electrónico.
 * @param {object} datosFila Un objeto con {nombreCompleto, nombrePreferido, email, fila}.
 */
function generarYEnviarPase(datosFila) {
  const CONFIG = getConfig();
  const hojaRespuestas = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Respuestas de formulario 2');
  const { nombreCompleto, nombrePreferido, email, fila } = datosFila;

  // Define las columnas para actualizar el estado del registro.
  const COLUMNA_ESTADO = 29;
  const COLUMNA_ENLACE_PDF = 30;

  try {
    if (!nombreCompleto || !email) {
      throw new Error(`El nombre o el correo están vacíos en la fila ${fila}`);
    }

    const nombreParaSaludo = nombrePreferido || nombreCompleto.split(' ')[0];
    Logger.log(`Procesando Fila ${fila}: ${nombreCompleto}`);

    // --- 1. Generación de Código QR ---
    const qrApiUrl = `https://api.qrserver.com/v1/create-qr-code/?size=150x150&data=${encodeURIComponent(nombreCompleto)}`;
    const qrBlob = UrlFetchApp.fetch(qrApiUrl).getBlob().setName('codigoQR.png');

    // --- 2. Creación y Personalización de la Presentación ---
    const plantilla = DriveApp.getFileById(CONFIG.slideTemplateId);
    const nombreArchivo = `Pase - ${nombreCompleto}`;
    const nuevaCopiaId = plantilla.makeCopy(nombreArchivo).getId();
    
    const presentacion = SlidesApp.openById(nuevaCopiaId);
    const diapositiva = presentacion.getSlides()[0];
    diapositiva.replaceAllText('{{NOMBRE}}', nombreCompleto);
    
    // Reemplaza la primera forma (shape) en la diapositiva con el QR.
    const formas = diapositiva.getShapes();
    if (formas.length > 0) {
        formas[0].replaceWithImage(qrBlob);
    }
    presentacion.saveAndClose();

    // --- 3. Conversión a PDF y Almacenamiento en Drive ---
    const archivoSlide = DriveApp.getFileById(nuevaCopiaId);
    const carpetaDestino = DriveApp.getFolderById(CONFIG.pdfFolderId);
    const archivoPDF = archivoSlide.getAs('application/pdf').setName(nombreArchivo + '.pdf');
    const pdfGuardado = carpetaDestino.createFile(archivoPDF);
    const pdfUrl = pdfGuardado.getUrl();
    
    // --- 4. Preparación y Envío del Correo Electrónico ---
    const asunto = '✅ ¡Tu lugar está confirmado! La aventura a Cristo Rey ⛰️ comienza hoy ⭐';
    const cuerpoHtml = `...`; // El cuerpo HTML permanece igual, se omite por brevedad.

    GmailApp.sendEmail(email, asunto, "Tu boleto ha llegado.", {
      htmlBody: cuerpoHtml,
      attachments: [pdfGuardado.getBlob()],
      name: 'Coordinación Jokmah 2026'
    });
    Logger.log(`Correo enviado a ${email}.`);

    // --- 5. Limpieza y Actualización de la Hoja de Cálculo ---
    archivoSlide.setTrashed(true); // Elimina la presentación temporal.
    hojaRespuestas.getRange(fila, COLUMNA_ESTADO).setValue('Pase Enviado').setBackground('#d9ead3');
    hojaRespuestas.getRange(fila, COLUMNA_ENLACE_PDF).setValue(pdfUrl);

  } catch (error) {
    // --- Manejo de Errores ---
    // Si algo falla, registra el error, actualiza la hoja y notifica al administrador.
    Logger.log(`Error en fila ${fila}: ${error.stack}`);
    hojaRespuestas.getRange(fila, COLUMNA_ESTADO).setValue(`ERROR: ${error.message}`).setBackground('#f4cccc');
    
    const asuntoError = `Error al generar pase PNJ para ${nombreCompleto}`;
    const cuerpoError = `Error procesando la fila ${fila}:\n\nNombre: ${nombreCompleto}\nEmail: ${email}\n\nDetalle: ${error.stack}`;
    GmailApp.sendEmail(CONFIG.adminEmail, asuntoError, cuerpoError);
  }
}

/**
 * Se activa automáticamente cuando se envía una nueva respuesta del formulario.
 * Extrae los datos del evento y llama a la función principal.
 * @param {object} e El objeto de evento que contiene los datos del formulario.
 */
function activadorOnFormSubmit(e) {
  try {
    const datos = {
      email: e.values[1],
      nombreCompleto: e.values[3],
      nombrePreferido: e.values[4],
      fila: e.range.getRow()
    };
    generarYEnviarPase(datos);
  } catch (error) {
    // Captura errores críticos que puedan ocurrir antes de llamar a la función principal.
    const CONFIG = getConfig();
    const asuntoError = "Error CRÍTICO en la automatización de Pases";
    const cuerpoError = `El script falló al leer los datos iniciales del formulario. Revisa la estructura de columnas.\n\nError: ${error.stack}`;
    GmailApp.sendEmail(CONFIG.adminEmail, asuntoError, cuerpoError);
    Logger.log(`Error CRÍTICO en activadorOnFormSubmit: ${error.stack}`);
  }
}

/**
 * Procesa las filas que el usuario selecciona manualmente en la hoja.
 * Se activa desde el menú personalizado.
 */
function procesarFilasManualmente() {
    const hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const rangoSeleccionado = hoja.getActiveRange();

    // Medida de seguridad para no exceder los límites de ejecución de Google.
    if (rangoSeleccionado.getHeight() > 50) {
        SpreadsheetApp.getUi().alert("Para evitar exceder los límites, por favor procesa menos de 50 filas a la vez.");
        return;
    }

    const filas = rangoSeleccionado.getValues();
    const encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];

    // Busca dinámicamente el índice de cada columna por su nombre.
    // Esto hace el script más robusto si las columnas cambian de orden.
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
        // Procesa la fila solo si contiene la información esencial.
        if (datos.nombreCompleto && datos.email) {
            generarYEnviarPase(datos);
        }
    });
}
