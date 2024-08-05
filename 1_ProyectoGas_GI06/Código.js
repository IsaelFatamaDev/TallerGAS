const PDF_FOLDER_ID = '1qC3EvNX0AVHBFTd81A6OsSDNmPkKEH-n';
const QR_FOLDER_ID = '1qkLnAk1yYrgDnBHS_KSoirMBMzbu9Cue';
const TEMPLATE_ID = '1UDj-nuDKl_HnZHAABi-XG_YKO7_rH7b2gdjsbU-K-TQ';
const COPY_FOLDER_ID = '17kyh6lKCZYItZSxeeBbg5G8HhfNQ2oYt';
const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Respuestas de formulario 1');

function onOpen() {
	const ui = SpreadsheetApp.getUi();
	ui.createMenu('Procesar Solicitud') 
		.addItem('Ingresar número de fila', 'mostrarDialogoFila')
		.addToUi();
}

function mostrarDialogoFila() {
	const ui = SpreadsheetApp.getUi();
	const response = ui.prompt('Ingresar número de fila', 'Por favor ingresa el número de fila que deseas procesar:', ui.ButtonSet.OK_CANCEL);

	if (response.getSelectedButton() == ui.Button.OK) {
		const fila = parseInt(response.getResponseText(), 10);

		if (!isNaN(fila) && fila > 1 && fila <= sheet.getLastRow()) {
			registrarDato(fila);
		} else {
			ui.alert('Número de fila no válido. Por favor, ingresa un número de fila entre 2 y ' + sheet.getLastRow());
		}
	}
}

function registrarDato(fila) {
	// Verificar si el número de fila es válido
	if (fila < 2 || fila > sheet.getLastRow()) {
		console.log('Número de fila no válido.');
		return;
	}

	var marcaTemporal = Utilities.formatDate(sheet.getRange(fila, 1).getValue(), "America/Bogota", "dd/MMM/yyyy");
	var email = sheet.getRange(fila, 2).getValue();
	var name = sheet.getRange(fila, 3).getValue();
	var lastName = sheet.getRange(fila, 4).getValue();
	var dni = sheet.getRange(fila, 5).getValue();
	var solicitud = sheet.getRange(fila, 6).getValue();
	var firmaImgUrl = sheet.getRange(fila, 7).getValue();

	// Extraer el ID de la imagen de firma 
	var firmaImgId = firmaImgUrl.match(/[-\w]{25,}/);
	if (!firmaImgId) {
		console.log('ID de imagen de firma no encontrado.');
		return;
	}

	var nombreNuevoArchivo = `Solicitud - ${solicitud}`;

	// Eliminar archivos existentes
	eliminarArchivos(nombreNuevoArchivo);

	var templateDoc = DriveApp.getFileById(TEMPLATE_ID).makeCopy();
	var nuevoArchivo = DriveApp.getFileById(templateDoc.getId());

	// Renombrar el archivo
	nuevoArchivo.setName(nombreNuevoArchivo);

	var folder = DriveApp.getFolderById(COPY_FOLDER_ID);
	folder.addFile(nuevoArchivo);

	var originalFolder = DriveApp.getRootFolder();
	originalFolder.removeFile(nuevoArchivo);

	var doc = DocumentApp.openById(nuevoArchivo.getId());
	var body = doc.getBody();

	body.replaceText('<<nombreSolicitud>>', solicitud);
	body.replaceText('<<name>>', name);
	body.replaceText('<<lastName>>', lastName);
	body.replaceText('<<dni>>', dni);

	var qrData = `Fecha: ${marcaTemporal}\nNombre: ${name} ${lastName}\nDNI: ${dni}`;
	var qrCodeUrl = generateQRCode(qrData);

	// Guardar QR como imagen 
	var qrBlob = UrlFetchApp.fetch(qrCodeUrl).getBlob();
	var qrFolder = DriveApp.getFolderById(QR_FOLDER_ID);
	var qrFileName = `${name} - ${dni}.png`;
	var qrFile = qrFolder.createFile(qrBlob).setName(qrFileName);
	var qrFileUrl = qrFile.getUrl();

	// Insertar el QR en el documento
	var qrPosition = body.findText('<<qr>>');
	if (qrPosition) {
		var qrElement = qrPosition.getElement();
		var qrParent = qrElement.getParent().asParagraph();
		qrParent.insertInlineImage(0, qrBlob).setWidth(100).setHeight(100);
		qrElement.asText().setText('');
	}

	// Insertar la firma en el documento
	var firmaFile = DriveApp.getFileById(firmaImgId[0]);
	var firmaBlob = firmaFile.getBlob();
	var firmaPosition = body.findText('<<firmaSol>>');
	if (firmaPosition) {
		var firmaElement = firmaPosition.getElement();
		var firmaParent = firmaElement.getParent().asParagraph();
		firmaParent.insertInlineImage(0, firmaBlob).setWidth(180).setHeight(100);
		firmaElement.asText().setText('');
	}

	doc.saveAndClose();

	// Convertir el documento a PDF
	var pdfBlob = DriveApp.getFileById(nuevoArchivo.getId()).getAs('application/pdf');
	var pdfFolder = DriveApp.getFolderById(PDF_FOLDER_ID);
	var pdfFile = pdfFolder.createFile(pdfBlob).setName(nombreNuevoArchivo + '.pdf');
	var pdfFileUrl = pdfFile.getUrl();

	sheet.getRange(fila, 8).setValue(qrFileUrl);
	sheet.getRange(fila, 9).setValue(pdfFileUrl);

	// Enviar el correo electrónico
	enviarCorreo(email, name, pdfFile, qrFileUrl);

	console.log('Documento generado: ' + doc.getUrl());
	console.log('QR guardado: ' + qrFileUrl);
	console.log('PDF generado: ' + pdfFileUrl);
}

// Función para generar QR 
function generateQRCode(data) {
	var qrServiceUrl = 'https://api.qrserver.com/v1/create-qr-code/';
	var qrCodeUrl = `${qrServiceUrl}?size=200x200&data=${encodeURIComponent(data)}`;
	return qrCodeUrl;
}

// Función para enviar correo electrónico
function enviarCorreo(email, nombre, pdfAdjunto, qrUrl) {
	const asunto = "Registro Exitoso";
	const cuerpo = `
     <div style="font-family: 'Poppins', Arial, sans-serif; color: #333; max-width: 600px; margin: 0 auto; padding: 20px; box-shadow: 0px 0px 15px rgba(0, 0, 0, 0.8); border-radius: 10px; border: 1px solid #ddd;">
          <div style="background-color: #182630; color: white; padding: 15px; border-radius: 10px 10px 0 0; text-align: center;">
               <h2 style="margin: 0; font-size: 24px;">REGISTRO EXITOSO</h2>
          </div>
          <div style="padding: 20px;">
               <p>Hola <strong>${nombre}</strong>,</p>
               <p>Espero que te encuentres bien. Queremos informarte que tu registro ha sido exitoso. A continuación, se adjunta el PDF y el código QR que se han generado según tu solicitud.</p>
               
               <p style="font-style: italic;">Por favor, revisa los documentos adjuntos y avísanos si tienes alguna pregunta o necesitas algún ajuste adicional. Estamos aquí para ayudarte.</p>
               
               <p>Gracias por tu atención y confianza en nosotros.</p>
               <p>Saludos cordiales,</p>
               <p style="font-weight: bold;">Empresa Valle Grande</p>
               
               <hr style="border: none; border-top: 1px solid #ddd; margin: 20px 0;">
               
               <div style="text-align: center;">
               <a href="${qrUrl}" style="display: inline-block; margin-top: 20px; padding: 10px 20px; background-color: #182630; color: white; text-decoration: none; border-radius: 5px; font-weight: bold;">Ver QR</a>
               </div>
          
               <br><br>
          
               <p style="font-size: 12px; color: #666; text-align: center;">P.D.: Si necesitas más información o asistencia, no dudes en ponerte en contacto con nosotros.</p>
          </div>
     </div>`;

	// Enviar correo usando el servicio de Gmail
	GmailApp.sendEmail(email, asunto, '', {
		htmlBody: cuerpo,
		attachments: [pdfAdjunto.getAs(MimeType.PDF)]
	});
}


function eliminarArchivos(nombreBase) {
	const pdfFolder = DriveApp.getFolderById(PDF_FOLDER_ID);
	const pdfFiles = pdfFolder.getFilesByName(nombreBase + '.pdf');
	while (pdfFiles.hasNext()) {
		const pdfFile = pdfFiles.next();
		pdfFile.setTrashed(true);
	}

	const qrFolder = DriveApp.getFolderById(QR_FOLDER_ID);
	const qrFiles = qrFolder.getFilesByName(`${nombreBase}.png`);
	while (qrFiles.hasNext()) {
		const qrFile = qrFiles.next();
		qrFile.setTrashed(true);
	}

	const copyFolder = DriveApp.getFolderById(COPY_FOLDER_ID);
	const copyFiles = copyFolder.getFilesByName(nombreBase);
	while (copyFiles.hasNext()) {
		const copyFile = copyFiles.next();
		copyFile.setTrashed(true);
	}
}
