const EMAIL_DESDE = "info@tecnokids.com";
const NOMBRE_HOJA_DATOS = "email_sender";
const TAMANIO_LOTE = 75;
const INTERVALO_MINUTOS = 10;
const LIMITE_CORREOS = 1500;
const MINUTOS_ESPERA_PRIMER_LOTE = 3;
const FILA_INICIO_LOTES = 5;

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Campaña Email')
    .addItem('Iniciar envíos', 'iniciarEnvioConTriggers')
    .addItem('Suspender envíos', 'suspenderEnvios')
    .addItem('Limpiar columna Estado', 'limpiarColumnaEstado')
    .addToUi();
}

function iniciarEnvioConTriggers() {
  const ui = SpreadsheetApp.getUi();
  const respuesta = ui.alert(
    "Confirmación",
    "Se iniciará el envío de correos por lotes con triggers temporizados.\n\nEsto borrará los estados anteriores. ¿Deseás continuar?",
    ui.ButtonSet.YES_NO
  );
  if (respuesta !== ui.Button.YES) return;

  SpreadsheetApp.getActiveSpreadsheet().toast("⏳ Programando los envíos, por favor esperá unos segundos...");

  // Limpiar triggers anteriores
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === 'enviarLoteTrigger') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOMBRE_HOJA_DATOS);
  const ultimaFila = sheet.getLastRow();

  // Limpiar y dejar pruebas marcadas
  sheet.getRange("D2:D" + ultimaFila).clearContent();
  sheet.getRange("D2:D4").setValue("Correo de prueba");

  enviarCorreosDePruebaDesdeFilas(sheet);

  const dataOriginal = sheet.getRange(FILA_INICIO_LOTES, 1, ultimaFila - FILA_INICIO_LOTES + 1, 4).getValues();
  const data = dataOriginal
    .map((row, i) => ({ rowIndex: i + FILA_INICIO_LOTES, nombre: row[0], email: row[1], curso: row[2] }))
    .filter(r => r.nombre || r.email);

  if (data.length > LIMITE_CORREOS) {
    ui.alert("Hay más de 1500 correos en la lista. No se iniciará el envío.");
    return;
  }

  const lotes = [];
  for (let i = 0; i < data.length; i += TAMANIO_LOTE) {
    lotes.push(data.slice(i, i + TAMANIO_LOTE));
  }

  const props = PropertiesService.getScriptProperties();
  props.deleteAllProperties();
  props.setProperty('TOTAL_LOTES', lotes.length);

  const now = new Date();
  for (let i = 0; i < lotes.length; i++) {
    const lote = lotes[i];
    props.setProperty('LOTE_' + i, JSON.stringify(lote));

    const offsetMin = MINUTOS_ESPERA_PRIMER_LOTE + i * INTERVALO_MINUTOS;
    const triggerTime = new Date(now.getTime() + offsetMin * 60000);
    ScriptApp.newTrigger('enviarLoteTrigger')
      .timeBased()
      .at(triggerTime)
      .create();

    lote.forEach((row) => {
      const estado = `Programado a las ${triggerTime.getHours().toString().padStart(2, '0')}:${triggerTime.getMinutes().toString().padStart(2, '0')}`;
      sheet.getRange(row.rowIndex, 4).setValue(estado);
    });
  }

  ui.alert("¡Envíos programados! Los correos se enviarán automáticamente por lotes.");
}

function enviarCorreosDePruebaDesdeFilas(sheet) {
  const filasTest = sheet.getRange("A2:C4").getValues();
  const asunto = obtenerAsuntoCorreo();
  const html = obtenerPlantillaHTML();

  filasTest.forEach(([nombre, email, curso], i) => {
    if (esEmailValido(email)) {
      const htmlPersonalizado = reemplazarVariables(html, {
        "{Nombre}": nombre,
        "{Curso}": curso
      });

      GmailApp.sendEmail(
        email,
        asunto,
        "",
        {
          htmlBody: htmlPersonalizado,
          from: EMAIL_DESDE,
          name: "Tecnokids (Test)"
        }
      );

      sheet.getRange(i + 2, 4).setValue("Test enviado - " + new Date().toLocaleString());
    }
  });
}

function enviarLoteTrigger() {
  const props = PropertiesService.getScriptProperties();
  const totalLotes = parseInt(props.getProperty('TOTAL_LOTES'), 10);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOMBRE_HOJA_DATOS);

  for (let i = 0; i < totalLotes; i++) {
    const loteKey = 'LOTE_' + i;
    const loteStr = props.getProperty(loteKey);
    if (!loteStr) continue;

    const lote = JSON.parse(loteStr);
    for (const row of lote) {
      const { rowIndex, nombre, email, curso } = row;
      if (!esEmailValido(email)) {
        sheet.getRange(rowIndex, 4).setValue("Error - Email inválido");
        continue;
      }

      try {
        const htmlTemplate = obtenerPlantillaHTML();
        const asunto = obtenerAsuntoCorreo();
        const htmlPersonalizado = reemplazarVariables(htmlTemplate, {
          "{Nombre}": nombre,
          "{Curso}": curso
        });

        GmailApp.sendEmail(
          email,
          asunto,
          "",
          {
            htmlBody: htmlPersonalizado,
            from: EMAIL_DESDE,
            name: "Tecnokids"
          }
        );

        sheet.getRange(rowIndex, 4).setValue("Enviado - " + new Date().toLocaleString());
      } catch (e) {
        sheet.getRange(rowIndex, 4).setValue("Error - " + e.message);
      }
    }

    props.deleteProperty(loteKey);
    break;
  }

  const restantes = Object.keys(props.getProperties()).filter(k => k.startsWith("LOTE_")).length;
  if (restantes === 0) {
    enviarResumenFinal();
  }
}

function suspenderEnvios() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOMBRE_HOJA_DATOS);
  const estadoRango = sheet.getRange("D2:D" + sheet.getLastRow());
  const estadoValores = estadoRango.getValues();

  const nuevos = estadoValores.map(([val]) => {
    if (val && val.toString().startsWith("Programado")) {
      return ["Envío suspendido"];
    }
    return [val];
  });
  estadoRango.setValues(nuevos);

  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === 'enviarLoteTrigger') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  PropertiesService.getScriptProperties().deleteAllProperties();
  SpreadsheetApp.getUi().alert("Envíos suspendidos y triggers eliminados.");
}

function limpiarColumnaEstado() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOMBRE_HOJA_DATOS);
  const ultimaFila = sheet.getLastRow();
  sheet.getRange("D2:D" + ultimaFila).clearContent();
  sheet.getRange("D2:D4").setValue("Correo de prueba");
  SpreadsheetApp.getUi().alert("Columna Estado limpiada. Las filas de prueba fueron marcadas.");
}

function obtenerAsuntoCorreo() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOMBRE_HOJA_DATOS);
  return sheet.getRange("G10").getValue().toString().trim();
}

function obtenerPlantillaHTML() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOMBRE_HOJA_DATOS);
  const fileUrl = sheet.getRange("G11").getValue().toString().trim();
  if (!fileUrl) return "";

  const fileId = extraerFileIdDesdeUrl(fileUrl);
  const file = DriveApp.getFileById(fileId);
  return file.getBlob().getDataAsString("UTF-8");
}

function extraerFileIdDesdeUrl(url) {
  const match = url.match(/[-\w]{25,}/);
  if (match) return match[0];
  throw new Error("No se pudo extraer el ID del archivo de la URL: " + url);
}

function reemplazarVariables(template, variables) {
  let result = template;
  for (const key in variables) {
    result = result.replaceAll(key, variables[key] || "");
  }
  return result;
}

function esEmailValido(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function enviarResumenFinal() {
  const usuario = Session.getActiveUser().getEmail();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOMBRE_HOJA_DATOS);
  const urlHoja = SpreadsheetApp.getActiveSpreadsheet().getUrl();

  const estados = sheet.getRange("D2:D" + sheet.getLastRow()).getValues();
  const enviados = estados.filter(val => val[0] && val[0].toString().startsWith("Enviado")).length;
  const errores = estados.filter(val => val[0] && val[0].toString().startsWith("Error")).length;

  const cuerpo = `
    <p>Se completó el envío por lotes.</p>
    <ul>
      <li>Enviados correctamente: ${enviados}</li>
      <li>Errores: ${errores}</li>
    </ul>
    <p><a href="${urlHoja}" target="_blank">📄 Abrir planilla</a></p>
  `;

  GmailApp.sendEmail(usuario, "Resumen de envío masivo", "", {
    htmlBody: cuerpo
  });
}
