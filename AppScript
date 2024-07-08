function fetchEmailData() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var searchQuery = 'label:custom-label';

// estas lineas son para definir la hoja cuando se llama por primera vez
  // borrar la hoja completa
  // sheet.getDataRange().clearContent();
  // definir la primera fila de las columnas
//  sheet.getRange("A1:E1").setValues([["Id", "Asunto", "LogSource", "Ubicacion", "Hora recepción"]]);

  // obtener los datos actuales en la hoja
  var existingData = sheet.getDataRange().getValues();
  var numRows = existingData.length;

  var threads = GmailApp.search(searchQuery);
  var data = [];

  // Definir la función extractSubject fuera del bucle
  function extractSubject(subject) {
    var regex = /([^|]+)\|([^|]+)\|([^|]+)\|([^|]+)/;
    var match = subject.match(regex);

    if (match) {
      // Asigna los campos a variables
      var subject_id = match[1].trim();
      var subject_asunto = match[2].trim();
      var subject_logSource = match[3].trim();
      var subject_ubicacion = match[4].trim();
      // Devuelve los campos en un array
      return [subject_id, subject_asunto, subject_logSource, subject_ubicacion];
    } else {
      // Si no hay coincidencia, devuelve null
      return null;
    }
  }

  for (var i = 0; i < threads.length; i++) {
    var firstMessage = threads[i].getMessages()[0]; 

    if (firstMessage) {
      var subject = firstMessage.getSubject();
      var subjectParseado = extractSubject(subject);

      if (subjectParseado) {
        var id = subjectParseado[0];
        var asunto = subjectParseado[1];
        var logSource = subjectParseado[2];
        var ubicacion = subjectParseado[3];
        var time = firstMessage.getDate();

        // Añadir los datos al array
        data.push([id, asunto, logSource, ubicacion, time]);
      } else {
        Logger.log("El asunto no coincide con el patrón esperado: " + subject);
      }
    } else {
      Logger.log("No se encontraron mensajes en el hilo.");
    }
  }

  // escribir data a la hoja, este codigo es solo cuando esta nuevo
//  if (data.length > 0) {
//    sheet.getRange(2, 1, data.length, 5).setValues(data);
//  } else {
//    Logger.log("No hay datos para escribir en la hoja.");
//  }

    // Escribir data al final de la hoja
  if (data.length > 0) {
    sheet.getRange(numRows + 1, 1, data.length, 5).setValues(data);
  } else {
    Logger.log("No hay nuevos datos para escribir en la hoja.");
  }
}
