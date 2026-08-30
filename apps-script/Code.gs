var MAIN_EMAIL = 'seminarios.circuloaikikai@gmail.com';

var HEADERS = [
  'Fecha', 'Nombre', 'Apellido', 'Email', 'Teléfono', 'Ciudad', 'Provincia',
  'DNI', 'Nacimiento', 'Rango', 'Años práctica', 'Dojo', 'Agrupación', 'Instructor',
  'Obra social', 'Nro. afiliado', 'Condición médica',
  'Contacto emergencia', 'Tel. emergencia',
  'Cupo', 'Precio cupo (USD)', 'Práctica especial', 'Método de pago', 'Total (USD)', 'Observaciones',
  'Posible duplicado', 'País', 'Género', 'Estado de pago'
];

function normalizeDni(v) {
  return String(v || '').replace(/[^0-9A-Za-z]/g, '').toUpperCase();
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Inscripciones');
    if (!sheet) sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('Inscripciones');

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
    } else {
      var currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      var missing = HEADERS.slice(currentHeaders.length);
      if (missing.length) {
        sheet.getRange(1, currentHeaders.length + 1, 1, missing.length).setValues([missing]);
      }
    }

    // Chequeo de DNI duplicado contra inscripciones ya cargadas
    var isDuplicate = false;
    var lastRow = sheet.getLastRow();
    var dniCol = HEADERS.indexOf('DNI') + 1;
    var newDni = normalizeDni(data.dni);
    if (lastRow > 1 && newDni) {
      var dniValues = sheet.getRange(2, dniCol, lastRow - 1, 1).getValues();
      for (var i = 0; i < dniValues.length; i++) {
        if (normalizeDni(dniValues[i][0]) === newDni) { isDuplicate = true; break; }
      }
    }

    var rowMap = {
      'Fecha': new Date(),
      'Nombre': data.nombre, 'Apellido': data.apellido, 'Email': data.email, 'Teléfono': data.telefono,
      'Ciudad': data.ciudad, 'Provincia': data.provincia,
      'DNI': data.dni, 'Nacimiento': data.nacimiento, 'Rango': data.rango, 'Años práctica': data.anios,
      'Dojo': data.dojo, 'Agrupación': data.agrupacion, 'Instructor': data.instructor,
      'Obra social': data.obraSocial, 'Nro. afiliado': data.nroAfiliado, 'Condición médica': data.medico,
      'Contacto emergencia': data.emergenciaNombre, 'Tel. emergencia': data.emergenciaTel,
      'Cupo': data.cupoNombre, 'Precio cupo (USD)': data.cupoPrecio,
      'Práctica especial': data.addon ? 'Sí' : 'No', 'Método de pago': data.pago,
      'Total (USD)': data.total, 'Observaciones': data.observaciones,
      'Posible duplicado': isDuplicate ? 'SÍ - revisar' : '',
      'País': data.pais, 'Género': data.genero || 'No especificado',
      'Estado de pago': 'Pendiente'
    };
    sheet.appendRow(HEADERS.map(function (h) { return rowMap[h] !== undefined ? rowMap[h] : ''; }));

    var body =
      (isDuplicate ? '⚠ POSIBLE DUPLICADO — ya existe una inscripción con este DNI. Revisar antes de confirmar el pago.\n\n' : '') +
      'Nueva inscripción recibida:\n\n' +
      'Nombre: ' + data.nombre + ' ' + data.apellido + '\n' +
      'Email: ' + data.email + '\n' +
      'Teléfono: ' + data.telefono + '\n' +
      'Ciudad: ' + data.ciudad + (data.provincia ? ', ' + data.provincia : '') + (data.pais ? ', ' + data.pais : '') + '\n' +
      'DNI: ' + data.dni + '\n' +
      'Rango: ' + data.rango + '\n' +
      'Dojo: ' + data.dojo + '\n' +
      'Cupo: ' + data.cupoNombre + ' (USD ' + data.cupoPrecio + ')\n' +
      'Práctica especial: ' + (data.addon ? 'Sí' : 'No') + '\n' +
      'Método de pago: ' + data.pago + '\n' +
      'Total: USD ' + data.total + '\n' +
      (data.medico ? 'Condición médica: ' + data.medico + '\n' : '') +
      (data.observaciones ? 'Observaciones: ' + data.observaciones + '\n' : '') +
      '\nContacto de emergencia: ' + data.emergenciaNombre + ' — ' + data.emergenciaTel +
      '\n\n---\nEste email es tu constancia de inscripción. Recordá que la inscripción se confirma una vez acreditado el pago.';

    var ccList = [data.contactEmail, data.email].filter(function (x) { return !!x; }).join(',');

    MailApp.sendEmail({
      to: MAIN_EMAIL,
      cc: ccList,
      subject: (isDuplicate ? '[POSIBLE DUPLICADO] ' : '') + 'Nueva inscripción — ' + data.nombre + ' ' + data.apellido,
      body: body
    });

    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// El id 'clase-especial' en "Precios" corresponde a la práctica especial
// (specialPractice en config.js), no a un cupo del seminario — se cuenta distinto
// (por la columna "Práctica especial" = "Sí", no por la columna "Cupo").
var CLASE_ESPECIAL_ID = 'clase-especial';

function esCerrado(v) {
  return v === true || String(v).trim().toUpperCase() === 'TRUE' || String(v).trim().toUpperCase() === 'VERDADERO';
}

// doGet ya NO lee "Inscripciones" — todo el conteo y la validación de cupos
// vive en fórmulas dentro de la pestaña "Precios" (columnas F "inscriptos" y
// G "cerrado calculado"), así el endpoint público solo expone precio, cupo
// total y si está cerrado — nunca cuántas inscripciones reales hay.
function doGet(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var precios = ss.getSheetByName('Precios');

    var result = [];
    if (precios) {
      var lastRow = precios.getLastRow();
      if (lastRow > 1) {
        precios.getRange(2, 1, lastRow - 1, 7).getValues().forEach(function (r) {
          if (!r[0]) return;
          result.push({
            id: r[0],
            priceUsd: Number(r[2]),
            slots: Number(r[3]),
            soldOut: esCerrado(r[6]) // columna G: "cerrado calculado"
          });
        });
      }
    }

    return ContentService.createTextOutput(JSON.stringify({ ok: true, tiers: result }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function setupPreciosSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Precios');
  if (sheet) return; // ya existe — no la toca, para no perder ediciones manuales

  sheet = ss.insertSheet('Precios');
  sheet.getRange('A1:E1').setValues([['id', 'nombre', 'precioUsd', 'cupos', 'cerrado manual']]).setFontWeight('bold');
  sheet.getRange('A2:E6').setValues([
    ['super-early-bird', 'Super Early Bird', 80, 20, false],
    ['early-bird', 'Early Bird', 90, 20, false],
    ['regular', 'Regular', 100, 20, false],
    ['last-minute', 'Last Minute', 110, 20, false],
    [CLASE_ESPECIAL_ID, 'Práctica privada', 30, 40, false],
  ]);
  sheet.getRange('E2:E6').insertCheckboxes();
  setPreciosFormulas(sheet);
  sheet.autoResizeColumns(1, 7);
}

// Agrega las columnas F ("inscriptos") y G ("cerrado calculado") a una pestaña
// "Precios" que ya existe (por ejemplo, si se armó a mano antes de que
// existieran estas columnas), sin tocar A-E. Correr una sola vez.
// Requiere que la columna E ("cerrado manual", casillas TRUE/FALSE) ya exista.
function upgradePreciosSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Precios');
  if (!sheet) throw new Error('No existe la pestaña "Precios". Correr primero setupPreciosSheet().');
  if (sheet.getLastColumn() < 5) throw new Error('Falta la columna E ("cerrado manual") — agregala primero (casillas TRUE/FALSE).');
  setPreciosFormulas(sheet);
  sheet.autoResizeColumns(1, 7);
}

function setPreciosFormulas(sheet) {
  sheet.getRange('F1').setValue('inscriptos').setFontWeight('bold');
  sheet.getRange('G1').setValue('cerrado calculado').setFontWeight('bold');

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    var row = i + 2;
    var id = ids[i][0];
    if (!id) continue;
    sheet.getRange('F' + row).setFormula(
      id === CLASE_ESPECIAL_ID
        ? '=COUNTIF(Inscripciones!$V$2:$V;"Sí")'
        : '=COUNTIF(Inscripciones!$T$2:$T;B' + row + ')'
    );
    sheet.getRange('G' + row).setFormula('=OR(E' + row + ';F' + row + '>=D' + row + ')');
  }
}

// Arma la pestaña "Resumen" con FÓRMULAS (no valores fijos) — se recalcula sola
// cada vez que entra una inscripción nueva. Correr esta función una sola vez;
// no hace falta volver a correrla después salvo que se rompa la estructura.
// Nota: esta planilla usa ";" como separador de argumentos en fórmulas (configuración
// regional), por eso todas las fórmulas de abajo usan ";" en vez de ",".
function setupResumenSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Resumen');
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet('Resumen');

  sheet.getRange('A1').setValue('RESUMEN DE INSCRIPCIONES').setFontWeight('bold').setFontSize(14);
  sheet.getRange('A2').setValue('Se actualiza solo — no hace falta volver a correr este script.')
    .setFontStyle('italic').setFontColor('#888888');

  var stats = [
    ['Total inscriptos', '=COUNTA(Inscripciones!B2:B)'],
    ['Total registrado (USD) — incluye lo pendiente de cobro', '=SUM(Inscripciones!X2:X)'],
    ['Total PAGADO / confirmado (USD)', '=SUMIF(Inscripciones!AC2:AC;"Pagado";Inscripciones!X2:X)'],
    ['Total PENDIENTE de cobro (USD)', '=B5-B6'],
    ['Inscriptos que ya pagaron', '=COUNTIF(Inscripciones!AC2:AC;"Pagado")'],
    ['Inscriptos que faltan pagar', '=B4-B8'],
    ['Con práctica especial', '=COUNTIF(Inscripciones!V2:V;"Sí")'],
    ['Sin práctica especial', '=B4-B10'],
    ['Recaudado práctica especial (USD, de lo registrado)', '=SUMPRODUCT((Inscripciones!V2:V="Sí")*(Inscripciones!X2:X-Inscripciones!U2:U))'],
    ['Posibles duplicados a revisar', '=COUNTIF(Inscripciones!Z2:Z;"SÍ - revisar")'],
  ];
  stats.forEach(function (row, i) {
    var r = 4 + i;
    sheet.getRange(r, 1).setValue(row[0]).setFontWeight('bold');
    sheet.getRange(r, 2).setFormula(row[1]);
  });
  [5, 6, 7, 12].forEach(function (row) { sheet.getRange(row, 2).setNumberFormat('$#,##0'); });

  // Nota: QUERY siempre agrega su propia fila de encabezado a partir de los "label",
  // así que esa fila (17) hace de encabezado — no se pone un encabezado manual aparte
  // para las columnas que vienen de QUERY (sí para "Pagaron"/"Faltan pagar", que no).
  // Todo envuelto en IFERROR para que, si todavía no hay datos para ese campo
  // (ej. Género en filas cargadas antes de que existiera esa columna), muestre
  // "Sin datos" en vez de romperse.

  sheet.getRange('A16').setValue('POR CUPO').setFontWeight('bold');
  sheet.getRange('D17').setValue('Pagaron');
  sheet.getRange('E17').setValue('Faltan pagar');
  sheet.getRange('A17').setFormula(
    '=IFERROR(QUERY(Inscripciones!T2:X;"select T, count(T), sum(X) where T is not null group by T label T \'Cupo\', count(T) \'Cantidad\', sum(X) \'Total USD\'";0);"Sin datos")'
  );
  // ARRAYFORMULA + COUNTIFS con dos rangos abiertos (columna entera) es inestable
  // en Sheets y tira #VALUE! — mejor una fórmula por fila (referencia relativa a
  // su propia celda A), copiada para hasta 10 cupos distintos (de sobra).
  for (var i = 18; i <= 27; i++) {
    sheet.getRange('D' + i).setFormula('=IF($A' + i + '="";"";COUNTIFS(Inscripciones!$T$2:$T;$A' + i + ';Inscripciones!$AC$2:$AC;"Pagado"))');
    sheet.getRange('E' + i).setFormula('=IF($A' + i + '="";"";COUNTIFS(Inscripciones!$T$2:$T;$A' + i + ';Inscripciones!$AC$2:$AC;"Pendiente"))');
  }

  sheet.getRange('G16').setValue('POR DOJO').setFontWeight('bold');
  sheet.getRange('G17').setFormula(
    '=IFERROR(QUERY(Inscripciones!L2:L;"select L, count(L) where L is not null group by L order by count(L) desc label L \'Dojo\', count(L) \'Cantidad\'";0);"Sin datos")'
  );

  sheet.getRange('J16').setValue('POR RANGO').setFontWeight('bold');
  sheet.getRange('J17').setFormula(
    '=IFERROR(QUERY(Inscripciones!J2:J;"select J, count(J) where J is not null group by J order by count(J) desc label J \'Rango\', count(J) \'Cantidad\'";0);"Sin datos")'
  );

  sheet.getRange('M16').setValue('POR MÉTODO DE PAGO').setFontWeight('bold');
  sheet.getRange('M17').setFormula(
    '=IFERROR(QUERY(Inscripciones!W2:X;"select W, count(W), sum(X) where W is not null group by W label W \'Método\', count(W) \'Cantidad\', sum(X) \'Total USD\'";0);"Sin datos")'
  );

  sheet.getRange('Q16').setValue('POR PAÍS').setFontWeight('bold');
  sheet.getRange('Q17').setFormula(
    '=IFERROR(QUERY(Inscripciones!AA2:AA;"select AA, count(AA) where AA is not null group by AA order by count(AA) desc label AA \'País\', count(AA) \'Cantidad\'";0);"Sin datos")'
  );

  sheet.getRange('T16').setValue('POR GÉNERO').setFontWeight('bold');
  sheet.getRange('T17').setFormula(
    '=IFERROR(QUERY(Inscripciones!AB2:AB;"select AB, count(AB) where AB is not null group by AB order by count(AB) desc label AB \'Género\', count(AB) \'Cantidad\'";0);"Sin datos")'
  );

  sheet.getRange(17, 1, 1, 21).setFontWeight('bold');
  sheet.autoResizeColumns(1, 21);
}
