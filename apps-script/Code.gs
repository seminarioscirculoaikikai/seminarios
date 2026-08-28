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

function doGet(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var precios = ss.getSheetByName('Precios');
    var inscripciones = ss.getSheetByName('Inscripciones');

    var tiers = [];
    if (precios) {
      var lastRow = precios.getLastRow();
      if (lastRow > 1) {
        precios.getRange(2, 1, lastRow - 1, 4).getValues().forEach(function (r) {
          if (!r[0]) return;
          tiers.push({ id: r[0], nombre: r[1], precioUsd: Number(r[2]), cupos: Number(r[3]) });
        });
      }
    }

    var counts = {};
    if (inscripciones) {
      var lr = inscripciones.getLastRow();
      if (lr > 1) {
        var headers = inscripciones.getRange(1, 1, 1, inscripciones.getLastColumn()).getValues()[0];
        var cupoCol = headers.indexOf('Cupo') + 1;
        inscripciones.getRange(2, cupoCol, lr - 1, 1).getValues().forEach(function (r) {
          if (!r[0]) return;
          counts[r[0]] = (counts[r[0]] || 0) + 1;
        });
      }
    }

    var result = tiers.map(function (t) {
      var registered = counts[t.nombre] || 0;
      return {
        id: t.id,
        priceUsd: t.precioUsd,
        slots: t.cupos,
        registered: registered,
        soldOut: registered >= t.cupos
      };
    });

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
  sheet.getRange('A1:D1').setValues([['id', 'nombre', 'precioUsd', 'cupos']]).setFontWeight('bold');
  sheet.getRange('A2:D5').setValues([
    ['super-early-bird', 'Super Early Bird', 80, 20],
    ['early-bird', 'Early Bird', 90, 20],
    ['regular', 'Regular', 100, 20],
    ['last-minute', 'Last Minute', 110, 20],
  ]);
  sheet.autoResizeColumns(1, 4);
}

function setupResumenSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var src = ss.getSheetByName('Inscripciones');
  var sheet = ss.getSheetByName('Resumen');
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet('Resumen');

  var lastRow = src.getLastRow();
  var lastCol = src.getLastColumn();
  var headers = src.getRange(1, 1, 1, lastCol).getValues()[0];
  var idx = {};
  headers.forEach(function (h, i) { idx[h] = i; });
  function get(r, name) { return idx[name] !== undefined ? r[idx[name]] : ''; }

  var rows = lastRow > 1 ? src.getRange(2, 1, lastRow - 1, lastCol).getValues() : [];

  var totalInscriptos = 0, totalRecaudado = 0;
  var totalPagado = 0, totalPendiente = 0, cantPagado = 0, cantPendiente = 0;
  var conAddon = 0, sinAddon = 0, recaudadoAddon = 0, duplicados = 0;
  var porCupo = {}, porDojo = {}, porRango = {}, porPago = {}, porPais = {}, porGenero = {};

  rows.forEach(function (r) {
    if (!get(r, 'Nombre')) return;
    totalInscriptos++;
    var total = Number(get(r, 'Total (USD)')) || 0;
    var precioCupo = Number(get(r, 'Precio cupo (USD)')) || 0;
    totalRecaudado += total;

    var estado = get(r, 'Estado de pago') || 'Pendiente';
    if (estado === 'Pagado') { totalPagado += total; cantPagado++; }
    else { totalPendiente += total; cantPendiente++; }

    if (get(r, 'Práctica especial') === 'Sí') { conAddon++; recaudadoAddon += (total - precioCupo); }
    else { sinAddon++; }

    if (get(r, 'Posible duplicado') === 'SÍ - revisar') duplicados++;

    var cupo = get(r, 'Cupo') || '(sin cupo)';
    if (!porCupo[cupo]) porCupo[cupo] = { cantidad: 0, total: 0, pagado: 0, pendiente: 0 };
    porCupo[cupo].cantidad++;
    porCupo[cupo].total += total;
    if (estado === 'Pagado') porCupo[cupo].pagado++; else porCupo[cupo].pendiente++;

    var dojo = get(r, 'Dojo') || '(sin dojo)';
    porDojo[dojo] = (porDojo[dojo] || 0) + 1;

    var rango = get(r, 'Rango') || '(sin rango)';
    porRango[rango] = (porRango[rango] || 0) + 1;

    var pago = get(r, 'Método de pago') || '(sin método)';
    if (!porPago[pago]) porPago[pago] = { cantidad: 0, total: 0 };
    porPago[pago].cantidad++;
    porPago[pago].total += total;

    var pais = get(r, 'País') || '(sin especificar)';
    porPais[pais] = (porPais[pais] || 0) + 1;

    var genero = get(r, 'Género') || 'No especificado';
    porGenero[genero] = (porGenero[genero] || 0) + 1;
  });

  sheet.getRange('A1').setValue('RESUMEN DE INSCRIPCIONES').setFontWeight('bold').setFontSize(14);
  sheet.getRange('A2').setValue('Generado: ' + new Date()).setFontStyle('italic').setFontColor('#888888');

  var stats = [
    ['Total inscriptos', totalInscriptos],
    ['Total registrado (USD) — incluye lo pendiente de cobro', totalRecaudado],
    ['Total PAGADO / confirmado (USD)', totalPagado],
    ['Total PENDIENTE de cobro (USD)', totalPendiente],
    ['Inscriptos que ya pagaron', cantPagado],
    ['Inscriptos que faltan pagar', cantPendiente],
    ['Con práctica especial', conAddon],
    ['Sin práctica especial', sinAddon],
    ['Recaudado práctica especial (USD, de lo registrado)', recaudadoAddon],
    ['Posibles duplicados a revisar', duplicados],
  ];
  sheet.getRange(4, 1, stats.length, 2).setValues(stats);
  sheet.getRange(4, 1, stats.length, 1).setFontWeight('bold');
  [5, 6, 7, 12].forEach(function (row) { sheet.getRange(row, 2).setNumberFormat('$#,##0'); });

  function toSortedTable(obj, cols) {
    var keys = Object.keys(obj).sort(function (a, b) {
      var av = cols ? obj[a].cantidad : obj[a];
      var bv = cols ? obj[b].cantidad : obj[b];
      return bv - av;
    });
    return keys.map(function (k) {
      if (!cols) return [k, obj[k]];
      return [k, obj[k].cantidad, obj[k].total, obj[k].pagado || 0, obj[k].pendiente || 0].slice(0, cols);
    });
  }

  sheet.getRange('A16').setValue('POR CUPO').setFontWeight('bold');
  sheet.getRange('A17:E17').setValues([['Cupo', 'Cantidad', 'Total USD', 'Pagaron', 'Faltan pagar']]).setFontWeight('bold');
  var cupoTable = toSortedTable(porCupo, 5);
  if (cupoTable.length) sheet.getRange(18, 1, cupoTable.length, 5).setValues(cupoTable);

  sheet.getRange('G16').setValue('POR DOJO').setFontWeight('bold');
  sheet.getRange('G17:H17').setValues([['Dojo', 'Cantidad']]).setFontWeight('bold');
  var dojoTable = toSortedTable(porDojo);
  if (dojoTable.length) sheet.getRange(18, 7, dojoTable.length, 2).setValues(dojoTable);

  sheet.getRange('J16').setValue('POR RANGO').setFontWeight('bold');
  sheet.getRange('J17:K17').setValues([['Rango', 'Cantidad']]).setFontWeight('bold');
  var rangoTable = toSortedTable(porRango);
  if (rangoTable.length) sheet.getRange(18, 10, rangoTable.length, 2).setValues(rangoTable);

  sheet.getRange('M16').setValue('POR MÉTODO DE PAGO').setFontWeight('bold');
  sheet.getRange('M17:O17').setValues([['Método', 'Cantidad', 'Total USD']]).setFontWeight('bold');
  var pagoTable = toSortedTable(porPago, 3);
  if (pagoTable.length) sheet.getRange(18, 13, pagoTable.length, 3).setValues(pagoTable);

  sheet.getRange('Q16').setValue('POR PAÍS').setFontWeight('bold');
  sheet.getRange('Q17:R17').setValues([['País', 'Cantidad']]).setFontWeight('bold');
  var paisTable = toSortedTable(porPais);
  if (paisTable.length) sheet.getRange(18, 17, paisTable.length, 2).setValues(paisTable);

  sheet.getRange('T16').setValue('POR GÉNERO').setFontWeight('bold');
  sheet.getRange('T17:U17').setValues([['Género', 'Cantidad']]).setFontWeight('bold');
  var generoTable = toSortedTable(porGenero);
  if (generoTable.length) sheet.getRange(18, 20, generoTable.length, 2).setValues(generoTable);

  sheet.autoResizeColumns(1, 21);
}
