/**
 * CONTROL DE VISITAS — Sincronización con Google Sheets
 * -------------------------------------------------------
 * Pega este código en Extensiones > Apps Script de tu Google Sheet.
 * Crea automáticamente dos pestañas: "Clientes" y "Visitas".
 *
 * doPost = recibe los datos desde la app y sobrescribe el Sheet (SUBIR)
 * doGet  = devuelve los datos del Sheet en formato JSON (DESCARGAR)
 */

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var clientsSheet = getOrCreateSheet_(ss, 'Clientes');
    var visitsSheet = getOrCreateSheet_(ss, 'Visitas');

    clientsSheet.clearContents();
    clientsSheet.appendRow(['id', 'nombre', 'ciudad', 'provincia', 'direccion', 'contacto', 'telefono', 'email', 'frecuencia', 'categoria', 'estado']);

    visitsSheet.clearContents();
    visitsSheet.appendRow(['clienteId', 'clienteNombre', 'fecha', 'tipo', 'observaciones']);

    (data.clients || []).forEach(function (c) {
      clientsSheet.appendRow([c.id, c.name, c.city, c.province, c.address, c.contact, c.phone, c.email, c.frequency, c.category, c.status]);
      (c.visits || []).forEach(function (v) {
        visitsSheet.appendRow([c.id, c.name, v.date, v.type, v.notes || '']);
      });
    });

    return ContentService.createTextOutput(JSON.stringify({ ok: true, clientes: (data.clients || []).length }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var clientsSheet = ss.getSheetByName('Clientes');
    if (!clientsSheet) {
      return ContentService.createTextOutput(JSON.stringify({ clients: [] }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    var visitsSheet = ss.getSheetByName('Visitas');
    var clientsData = clientsSheet.getDataRange().getValues();
    var visitsData = visitsSheet ? visitsSheet.getDataRange().getValues() : [];

    var map = {};
    for (var i = 1; i < clientsData.length; i++) {
      var r = clientsData[i];
      if (!r[0]) continue;
      map[r[0]] = {
        id: String(r[0]), name: r[1], city: r[2], province: r[3], address: r[4],
        contact: r[5], phone: r[6], email: r[7], frequency: Number(r[8]) || 30,
        category: r[9], status: r[10], visits: []
      };
    }
    for (var j = 1; j < visitsData.length; j++) {
      var v = visitsData[j];
      if (!v[0] || !map[v[0]]) continue;
      var d = v[2] instanceof Date ? v[2] : new Date(v[2]);
      map[v[0]].visits.push({ id: Utilities.getUuid(), date: d.toISOString(), type: v[3], notes: v[4] || '' });
    }

    return ContentService.createTextOutput(JSON.stringify({ clients: Object.keys(map).map(function (k) { return map[k]; }) }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ clients: [], error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getOrCreateSheet_(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}
