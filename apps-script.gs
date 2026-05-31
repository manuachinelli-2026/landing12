function doGet(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Timestamp", "Nombre", "Apellido", "Asistencia", "Dias", "Transporte", "Personas", "Menu"]);
  }

  sheet.appendRow([
    new Date(),
    e.parameter.nombre     || "",
    e.parameter.apellido   || "",
    e.parameter.asistencia || "",
    e.parameter.dias       || "",
    e.parameter.transporte || "",
    e.parameter.personas   || "",
    e.parameter.menu       || ""
  ]);

  return ContentService
    .createTextOutput(JSON.stringify({ok: true}))
    .setMimeType(ContentService.MimeType.JSON);
}
