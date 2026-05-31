function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = JSON.parse(e.postData.contents);

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Timestamp", "Nombre", "Apellido", "Asistencia", "Dias", "Transporte", "Personas", "Menu"]);
  }

  sheet.appendRow([
    new Date(),
    data.nombre || "",
    data.apellido || "",
    data.asistencia || "",
    data.dias || "",
    data.transporte || "",
    data.personas || "",
    data.menu || ""
  ]);

  return ContentService
    .createTextOutput(JSON.stringify({ok: true}))
    .setMimeType(ContentService.MimeType.JSON);
}
