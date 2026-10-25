function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Auditor de Facturas');
}
