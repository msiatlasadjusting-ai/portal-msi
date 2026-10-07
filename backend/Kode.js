// =========================================================
// APPLICATION ROUTING & ENTRY POINT
// Penanganan request HTTP GET dan helper pemuatan komponen view/script
// =========================================================

function doGet() {
  return HtmlService.createTemplateFromFile('views/Main')
    .evaluate()
    .setTitle('Dashboard MSI')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setFaviconUrl('https://raw.githubusercontent.com/Akhfabgss/MSI/main/Logo-MSI.png');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}