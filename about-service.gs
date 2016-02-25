function about() {
    var htmlForm = HtmlService.createTemplateFromFile('about-gui').evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME);
    htmlForm.setHeight(400);
    htmlForm.setWidth(540);
    SpreadsheetApp.getUi().showModelessDialog(htmlForm, PropertiesService.getUserProperties().getProperty("app_version"));
}
