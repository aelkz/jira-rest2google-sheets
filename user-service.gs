function openUserDialog() {
  var myself = PropertiesService.getUserProperties().getProperty("myself");
  if (myself == null || myself == '') {
      Browser.msgBox("User login/password incorrect.\\nPlease try again.");
      openDialog();
  }else {
    var htmlForm = HtmlService.createTemplateFromFile('user-gui').evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME);
    htmlForm.setHeight(400);
    htmlForm.setWidth(540);
    SpreadsheetApp.getUi().showModelessDialog(htmlForm, PropertiesService.getUserProperties().getProperty("app_version"));
    //SpreadsheetApp.getUi().showSidebar(htmlForm);
  }
}

function userAction(formObject) {
  return true;
}

function getUserInfo() {
  // https://jira.indra.es/rest/api/2/myself
  // https://json-indent.appspot.com/indent
  // https://jsonformatter.curiousconcept.com/
  // http://us.battle.net/wow/en/forum/topic/13507591084
  // https://developers.google.com/apps-script/guides/services/external
  var data = {};
  
  var query = "myself";
  data = JSON.parse(getDataForAPI(query));  
  PropertiesService.getUserProperties().setProperty("myself", data);
  query = "serverInfo";
  data = JSON.parse(getDataForAPI(query));  
  PropertiesService.getUserProperties().setProperty("server_info", data);  
  //Logger.log(data.emailAddress);
}  

function getMyselfData() {
  var data = {};
  var query = "myself";
  data = JSON.parse(getDataForAPI(query)); 
  PropertiesService.getUserProperties().setProperty("current_user", data.key);    
  return data;
}

function getServerData() {
  var data = {};
  var query = "serverInfo";
  data = JSON.parse(getDataForAPI(query)); 
  return data;
}

function getSessionData() {
  var data = {};
  // parse nao funcionou pois não é um json válido.
  //data = JSON.parse(getUserSession());
  data = JSON.parse(getUserSession());
  return data;
}