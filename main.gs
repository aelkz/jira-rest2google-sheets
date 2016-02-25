// ---------------------------------------------------------------------------------------------------------------------------------------------------
//The MIT License (MIT)

//Copyright (c) 2016 RAPHAEL ALEX SILVA ABREU 
//https://github.com/aelkz

//Permission is hereby granted, free of charge, to any person obtaining a copy
//of this software and associated documentation files (the "Software"), to deal
//in the Software without restriction, including without limitation the rights
//to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
//copies of the Software, and to permit persons to whom the Software is
//furnished to do so, subject to the following conditions:

//The above copyright notice and this permission notice shall be included in all
//copies or substantial portions of the Software.

//THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
//IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
//FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
//AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
//LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
//OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
//SOFTWARE.      

var C_MAX_RESULTS = 1000;

/*
* This function loads added menu in the spreadsheet
* Intialize project url and project API key while openning or refreshing spreadsheet
*/
function onOpen(){
  ScriptProperties.setProperty("APIkey", "");
  ScriptProperties.setProperty("projectURL", "");
  PropertiesService.getUserProperties().setProperty("myself", '');
  PropertiesService.getUserProperties().setProperty("app_version", "INDRA MIND JIRA backlog - v0.4b25022016");

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var spread=SpreadsheetApp.getActiveSpreadsheet();

  var menuEntries = [
    //var menuEntries = [{name: "Get Data",functionName: "selectUI"}];
    {name: "Configure Jira User Login", functionName: "jiraConfigure"},
    {name: "Configure Jira Search Params", functionName: "jiraConfigureParams"},
    {name: "Refresh Backlog", functionName: "jiraPullManual"},
    {name: "Refresh Backlog Assigned to me", functionName: "jiraPullManualAssigned"},
    {name: "Refresh Timesheet", functionName: "jiraPullManual"},
    {name: "Jira User Details", functionName: "openUserDialog"},
    {name: "Schedule 1 Hour Automatic Refresh", functionName: "scheduleRefresh"},
    {name: "Stop Automatic Refresh", functionName: "removeTriggers"}]; 
  ss.addMenu("[JIRA]", menuEntries);
                     
  menuEntries = [ 
    {name: "Format cells", functionName: "formatCells"}];
  ss.addMenu("[CONFIGURATION]", menuEntries);

  menuEntries = [ 
    {name: "Contents", functionName: "helpContents"},
    {name: "About", functionName: "about"}];
  ss.addMenu("[HELP]", menuEntries);
}

function formatCells() {
  setBold("A2:A1000");
  setUnBold("B2:V1000");
  formatFont("A2:V1000");
  SpreadsheetApp.flush();
}

function onEdit() {
  setBold("A2:A1000");
  setUnBold("B2:V1000");
  formatFont("A2:V1000");
  //setCreationDateSingleRow();
  SpreadsheetApp.flush();
};

function jiraConfigure(formObject) {
  var myself = PropertiesService.getUserProperties().getProperty("myself");
  if (myself == null || myself == '' || formObject == null || formObject == undefined) {
      Browser.msgBox("User login/password incorrect.\\nPlease try again.");
      openDialog();
  }else {
    var prefix = "MEBAMG";
    PropertiesService.getUserProperties().setProperty("prefix", prefix.toUpperCase());
    var host = "jira.indra.es";
    PropertiesService.getUserProperties().setProperty("host", host);
    var userAndPassword = formObject.username+":"+formObject.password;
    var x = Utilities.base64Encode(userAndPassword);
    PropertiesService.getUserProperties().setProperty("digest", "Basic " + x);
    var issueTypes = "%22Functional%20Support%22";
    PropertiesService.getUserProperties().setProperty("issueTypes", issueTypes);
    Browser.msgBox("Jira configuration saved successfully:\\n");
  }
}  

function removeTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  
  Browser.msgBox("Spreadsheet will no longer refresh automatically.");
}  

function scheduleRefresh() {
  var myself = PropertiesService.getUserProperties().getProperty("myself");
  if (myself == null || myself == '') {
      Browser.msgBox("User login/password incorrect.\\nPlease try again.");
      openDialog();
  }else {
    var triggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < triggers.length; i++) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
    
    ScriptApp.newTrigger("jiraPull").timeBased().everyHours(1).create();
    Browser.msgBox("Spreadsheet will refresh automatically every 1 hour.");
  }
}  

function jiraPullManual() {
  var myself = PropertiesService.getUserProperties().getProperty("myself");
  if (myself == null || myself == '') {
      Browser.msgBox("User login/password incorrect.\\nPlease try again.");
      openDialog();
  }else {
    jiraPull();
  }
}  

function jiraPullManualAssigned() {
  var myself = PropertiesService.getUserProperties().getProperty("myself");
  if (myself == null || myself == '') {
      Browser.msgBox("User login/password incorrect.\\nPlease try again.");
      openDialog();
  }else {
    jiraPullAssigned();
  }
}  

function getFields() {
  return JSON.parse(getDataForAPI("field"));
}  

function getStories(type) {
  var allData = {issues:[]};
  var data = {startAt:0,maxResults:0,total:1};
  var startAt = 0;
 
  if (type == 1) {
    while (data.startAt + data.maxResults < data.total) {
      Logger.log("Making request for %s entries", C_MAX_RESULTS);
      
      var query = "search?jql=";
      // parameters
      query = query + "project%20%3D%20" + PropertiesService.getUserProperties().getProperty("prefix");
      query = query + "%20and%20status%20!%3D%20resolved%20";
      query = query + "and%20type%20in%20("+ PropertiesService.getUserProperties().getProperty("issueTypes") + ")";
      query = query + "%20and%20assignee%20%3D%20goliveiraa%20";
      // order by and pagination parameters
      query = query + "%20order%20by%20rank%20";
      query = query + "&maxResults=" + C_MAX_RESULTS;
      query = query + "&startAt=" + startAt;
      
      Logger.log(query);
      data =  JSON.parse(getDataForAPI(query));  
      allData.issues = allData.issues.concat(data.issues);
      startAt = data.startAt + data.maxResults;
    }  
  }else if (type == 2) {
    getMyselfData();
    while (data.startAt + data.maxResults < data.total) {
      Logger.log("Making request for %s entries", C_MAX_RESULTS);
      
      var query = "search?jql=";
      // parameters
      query = query + "project%20%3D%20" + PropertiesService.getUserProperties().getProperty("prefix"); // MEBAMG
      query = query + "%20and%20status%20IN%20%28%22In%20Progress%22%2C%22Open%22%2C%22Reopened%22%29"; // In Progress, Open, Reopened
      query = query + "%20and%20type%20in%20("+ PropertiesService.getUserProperties().getProperty("issueTypes") + ")";
      query = query + "%20and%20assignee%20%3D%20"+PropertiesService.getUserProperties().getProperty("current_user")+"%20";
      // order by and pagination parameters
      query = query + "%20order%20by%20rank%20";
      query = query + "&maxResults=" + C_MAX_RESULTS;
      query = query + "&startAt=" + startAt;
      
      Logger.log(query);
      data =  JSON.parse(getDataForAPI(query));  
      allData.issues = allData.issues.concat(data.issues);
      startAt = data.startAt + data.maxResults;
    }  
  }

  return allData;
}  

function getDataForAPI(path) {
   var url = "https://" + PropertiesService.getUserProperties().getProperty("host") + "/rest/api/2/" + path;
   var digestfull = PropertiesService.getUserProperties().getProperty("digest");
   var headers = { "Accept":"application/json", "Content-Type":"application/json", "method": "GET", "headers": {"Authorization": digestfull}, "muteHttpExceptions": true};
   var resp = UrlFetchApp.fetch(url,headers);
   if (resp.getResponseCode() != 200) {
     Browser.msgBox("Error retrieving data for url:\n" + url + "\n\n\n" + resp.getContentText() + "\nResponse code status:" + resp.getResponseCode());
      return "";
   }else {
      return resp.getContentText();
   }  
}  

function getUserSession() {
   // https://jira.indra.es/rest/auth/1/session
   var url = "https://" + PropertiesService.getUserProperties().getProperty("host") + "/rest/auth/1/session";
   var digestfull = PropertiesService.getUserProperties().getProperty("digest");
   var headers = { "Accept":"application/json", "Content-Type":"application/json", "method": "GET", "headers": {"Authorization": digestfull}, "muteHttpExceptions": true};
   var resp = UrlFetchApp.fetch(url,headers);
   if (resp.getResponseCode() != 200) {
     return "";
   }else {
     return resp.getContentText();
   }  
}  

function jiraPull() {
  var spreadsheet_name = "BACKLOG";
  var allFields = getAllFields();
  var data = getStories(1);
  
  if (allFields === "" || data === "") {
    Browser.msgBox("Error pulling data from Jira - aborting now.");
    return;
  }  
  
  // GENERAL BACKLOG
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(spreadsheet_name);
  var headings = ss.getRange(1, 1, 1, ss.getLastColumn()).getValues()[0];
  var y = new Array();
  var last = ss.getLastRow();

  for (i=0;i<data.issues.length;i++) {
    var d=data.issues[i];
    y.push(getStory(d,headings,allFields));
  }  

  if (last >= 2) {
    ss.getRange(2, 1, ss.getLastRow()-1,ss.getLastColumn()).clearContent();  
  }  
  
  if (y.length > 0) {
    ss.getRange(2, 1, data.issues.length,y[0].length).setValues(y);
  }

  // seconds to minutes to hours
  convertToMinutes(null,'M',spreadsheet_name);
  convertToMinutes(null,'O',spreadsheet_name);
  convertToMinutes(null,'P',spreadsheet_name);
  // cleanup
  cleanCell("I",spreadsheet_name);
  cleanCell("L",spreadsheet_name);
  cleanCell("R",spreadsheet_name);
  // date format
  cleanCell("B",spreadsheet_name);
  cleanCell("C",spreadsheet_name);
  cleanCell("D",spreadsheet_name);
  cleanCell("E",spreadsheet_name);
  cleanCell("F",spreadsheet_name);
}

function jiraPullAssigned() {
  var spreadsheet_name = "ASSIGNED TO ME";
  var allFields = getAllFields();
  var data = getStories(2);
  
  if (allFields === "" || data === "") {
    Browser.msgBox("Error pulling data from Jira - aborting now.");
    return;
  }  
  
  // BACKLOG ASSIGNED TO ME
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(spreadsheet_name);
  var headings = ss.getRange(1, 1, 1, ss.getLastColumn()).getValues()[0];
  var y = new Array();
  var last = ss.getLastRow();

  for (i=0;i<data.issues.length;i++) {
    var d=data.issues[i];
    y.push(getStory(d,headings,allFields));
  }  

  if (last >= 2) {
    ss.getRange(2, 1, ss.getLastRow()-1,ss.getLastColumn()).clearContent();  
  }  
  
  if (y.length > 0) {
    ss.getRange(2, 1, data.issues.length,y[0].length).setValues(y);
  }

  // seconds to minutes to hours
  convertToMinutes(null,'M',spreadsheet_name);
  convertToMinutes(null,'O',spreadsheet_name);
  convertToMinutes(null,'P',spreadsheet_name);
  // cleanup
  cleanCell("I",spreadsheet_name);
  cleanCell("L",spreadsheet_name);
  cleanCell("R",spreadsheet_name);
  // date format
  cleanCell("B",spreadsheet_name);
  cleanCell("C",spreadsheet_name);
  cleanCell("D",spreadsheet_name);
  cleanCell("E",spreadsheet_name);
  cleanCell("F",spreadsheet_name);
}

function getAllFields() {
  var theFields = getFields();
  var allFields = new Object();
  allFields.ids = new Array();
  allFields.names = new Array();
  
  for (var i = 0; i < theFields.length; i++) {
      allFields.ids.push(theFields[i].id);
      allFields.names.push(theFields[i].name.toLowerCase());
  }  
  
  return allFields;
}  

function getStory(data,headings,fields) {
  var story = [];
  for (var i = 0;i < headings.length;i++) {
    if (headings[i] != "" || headings[i] != null || headings[i] != undefined) {
      story.push(getDataForHeading(data,headings[i].toLowerCase(),fields));
    }  
  }        
  
  return story;
}  

function getDataForHeading(data,heading,fields) {
    if (heading == "Planned Start Date") {
      heading = "customfield_11801";
    }else if (heading == "Planned End Date") {
      heading = "customfield_10313";
    }else if (heading == "Due Delivery Date") {
      heading = "customfield_16600";
    }else if (heading == "Real Start Date") {
      heading = "customfield_12005";
    }else if (heading == "Attended") {
      heading = "customfield_17108";
    }
  
    if (data.hasOwnProperty(heading)) {
      return data[heading];
    }  
    else if (data.fields.hasOwnProperty(heading)) {
      return data.fields[heading];
    }  

    var fieldName = getFieldName(heading,fields);

    if (fieldName !== "") {
      if (data.hasOwnProperty(fieldName)) {
        return data[fieldName];
      }  
      else if (data.fields.hasOwnProperty(fieldName)) {
        return data.fields[fieldName];
      }  
    }

    var splitName = heading.split(" ");

    if (splitName.length == 2) {
      if (data.fields.hasOwnProperty(splitName[0]) ) {
        if (data.fields[splitName[0]] && data.fields[splitName[0]].hasOwnProperty(splitName[1])) {
          return data.fields[splitName[0]][splitName[1]];
        }
        return "";
      }  
    }  
    //return "Could not find value for " + heading;
    return "";
}  