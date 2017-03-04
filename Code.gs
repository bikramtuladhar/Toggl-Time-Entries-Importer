/**
 * @OnlyCurrentDoc
 */
 
/**
 * @param {object} e 
 */
function onOpen(e) {
  DocumentApp.getUi().createAddonMenu()
      .addItem('Start', 'showSidebar')
      .addToUi();
}
/**
 * @param {object} e 
 */

function onInstall(e) {
  onOpen(e);
}

function showSidebar() {
  var ui = HtmlService.createHtmlOutputFromFile('SideBar')
      .setTitle('Toggl Importer');
  DocumentApp.getUi().showSidebar(ui);
}

function getBasicInfo() {
  updateApiUrl();
  var TOKEN = PropertiesService.getScriptProperties().getProperty('token');
  var url = PropertiesService.getScriptProperties().getProperty('profileUrl');
  var headers = {"Authorization" : "Basic " + Utilities.base64Encode(TOKEN + ':' + 'api_token')};
  var params = {
    "method":"GET",
    "headers":headers
  };
  var response = UrlFetchApp.fetch(url, params);
  var togglData = response.getContentText();
  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('basicInfo',togglData );
  var data=scriptProperties.getProperty('basicInfo' );
  createTable();
  return JSON.parse(data).data;
}

function processForm(formObject) {
  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('token', formObject.token);
  return "sucessfully update token";
}

function getTodaysTask(){
  var TOKEN = PropertiesService.getScriptProperties().getProperty('token');  
  var dateObj = new Date();
  
  Date.prototype.getMonthFormatted = function() {
	var month = this.getMonth() + 1;
	return month < 10 ? '0' + month : month;
  }
  
  Date.prototype.getDateFormatted = function() {
    var date = this.getDate();
    return date < 10 ? '0' + date : date;
  }

  var month = dateObj.getMonthFormatted(); 
  var day = dateObj.getDateFormatted();
  var year = dateObj.getUTCFullYear();
  
  var startTime =year+"-"+month+"-"+day+"T03:15:00Z";
  var endTime =year+"-"+month+"-"+day+"T11:45:00Z";
  var url = "https://www.toggl.com/api/v8/time_entries?start_date="+startTime+"&end_date="+endTime;
  var headers = {"Authorization" : "Basic " + Utilities.base64Encode(TOKEN + ':' + 'api_token')};
  var params = {
    "method":"GET",
    "headers":headers
  };
  var response = UrlFetchApp.fetch(url, params);
  var timeEntries = response.getContentText();
  return timeEntries;
}

function createTable(){
  
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var table = body.appendTable();
  var tr = table.appendTableRow();
  var cell= tr.appendTableCell('Toggl Time Entries ('+ currrentDate() +') \n');

  JSON.parse(getTodaysTask()).forEach(function(obj){ 
   var data=obj.description+" ("+secondsToHms(obj.duration)+")";
    cell.appendListItem(data).setGlyphType(DocumentApp.GlyphType.BULLET);
  });

  doc.saveAndClose();
}

function currrentDate()
{
var dateObj = new Date();
var month = dateObj.getUTCMonth() + 1; 
var day = dateObj.getUTCDate();
var year = dateObj.getUTCFullYear();

var newdate = year + "/" + month + "/" + day;
  return newdate;
}

function getToken()
{
  var scriptProperties = PropertiesService.getScriptProperties();
  var token =scriptProperties.getProperty('token');
  return token;
}

function updateApiUrl(){
  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('profileUrl','https://www.toggl.com/api/v8/me');
}


function getFullname(){
   var scriptProperties = PropertiesService.getScriptProperties();
  var basicDetail =scriptProperties.getProperty('basicInfo');
  var fullName = JSON.parse(basicDetaul).data.fullname;
  return fullName;
}

function secondsToHms(d) {
    d = Number(d);
    var h = Math.floor(d / 3600);
    var m = Math.floor(d % 3600 / 60);
    var s = Math.floor(d % 3600 % 60);

    var hDisplay = h > 0 ? h + (h == 1 ? " hour " : " hours ") : "";
    var mDisplay = m > 0 ? m + (m == 1 ? " minute " : " minutes ") : "";
    return hDisplay + mDisplay; 

}
