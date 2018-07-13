// ********************** //
// general event handlers //
// ********************** //

function onOpen(e) {
  SpreadsheetApp.getUi().createAddonMenu()
      .addItem('Kitöltés', 'showSidebar')
      .addToUi();
}

function onInstall(e) {
  onOpen(e);
}

// **************** //
// server functions //
// **************** //

function showSidebar() {
  var ui = HtmlService.createHtmlOutputFromFile('sidebar')
      .setTitle('Sablon kitöltés')
  SpreadsheetApp.getUi().showSidebar(ui);
}

// ******************************* //
// functions called by client side //
// ******************************* //

var maxTemplates = 5;
function getPreferences() {
  var preferences = {
    oAuthToken: getOAuthToken(),
    templateDir: getUserProperty('templateDir'),
    outputDir: getUserProperty('outputDir'),
    maxTemplates:maxTemplates,
    templates: []
  };
  if (preferences.templateDir) {
    preferences.templateDirName = DriveApp.getFolderById(getUserProperty('templateDir')).getName();
  }
  if (preferences.outputDir) {
    preferences.outputDirName = DriveApp.getFolderById(getUserProperty('outputDir')).getName();
  }
  for (var i = 0; i < maxTemplates; ++i) {
    if (getUserProperty("template" + i)) {
      preferences.templates[i] = {
        id: getUserProperty("template" + i),
        name: DriveApp.getFileById(getUserProperty("template" + i)).getName()
      }
    } else {
      preferences.templates[i] = null;
    }
  }
  return preferences;
}

function savePreferences(preferences) {
  return message("savePreferences not implemented");
}

function fillTemplate(id) {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getActiveSheet()
  var row = sheet.getActiveCell().getRow()
  var dataId = sheet.getRange(row, 1).getValue();
  var columns = 1;
  while (sheet.getRange(1, columns).getValue() != "") {
    columns = columns + 1;
  }
  
  var template = DocumentApp.openById(id);
  
  var newFile = DriveApp.getFileById(id).makeCopy(DriveApp.getFolderById(getTemplateDir()));
  newFile.setName(baseName(template.getName()) + " " + dataId);
  var newId = newFile.getId();
  
  var document = DocumentApp.openById(newId);
  for (var col = 1; col < columns; col = col + 1) {
    document.getHeader().replaceText("<" + sheet.getRange(1, col).getDisplayValue() + ">", sheet.getRange(row, col).getDisplayValue())
    document.getBody().replaceText("<" + sheet.getRange(1, col).getDisplayValue() + ">", sheet.getRange(row, col).getDisplayValue())
    document.getFooter().replaceText("<" + sheet.getRange(1, col).getDisplayValue() + ">", sheet.getRange(row, col).getDisplayValue())
  }
  document.saveAndClose();
  
  return message("Kész: <a href='"+newFile.getUrl()+"'>" + newFile.getName()+"</a>")
}

// ***************** //
// utility functions //
// ***************** //

function getOAuthToken() {
  DriveApp.getRootFolder();
  return ScriptApp.getOAuthToken();
}

function getTemplateDir() {
  var templateDir = getUserProperty("templateDir")
  if (templateDir) {
    return templateDir;
  }
  
  var currentId = SpreadsheetApp.getActive().getId();
  var parents = DriveApp.getFileById(currentId).getParents();
  if (parents.hasNext()) {
    var parent = parents.next();
    return parent.getId();
  }
  return DriveApp.getRootFolder().getId();
}

function baseName(name) {
  return name
      .replace("template", "")
      .replace("Template", "")
      .replace("Sablon", "")
      .replace("sablon", "")
      .trim();
}

var tempLog = ""
function debug(str) {
  if (tempLog.length > 0) tempLog = tempLog + "<br>"
  tempLog = tempLog + str
}

function getUserProperty(key) {
  return null
  try {
    return PropertiesService.getUserProperties().getProperty(key)
  } catch (err) {
    return null
  }
}

function setUserProperty(key, value) {
  try {
    PropertiesService.getUserProperties().setProperty(key, value)
  } catch (err) {
    return
  }
}

function message(msg) {
  return {
    message: msg,
    log: tempLog
  }
}
