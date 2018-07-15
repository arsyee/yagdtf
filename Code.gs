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
    maxTemplates:maxTemplates,
    templates: []
  };
  
  try {
    var userProperties = PropertiesService.getUserProperties().getProperties()
    for (var key in userProperties) {
      debug(key + " : " + userProperties[key])
    }
    if (userProperties.templateDir) {
      preferences.templateDir = {
        id: userProperties.templateDir,
        name: DriveApp.getFolderById(userProperties.templateDir).getName()
      }
    }
    if (userProperties.outputDir) {
      preferences.outputDir = {
        id: userProperties.outputDir,
        name: DriveApp.getFolderById(userProperties.outputDir).getName()
      }
    }
    for (var i = 0; i < maxTemplates; ++i) {
      if (userProperties["template" + i]) {
        preferences.templates[i] = {
          id: userProperties["template" + i],
          name: DriveApp.getFileById(userProperties["template" + i]).getName()
        }
      } else {
        preferences.templates[i] = null;
      }
    }
  } catch (err) {
    debug("user properties cannot be read: " + str(err))
  }
  preferences.message = tempLog
  return preferences;
}

function savePreferences(preferences) {
  var userProperties = {}
  if (preferences.templateDir) {
    userProperties.templateDir = preferences.templateDir.id
  }
  if (preferences.outputDir) {
    userProperties.outputDir = preferences.outputDir.id
  }
  var j = 0;
  for (var i = 0; i < maxTemplates; ++i) {
    if (preferences.templates[i]) {
      userProperties["template" + j] = preferences.templates[i].id
      ++j
    }
  }
  debug(JSON.stringify(userProperties))
  try {
    PropertiesService.getUserProperties().deleteAllProperties()
    PropertiesService.getUserProperties().setProperties(userProperties)
    return message("Beállítások elmentve.")
  } catch (err) {
    return message("Hiba a mentés során.")
  }
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
  
  var newFile = DriveApp.getFileById(id).makeCopy(DriveApp.getFolderById(getOutputDir()))
  newFile.setName(baseName(template.getName()) + " " + dataId);
  var newId = newFile.getId();
  
  var document = DocumentApp.openById(newId);
  for (var col = 1; col < columns; col = col + 1) {
    if (document.getHeader()) {
      document.getHeader().replaceText("<" + sheet.getRange(1, col).getDisplayValue() + ">", sheet.getRange(row, col).getDisplayValue())
    }
    if (document.getBody()) {
      document.getBody().replaceText("<" + sheet.getRange(1, col).getDisplayValue() + ">", sheet.getRange(row, col).getDisplayValue())
    }
    if (document.getFooter()) {
        document.getFooter().replaceText("<" + sheet.getRange(1, col).getDisplayValue() + ">", sheet.getRange(row, col).getDisplayValue())
    }
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

function getOutputDir() {
  var outputDir = PropertiesService.getUserProperties().getProperty("outputDir")
  if (outputDir) {
    return outputDir;
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

function message(msg) {
  return {
    message: msg,
    log: tempLog
  }
}
