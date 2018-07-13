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
      .setTitle('Sablon kitöltés');
  SpreadsheetApp.getUi().showSidebar(ui);
}

function showPicker() {
  var html = HtmlService.createHtmlOutputFromFile('picker.html')
      .setWidth(600)
      .setHeight(425)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showModalDialog(html, 'Select a document');
}

// ******************************* //
// functions called by client side //
// ******************************* //

// picker interface

function getPickerConfiguration() {
  var pickerConfiguration = {
    oAuthToken: getOAuthToken(),
    lastPicker: PropertiesService.getUserProperties().getProperty("lastPicker"),
    startDir: DriveApp.getRootFolder().getId(),
  }
  if (pickerConfiguration.lastPicker == "template") {
    pickerConfiguration.startDir = PropertiesService.getUserProperties().getProperty("templateDir");
  }
  return pickerConfiguration;
}

function selectItem(id) { // callback for showPicker
  switch (PropertiesService.getUserProperties().getProperty("lastPicker")) {
      case 'templateDir':
          PropertiesService.getUserProperties().setProperty("templateDir", id)
          break;
      case 'outputDir':
          PropertiesService.getUserProperties().setProperty("outputDir", id)
          break;
      default:
          addTemplate(id)
          break;
  }
  return PropertiesService.getUserProperties().getProperty("lastPicker") + " set to " + id;
}

// sidebar interface

var maxTemplates = 5;
function getPreferences() {
  var userProperties = PropertiesService.getUserProperties();
  var preferences = {
    templateDir: userProperties.getProperty('templateDir'),
    outputDir: userProperties.getProperty('outputDir'),
    templates: []
  };
  if (preferences.templateDir) {
    preferences.templateDirName = DriveApp.getFolderById(userProperties.getProperty('templateDir')).getName();
  }
  if (preferences.outputDir) {
    preferences.outputDirName = DriveApp.getFolderById(userProperties.getProperty('outputDir')).getName();
  }
  for (var i = 0; i < maxTemplates; ++i) {
    if (userProperties.getProperty("template" + i)) {
      preferences.templates[i] = {
        id: userProperties.getProperty("template" + i),
        name: DriveApp.getFileById(userProperties.getProperty("template" + i)).getName()
      }
    } else {
      preferences.templates[i] = null;
    }
  }
  return preferences;
}

function selectTemplateDir() {
  PropertiesService.getUserProperties().setProperty("lastPicker", "templateDir")
  showPicker()
  return message("Válassz egy könyvtárat!")
}

function selectOutputDir() {
  PropertiesService.getUserProperties().setProperty("lastPicker", "outputDir")
  showPicker()
  return message("Válassz egy könyvtárat!")
}

function selectTemplate() {
  PropertiesService.getUserProperties().setProperty("lastPicker", "template")
  showPicker()
  return message("Válassz egy dokumentumot!")
}

function removeTemplate(id) {
  var userProperties = PropertiesService.getUserProperties();
  for (var i = 0; i < maxTemplates; ++i) {
    if (userProperties.getProperty("template" + i)) {
      if (userProperties.getProperty("template" + i) == id) {
        userProperties.deleteProperty("template" + i)
        debug("Template " + i + " was deleted.")
      }
    }
  }
  return message("Törlés végrehajtva.")
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
    document.getHeader().replaceText("<" + sheet.getRange(1, col).getValue() + ">", sheet.getRange(row, col).getValue())
    document.getBody().replaceText("<" + sheet.getRange(1, col).getValue() + ">", sheet.getRange(row, col).getValue())
    document.getFooter().replaceText("<" + sheet.getRange(1, col).getValue() + ">", sheet.getRange(row, col).getValue())
  }
  document.saveAndClose();
  
  return message("Kész: <a href='"+newFile.getUrl()+"'>" + newFile.getName()+"</a>")
}

// ***************** //
// utility functions //
// ***************** //

function addTemplate(id) {
  var userProperties = PropertiesService.getUserProperties();
  for (var i = 0; i < maxTemplates; ++i) {
    if (!userProperties.getProperty("template" + i)) {
      userProperties.setProperty("template" + i, id)
      return;
    }
  }
}

function getOAuthToken() {
  DriveApp.getRootFolder();
  return ScriptApp.getOAuthToken();
}

function getTemplateDir() {
  var templateDir = PropertiesService.getUserProperties().getProperty("templateDir")
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

function message(msg) {
  return {
    message: msg,
    log: tempLog
  }
}
