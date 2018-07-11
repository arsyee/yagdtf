function onOpen(e) {
  SpreadsheetApp.getUi().createAddonMenu()
      .addItem('Kitöltés', 'showSidebar')
      .addToUi();
}

function onInstall(e) {
  onOpen(e);
}

function showSidebar() {
  Logger.log("showSidebar called")
  var ui = HtmlService.createHtmlOutputFromFile('sidebar')
      .setTitle('Sablon kitöltés');
  SpreadsheetApp.getUi().showSidebar(ui);
}

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

function selectTemplate() {
  PropertiesService.getUserProperties().setProperty("lastPicker", "template")
  showPicker()
  return "Code::selectTemplate"
}

function selectTemplateDir() {
  PropertiesService.getUserProperties().setProperty("lastPicker", "templateDir")
  showPicker()
  return "Code::selectTemplateDir"
}

function selectOutputDir() {
  PropertiesService.getUserProperties().setProperty("lastPicker", "outputDir")
  showPicker()
  return "Code::selectOutputDir"
}

function showPicker() {
  var html = HtmlService.createHtmlOutputFromFile('picker.html')
      .setWidth(600)
      .setHeight(425)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showModalDialog(html, 'Select a document');
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

function addTemplate(id) {
  var userProperties = PropertiesService.getUserProperties();
  for (var i = 0; i < maxTemplates; ++i) {
    if (!userProperties.getProperty("template" + i)) {
      userProperties.setProperty("template" + i, id)
      return;
    }
  }
}

function removeTemplate(id) {
  var userProperties = PropertiesService.getUserProperties();
  for (var i = 0; i < maxTemplates; ++i) {
    if (userProperties.getProperty("template" + i)) {
      if (userProperties.getProperty("template" + i) == id) {
        userProperties.deleteProperty("template" + i)
      }
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

function fillTemplate(id) {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getActiveSheet()
  var row = sheet.getActiveCell().getRow()
  var dataId = sheet.getRange(row, 1).getValue();
  var columns = 1;
  while (sheet.getRange(1, columns).getValue() != "") {
    if (columns == 10) {
      return "error";
    }
    columns = columns + 1;
  }
  
  var template = DocumentApp.openById(id);
  
  var newFile = DriveApp.getFileById(id).makeCopy(DriveApp.getFolderById(getTemplateDir()));
  newFile.setName(baseName(template.getName()) + " " + dataId + " " + (new Date().toISOString()));
  var newId = newFile.getId();
  
  var document = DocumentApp.openById(newId);
  var body = document.getBody();
  for (var col = 1; col < columns; col = col + 1) {
    body.replaceText("<" + sheet.getRange(1, col).getValue() + ">", sheet.getRange(row, col).getValue())
  }
  body.replaceText("MAI_DATUM", new Date().toISOString())
  body.appendParagraph("This template was filled at " + (new Date().toISOString()))
  document.saveAndClose();
  
  return "Created: " + newFile.getName();
}

function baseName(name) {
  return name
      .replace("template", "")
      .replace("Template", "")
      .replace("Sablon", "")
      .replace("sablon", "")
      .trim();
}