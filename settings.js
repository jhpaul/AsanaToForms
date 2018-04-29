
// function getSettings(settings) {
//     var scriptProperties = PropertiesService.getScriptProperties()
//     if (settings) {
//         scriptProperties.setProperties(settings, true)
//     }
//     return scriptProperties.getProperties()
// }

// function test() {
//     Logger.log(getSettings({
//         "a": "B"
//     }))
// }


var Settings = {}
function installSettings(){
    Settings.install()
}
Settings.install = function (){
   var scriptProperties = PropertiesService.getScriptProperties()
   if (!scriptProperties){
   scriptProperties.setProperties({
        settingsTemplate: "1fqFoBN1T_T78ME4XalCddEIosqFIghM9San3p6Fe5Ag",
        settingsName: "SETTINGS"
    });
    }
  buildSettings(Settings)
  runLog("Default Settings Added Successfully", SpreadsheetApp.getActiveSpreadsheet())
//  return Settings.get()

}


Settings.get = function(settings){
  		var scriptProperties = PropertiesService.getScriptProperties()
        if (scriptProperties) {
        return      scriptProperties.getProperties()
        }
  return null
}

function setSettings(settingsObj){
    Settings.set(settingsObj)
}

Settings.set = function(settings) {
  if (settings){
    var scriptProperties = PropertiesService.getScriptProperties()
          scriptProperties.setProperties(settings)
        }
    }

Settings.clear = function(){
    PropertiesService.getScriptProperties().deleteAllProperties

}



function test(){
  PropertiesService.getScriptProperties().deleteAllProperties()
  Logger.log(Settings.get())
  Logger.log(Settings.install())
}



function loadSettings(activeSpreadsheet, settingsName) {
    ///// Load Settings Sheet
    var scriptSettings = Settings.get()
    var settingsSheet = activeSpreadsheet.getSheetByName(scriptSettings.settingsName)

    if (!scriptSettings || scriptSettings === {}){
        Settings.set({
            settingsName: "SETTINGS"
        })
        var scriptSettings = Settings.get()
    }
    if (!scriptSettings.settingsName){
        Settings.set({
            settingsName: "SETTINGS"
        })
        var scriptSettings = Settings.get()
    }
    if (!scriptSettings.settingsTemplate){
        Settings.set({
            settingsTemplate: "1fqFoBN1T_T78ME4XalCddEIosqFIghM9San3p6Fe5Ag"
        })
        var scriptSettings = Settings.get()
    }               

    Logger.log(scriptSettings)
    if (!settingsSheet) {
      Logger.log("Run")
        var set = SpreadsheetApp.openById(scriptSettings.settingsTemplate)
        var settingsSheet = set.getSheets()[0].copyTo(activeSpreadsheet).setName(scriptSettings.settingsName)
    }
    //        Logger.log(settings)
    var setLength = settingsSheet.getLastRow()
    //    Logger.log("last row " + setLength)
    var settingsRange = settingsSheet.getRange(1, 1, setLength, 2)
    //        Logger.log(settings.getRange(1,1,setLength, settings.getLastColumn()).getValues())
    var settingsValues = settingsRange.getValues()
    //		Logger.log(settingsValues)
    var settingsObj = {}
    for (var each in settingsValues) {
        var key = settingsValues[each][0]
        var value = settingsValues[each][1]
        if (value !== "") {
            settingsObj[key] = value
        }
    }
    PERSONAL_ACCESS_TOKEN = settingsObj.personalAccessToken; // Put your unique Personal access token here
    WORKSPACE_ID = settingsObj.workspaceId; // Put in the main workspace key you want to access (you can copy from asana web address)
    ASSIGNEE = settingsObj.defaultAssignee; // put in the e-mail addresss you use to log into asana
    PREMIUM = settingsObj.asanaPremium
    PREMIUM_FIELDS = JSON.parse(settingsObj.premiumFields)
    return {settingsObj: settingsObj, settingsRange: settingsRange, settingsSheet: settingsSheet}
}

function buildSettings(Settings) {
//    var scriptProperties = PropertiesService.getScriptProperties();
//    scriptProperties.setProperties({
//        settingsTemplate: "",
//        settingsName: "SETTINGS"
//    });
//    Logger.log(scriptProperties.getProperties())
  
    var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet()
    var settingsObj = loadSettings(activeSpreadsheet, Settings)
    Logger.log(settingsObj)
    var sheet = settingsObj.settingsSheet
    var range = settingsObj.settingsRange
    var protection = sheet.protect()
    activeSpreadsheet.setNamedRange(Settings.settingsName, range)
    protection.setWarningOnly(true)
    //function showSidebar() {
//    var html = HtmlService.createHtmlOutputFromFile('page')
//        .setTitle('My custom sidebar')
//        .setWidth(300);
//    SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
//        .showSidebar(html);
}




function run() {
    processEntries()
}

function emailSplitJoin(value, split, join) {
    //  Logger.log(value)
    if (value) {
        return value.split(split).join(join)
    } else {
        return null
    }
}
