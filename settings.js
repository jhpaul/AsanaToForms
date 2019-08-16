
var Settings = {}
function run() {
    processEntries()
}

function installSettings(){
    Settings.install()
}

function setSettings(settingsObj){
    return Settings.set(settingsObj)
}

function loadSettings(activeSpreadsheet, settingsName) {
    return Settings.load(activeSpreadsheet, settingsName);
}

function buildSettings(Settings) {
    return Settings.build(Settings)
}

function test(){
    PropertiesService.getScriptProperties().deleteAllProperties()
    Logger.log(Settings.get())
    Logger.log(Settings.install())
  }


Settings.install = function (){
   var scriptProperties = PropertiesService.getScriptProperties()
   if (!scriptProperties){
   scriptProperties.setProperties({
        settingsTemplate: "1fqFoBN1T_T78ME4XalCddEIosqFIghM9San3p6Fe5Ag",
        settingsName: "SETTINGS"
    });
    }
  Settings.build(Settings)
  runLog("Default Settings Added Successfully", SpreadsheetApp.getActiveSpreadsheet())


}

Settings.build = function(Settings) {
    var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet()
    var settingsObj = Settings.load(activeSpreadsheet, Settings)
    Logger.log(settingsObj)
    var sheet = settingsObj.settingsSheet
    var range = settingsObj.settingsRange
    var protection = sheet.protect()
    activeSpreadsheet.setNamedRange(Settings.settingsName, range)
    protection.setWarningOnly(true)
}





Settings.get = function(settings){
  		var scriptProperties = PropertiesService.getScriptProperties()
        if (scriptProperties) {
        return      scriptProperties.getProperties()
        }
  return null
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


/**
 * @description
 * @param {*} activeSpreadsheet
 * @param {*} settingsName
 * @returns {settingsObj: settingsObj, settingsRange: settingsRange, settingsSheet: settingsSheet}
 */
Settings.load = function(activeSpreadsheet, settingsName) {
    ///// Load Settings Sheet
    var scriptSettings = Settings.get()
    //var settingsSheet = activeSpreadsheet.getSheetByName("Settings")
    Logger.log("Active Spreadsheet \n" + toString(activeSpreadsheet));
    if (activeSpreadsheet) {
      var settingsSheet = activeSpreadsheet.getSheetByName("SETTINGS");
      } else var settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SETTINGS");
  
    
      if (!scriptSettings || scriptSettings === {}){
        Settings.set({
            settingsName:  defaults.settingsSheetName
        })
        var scriptSettings = Settings.get()
    }
    if (!scriptSettings.settingsName){
        Settings.set({
            settingsName: defaults.settingsSheetName
        })
        var scriptSettings = Settings.get()
    }
    if (!scriptSettings.settingsTemplate){
        Settings.set({
            settingsTemplate: defaults.settingsTemplate
        })
        var scriptSettings = Settings.get()
    }               
    Logger.log(scriptSettings)
//    if (!settingsSheet) {
        Logger.log("Run")
//        var set = SpreadsheetApp.openById(scriptSettings.settingsTemplate)
//        var settingsSheet = set.getSheets()[0].copyTo(activeSpreadsheet).setName(scriptSettings.settingsName)
//    }
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
    Logger.log (settingsObj)
    PERSONAL_ACCESS_TOKEN = settingsObj.personalAccessToken; // Put your unique Personal access token here
    WORKSPACE_ID = settingsObj.workspaceId; // Put in the main workspace key you want to access (you can copy from asana web address)
    ASSIGNEE = settingsObj.defaultAssignee; // put in the e-mail addresss you use to log into asana
    PREMIUM = settingsObj.asanaPremium
    PREMIUM_FIELDS = JSON.parse(settingsObj.premiumFields)
    RUNLOG_ID = settingsObj.logSheet
    Settings.object = {settingsObj: settingsObj, settingsRange: settingsRange, settingsSheet: settingsSheet}
    return Settings.object
}




