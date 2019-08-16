// Require Moment from Apps Script Libraries
// Project Key = MHMchiX6c1bwSqGM1PZiW_PxhMjh3Sh48

var moment = Moment.moment
//Requre Template Reader
//MD6VoUXdwcQpfeNDJUagDfFjzL90iNPPq
//var MailMerge = MailMerge.createTextFromTemplate(template, rowValues, headers, dateFormat, dateTimeZone)

/**
 * Creates a form driven trigger.
 */
function createTimeDrivenTriggers() {
    removeTriggers();
    var ss = SpreadsheetApp.getActiveSpreadsheet()
    // Trigger every 1 min.
    ScriptApp.newTrigger('processEntries()').timeBased().atHour(5).everyDays(1).create()
    //   ScriptApp.newTrigger('processEntries()').timeBased().atHour(15).everyDays(1).create()
    ScriptApp.newTrigger('processEntries').forSpreadsheet(ss).onFormSubmit().create()
    //  ScriptApp.newTrigger(functionName).forSpreadsheet(ss).

}

function removeTriggers() {
    // Deletes all triggers in the current project.
    var triggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < triggers.length; i++) {
        ScriptApp.deleteTrigger(triggers[i]);
    }
}

function onInstall(e) {
    Settings.install()
    onOpen(e);

}

function onOpen() {
    var ui = SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
    ui.createMenu('Asana Forms').addItem('Process Now', 'processEntries')
        .addSubMenu(
            ui
            .createMenu('Advanced - 21')
            .addItem('Add Trigger', 'createTimeDrivenTriggers')
            .addItem('Remove Triggers', 'removeTriggers')
            // .addItem('Settings', 'showDialog')
            .addItem('Install', 'installSettings')
        ).addToUi();
    loadSettings(SpreadsheetApp.getActiveSpreadsheet(), Settings.get.settingsName)
}

/**
 * @description
 * @param {*} dialogTemplate
 */
function showDialog(dialogTemplate) {
    var dialogTemplate = "SettingsDialog.html"
    var ui = HtmlService.createTemplateFromFile(dialogTemplate)
        .evaluate()
        .setWidth(700)
        .setHeight(600)
        .setSandboxMode(HtmlService.SandboxMode.IFRAME);
    SpreadsheetApp.getUi().showModalDialog(ui, "Installation Settings");
}


/**
 * @description
 */
function processEntries(activeSpreadsheet) {
    runLog("Start Merge")
    if (activeSpreadsheet) {}
    else { activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet(); }
    var settings = loadSettings(activeSpreadsheet, "SETTINGS")
    var settingsObj = settings.settingsObj
    var formTable = loadFormTable(settingsObj["responseSheet"], activeSpreadsheet)
    var statusCols = loadStatusCols(settingsObj, formTable)
    //  Get and merge any row with no status value
    var mergeCount = 0
    for (var index in statusCols.statusRangeValues) {
        if (statusCols.statusRangeValues[index] == "" && statusCols.verificationRangeValues[index] != "") {
            var rowInt = parseInt(index) + 2
            process(settingsObj, formTable, rowInt)
            runLog("Sucessfully Merged Row: " + rowInt)
            mergeCount++
        }
    }
    Logger.log(mergeCount)
    if (mergeCount !== 1) {
        runLog(mergeCount + " Merges Completed Successfully", activeSpreadsheet)
    }
    else {
        runLog(mergeCount + " Merge Completed Successfully", activeSpreadsheet)
    }
}

/**
 * @description
 * @param {*} settingsObj
 * @param {*} formTable
 * @param {*} row
 */
function process(settingsObj, formTable, row) {
    runLog("Merging Row: " + row)
    var contents = getRowCells_(formTable.sheet, formTable.cols, row)[0];

    try {
        var valuesObj = textFormat(JSON.parse(settingsObj["dateFormatArray"]), JSON.parse(settingsObj["replaceTextArray"]), formTable, contents)
    }
    catch (e) {
        errorLog(e, "Text & Date Replacement Failed")
    }
    //    console.log(valuesObj)
    var rowValues = valuesObj.rowValues
    var rowValuesArray = valuesObj.rowValuesArray

    try {
        var Asana = asanaProcess(settingsObj, formTable, rowValues, rowValuesArray)
    }
    catch (e) {
        errorLog(e, "Sending to Asana Failed")
    }
    //        Logger.log(Asana.response.result)
    //        alertTeam(title, text, url, id, webhookURI)
    if (settingsObj.alertTeam) {
        alertTeam("Summer 2018: " + Asana.response.result["name"], Asana.response.result["notes"], Asana.response["url"],
            Asana.response.result["id"].toString(), settingsObj["teamsWebHookUri"])
    }
    //      Build Final Body
    Body.prefix = merge(settingsObj["emailHeader"], rowValuesArray, formTable.headerArray, settingsObj.titleDateFormat, settingsObj.timeZone)

    Body.suffix = merge(settingsObj["emailFooter"], rowValuesArray, formTable.headerArray, settingsObj.titleDateFormat, settingsObj.timeZone)

    Body.html = Body.prefix + '<a href="' + Asana.response["url"] + '">' + Asana.response["url"] + '</a><br><br>' + Body.html +
        Body.suffix
    // Logger.log(Body.html)
    //        Send Email
    if (settingsObj.emailEnabled) {
        Email.build(settingsObj, rowValuesArray, formTable, Body.html)
    }
    if (settingsObj.updateStatus) {
        setStatus(settingsObj, formTable.sheet, formTable.cols, row, Asana.response)
    }
    if (settingsObj.appendToDestination) {
        addToTracker(settingsObj, valuesObj.cleanValuesArray, valuesObj.originalFormTable.headerArray, Asana.response)
    }
    //        
}
