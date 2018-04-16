// TODO: Cleanup, add section logic, switch children and tags to batch processing
// Require Moment from Apps Script Libraries
// Project Key = MHMchiX6c1bwSqGM1PZiW_PxhMjh3Sh48
var moment = Moment.moment
//Requre Template Reader
//MD6VoUXdwcQpfeNDJUagDfFjzL90iNPPq
//var MailMerge = MailMerge.createTextFromTemplate(template, rowValues, headers, dateFormat, dateTimeZone)


function merge(template, rowValues, headers, dateFormat, dateTimeZone) {
    Logger.log(typeof template)
    if (template === undefined) {
        return
    }
    if (typeof template === "number") {
        template = template.toString()
    }
    if (JSON.stringify(template).indexOf("${") >= 0) {
        return MailMerge.createTextFromTemplate(template, rowValues, headers, dateFormat, dateTimeZone)
    } else if (typeof template === "string") {
        return template
    } else {
        return ""
    }
}

function onInstall(e) {
    onOpen(e);

}

function onOpen() {
    SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
        .createMenu('Asana Merge').addItem('Merge Now', 'processEntries')
        .addItem('Settings', 'showDialog').addToUi();
}

function showDialog() {
  var ui = HtmlService.createTemplateFromFile('Dialog')
      .evaluate()
      .setWidth(700)
      .setHeight(600)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showModalDialog(ui, "Installation Settings");
}



function buildSettings(){
  var scriptProperties = PropertiesService.getScriptProperties();
scriptProperties.setProperties({
  settingsTemplate: "1fqFoBN1T_T78ME4XalCddEIosqFIghM9San3p6Fe5Ag",
  settingsName: "SETTINGS"
});
Logger.log(scriptProperties.getProperties())
    var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet()
    
    loadSettings(activeSpreadsheet, settingsName)
    var sheet = activeSpreadsheet.getSheetByName(settingsName)
    var range = sheet.getActiveRange()
    var protection = sheet.protect()
    activeSpreadsheet.setNamedRange(settingsName, range)
    protection.setWarningOnly(true)
//function showSidebar() {
 var html = HtmlService.createHtmlOutputFromFile('page')
     .setTitle('My custom sidebar')
     .setWidth(300);
 SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
     .showSidebar(html);
}



// Set multiple script properties in one call.


function processEntries() {
    try {
        runLog("Start Merge")
        var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
        var settingsObj = loadSettings(activeSpreadsheet, "SETTINGS")
        var formTable = loadFormTable(settingsObj, activeSpreadsheet)
        var statusCols = loadStatusCols(settingsObj, formTable)
        //  Get and merge any row with no status value
        var mergeCount = 0
        for (var index in statusCols.statusRangeValues) {
            if (statusCols.statusRangeValues[index] == "" && statusCols.verificationRangeValues[index] != "") {
                var rowInt = parseInt(index) + 2
                process_(settingsObj, formTable, rowInt)
                runLog("Sucessfully Merged Row: " + rowInt)
                mergeCount++
            }
        }
        if (mergeCount !== 1) {
            runLog(mergeCount + " Merges Completed Successfully", activeSpreadsheet)
        } else {
            runLog(mergeCount + " Merge Completed Successfully", activeSpreadsheet)
        }
    } catch (e) {
        errorLog(e)
    }
}

function process_(settingsObj, formTable, row) {
    try {
        runLog("Merging Row: " + row)
        var contents = getRowCells_(formTable.sheet, formTable.cols, row)[0];
        var valuesObj = dateFormat(JSON.parse(settingsObj["dateFormatArray"].trim()), formTable, contents)
        var rowValues = valuesObj.rowValues
        var rowValuesArray = valuesObj.rowValuesArray
        var task = {}
        var taskName = merge(settingsObj["taskName"], rowValuesArray, formTable.headerArray, settingsObj["dateFormat"],
            settingsObj["timeZone"]) //, dateFormat, dateTimeZone
        task.name = taskName
        if (PROCESS_TAGS) {
            task.tags = processTags(settingsObj, rowValues)
        }
        task.due_on = dueDate(settingsObj.dueDateDuration)
        task.assignee = settingsObj["assignee"]
        task.followers = settingsObj["followers"].split(",").map(function(item) {
            return item.trim();
        });
        task.projects = settingsObj["project IDs"].split(",").map(function(item) {
            return item.trim();
        });
        task.hearted = settingsObj["liked"]
        task.assignee_status = settingsObj["status"]
        var body = createBody(settingsObj, rowValues)
        task.html_notes = body.asana
        var taskResults = createAsanaTask(settingsObj, task)
        if (PROCESS_CHILDREN) {
            processChildren(settingsObj, rowValues, taskResults)
        }
        //        Logger.log(taskResults.result)
        //        alertTeam(title, text, url, id, webhookURI)
        if (settingsObj["alertTeam"] === true) {
            alertTeam("Summer 2018: " + taskResults.result["name"], taskResults.result["notes"], taskResults["url"],
                taskResults.result["id"].toString(), settingsObj["teamsWebHookUri"])
        }
        //      Build Final Body
        if (settingsObj["emailHeader"]) {
            var prefix = merge(settingsObj["emailHeader"], rowValuesArray, formTable.headerArray, settingsObj[
                "dateFormat"], settingsObj["timeZone"])
        }
        if (settingsObj["emailFooter"]) {
            var suffix = merge(settingsObj["emailFooter"], rowValuesArray, formTable.headerArray, settingsObj[
                "dateFormat"], settingsObj["timeZone"])
        }
        body.html = prefix + '<a href="' + taskResults["url"] + '">' + taskResults["url"] + '</a><br><br>' + body.html +
            suffix
        Logger.log(body.html)
        //        Send Email
        if (settingsObj.emailEnabled) {
            buildEmail(settingsObj, rowValuesArray, formTable, body.html)
        }
        if (UPDATE_STATUS) {
            setStatus(settingsObj, formTable.sheet, formTable.cols, row, taskResults)
        }
        //        
    } catch (e) {
        errorLog(e)
    }
}