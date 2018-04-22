/****************************************************************** 
### TODO: 
- Cleanup, add section logic
- switch children and tags to batch processing
- append row to sheet
- allow replaceText to take multiple repeats of the same column 
- dateFormat should not error when given the same column twice
******************************************************************/



// Require Moment from Apps Script Libraries
// Project Key = MHMchiX6c1bwSqGM1PZiW_PxhMjh3Sh48
var moment = Moment.moment
//Requre Template Reader
//MD6VoUXdwcQpfeNDJUagDfFjzL90iNPPq
//var MailMerge = MailMerge.createTextFromTemplate(template, rowValues, headers, dateFormat, dateTimeZone)




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



function buildSettings() {
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
        var Asana = {
            taskName: "",
            task: {},
            response: {}
        }
        runLog("Merging Row: " + row)
        var contents = getRowCells_(formTable.sheet, formTable.cols, row)[0];
        var valuesObj = dateFormat(JSON.parse(settingsObj["dateFormatArray"].trim()), formTable, contents)
        valuesObj = replaceText(JSON.parse(settingsObj["replaceTextArray"]),valuesObj)
        Logger.log(valuesObj)
        var rowValues = valuesObj.rowValues
        var rowValuesArray = valuesObj.rowValuesArray
 
        Asana.taskName = merge(settingsObj["taskName"], rowValuesArray, formTable.headerArray, settingsObj["dateFormat"],
            settingsObj["timeZone"]) //, dateFormat, dateTimeZone
        Asana.task.name = Asana.taskName
        if (PROCESS_TAGS) {
            Asana.task.tags = processTags(settingsObj, rowValues)
        }
        Asana.task.due_on = dueDate(settingsObj.dueDateDuration)
        Asana.task.assignee = settingsObj["assignee"]
        Asana.task.followers = ifSplit(settingsObj.followers,",")

        Asana.task.projects = ifSplit(settingsObj["project IDs"],",")
        Asana.task.hearted = settingsObj["liked"]
        Asana.task.assignee_status = settingsObj["status"]
        Body.create(settingsObj, rowValues)
        Asana.task.html_notes = Body.asana
        Asana.response = createAsanaTask(settingsObj, Asana.task)
        if (PROCESS_CHILDREN) {
            processChildren(settingsObj, rowValues, Asana.response)
        }
        //        Logger.log(Asana.response.result)
        //        alertTeam(title, text, url, id, webhookURI)
        if (settingsObj["alertTeam"] === true) {
            alertTeam("Summer 2018: " + Asana.response.result["name"], Asana.response.result["notes"], Asana.response["url"],
                Asana.response.result["id"].toString(), settingsObj["teamsWebHookUri"])
        }
        //      Build Final Body
        Body.prefix = merge(settingsObj["emailHeader"], rowValuesArray, formTable.headerArray, settingsObj[
                "dateFormat"], settingsObj["timeZone"])
        
        Body.suffix = merge(settingsObj["emailFooter"], rowValuesArray, formTable.headerArray, settingsObj[
                "dateFormat"], settingsObj["timeZone"])
        
        Body.html = Body.prefix + '<a href="' + Asana.response["url"] + '">' + Asana.response["url"] + '</a><br><br>' + Body.html +
            Body.suffix
        // Logger.log(Body.html)
        //        Send Email
        if (settingsObj.emailEnabled) {
            buildEmail(settingsObj, rowValuesArray, formTable, Body.html)
        }
        if (UPDATE_STATUS) {
            setStatus(settingsObj, formTable.sheet, formTable.cols, row, Asana.response)
        }
        //        
    } catch (e) {
        errorLog(e)
    }
}

