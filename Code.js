/****************************************************************** 
### TODO: 
- Cleanup
- add section logic
- switch children and tags to batch processing
- append row to sheet
- replace range in google sheet (intake form)
- allow replaceText to take multiple repeats of the same column 
- dateFormat should not error when given the same column twice
- add additional projects to children 
- trim all fields as they come in
- refactor datereplaclement and textreplacement
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
                process(settingsObj, formTable, rowInt)
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

function process(settingsObj, formTable, row) {
    try {
        runLog("Merging Row: " + row)
        var contents = getRowCells_(formTable.sheet, formTable.cols, row)[0];
        var valuesObj = textFormat(JSON.parse(settingsObj["dateFormatArray"]), JSON.parse(settingsObj["replaceTextArray"]), formTable, contents)
        var rowValues = valuesObj.rowValues
        var rowValuesArray = valuesObj.rowValuesArray
        var Asana = asanaProcess(settingsObj, formTable, rowValues, rowValuesArray)
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
        if (settingsObj.updateStatus) {
            setStatus(settingsObj, formTable.sheet, formTable.cols, row, Asana.response)
        }
        //        
    } catch (e) {
        errorLog(e)
    }
}