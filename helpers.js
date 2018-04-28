function loadSettings(activeSpreadsheet, settingsSheetName, settingsTemplateId) {
    ///// Load Settings Sheet
    var settings = activeSpreadsheet.getSheetByName(settingsSheetName)
    if (settings === null) {
        var set = SpreadsheetApp.openById(settingsTemplateId)
        var settings = set.getSheets()[0].copyTo(activeSpreadsheet).setName(settingsSheetName)
    }
    //        Logger.log(settings)
    var setLength = settings.getLastRow()
    //    Logger.log("last row " + setLength)
    var settingsRange = settings.getRange(1, 1, setLength, 2)
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
    return settingsObj
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


function loadFormTable(settingsObj, activeSpreadsheet) {
    var sheet = activeSpreadsheet.getSheetByName(settingsObj["responseSheet"])
    // Load Headers
    var cols = sheet.getLastColumn();
    var rows = sheet.getLastRow() - 1;
    var headerArray = getRowCells_(sheet, cols, 1)[0]

    var headerKey = {}
    for (var each in headerArray) {
        //      Logger.log(headers[0][each])
        headerKey[each] = headerArray[each]
    }
    runLog("Processing " + cols + " columns and " + rows + " rows");
    //		Logger.log("HeaderKey: " + JSON.stringify(headerKey))
    return {
        sheet: sheet,
        headerKey: headerKey,
        rows: rows,
        cols: cols,
        headerArray: headerArray
    }
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

function buildEmail(settingsObj, rowValuesArray, formTable, body, attachments) {
    var subjectName = merge(settingsObj["emailSubject"], rowValuesArray, formTable.headerArray,
        settingsObj["dateFormat"], settingsObj["timeZone"]) //, dateFormat, dateTimeZone
    var subHead = merge(settingsObj["emailSubHead"], rowValuesArray, formTable.headerArray,
        settingsObj["dateFormat"], settingsObj["timeZone"]) //, dateFormat, dateTimeZone
    var to = emailSplitJoin(settingsObj.emailTo, ",", ";")
    var body = body
    var cc = emailSplitJoin(settingsObj.emailCc, ",", ";")
    var bcc = emailSplitJoin(settingsObj.emailBcc, ",", ";")
    var from = settingsObj["emailFrom"]

    Email.send(to, subjectName, body, subHead, cc, bcc, from, attachments, settingsObj["emailWebhook"])

}



function search(nameKey, myArray) {
    for (var i = 0; i < myArray.length; i++) {
        Logger.log([myArray[i].name, nameKey])
        if (myArray[i].name === nameKey) {
            return myArray[i];
        }
    }
}


function alertTeam(title, text, url, id, webhookURI) {
    var body = "{data:{text:{{comment.value}}}}"
    var message = {
        "@context": "http://schema.org/extensions",
        "@type": "MessageCard",
        "themeColor": "02ff49",
        "title": title,
        "text": text,
        "potentialAction": [{
            "@type": "OpenUri",
            "name": "Open Task",
            "targets": [{
                "os": "default",
                "uri": url
            }]
        }]

    }

    var options = {
        'method': 'post',
        'contentType': 'application/json',
        // Convert the JavaScript object to a JSON string.
        'payload': JSON.stringify(message)
    };
    var results = UrlFetchApp.fetch(webhookURI, options);
    //	Logger.log(results)
}


function setStatus(settingsObj, sheet, cols, row, result) {
    try {
        var statusCols = {
            url: getByName(settingsObj["urlColumnName"], sheet) +
                1,
            status: getByName(settingsObj["statusColumnName"],
                sheet) + 1,
            id: getByName(settingsObj["idColumnName"], sheet) + 1
        }
        var statusRanges = {
            url: sheet.getRange(row, statusCols.url),
            status: sheet.getRange(row, statusCols.status),
            id: sheet.getRange(row, statusCols.id)
        }
        //        var date = moment()
        var date = moment().format("YYYY-MM-DDTHH:mm:ssZ")
        statusRanges.status.setValue("SENT: " + date)
        //                                     Utilities.formatDate(
        //			new Date(), "EST", "yyyy-MM-dd'T'HH:mm:ss'Z'"))
        statusRanges.url.setValue(result["url"])
        statusRanges.id.setValue(result["result"]["id"])
    } catch (e) {
        errorLog(e)
    }
}

var Email = {}
Email.send = function Email(to, subject, body, subhead, cc, bcc, from, attachments, webhook) {
    var payload = {}
    payload.to = to
    payload.subject = subject

    if (subhead) {
        payload.body = "<small>" + subhead + "</small><br><br>" + body
    } else {
        payload.body = body
    }
    payload.cc = cc
    payload.bcc = bcc
    payload.from = from
    if (attachments) {
        attachments.forEach(function(each) {
            each.ContentBytes = Utilities.base64Encode(each.ContentBytes)
        })
        payload.attachments = attachments
    }

    var options = {
        'method': 'post',
        'contentType': 'application/json',
        // Convert the JavaScript object to a JSON string.
        'payload': JSON.stringify(payload)
    };
    //  Logger.log(payload)
    var results = UrlFetchApp.fetch(webhook, options)
}

var Body = {}
Body.create = function(settingsObj, rowValues) {
    asanaBody = ""
    htmlBody = ""
    //  Logger.log(rowValues)
    for (var each in rowValues) {

        //      Logger.log(each)
        var excludeArray = JSON.parse(settingsObj["excludeFromBody"].trim())
        var found = ""
        for (var item in excludeArray) {
            if (excludeArray[item] === each)
                found = found + item
        }
        if (found) {
            Logger.log("Excluding " + each)
        } else {

            asanaBody += "<b>" + each + ":    </b>" + rowValues[each] +
                " \n \n"
            htmlBody += "<b>" + each + ":</b>&nbsp;&nbsp;&nbsp;&nbsp;" + rowValues[each] + "<br><br>"

        }
    }
    Body.asana = asanaBody
    Body.html = htmlBody
    return {
        asana: Body.asana,
        html: Body.html
    }
}


function ifSplit(item, splitBy) {
    if (item) {
        if (item.indexOf(splitBy) > 0) {
            item.split(splitBy).map(function(i) {
                return i.trim()
            })
        } else {
            return item.trim()
        }
    }
}

function merge(template, rowValues, headers, dateFormat, dateTimeZone) {
    Logger.log(template)
    if (template === undefined) {
        return ""
    }
    if (typeof template === "number") {
        template = template.toString()
    }
    if (JSON.stringify(template).indexOf("${") > 0) {
        return MailMerge.createTextFromTemplate(template, rowValues, headers, dateFormat, dateTimeZone)
    } else if (typeof template === "string") {
        return template
    } else {
        return ""
    }
}


function dateFormat(dateFormatArray, formTable, contents) {
    var rowValues = {}
    var rowValuesArray = []
    Logger.log(dateFormatArray)
    for (var each in contents) {
        //    Logger.log(each)
        var headerVal = formTable.headerKey[each].trim()
        var cell = contents[each]
        //            runLog("Processing "+dateFormatArray.length+" custom dates")
        //        Logger.log([headerVal, dateFormatArray, cell])
        for (var header in dateFormatArray) {
            if (headerVal === dateFormatArray[header].columnName && cell) {
                Logger.log(["MATCH", dateFormatArray[header].columnName, cell])
                var date = moment(cell).format(dateFormatArray[header].dateFormat)
                rowValues[headerVal] = date
                rowValuesArray.push(date)
            } else {
                rowValues[headerVal] = cell
                rowValuesArray.push(cell)
            }

            //        if (found) {
            //              Logger.log("Found: "+JSON.stringify(found))
            //            var date = moment(cell).format(found.dateFormat)
            //              Logger.log(date)
            //                rowValues[headerVal] = date
            //                rowValuesArray.push(date)
            ////              runLog("Tag ID for "+each.name+ " is "+each.id)
            //        } else {

            //            dateFormatArray.forEach( function (object, cell, headerVal) {
            //              Logger.log(object)
            //              
            //              if (headerVal === each.columnName) {
            //                
            //                var date = moment(cell).format(object.dateFormat)
            //                rowValues[headerVal] = date
            //                rowValuesArray.push(date)

            //
            //            rowValues[headerVal] = contents[each]
            //            rowValuesArray.push(contents[each])
        }
    }




    return {
        rowValues: rowValues,
        rowValuesArray: rowValuesArray,
        formTable: formTable
    }
}

function replaceText(replaceTextArray, valuesObj) {
    var rowValues = {}
    var rowValuesArray = []
    //    Logger.log(valuesObj)
    for (var each in valuesObj.rowValuesArray) {
        var headerVal = valuesObj.formTable.headerKey[each].trim()
        var cell = valuesObj.rowValuesArray[each]
        Logger.log([headerVal, cell])
        for (var header in replaceTextArray) {
            if (headerVal === replaceTextArray[header].columnName) {
                Logger.log(["MATCH", replaceTextArray[header].columnName, cell])
                if (cell) {
                    cell = cell.split(replaceTextArray[header].find).join(replaceTextArray[header].replace)
                    rowValues[headerVal] = cell
                    rowValuesArray.push(cell)
                }
            } else {
                rowValues[headerVal] = cell
                rowValuesArray.push(cell)
            }
        }
    }





    return {
        rowValues: rowValues,
        rowValuesArray: rowValuesArray,
        formTable: valuesObj.formTable
    }
}

function getSettings(settings) {
    var scriptProperties = PropertiesService.getScriptProperties()
    if (settings) {
        scriptProperties.setProperties(settings, true)
    }
    return scriptProperties.getProperties()
}

function test() {
    Logger.log(getSettings({
        "a": "B"
    }))
}

function loadStatusCols(settingsObj, formTable) {
    var statusCols = {}
    statusCols.statusColName = settingsObj["statusColumnName"];
    statusCols.statusCol = getByName(statusCols.statusColName, formTable.sheet) + 1;
    statusCols.statusRange = formTable.sheet.getRange(2, statusCols.statusCol, formTable.rows);
    statusCols.statusRangeValues = statusCols.statusRange.getValues();
    statusCols.verificationColName = "Timestamp" //settingsObj["statusColumnName"];
    statusCols.verificationCol = getByName(statusCols.verificationColName, formTable.sheet) + 1;
    statusCols.verificationRange = formTable.sheet.getRange(2, statusCols.verificationCol, formTable.rows);
    statusCols.verificationRangeValues = statusCols.verificationRange.getValues();
    //          Logger.log(statusCols.verificationCol)
    return statusCols

}


function dueDate(dueDateDuration) {
    //  Logger.log(dueDateDuration)
    if (typeof dueDateDuration === "number") {
        var date = moment()
        var dueDate = moment(date, "DD-MM-YYYY").add(dueDateDuration, 'days');
        Logger.log("Due Date: " + dueDate.format("YYYY-MM-DD"))
        return dueDate
    } else {
        return ""
    }
}



function getByName(colName, sheet) {
    //  var sheet = SpreadsheetApp.getActiveSheet();
    var data = sheet.getDataRange().getValues();
    var col = data[0].indexOf(colName);
    if (col != -1) {
        return col;
    }
}

function addDays(date, days) {
    var result = new Date(date);
    result.setDate(result.getDate() + days);
    return result;
}

function runLog(message, activeSheet) {
    try {
        var ID = RUNLOG_ID;
        if (activeSheet) {
            activeSheet.toast(message)
        }
        var sheet = SpreadsheetApp.openById(ID).getActiveSheet()
        sheet.appendRow([Utilities.formatDate(new Date(), "EST",
            "MM/dd/YYYY HH:mm"), message])
        Logger.log(message)
    } catch (e) {
        errorLog(e)
    }
}

function errorLog(e) {
    scriptUrl = SCRIPT_URL
    if (ERROR_EMAIL_ENABLED) {
        MailApp.sendEmail(ERROR_EMAIL_ADDRESS, ERROR_EMAIL_SUBJECT, "Error: " + e.message + "\nErrorType: " + e.filename + "\nLineNumber: " + e.lineNumber + "\n" + scriptUrl)
    }
    throw new Error(e.message + " | ErrorType: " + e.filename +

        " | LineNumber: " + e.lineNumber)
}

function getRowCells_(sheet, cols, row) {
    try {
        var rowContents = sheet.getRange(row, 1, 1, cols);
        var rowContentsValues = rowContents.getValues();
        return rowContentsValues
    } catch (e) {
        errorLog(e)
    }

}

function asanaProcess(settingsObj, formTable, rowValues, rowValuesArray) {
    var Asana = {
        taskName: "",
        task: {},
        response: {}
    }
    // runLog("Merging Row: " + row)
    // var contents = getRowCells_(formTable.sheet, formTable.cols, row)[0];
    // var valuesObj = dateFormat(JSON.parse(settingsObj["dateFormatArray"].trim()), formTable, contents)
    // valuesObj = replaceText(JSON.parse(settingsObj["replaceTextArray"]),valuesObj)
    // Logger.log(valuesObj)
    // var rowValues = valuesObj.rowValues
    // var rowValuesArray = valuesObj.rowValuesArray

    Asana.taskName = merge(settingsObj["taskName"], rowValuesArray, formTable.headerArray, settingsObj["dateFormat"],
        settingsObj["timeZone"]) //, dateFormat, dateTimeZone
    Asana.task.name = Asana.taskName
    if (PROCESS_TAGS) {
        Asana.task.tags = processTags(settingsObj, rowValues)
    }
    Asana.task.due_on = dueDate(settingsObj.dueDateDuration)
    Asana.task.assignee = settingsObj["assignee"]
    Asana.task.followers = ifSplit(settingsObj.followers, ",")

    Asana.task.projects = ifSplit(settingsObj["project IDs"], ",")
    Asana.task.hearted = settingsObj["liked"]
    Asana.task.assignee_status = settingsObj["status"]
    Body.create(settingsObj, rowValues)
    Asana.task.html_notes = Body.asana
    Asana.response = createAsanaTask(settingsObj, Asana.task)
    if (settingsObj.processChildren) {
        processChildren(settingsObj, rowValues, Asana.response)
    }
    return Asana
}