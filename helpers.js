/**
 * @description Loads the response sheet into an object variable and sets it as activeSpreadsheet
 * @param {*} settingsObj
 * @param {*} activeSpreadsheet
 * @returns formTable
 */
function loadFormTable(dataSheet, activeSpreadsheet) {
    var sheet = activeSpreadsheet.getSheetByName(dataSheet)

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
		Logger.log("HeaderKey: " + JSON.stringify(headerKey))
    var formTable = {
        sheet: sheet,
        headerKey: headerKey,
        rows: rows,
        cols: cols,
        headerArray: headerArray
    }
    return formTable
}

/**
 * @description
 * @param {*} sheet
 * @param {*} cols
 * @param {*} row
 * @returns
 */
function getRowCells_(sheet, cols, row) {
    try {
        var rowContents = sheet.getRange(row, 1, 1, cols);
        var rowContentsValues = rowContents.getValues();
        return rowContentsValues
    } catch (e) {
        errorLog(e)
    }

}


/**
 * @description remove commas from email addresses
 * @param {*} value
 * @param {*} split
 * @param {*} join
 * @returns replaces a charachter in a list of strings 
 */
function emailSplitJoin(value, split, join) {
    //  Logger.log(value)
    if (value) {
        return value.split(split).join(join)
    } else {
        return null
    }
}

function swap(json){
    var ret = {};
    for(var key in json){
      ret[json[key]] = key;
    }
    return ret;
  }


/**
 * @description Find Value in array of objects [https://stackoverflow.com/questions/12462318/find-a-value-in-an-array-of-objects-in-javascript]
 * @param {*} nameKey
 * @param {*} myArray
 * @returns
 */
function search(nameKey, myArray) {
    for (var i = 0; i < myArray.length; i++) {
        // Logger.log([myArray[i].name, nameKey])
        if (myArray[i].name === nameKey) {
            // Logger.log(myArray[i])
            return myArray[i];
        }
    }
}



/**
 * @description send an alert to microsoft teams feed
 * @param {*} title
 * @param {*} text
 * @param {*} url
 * @param {*} id
 * @param {*} webhookURI
 */
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




/**
 * @description Add status to row
 * @param {*} settingsObj
 * @param {*} sheet
 * @param {*} cols
 * @param {*} row
 * @param {*} result
 */
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

var Body = {}
Body.create = function(settingsObj, rowValues) {
    asanaBody = "<body>"
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
    Body.asana = asanaBody + '</body>'
    Body.html = htmlBody
    return {
        asana: Body.asana,
        html: Body.html
    }
}


/**
 * @description
 * @param {*} item
 * @param {*} splitBy
 * @returns array of string contents split by a delimiter
 */
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

/**
 * @description uses MailMerge library to return a template with all curly brace 
 *              values replaced with their appropriate entry variables
 * @param {*} template
 * @param {*} rowValues
 * @param {*} headers
 * @param {*} dateFormat
 * @param {*} dateTimeZone
 * @param {*} 
 * @returns either the template or the template with values replaced
 */
function merge(template, rowValues, headers, dateFormat, dateTimeZone) {
    Logger.log("Merging: "+template)
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



/**
 * @description Replace text and date strings within form values.
 * @param {*} dateFormatArray
 * @param {*} replaceTextArray
 * @param {*} formTable
 * @param {*} contents
 * @returns
 */
function textFormat(dateFormatArray, replaceTextArray, formTable, contents) {
    var originalFormTable = formTable
    var rowValues = {}
    var rowValuesArray = []
    var cleanValues = {}
    var cleanValuesArray = []
    Logger.log(dateFormatArray)
    for (var each in contents) {
        //    Logger.log(each)
        var headerVal = formTable.headerKey[each].trim()
        var cell = contents[each]

        if (typeof cell === "string") { cell = cell.trim() }
        if (cell < 10000 && (headerVal.toLowerCase() === "zip" || headerVal.toLowerCase() === "zip code" || headerVal.toLowerCase() === "zipcode") ) {cell = "0"+cell}
        Logger.log(cell)
        //            runLog("Processing "+dateFormatArray.length+" custom dates")
        //        Logger.log([headerVal, dateFormatArray, cell])
        var clean = cell
        for (var header in dateFormatArray) {
          
          
            if (headerVal === dateFormatArray[header].columnName && cell) {
                Logger.log(["MATCH", dateFormatArray[header].columnName, cell])
                var date = moment(cell).format(dateFormatArray[header].dateFormat)
                var dateClean = Utilities.formatDate(moment(cell).toDate(),"EST", "MM/dd/yyyy HH:mm:ss")
                cell = date
                clean = dateClean

            } 

        }
        for (var header in replaceTextArray) {
            if (headerVal === replaceTextArray[header].columnName) {
                Logger.log(["MATCH", replaceTextArray[header].columnName, cell])
                if (cell) {
                    cell = cell.split(replaceTextArray[header].find).join(replaceTextArray[header].replace)

                }
            } 
         }
                rowValues[headerVal] = cell
                rowValuesArray.push(cell)
                cleanValues[headerVal] = clean
                cleanValuesArray.push(clean)
                
    }

  var valuesObj = {
        rowValues: rowValues,
        rowValuesArray: rowValuesArray,
        formTable: formTable,
        cleanValues: cleanValues,
        cleanValuesArray: cleanValuesArray,
        originalFormTable: originalFormTable
    }
//Logger.log(valuesObj)
    return valuesObj
 
}



/**
 * @description
 * @param {*} settingsObj
 * @param {*} formTable
 * @returns range of status columns
 */
function loadStatusCols(settingsObj, formTable) {
  try {
    var statusCols = {}
    statusCols.statusColName = settingsObj["statusColumnName"];
    statusCols.statusCol = getByName(statusCols.statusColName, formTable.sheet) + 1;
    statusCols.statusRange = formTable.sheet.getRange(2, statusCols.statusCol, formTable.rows);
    statusCols.statusRangeValues = statusCols.statusRange.getValues();
    statusCols.verificationColName = "Timestamp" //settingsObj["statusColumnName"];
    statusCols.verificationCol = getByName(statusCols.verificationColName, formTable.sheet) + 1;
    statusCols.verificationRange = formTable.sheet.getRange(2, statusCols.verificationCol, formTable.rows);
    statusCols.verificationRangeValues = statusCols.verificationRange.getValues();
          Logger.log(statusCols.verificationCol)
    return statusCols
  } catch (e) {
    errorLog(e ,"Unable to load Status Columns");
    
    }
}


/**
 * @description return column number for a given header name
 * @param {*} colName
 * @param {*} sheet
 * @returns
 */
function getByName(colName, sheet) {
    var data = sheet.getDataRange().getValues();
    var col = data[0].indexOf(colName);
    if (col != -1) {
        return col;
    }
}


/**
 * @description add days to a date
 * @param {*} date
 * @param {*} days
 * @returns
 */
function addDays(date, days) {
    var result = new Date(date);
    result.setDate(result.getDate() + days);
    return result;
}

function addToTracker(settingsObj, rowValues, headers, response){
    try {
    trackerWorkbook = SpreadsheetApp.openById(settingsObj["appendDestinationId"]);
    trackerSheet = trackerWorkbook.getSheetByName(settingsObj["appendDestinationSheet"]);
    targetFormTable = loadFormTable(settingsObj["appendDestinationSheet"], SpreadsheetApp.openById(settingsObj["appendDestinationId"]) )
    urlName = settingsObj["urlColumnName"]
    rowValues[headers.indexOf(urlName)] = response.url
    mergedObj = merge(settingsObj["appendColumns"], rowValues, headers, settingsObj.titleDateFormat, settingsObj.timeZone)
    var trackArray = JSON.parse(mergedObj)
    // Logger.log(JSON.stringify(targetFormTable["headerArray"]) +'\n====\n'+JSON.stringify(trackArray))
    
    trackArray.forEach(function(each) {
        var headerIndex = swap(targetFormTable.headerKey)
        var index = headerIndex[each["targetColumn"]]
        if (index) {
        // Logger.log('\n=='+parseInt(index))

            var col = parseInt(index) + 1
            var row = targetFormTable.rows + 2
     
            // Logger.log(col+"\n"+row)
            var cell = trackerSheet.getRange(row, col)
            cell.setValue(each["sourceColumn"])
        }

    });


    CopyFormulasDown.copyFormulasDown(trackerSheet, settingsObj["appendCopyFormulaRow"]);
    } catch(e) { errorLog(e) }
  }
  







function runLog(message, activeSheet) {
    try {
//        var ID = RUNLOG_ID;
//        if (activeSheet) {
//            activeSheet.toast(message)
//            console.log(message)
//        }
//        var sheet = SpreadsheetApp.openById(ID).getActiveSheet()
//        sheet.appendRow([Utilities.formatDate(new Date(), "EST",
//            "MM/dd/YYYY HH:mm"), message])
        Logger.log(message)
    } catch (e) {
        errorLog(e, message)
    }
}

function errorLog(e, message) {

  throw new Error(message + ": "+e.name)//+" | "+e.message+" | " + e.lineNumber+ " | " + e.fileName)
}