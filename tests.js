var Test = {}


function testFunction() { return Test.dateFormat(Test) }




Test.dateFormat = function(Test){
  Test.sheet = HtmlService.createTemplateFromFile('testSheet').getRawContent()
  Test.data = Utilities.parseCsv(Test.sheet);
    Logger.log(Test.data)
    var dateFormatArray =     [{
        "columnName": "Date of Birth", 
        "dateFormat": "YYYY-MM-DD"
    }]
    var formTable =    {
        sheet: sheet,
        headerKey: headerKey,
        rows: rows,
        cols: cols,
        headerArray: headerArray
    }
    var contents = []
    return dateFormat(dateFormatArray, formTable, contents)
}


    