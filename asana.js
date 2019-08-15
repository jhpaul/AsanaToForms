/**
 * @description build an asana task
 * @param {*} settingsObj
 * @param {*} formTable
 * @param {*} rowValues
 * @param {*} rowValuesArray
 * @returns
 */
function asanaProcess(settingsObj, formTable, rowValues, rowValuesArray) {
    var Asana = {
        taskName: "",
        task: {},
        response: {}
    }
    Asana.taskName = merge(settingsObj["taskName"], rowValuesArray, formTable.headerArray, settingsObj.titleDateFormat, settingsObj.timeZone) //, dateFormat, dateTimeZone
    Asana.task.name = Asana.taskName
    if (settingsObj.processTags) {
        Asana.task.tags = processTags(settingsObj, rowValues)
    }
    if ( settingsObj["dueTime"] === "" || settingsObj["dueTime"] === null ||  settingsObj["dueTime"] === undefined) {
        Asana.task.due_on = dueDate(settingsObj["dueDateDuration"])
    } else {

        Asana.task.due_at = dueTime(settingsObj["dueDateDuration"], settingsObj["dueTime"])        
    }
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


/**
 * @description
 * @param {*} dueDateDuration
 * @returns date-formatted value of relative due date for asana task
 */
function dueDate(dueDateDuration) {
    dueDateDuration = parseInt(dueDateDuration)
    Logger.log("Due relative: "+ dueDateDuration)
    if (dueDateDuration >= 0) {
        var date = moment()
        var dueDate = moment(date, "DD-MM-YYYY").add(dueDateDuration, 'days');
        Logger.log("Due Date: " + dueDate.format("YYYY-MM-DD"))
        return dueDate
    } else {
        return 
    }
}


/**
 * @description 
 * @param {*} dueDateDuration
 * @param {*} dueTime
 * @returns datetime-formatted value of relative due date for asana task
 */
function dueTime(dueDateDuration, dueTime) {
    dueDateDuration = parseInt(dueDateDuration)
    Logger.log("Due relative: "+ dueDateDuration)
    if (dueDateDuration >= 0) {
        var now = moment().add(dueDateDuration,"days").format('MM/DD/YYYY');
        var time = moment(dueTime,"HH:mm").format("HH:mm")
        Logger.log(time)
        var dueDate = moment(now+time,"MM/DD/YYYYHH:mm")
        // var dueDate = moment(date, "DD-MM-YYYY").add(dueDateDuration, 'days').add(moment.duration(dueTime).asHours(),'hours');
        Logger.log("Due Time: " + dueDate.format("YYYY-MM-DD hh:mm A"))
        return dueDate
    } else {
        return 
    }
}


///////////////Task Object
//{
//  "name": "",
//  "parent": "",
//  "notes": "",
//  "assignee_status": "",
//  "assignee": "",
//  "completed": false,
//  "followers": [],
//  "workspace": WORKSPACE_ID,
//  "due_on": "",
//  "due_at": "",
//  "start_on": "2015-01-01",
//  "projects": [],
//  "custom_fields": [],
//  "hearted": false,
//  "memberships": [{"project":"","section":""}],
//  "tags":[]
//}
/*************************
 * Asana     Functions    *
 *************************/
//// first Global constants ... Key Ids / tokens etc.
//PERSONAL_ACCESS_TOKEN = "0/"; // Put your unique Personal access token here
//WORKSPACE_ID = ""; // Put in the main workspace key you want to access (you can copy from asana web address)
//ASSIGNEE = ""; // put in the e-mail addresss you use to log into asana
//PREMIUM = false
//PREMIUM_FIELDS = ["start_on"]
// ** testTask() **  is useful for using as a Debug start point.  "select function" on script editor menu
// choose "testTask" then debug functionality is enabled
function testTask() {
    quickTask("a quick task")
};

// ** quickTask(taskName) ** Made a short function so I could just add simple tasks easily
function quickTask(taskName) {
    var newTask = {
        name: taskName,
        workspace: "",
        project: "",
        assignee: "me"
    }
    createAsanaTask(newTask);
};


/******************************************************************************************
 **  createAsanaTask(task) **
 ************************ 
 * creates a new asana task with information (like task name, project, notes etc.) contained in  
 * the  object 'newTask" passed to it.
 * 'task' should be of the format an object with option pairs that match the Asana task
 * key parameters, as many or as few as you want.
 * e.g. 
 * var newTask = {
 *   name: taskName,
 *   workspace: WORKSPACE_ID,
 *   project: "My Project",       // if you have a project you like to add add it here
 *   assignee: "JohnDoe@madeupmail.com"     // person the task should be assigned to.
 * }
 *  you could add other info like due dates etc.
 * it returns a "task" object containing all asana task elements of the one task created including the id.
 var task = {
	"name": "Hello, world!",
	"parent": "",
	"notes": "How are you today?",
	"assignee_status": "",
	"assignee": "me",
	"completed": false,
	"followers": ["email@example.com"],
	"workspace": "1234",
	"due_on": "2015-01-01",
	"due_at": "2018-03-24T19:45:12-05:00",
	"start_on": "2015-01-01",
	"projects": [],
	"custom_fields": [],
	"hearted": false,
	"memberships": [{"project":"","section":""}],
	"tags":[]
} 
SOURCE = https://stackoverflow.com/questions/30188226/creating-asana-tasks-from-google-apps-script?rq=1
 *************************************************************************************************/
function createAsanaTask(settingsObj, task) {
    // when creating an Asana task you must have at least a workspace id and an assignee
    // this routine checks if you defined one in the argument you passed
    if (task.workspace == null) {
        task.workspace = WORKSPACE_ID
    }
    if (task.assignee == null) {
        task.assignee = "me";
    }
    /* first setup  the "options" object with the following key elements:
     *
     *   method: can be GET,POST typically
     *
     *   headers: object containing header option pairs
     *                    "Accept": "application/json",        // accept JSON format
     *                    "Content-Type": "application/json",  //content I'm passing is JSON format
     *                    "Authorization": "Bearer " + PERSONAL_ACCESS_TOKEN // authorisation
     *  the authorisation aspect took me ages to figure out.
     *  for small apps like this use the Personal Access Token method.
     *  the important thing is to use the Authorization option in the header with the 
     *  parameter of  "Bearer " + PERSONAL_ACCESS_TOKEN
     *  the PERSONAL_ACCESS_TOKEN  is exactly the string as given to you in the Asana Web app at
     *  the time of registering a Personal Access Token.  it DOES NOT need any further authorisation / exchanges
     *  NOR does it needo any encoding in base 64 or any colon.
     *
     *  payload: this can be an object with option pairs  required for each element to be created... in this case 
     *           its the task elements as passed to this function in the argument "task" object.
     *            I found it doesn't need stringifying or anything.   
     *
     ********************************************************************************************************/
    //  Logger.log(JSON.stringify(task))
    //	Logger.log(task["memberships"])
    for (var each in task) {
        //		Logger.log(each)
        if (task[each] === "" || task[each] === []) {
            delete task[each]
        }
        if (each === "memberships") {
            if (task[each][0]["project"] === "" && task[each][0]["section"] === "") {
                delete task[each]
            }
        }
        if (each === "due_on" && task["due_at"] !== "" && task["due_at"] !== undefined) {
            delete task[each]
        }
        if (PREMIUM === false && PREMIUM_FIELDS.indexOf(each) > -1) {
            delete task[each]
        }
    }
    // Logger.log(JSON.stringify(task))
    var options = {
        "method": "POST",
        "headers": {
            "Accept": "application/json",
            "Content-Type": "application/json",
            "Authorization": "Bearer " + PERSONAL_ACCESS_TOKEN
        },
        "payload": JSON.stringify({
            "data": task
        })
    };
    // using try to capture errors 
    try {
        // set the base url to appropriate endpoint - 
        // this case is "https://app.asana.com/api/1.0"  plus "/tasks"
        // note workspace id or project id not in base url as they are in the payload options
        // use asana API reference for the requirements for each method
        //		var APIurl = "https://app.asana.com/api/1.0/tasks";
        //		// using url of endpoint and options object do a urlfetch.
        //		// this returns an object that contains JSON data structure into the 'result' variable 
        //		// see below for sample structure
        //		var result = UrlFetchApp.fetch(APIurl, options);
        //		// 
        //		var taskJSON = JSON.parse(result.getContentText());
        var taskJSON = callAsanaApi("POST", "tasks", JSON.stringify({
            "data": task
        }))
        //		Logger.log(task);
        var url = "https://app.asana.com/0/0/" + taskJSON["id"] + "/f"
        //        Logger.log("Task Created: \n"+task.name + "\n    " +url+"\n")
        runLog("Task Created: \n" + task.name + "\n    " + url + "\n")
        return {
            url: url,
            result: taskJSON
        }
    } catch (e) {
        Logger.log(e.message);
        throw new Error(e.message);
        return null;
    } finally {
        // parse the result text with JSON format to get object, then get the "data" element from that object and return it.
        // this will be an object containing all the elements of the task.
        //  try {
        //    Logger.log(taskJSON);
        //    return JSON.parse(taskJSON).data;
        //    } catch (e, taskJSON) {
        //        Logger.log(e);
        //        throw new Error(e);
        //        return null;
        //    }
    }
};

/**
 * @description
 * @param {*} method
 * @param {*} endpoint
 * @param {*} payload
 * @returns
 */
function callAsanaApi(method, endpoint, payload) {
    var options = {
        method: method,
        headers: {
            Authorization: "Bearer " + PERSONAL_ACCESS_TOKEN
        }
    };
    if (payload) {
        options.payload = payload
    }
    if (method === "POST") {
        options.headers["Accept"] = "application/json"
        options.headers["Content-Type"] = "application/json"
    }
    var APIurl = "https://app.asana.com/api/1.0/" + endpoint;
    //        runLog("Calling "+ APIurl +"\n" + JSON.stringify(options))
    var result = UrlFetchApp.fetch(APIurl, options);
    var resultData = JSON.parse(result.getContentText()).data;
    return resultData
}

/**
 * @description
 * @param {*} settingsObj
 * @param {*} rowValues
 * @returns
 */
function processTags(settingsObj, rowValues) {
    try {
        var tagArray = JSON.parse(settingsObj.customTags)
        console.log("Processing " + tagArray.length + " tags")
        runLog("Processing " + tagArray.length + " tags")
        var tags = []
        tagArray.forEach(function(each) {
            var value = rowValues[each.columnName]
            Logger.log("tag: " + JSON.stringify(each))



            if (each.id === "" || each.id === null || each.id === undefined) {

                if (each.name === "" || each.name === null || each.name === undefined) {

                    return

                } 
                else {
                    Logger.log("==================")
                    var tagsArray = callAsanaApi("GET", "tags")
                    var found = search(each.name, tagsArray)
                    
                    if (found) {
                        each.id = found.id
                        runLog("Tag ID for " + each.name + " is " + each.id)
                    } 
                    else {
                        each.id = ""
                        runLog("No ID Found \n Creating tag: " + each.name)
                        var payload = {}
                        payload.data = {}
                        payload.data.name = each.name
                        payload.data.workspace = WORKSPACE_ID
                        var tagJSON = callAsanaApi("POST", "tags", JSON.stringify(payload))
                        // Logger.log(JSON.stringify(tagJSON))
                        each.id = tagJSON.id
                    }
                }
            }
            if (value !== "" && value !== undefined) {
                if (each.columnValue === undefined || each.columnValue === null || each.columnValue === value) {
                    //            if (each.columnName === undefined || each.columnName === null || each.columnName ===

                    tags.push(each.id)
                }
            }
        })
        return tags
    } catch (e) {
        errorLog(e)
    }
}

// Takes an object and creates subtasks of main task. If provided columnName and columnValue parameters, tasks can be selectively created.
/**
 * @description
 * @param {*} settingsObj
 * @param {*} rowValues
 * @param {*} taskResults
 */
function processChildren(settingsObj, rowValues, taskResults) {
    try {
        // todo: add due date processing
        var children = JSON.parse(settingsObj.customChildren).reverse()
        runLog("Processing " + children.length + " children")
        children.forEach(function(each) {
            var value = rowValues[each.columnName]
                //  Logger.log(value)
            //      child = {}
            each.parent = taskResults.result["id"].toString()
            //      each.name = each.name
            //      child.assignee = each.assignee
            if (value !== "") {
                if (each.columnValue === undefined || each.columnValue === null || each.columnValue === value) {
                    //            if (each.columnName === undefined || each.columnName === null || each.columnName ===
                    if (each.dueDateDuration === "" || (each.dueDateDuration === undefined || each.dueDateDuration ===
                        null)) {
                            return
                        } else {
                            Logger.log(each.dueTime)
                            if ( each.dueTime === "" || each.dueTime === null ||  each.dueTime === undefined) {
                                each.due_on = dueDate(each.dueDateDuration)
                            } 
                            else {
                        
                                each.due_at = dueTime(each.dueDateDuration, each.dueTime)        
                            }
                        }
                    createAsanaTask(settingsObj, each)
                }
            }
        })
    } catch (e) {
        errorLog(e)
    }
}

/**
 * @description
 * @param {*} taskId
 * @returns
 */
function asanaCommentURI(taskId) {
    var method = "POST"
    var endpoint = "tasks/" + taskId + "/stories"
    var payload
    var options = {
        method: method,
        headers: {
            Authorization: "Bearer " + PERSONAL_ACCESS_TOKEN
        }
    };
    if (payload) {
        options.payload = payload
    }
    if (method === "POST") {
        options.headers["Accept"] = "application/json"
        options.headers["Content-Type"] = "application/json"
    }
    var APIurl = "https://app.asana.com/api/1.0/" + endpoint;
    runLog("Calling " + APIurl + "\n" + JSON.stringify(options))
    var result = UrlFetchApp.fetch(APIurl, options);
    var resultData = JSON.parse(result.getContentText()).data;
    return resultData
}