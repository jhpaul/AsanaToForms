var Email = {}


/**
 * @description Creates and sends an email containing entry data
 * @param {*} settingsObj
 * @param {*} rowValuesArray
 * @param {*} formTable
 * @param {*} body
 * @param {*} attachments
 */

Email.build = function (settingsObj, rowValuesArray, formTable, body, attachments) {
    var subjectName = merge(settingsObj["emailSubject"], rowValuesArray, formTable.headerArray) //, dateFormat, dateTimeZone
    var subHead = merge(settingsObj["emailSubHead"], rowValuesArray, formTable.headerArray) //, dateFormat, dateTimeZone
    var to = emailSplitJoin(settingsObj.emailTo, ",", ";")
    var body = body
    var cc = emailSplitJoin(settingsObj.emailCc, ",", ";")
    var bcc = emailSplitJoin(settingsObj.emailBcc, ",", ";")
    var from = settingsObj["emailFrom"]
    Email.send(to, subjectName, body, subHead, cc, bcc, from, attachments, settingsObj["emailWebhook"])

}


/**
 * @description generates an email via the microsoft flow webhook api
 * @param {*} to
 * @param {*} subject
 * @param {*} body
 * @param {*} subhead
 * @param {*} cc
 * @param {*} bcc
 * @param {*} from
 * @param {*} attachments
 * @param {*} webhook
 */
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
    return results
}