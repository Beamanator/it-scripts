// -> Script used in StARS Staff Departure Form
// -> Form name changed to StARS - New / Promoted / Transferred Staff Members

// ============================ DEBUGGING NOTES =============================
// NOTE [from 11 Feb 2018]: If form submissions stop getting into
// -> new tab, try re-creating / re-running createFormFormSubmitTrigger()

// ============================== CONSTANTS ================================
function getNewFormSheetName()       { return 'Form Responses 1' }
// function getFormReasonQuestion()     { return 'Reason for filling out form' }
function getSubmitterEmailAddress()  { return 'Email Address' }

function getHREmail()         { return 'hr@stars-egypt.org' }

// ======================== ONE-TIME TRIGGER SETUP =========================
// set up trigger that listens for form submission
// Note: only needs to be set up once
// -> to check if trigger is set up, visit:
// -> Edit => Current Project's Triggers in Script editor page
function createFormFormSubmitTrigger() {
    var ss = SpreadsheetApp.getActive();
    ScriptApp.newTrigger('submit')
        .forSpreadsheet(ss).onFormSubmit().create();
}

// =========================== TRIGGER RESPONSE =============================
// callback for form submission.
function submit(eventObj) {
    // grab data from triggerObj
    var namedValues = eventObj.namedValues;
    var range = eventObj.range;
    // Unused props: authMode, values, source, triggerUid

    // get source's sheet name
    const sourceSheetName = range.getSheet().getName().trim().toUpperCase();

    // if source sheet name doesn't match new form sheet name, quit
    if (sourceSheetName !== getNewFormSheetName().toUpperCase()) return;

    // get data from namedValues
    // Note: have to use [0] to get the string value out of an array
    const formFilledBy = namedValues[ getSubmitterEmailAddress() ][0];
    const auditData = {
        name: namedValues[ 'Staff Name' ][0],
        position: namedValues[ 'Position' ][0],
        startDate: namedValues[ 'Staff Start Date' ][0],
        probationPassedDate: namedValues[ 'Probation End Date' ][0],
        endDate: namedValues[ 'Last Day at StARS' ][0],
    };

    // create basic email template
    const email = {
        // name: 'New StARS Staff Form', // (not used if noReply is true)
        noReply: true,
        to: formFilledBy,
        cc: getHREmail(),
        subject: 'Staff departure form submitted',
        htmlBody: emailBodyIntro(formFilledBy) +
            programEmailBody() +
            auditEmailBody(auditData) +
            HRemailBody(),
    };

    // aaand send!
    MailApp.sendEmail(email);
    Logger.log('Done notifying HR of form filled');
}

// ================================ Email Templates ================================
function emailBodyIntro(formFilledBy) {
    return 'Staff Departure Form just filled out by: ' + formFilledBy
        + '<br><br>';
}
function programEmailBody() {
    return '<strong>Reminders for staff member who filled out this form:</strong>'
        + '<br><ul>'
            + '<li>Complete final pay</li>'
            + '<li>Return IT equipment</li>'
        + '</ul>';
}
function auditEmailBody(data) {
    return 'For auditing purposes, make sure the following information that '
        + 'was filled in the Staff Departure Form is accurate:'
        + '<br><ul>'
            + '<li>Name of Staff: ' + data.name + '</li>'
            + '<li>Position: ' + data.position + '</li>'
            + '<li>Start Date: ' + data.startDate + '</li>'
            + '<li>Date probation passed: ' + data.probationPassedDate + '</li>'
            + '<li>Last Day: ' + data.endDate + '</li>'
        + '</ul>';
}
function HRemailBody() {
    return '<strong>Reminder for HR / Finance staff:</strong>'
        + '<br><ul>'
            + '<li>Check time-sheet & annual leave</li>'
            + '<li>Complete final pay</li>'
            + "<li>Record information in employee personnel file</li>"
        + '</ul>'
        + '<i>P.S. Email "it.team@stars-egypt.org" if you would like to edit this message</i>';
}