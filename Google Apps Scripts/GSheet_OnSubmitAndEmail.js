// ============================ DEBUGGING NOTES =============================
// NOTE [from 11 Feb 2018]: If form submissions stop getting into
// -> new tab, try re-creating / re-running createFormFormSubmitTrigger()

// ============================== CONSTANTS ================================
function getNewFormSheetName()       { return 'New Staff' }
function getFormReasonQuestion()     { return 'Reason for filling out form' }
function getSubmitterEmailAddress()  { return 'Email Address' }

function getHREmail()         { return 'hr@stars-egypt.org' }
function getVolunteerEmail()  { return 'volunteer@stars-egypt.org' }

// ======================== ONE-TIME TRIGGER SETUP =========================
// set up trigger that listens for form submission
// Note: only needs to be set up once
// -> to check if trigger is set up, visit Edit -> Current Project's Triggers in Script editor page
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
    // Note: unused props: authMode, values, source, triggerUid

    // get source's sheet name
    const sourceSheetName = range.getSheet().getName().trim().toUpperCase();

    // if source sheet name doesn't match new form sheet name, quit
    if (sourceSheetName !== getNewFormSheetName().toUpperCase()) return;

    // get data from namedValues
    const formReason = namedValues[ getFormReasonQuestion() ];
    const formFilledBy = namedValues[ getSubmitterEmailAddress() ];

    const email = {
        // name: 'New StARS Staff Form', // (not used if noReply is true)
        noReply: true,
        to: getHREmail(),
        // cc: '...', // conditionally set below
        subject: 'New staff form submitted!',
        // htmlBody: '...', // conditionally set below
    };
  
    // if promotion / transferred, send extra body text
    if (formReason === 'Staff member being promoted OR transferring between programs') {
        email['htmlBody'] = emailBodyIntro(formFilledBy, formReason) + programEmailBody() + HRemailBody();
    } else {
        email['htmlBody'] = emailBodyIntro(formFilledBy, formReason) + HRemailBody();
    }

    // if volunteer, cc volunteer officer
    if (formReason === 'Volunteer') {
        email['cc'] = getVolunteerEmail() + ',abeaman@stars-egypt.org';
    } else {
        email['cc'] = 'abeaman@stars-egypt.org' ;
    }

    // aaand send!
    MailApp.sendEmail(email);

    Logger.log('done notifying HR of form filled');
}

// ================================ Email Templates ================================
function emailBodyIntro(formFilledBy, formReason) {
    return 'New staff form just filled out by: ' + formFilledBy
        + '<br><br>Reason for filling out form: ' + formReason
        + '<br><br>';
}
function programEmailBody() {
    return '<strong>Reminder for program promoting / transferring staff member:</strong>'
        + '<br><ul>'
            + '<li>Job Description</li>'
            + '<li>Start Date of new position</li>'
            + '<li>Fill <a href="https://drive.google.com/open?id=1heYm2qy4Gog_usQbqRLmgnwA5ANcRg3Y">promotion / demotion form</a>'
        + '</ul>';
}
function HRemailBody() {
    return '<strong>Reminder for HR staff: Follow this checklist:</strong>'
        + '<br><ul>'
            + '<li>Prepare Contract to be signed</li>'
            + '<li>Check references</li>'
            + "<li>Start employee personnel file (including ID's)</li>"
        + '</ul>'
        + '<i>P.S. Email "it.team@stars-egypt.org" if you would like to edit this message :)</i>';
}
function testEmail() {
    MailApp.sendEmail({
        noReply: true,
        to: 'abeaman@stars-egypt.org',
        subject: 'New staff form submitted!',
        htmlBody: emailBodyIntro('bob', 'cuz i want to') + programEmailBody() + HRemailBody(),
    });
}