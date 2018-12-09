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
        name: 'New StARS Staff Form',
        noReply: true,
        to: getHREmail(),
        // cc: '...',
        subject: 'New staff form submitted!',
        body: 'New staff form just filled out by: ' + formFilledBy +
            '\n' +
            '\nReason for filling out form: ' + formReason +
            '\n' +
            '\nRemember to ask this person for an updated role description!',
    };

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