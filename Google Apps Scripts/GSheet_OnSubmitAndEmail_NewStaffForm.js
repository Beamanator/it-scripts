// -> Script used in StARS New Staff Form
// -> Form name changed to StARS - New / Promoted / Transferred Staff Members

// ============================ DEBUGGING NOTES =============================
// NOTE [from 11 Feb 2018]: If form submissions stop getting into
// -> new tab, try re-creating / re-running createFormFormSubmitTrigger()

// ============================== CONSTANTS ================================
function getNewFormSheetName()       { return 'New Staff' }
function getFormReasonQuestion()     { return 'Reason for filling out form' }
function getSubmitterEmailAddress()  { return 'Email Address' }

function getHREmail()         { return 'hr@stars-egypt.org' }
function getVolunteerEmail()  { return 'volunteer@stars-egypt.org' }

function getPromoteDeleteFormURL() { return 'https://drive.google.com/open?id=1heYm2qy4Gog_usQbqRLmgnwA5ANcRg3Y' }

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
    const formReason = namedValues[ getFormReasonQuestion() ][0];
    const formFilledBy = namedValues[ getSubmitterEmailAddress() ][0];
    var error = false;

    // create basic email template
    const email = {
        // name: 'New StARS Staff Form', // (not used if noReply is true)
        noReply: true,
        to: getHREmail(),
        // cc: null, // default is empty (may change below)
        subject: 'New staff form submitted!',
        htmlBody: emailBodyIntro(formFilledBy, formReason) + HRemailBody(), // default (may change below)
    };

    // change email based on form reason
    switch (formReason) {
        case 'New staff':
            Logger.log('new staff');
            break;
        case 'Interpreter':
            Logger.log('interpreter');
            break;

        // handle special cases below...
        case 'Staff member being promoted OR transferring between programs':
            Logger.log('promoted / transferred');
            email['to'] = formFilledBy;
            email['cc'] = getHREmail();
            email['htmlBody'] = emailBodyIntro(formFilledBy, formReason) + programEmailBody()
                + HRemailBody();
            break;
            
        case 'Volunteer':
            Logger.log('volunteer');
            email['cc'] = getVolunteerEmail();
            break;
          
        // default (shouldn't get here unless we're not handling cases properly!)
        default:
            Logger.log('Error: Form reason not matched in switch statement! Fix!');
            Logger.log('"Form reason" that doesn\'t match: ' + formReason);
            error = true;
    }

    // aaand send!
    if (!error) {
        MailApp.sendEmail(email);
        Logger.log('Done notifying HR of form filled');
    }
}

// ================================ Email Templates ================================
function emailBodyIntro(formFilledBy, formReason) {
    return 'New staff form just filled out by: ' + formFilledBy
        + '<br><br>Reason for filling out form: ' + formReason
        + '<br><br>';
}
function programEmailBody() {
    return '<strong>Reminder for Program that is promoting / transferring staff member:</strong>'
        + ' Please send the following to HR:'
        + '<br><ul>'
            + '<li>An updated job description</li>'
            + '<li>Start Date of new position</li>'
            + '<li>Fill out the <a href="' + getPromoteDeleteFormURL() + '">Staff Promotion / Demotion Form</a>'
        + '</ul>';
}
function HRemailBody() {
    return '<strong>Reminder for HR staff:</strong> Follow this checklist:'
        + '<br><ul>'
            + '<li>Prepare Contract to be signed</li>'
            + '<li>Check references</li>'
            + "<li>Start employee personnel file (including ID's)</li>"
        + '</ul>'
        + '<i>P.S. Email "it.team@stars-egypt.org" if you would like to edit this message :)</i>';
}