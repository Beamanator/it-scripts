// ============================ DEBUGGING NOTES =============================
// remember, the 'onSubmit' function needs to be 'activated' each time a
// -> submission is being made. In the script editor, under Edit, then under
// -> Current Project's Triggers, select the script to run on Form Submit.
// from: https://webapps.stackexchange.com/questions/56537/setting-maximum-number-of-sign-ups-in-google-forms
// Update from 24 Sept 2019:
// -> Note: May need to press Play button "Run" to give script permissions to run in your form

// ============================== CONSTANTS ================================
var QUESTION_TITLE = "<YOUR QUESTION TITLE>"; // must match exactly!
var EMAIL_MAP = {
	// note: keys should be in ALL LOWERCASE!! (prevent mix-case errors)
	"<PERSON 1 NAME>": "<PERSON 1 EMAIL ADDRESS>",
	"<PERSON 2 NAME>": "<PERSON 2 EMAIL ADDRESS>",
	"<PERSON 3 NAME>": "<PERSON 3 EMAIL ADDRESS>",
	cc: "<EMAIL ADDRESS TO CC>",
};

var EMAIL_SUBJECT = "<YOUR EMAIL SUBJECT>";
var EMAIL_BODY = "<YOUR EMAIL BODY>";

// =========================== TRIGGER RESPONSE =============================
// function triggered on form submission:
function onSubmit(eventObj) {
	// get active form
	var form = FormApp.getActiveForm();

	// get all multiple-choice items in form (assuming question we want is radio buttons!)
	// https://developers.google.com/apps-script/reference/forms/item-type.html
	var formFields = form.getItems(FormApp.ItemType.MULTIPLE_CHOICE);

	// temporary vars
	var targetField = null;

	// loop through multiple-choice to find fields with question title we want
	for (var i = 0; i < formFields.length; i++) {
		var thisField = formFields[i];
		var thisTitle = thisField.getTitle();

		// check if title of current question matches target question title
		if (thisTitle === QUESTION_TITLE) {
			targetField = thisField;
		}
	}

	// if this field doesn't exist, throw error and quit
	if (targetField === null) {
		Logger.log("Couldn't find question: " + QUESTION_TITLE + "! Quitting.");
		return;
	}
	// else, success! move forward.

	// grab data we care about from event Obj
	// -> disregard eventObj's authMode, source, triggerUid
	var formResponse = eventObj.response;

	// get response object for target field (if not answered, returns null)
	var responseObj = formResponse.getResponseForItem(targetField);

	// if the field wasn't answered, throw error (should be required, right??)
	if (responseObj === null) {
		Logger.log(
			'Response to "' + QUESTION_TITLE + '" is empty :( Quitting.'
		);
		return;
	}
	// else, success! move forward.

	var responseText = responseObj
		.getResponse()
		.toString()
		.toLowerCase();

	// get map response text to email address?
	var targetPersonEmail = EMAIL_MAP[responseText];
	var ccEmail = EMAIL_MAP["cc"];

	// if emails not mapped correctly, throw error and quit
	if (!targetPersonEmail) {
		Logger.log(
			'Email address for "' + responseText + '" not found :( Quitting.'
		);
		return;
	}
	// else, success! send email!

	// create basic email template
	const email = {
		// name: 'New StARS Staff Form', // (not used if noReply is true)
		noReply: true,
		to: targetPersonEmail,
		// cc: ccEmail,
		subject: EMAIL_SUBJECT,
		htmlBody: EMAIL_BODY,
	};

	// add email address to cc, if mapped
	if (ccEmail) {
		email["cc"] = ccEmail;
	}

	// aaand send!
	MailApp.sendEmail(email);
}
