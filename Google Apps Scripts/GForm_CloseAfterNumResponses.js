// Note: this script needs to be 'activated' each time a submission is being made.
// -> a.k.a. whenever the number needs to be reset.
// -> In the script editor, under Resources, select the script to run On form submit.
// from: https://webapps.stackexchange.com/questions/56537/setting-maximum-number-of-sign-ups-in-google-forms
function closeForm() {
    // get active form
    var form = FormApp.getActiveForm();
  
    // retrieve number of responses thusfar
    var responses = form.getResponses().length;
  
    // set close message
    var msg = "Maximum number of respondents has been reached";
  
    // set max 
    var maxResponse = 22;
  
    // do the math
    if(responses >= maxResponse) {
        form.setAcceptingResponses(false).setCustomClosedFormMessage(msg);
    }
}