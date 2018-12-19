/**
 * Created: 11 Sept 2017
 * Latest updates:
 *** 28 Nov 2018 - Rearranged some columns, so updated getEmailColumnLetters() & getEmailSentColumnLetter().
 *** 19 Dec 2018 - Transitioning RIPS guys (adding Abdullah)
 * Source = https://github.com/Beamanator/stars-scripts/blob/master/Google%20Apps%20Scripts/GSheet_SendEmailFromMenu.js
 * 
 * Purpose = Send email to RIPS email account with new user information, to make things nice and simple.
 * 
 * NOTE: Make sure you set up your onOpen trigger before expecting the code to run!
 * Edit -> Current project's triggers -> 
 * 
 * After menu item is clicked, email list to RIPS email account
 */ 

function getEmailColumnLetters() {
  return {
    // required columns:
    'N': 'New StARS email:',
    // 'Q': 'RIPS account needed?',
    'E': 'Program of new staff:',
    'S': 'Caseworker needed?',
    
    // optional columns:
    'D': 'Name of new staff:',
    'R': 'Temporary password',
    'B': 'New user request from:'
  };
}

function getEmailRIPS()   { return 'RIPS@stars-egypt.org';  }
function getEmailBackup() { return 'asadaqa@stars-egypt.org'; }
function getEmailSentColumnLetter()   {  return 'T';  } // actual col #, not index
function getEmailSubject()   {  return 'New RIPS Account - Login Details';  }
/*
Hello <NAME>,

You have a new RIPS account!

RIPS is the main system StARS uses to track the work we do with individual clients. During your time at StARS, you will learn how to access that information, and you may even learn how to add new information to the system. Your coordinator should schedule a training for you soon.

Below is the login information for your new RIPS account! Get excited!

Here is the link to RIPS: http://rips.247lib.com/Stars/
-    Your username is: The first part of your email address before the "@" symbol.
-    Your password is: <PASSWORD>

Please make sure that you complete the following steps after receiving this email:
1) Create a new password for your RIPS account. Here are instructions: http://bit.ly/2Bn7rcK
2) Install the RIPS validation extension. Here are the installation instructions: http://bit.ly/2x54v4T

Let me know if you have any questions!

Thanks,

Abdullah & Alex "The RIPS Guys"';
 */
function getTemplateEmailText() {
  return 'Hello <NAME>,' +
    '\n' +
    '\nYou have a new RIPS account!' +
    '\n' +
    '\nRIPS is the main system StARS uses to track the work we do with individual clients. ' +
      'During your time at StARS, you will learn how to access that information, ' +
      'and you may even learn how to add new information to the system. ' +
      'Your coordinator should schedule a training for you soon.' +
    '\n' +
    '\nBelow is the login information for your new RIPS account. Get excited!' +
    '\n' +
    '\nHere is the link to RIPS: http://rips.247lib.com/Stars/' +
    '\n-\tYour username is: The first part of your email address before the "@" symbol.' +
    '\n-\tYour password is: <PASSWORD>' +
    '\n' +
    '\nPlease make sure that you complete the following steps after receiving this email:' +
    '\n1) Create a new password for your RIPS account. Here is a step-by-step guide: http://bit.ly/2Bn7rcK' + 
    '\n2) Install the RIPS validation extension. Here are the installation instructions: http://bit.ly/2x54v4T' +
    '\n' +
    '\nLet me know if you have any questions!' +
    '\n' +
    '\nThanks,' +
    '\n' +
    '\nAbdullay & Alex "The RIPS Guys"';
}

// Add a custom menu to the active spreadsheet, including a separator and a sub-menu.
function onOpen(e) {
  SpreadsheetApp.getUi()
    .createMenu('RIPS')
    .addItem('Send new RIPS users', 'sendUsers')
    .addToUi();
}

/**
 * Main function called by menu item click. Goal:
 * 1) Read column "Email Sent?" and find all rows w/out 'Y' or 'Yes'
 * 2) For each row, get data from columns in getEmailColumnLetters()
 * 3) Send email to RIPS email account
 * 
 */
function sendUsers() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var dataRange = sheet.getDataRange();

  // Fetch values & backgrounds in data range (note: gets 2D array)
  var data = dataRange.getValues();

  // get array of row numbers that have staff info that will be sent in email
  var rowIndexArray = getRowIndicesToSend(data);

  if (rowIndexArray.length > 0) {
    // send emails
    emailStaffData(rowIndexArray, data);

    // set rows to 'Yes' now that email has been sent.
    setEmailSentColToYes(rowIndexArray, sheet);
  }

  // else, do nothing since no rows should be sent!
}

/**
 * Function loops through email sent column of data and finds all columns
 * without 'Y' or 'Yes' and sets these up for an email to be sent
 * 
 * @param {object} data - array of data in spreadsheet
 * @returns - array of row numbers that need the email sent
 */
function getRowIndicesToSend(data) {
  // get email sent column number
  var emailSentColumnNumber = getColNumbersFromLetters( getEmailSentColumnLetter() );

  var rowIndexHolder = [];

  // loop through rows of data, pulling out 'email sent' values
  for (var row = 1; row < data.length; row++) {
    var val = data[row][emailSentColumnNumber - 1];

    if (val !== 'Y' && val !== 'Yes' && val !== 'N/A') {
      Logger.log('Actually: <' + val + '>');
      rowIndexHolder.push(row);
    }
  }

  return rowIndexHolder;
}

/**
 * Function converts array of column letters to column numbers.
 * Ex: [A,B,C] -> [1,2,3]
 * 
 * If passed-in 'cols' is only one letter, return one number (not in array form)
 * 
 * returns col NUMBER not INDEX
 */
function getColNumbersFromLetters(cols) {
  var objFlag = true;
  
  // set flag stating passed-in variable is not an array
  if (typeof(cols) !== 'object') {
    objFlag = false;
    cols = [cols];
  }

  var colNums = [];
  
  for (var i = 0; i < cols.length; i++) {
    var letter = cols[i].toUpperCase();
    
    if (letter.length !== 1) {
      Logger.log('Skipping letter in Column #s array - not only one char:');
      Logger.log(letter);
      continue;
    }
    
    // char code of 'A' is 65, so remove 64 to get column number
    colNums.push(letter.charCodeAt(0) - 64);
  }

  // if we passed in an array, return an array. otherwise, return the only number
  if (objFlag) {
    return colNums;
  } else if (!objFlag && colNums.length === 1) {
    return colNums[0];
  } else {
    Logger.log('Some error in getColNumbersFromLetters function');
  }
}

/**
 * Function sends email w/ formatted list of client data
 */
function emailStaffData(rowIndexArray, data) {
  var message = 'New staff members to email:';
  var emailRIPS = getEmailRIPS();
  var emailBackup = getEmailBackup();
  var emailColumnLetters = getEmailColumnLetters();

  // output headers
  message += '\nRow #s:\t';
  Object.keys(emailColumnLetters).forEach(function(key) {
    var colDesc = emailColumnLetters[key];
    
    message += '\t' + colDesc + '\t|';
  });
  
  // output staff data
  for (var i = 0; i < rowIndexArray.length; i++) {
    var rowIndex = rowIndexArray[i];

    message += '\nRow #' + (rowIndex + 1) + ': ';

    // loop through column letters that have needed data
    Object.keys(emailColumnLetters).forEach(function(key_colLetter, index) {
      // get column number from column letter:
      var colNumber = getColNumbersFromLetters(key_colLetter);

      // get data from data array:
      var colData = data[ rowIndex ][ colNumber - 1 ];

      // add new row of text to message:
      message += '\t' + colData + '\t|';
    });
  }

  // add some spacing
  message += '\n\n';

  // add template email text
  message += getTemplateEmailText();
  
  // send email to RIPS guy / gal + backup
  MailApp.sendEmail(emailRIPS, getEmailSubject(), message);
  MailApp.sendEmail(emailBackup, getEmailSubject(), message);
}

/**
 * Function sets 'email sent' column to 'Yes' values if email was sent
 * for this row
 * 
 * @param {object} rowIndexArray - array of row indices to set to 'Yes'
 * @param {object} sheet - spreadsheet sheet object
 */
function setEmailSentColToYes(rowIndexArray, sheet) {
  var emailSentColumnLetter = getEmailSentColumnLetter();

  for (var i = 0; i < rowIndexArray.length; i++) {
    var cellRange = emailSentColumnLetter + (rowIndexArray[i] + 1);

    // get cell of 'email sent column'
    var cell = sheet.getRange( cellRange );

    // set cell to 'Yes'
    cell.setValue('Yes');
  }
}
