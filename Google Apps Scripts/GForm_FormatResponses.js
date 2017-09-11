// Created 15-Aug-2017

// ============================== CONSTANTS ================================
function getPendingOnCallInquiriesSheetName()
    { return 'RLAP RSD Team On-Call Inquiries (Pending)'; }

function F_RIPSAIDate() {   return 'dd-MMM-yyyy';   }

function getColumnsToFormat() {
    return [
        { name: 'Timestamp',                     format: F_RIPSAIDate() },
        { name: 'Date of Birth\n(MM/DD/Year)',   format: F_RIPSAIDate() }
    ];
}

// set up trigger that listens for form submission
// only needs to be set up once
// -> to check if trigger is set up, visit Edit -> Current Project's Triggers in Script editor page
function createFormFormSubmitTrigger() {
    var ss = SpreadsheetApp.getActive();
    ScriptApp.newTrigger('submit').forSpreadsheet(ss).onFormSubmit().create();
}

// =========================== TRIGGER RESPONSE =============================

/**
 * Function is called whenever a form attached to the spreadsheet is submitted.
 * The function will decide how to handle the form submit based on where the
 * data was sent.
 * 
 * @param {object} eventObj - information about each form submission
 */
function submit(eventObj) {
    // grab data from eventObj
    var authMode = eventObj.authMode;
    var values = eventObj.values;
    var namedValues = eventObj.namedValues; // -> Object - key: question names, value: data from form for that question
    var range = eventObj.range;
    var source = eventObj.source;
    var triggerUid = eventObj.triggerUid;

    // get sheet name where data (from form submission) was stored
    var sourceSheetName = range.getSheet().getName();
    // can also use: var sourceSheetName = source.getActiveSheet().getName();

    var pendingOnCallInquiriesSheetName = getPendingOnCallInquiriesSheetName()
        .toUpperCase();

    // Figure out what function to call based on sheet name
    switch(sourceSheetName.toUpperCase()) {
        case pendingOnCallInquiriesSheetName:
            Format_PendingOnCallInquiries( eventObj );
            break;

        default:
            Logger.log('Sheet not found! Source sheet:');
            Logger.log( sourceSheetName );
    }
}

/**
 * Function formats form submissions to PendingOnCallInquiries sheet
 * eventObj.namedValues looks like this:
    {
        Summary of Issue    =   [test],
        Date of UNHCR Card Expiry = [],
        UNHCR File Status   =   [Blue card],
        Languages           =   [Arabic],
        UNHCR Number        =   [test],
        Gender              =   [Female],
        Date UNHCR Card Issued = [],
        Timestamp           =   [8/15/2017 21:26:47],  // Formatting Needed
        Nationality         =   [Eritrean],
        Applicant Name      =   [test],
        Name of On-Call Worker = [test],
        Date of Birth\n(MM/DD/Year) = [12/31/2017],
        Case Size           =   [test],
        Type of Inquiry     =   [RSD (other than waiting for results or rescheduling)],
        Phone Number        =   [test]
    }
    -> access like normal object (namedValues['Timestamp'] or namedValues.Timestamp)
 * @param {object} eventObj - information about specific form submission
 */
function Format_PendingOnCallInquiries(eventObj) {
    // grab needed data from eventObj
    var values = eventObj.values; // -> in same order as delivered to sheet
    var namedValues = eventObj.namedValues; // -> in whacky order -> Object - key: question names, value: data from form for that question
    var range = eventObj.range;
    var source = eventObj.source;

    // Next:
    // 1) Loop through column names that need formatting
    // 2) Get formatting needed for column
    // 3) Get column range from columnName
    // 4) Format the data
    // 5) set format for column
    var columnsToFormat = getColumnsToFormat();

    // 1) Loop through columns that need formatting
    for (var i = 0; i < columnsToFormat.length; i++) {
        var columnName = columnsToFormat[i].name;
        
        // 2) get formatting for column
        var format = columnsToFormat[i].format;

        // 3) get column range by columnName (B2:B is full col B except header row)
        var colRange = getColumnRange( columnName, range );

        if (colRange === '') {
            Logger.log('error in column:');
            Logger.log(columnName);
            continue;
        }

        // 4) format the data
        var column = range.getSheet().getRange(colRange);
 
        // 5) set specified date format on column
        column.setNumberFormat(format);
    };
}

/**
 * Function gets column letter based on specific columnName
 * 
 * @param {string} columnName - name of header column to look for in spreadsheet
 * @param {object} range - range of source spreadsheet
 * @returns - column Letter corresponding to given columnName
 */
function getColumnRange( columnName, range ) {
    var colRange = '';
    var colLetter = '';

    // get header range (starting at row 1, col 1, numRows = 1, numCols = getLastColumn())
    var headerRange = range.getSheet().getRange(1, 1, 1, range.getNumColumns() );

    var headerArray = headerRange.getValues()[0];

    // try to match up columnName with a column in the header
    for (var i = 0; i < headerArray.length; i++) {
        var headerText = headerArray[i];

        // if header text matches column name, calculate column letter
        if (headerText.toUpperCase() === columnName.toUpperCase()) {
            colLetter = columnToLetter(i);
            break;
        }
    }

    if (colLetter === '') {
        Logger.log('oh no! column not found:');
        Logger.log(columnName);
        return '';
    }

    // return full column range (ex: B2:B is full column range)
    return colLetter + '2:' + colLetter;
}

// from: https://stackoverflow.com/questions/21229180/convert-column-index-into-corresponding-column-letter
/**
 * Function takes a column number (1) and turns into a letter (A)
 * 
 * @param {number} column - column number
 * @returns - letter corresponding to column number
 */
function columnToLetter(column) {
    var temp, letter = '';
    while (column > 0)
    {
        temp = (column) % 26;
        letter = String.fromCharCode(temp + 65) + letter;
        column = (column - temp) / 26;
    }
    return letter;
}