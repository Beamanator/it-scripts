// Last updated Early July 2017

// ============================== DEBUGGING ================================
// NOTE [from 11 Feb 2018]: If form submissions stop getting into new tab, try re-creating / re-running createFormFormSubmitTrigger()

// ============================== CONSTANTS ================================
function getSourceSheetName()       { return 'RSD Workshop New entries'; }
function getTargetSheetName()       { return "FI Workshop";              }
function getTargetRowMethod()       { return "DESC";                     } // options: "DESC" OR "ASC"
function getTargetRowStart()        { return 3;                          } // starting row where responses are stored
function getFirstColumnLetter()     { return "B";                        } // must be capital letters!!!

function getOrangeRowColor()        { return '#ff9900';                  } // last row where script should look for adding data - if this color is hit, add cell above and past data there
function getGreenRowColor()         { return '#00ff00';                  } // same as orange, but green (don't add data after this)

function UYBP_getFlagColor()        { return "#00ffff";                  } // Light blue
function UYBP_getFormQuestionName() { return 'Date of Birth';            } // this question name stores DOB
function UYBP_getColumnLetToFlag()  { return "C";                        } // this column will be flagged if potential UYBP

// =========================== MAIN TRIGGER SETUP =============================
// set up trigger that listens for form submission
// only needs to be set up once
// -> to check if trigger is set up, visit Edit -> Current Project's Triggers in Script editor page
function createFormFormSubmitTrigger() {
  var ss = SpreadsheetApp.getActive();
  ScriptApp.newTrigger('submit').forSpreadsheet(ss).onFormSubmit().create();
}

// =========================== TRIGGER RESPONSE =============================

// callback for form submission.
function submit(eventObj) {
  // grab data from triggerObj
  var authMode = eventObj.authMode;
  var values = eventObj.values;
  var namedValues = eventObj.namedValues; // -> Object - key: question names, value: data from form for that question
  var range = eventObj.range;
  var source = eventObj.source;
  var triggerUid = eventObj.triggerUid;
  
  // get source sheet name (from form submission sheet)
  var sourceSheetName = range.getSheet().getName();
  var givenSourceSheetName = getSourceSheetName();
 
  // another way to get source's sheet name: source.getActiveSheet().getName()
  
  // if the source sheet (from form response sheet) is not the same as the specified form response sheet, quit (return) so extra forms don't trigger the code
  if (sourceSheetName.toUpperCase() !== givenSourceSheetName.toUpperCase() ) {
    Logger.log('A different form just submitted data - not added to specified sheet!');
    Logger.log('Specified sheet (that should receive your form data): ' + givenSourceSheetName);
    return;
  }
  
  // get spreadsheet to add new responses to:
  var responseSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName( getTargetSheetName() );
  
  // get target row number. This is the row where the client data will be pasted.
  var targetRow = getTargetRow(responseSheet);
  
  // if targetRow is -1, there was an error in getTargetRow.
  if (targetRow === -1) {
    Logger.log('error getting targetRow. Check other logs');
    return;
  }
  
  // setup range where response data will be shoved into:
  var targetRange = getFirstColumnLetter() + targetRow + ':' + getLastColumnLetter(values) + targetRow;
  
  // insert row before position targetRow:
  responseSheet.insertRowBefore(targetRow);
  
  // New: insert row after position targetRow -1 (to avoid same colors)
//  responseSheet.insertRowAfter(targetRow - 1);
  
  // sets client values in target Sheet (NOTE: setValues takes 2-D array!!!)
  var cellRng = responseSheet.getRange(targetRange);
  cellRng.setValues([values]);
  
  // clear color from target range
  cellRng.setBackground('white');
  
  // check if client is <= 18 years old. If yes, maybe should be UYBP -> Color the name cell.
  // Logger.log(namedValues['Date of Birth']);
  var age = calculateClientAge(namedValues);
  if (age !== undefined && age !== '' && age <= 18) colorCell(responseSheet, (UYBP_getColumnLetToFlag() + targetRow), UYBP_getFlagColor());
}

// =========================== OTHER FUNCTIONS =============================

// colors a target cell a specific color:
function colorCell(sheet, cellRange, color) {
  Logger.log('inside color cell function');
  if (!sheet || !cellRange || !color) {
    Logger.log('error in colorCell function? some value is undefined. here they are:');
    Logger.log(sheet);
    Logger.log(cellRange);
    Logger.log(color);
    return;
  }
  
  // create cell obj
  var cell = sheet.getRange(cellRange);
  
  // color it with specified color!
  cell.setBackground( color );
}

// calculates and returns the age of the client
// -> Note: age only based off years, always calculate from 1/1 of 18 years ago 
function calculateClientAge(valuesObj) {
  // get date of birth of current client from form questions.
  var dobValueArr = valuesObj[ UYBP_getFormQuestionName() ];
  
  if (dobValueArr.length > 1) {
    Logger.log('why is the DOB array longer than one??');
    Logger.log(dobValueArr);
    return;
  }
  var dobString = dobValueArr[0];
  
  // get current year from current date
  var thisYear = (new Date()).getFullYear();
  
  // get client birth year from dobString
  // -> if string contains '-' characters, it's probably in the form of 'day-mon-year' -> error
  // -> if string contains '/' characters, it's probably in the form of 'day/mon'year' -> take the last piece as year.
  // -> if string doesn't contain either, it's probably just 'year' -> if 4 numbers and between 1900 and 2100, take it.
  var dobYear = '';
  
  if ( dobString.indexOf('-') !== -1 ) {
    Logger.log('DOB string has \'-\' character -> not sure what to do with year!');
  } else if ( dobString.indexOf('/') !== -1 ) {
    
    var dobArr = dobString.split('/');
    
    if ( dobArr.length === 3 ) {
      
      if ( dobArr[2].length === 4 ) dobYear = dobArr[2];
      else {
        Logger.log('Why isnt the 3rd component the year? Handle it here :( dobString:');
        Logger.log(dobString);
      }
    }
    // else -> not the right length, so don't set dobYear
  } else {
    // 4 characters?
    if ( dobString.length !== 4 ) Logger.log('DOB string has no common delimiters and isnt a 4 character year :(');
    else if ( parseInt(dobString) > 1900 && parseInt(dobString) < 2100 ) dobYear = dobString;
  }
  
  if (dobYear === '') {
    Logger.log('Couldnt figure out what to do with DOBs!');
    return;
  }
  
  // we have dob Year, now age (difference between this year and dob year).
  return thisYear - parseInt(dobYear);
}

// determines target row based off getTargetRowMethod() - ascending or descending submissions
function getTargetRow(sheet) {
  var method = getTargetRowMethod();
  
  // get target cell coords
  //  Column = letter - 64 ("A" is 65, so "A" turns into 1)
  var targetRow = getTargetRowStart();
  
  if (method === "DESC") {
    // descending = bottom value in target sheet is the most recent entry
    
    // get top - left coordinates: (subtract 1 b/c indices are 0,1,2 not 1,2,3)
    var rowCoord     = targetRow - 1;
    
    var range = sheet.getDataRange();
    var colors = range.getBackgrounds();
    var rangeData = range.getDisplayValues();
    
    var found = false;
    
    while (!found) {
      var row = rangeData[rowCoord];
      
      if (rowCoord > 5000) {
        // rowCoord > 5000 means there's probably an error.
        Logger.log("rowCoord > 5000 - are there really more than 5000 clients? If yes, change the code!");
        
        found = true;
        return -1;
      } else if (rowCoord >= rangeData.length) {
        // rowCoord is out of bounds (all rows have data)
        //  -> stop looping and return an error.
        Logger.log("rowCoord reached final row of data in spreadsheet. This should happen in Docket SS");
        
        found = true;
        return -1;
        
      } else if ( isRowEmpty(row) ) {
        // found the first empty row ->
        found = true;
        
      } else if ( colors[rowCoord][1] === getOrangeRowColor() || colors[rowCoord][1] === getGreenRowColor()) {
        // found orange or gree row -> insert row above and paste data! (added 6 July 2017)
        found = true;
        
      } else {
        // row has data and rowCoord in range, so look at next row.
        rowCoord++;
        
      }
    }
    
    // return row # (not index number)
    return rowCoord + 1;

  } else if (method === "ASC") {
    // ascending = top value in target sheet is the most recent entry
    return targetRow;
  }
}

// function loops throw row values & returns if row is all empty strings ('')
//  DEFAULT -> loops through entire row of data
//  OPTIONAL -> loops through TARGET range only
function isRowEmpty(rowArr) {
  // OPTIONAL -> var targetColumnStart = getFirstColumnLetter().charCodeAt(0) - 64;
  // OPTIONAL -> var targetColumnEnd = getLastColumnLetter().charCodeAt(0) - 64;
  
  // OPTIONAL -> var columnCoordStart  = targetColumnStart - 1; // inclusive
  // OPTIONAL -> var columnCoordEnd = targetColumnEnd - 1; // exclusive
  
  // OPTIONAL -> for (var i = columnCoordStart; i < (columnCoordEnd + 1); i++) {

  // DEFAULT:
  for (var i = 0; i < rowArr.length; i++) {
    var val = rowArr[i];
    
    if (val !== '') return false;
  }
  
  return true;
}

// gets the last column letter based on:
//  1. the first column letter
//  2. the length of the data array parameter (from form submission)
function getLastColumnLetter(dataArr) {
  // get column 1 number from first column letter
  var col1num = getFirstColumnLetter().charCodeAt(0);
  
  // add length of dataArr to first column to get final column number:
  var col2num = col1num + dataArr.length;
  
  col2num--; // subtract 1 for science (i don't want to explain)
  
  return String.fromCharCode(col2num);
}
