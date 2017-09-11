/**
 * Last Updated: 15 Aug 2017
 * 
 * Purpose = Generate a list of clients from spreadsheet (with white background) in this order:
 * 1) Random from list
 * 2) Random from list
 * 3) Random from list
 * etc... 
 * 
 * Old formats:
 * -- Before August 10:
 * 1) Next on list
 * 2) Next on list
 * 3) Random from list
 * etc...
 * 
 * NOTE: Make sure you set up your onOpen trigger before expecting the code to run!
 * Edit -> Current project's triggers -> 
 * 
 * After list is generated, email list to specified person (in spreadsheet)
 */ 

function getSettingsRow()    {  return 2;  }
function getStartDataRow()   {  return 5;  }

function getEmailColumn()    {  return 2;  } // actual col #, not index
function getListSizeColumn() {  return 4;  }
function getColumnsColumn()  {  return 6;  }

function getColorOrange()    {  return '#ff9900';  }
function getColorGreen()     {  return '#00ff00';  }
function getColorWhite()     {  return '#ffffff';  }
function getColorRed()       {  return '#ff0000';  }

function getEmailSubject()   {  return 'RST Self-Referral - Generated Client List';  }

// Add a custom menu to the active spreadsheet, including a separator and a sub-menu.
function onOpen(e) {
  SpreadsheetApp.getUi()
    .createMenu('Generate')
    .addItem('Generate Client List', 'generateList')
    .addToUi();
}

/**
 * Main function called by button click
 * 
 */
function generateList() {
  var sheet = SpreadsheetApp.getActiveSheet();

  var dataRange = sheet.getDataRange();

  // Fetch values & backgrounds in data range
  var data = dataRange.getValues();
  var bgColors = dataRange.getBackgrounds();

  var email =    data[ getSettingsRow() - 1 ][ getEmailColumn()    - 1 ];
  var listSize = data[ getSettingsRow() - 1 ][ getListSizeColumn() - 1 ];
  var columns  = data[ getSettingsRow() - 1 ][ getColumnsColumn()  - 1 ].split(',');

  var columns = getColNumbersFromLetters(columns);

  // Get full list of possible clients:
  var fullIndexList = getFullIndexList(data, bgColors);

  // 
  var finalIndexList = getFinalIndexList(fullIndexList, listSize, data);

  // TODO: color rows here if desired?

  // send emails
  emailClientData(email, finalIndexList, data, columns);
}

/**
 * Function converts array of column letters to column numbers.
 * Ex: [A,B,C] -> [1,2,3]
 */
function getColNumbersFromLetters(cols) {
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
  
  return colNums;
}

/**
 * Function sends email w/ formatted list of client data
 */
function emailClientData(email, indexList, data, cols) {
  var message = 'Generated list of clients to contact:';
  
  for (var i = 0; i < indexList.length; i++) {
    var clientIndex = parseInt(indexList[i]);
    var rowNum = clientIndex + 1;
    
    message += '\nRow ' + rowNum + ':\t\t| ';
    
    for (var j = 0; j < cols.length; j++) {
      var colNum = cols[j];
      var cellData = data[ clientIndex ][ colNum ];
      
      message += cellData + ' | ';
    }
  }
  
  MailApp.sendEmail(email, getEmailSubject(), message);
}

/**
 * Function gets the final array of clients
 */
function getFinalIndexList(list, numClients, clientData) {
  var finalIndices = [];
  
  for (var i = 0; i < numClients; i++) {
    // old logic when order was: Next, Next, Random, Next, Next, Random, etc...
    // take 2 items from the beginning of list 'list'
    // if (i % 3 === 0 || i % 3 === 1) {
    //   finalIndices.push( list.splice(0, 1) );
    // }
    
    // new logic = only pick random names from list
    // take a random client from list
    var nextIndex = Math.floor(Math.random() * list.length);
    
    // in case Math.random() can get close enough to 1...
    if (nextIndex === list.length) {
      nextIndex--;
    }
    
    // add client row index to array, and remove this index from list so no duplicates
    finalIndices.push( list.splice(nextIndex, 1) );
  }
  
  return finalIndices;
}

/**
 * Get all row indices with white background
 */
function getFullIndexList(data, colors) {
  var availRows = [];
  
  var colorObj = {};
  colorObj[ getColorWhite()  ] = 'white';
  colorObj[ getColorGreen()  ] = 'green';
  colorObj[ getColorOrange() ] = 'orange';
  colorObj[  getColorRed()   ] = 'red';
  
  // fill variable availRows with indices of rows with white background
  for (var i = getStartDataRow() - 1; i < colors.length; i++) {
    var color = colorObj[ colors[i][0] ];
    if (color === 'white') {
      availRows.push(i);
    }
  }
  
  return availRows;
}