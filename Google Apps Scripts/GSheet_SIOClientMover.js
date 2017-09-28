// last updated: 28 May 2017
// ====================================================================================================
//                                   START - SIO Copy & Move to On-Call Sheets
// ====================================================================================================

// TODO: potentially beneficial enhancements:
// color an issue red [background] if it isn't handled by map object

// separate function for every SIO (and different buttons)
function clickedSIOAbdelfatah() {     mainFn();   }
function clickedSIOAdemFrench() {     mainFn();   }
function clickedSIOInterpreter(){     mainFn();   }

function clickedSIOJimie()      {     Logger.log('clicked Jimie');        }
function clickedSIORasha()      {     Logger.log('clicked Rasha');        }

// ------------------------------------- constants ------------------------------------
function getStartRowRow()  {  return   0; } // index - so first row = 0
function getStartRowCol()  {  return  12; } // also index - so 12 = 13th column
function getHeaderRowRow() {  return   0; } // also index - so first row = 0

function getServiceColIndex(){ return   4; } // service that determines where the row should go! - also index, not actual row #

function getTargetColIndex() { return 1; } // column index (1 = 2nd column) for client to be moved to, in target sheet

function getTimeColIndex() {return   8; } // column index (0 = first) where time should be placed

function transformRowLocation(service) {
  var map = {
    'RSD':              'RLAP On-call RSD',
    'RSD-registration': 'RLAP On-call RSD',
    'RSD-IV':           'RLAP On-call RSD',
    'RSD-Workshop':            'RLAP On-call RSD',
    'RSD-result':              'RLAP On-call RSD',
    'RSD-rejection':           'RLAP On-call RSD',
    'RSD-PRT - RSD-Residency': 'RLAP On-call RSD',
    'RSD-reop':         'RLAP On-call RSD',
    'RSD-followup':     'RLAP On-call RSD',
    
    'RST-registration': 'RLAP On-call RST',
    'RST-workshop':     'RLAP On-call RST',
    'RST-IV':           'RLAP On-call RST',
    'RST-rejection':    'RLAP On-call RST',
    'RST-followup':     'RLAP On-call RST',
    'RST-IV-Prep':      'RLAP On-call RST',
    'RST-PRT':          'RLAP On-call RST',
    
    'PS-housing':    'PS-AFP On-call',
    'PS-FA':         'PS-AFP On-call',
    'PS-mental':     'PS-AFP On-call',
    'PS-followup':   'PS-AFP On-call',
    'PS-SGBV':       'PS-AFP On-call',
    'PS-emergency':  'PS-AFP On-call',
    'PS-CRS':        'PS-AFP On-call',
    
    // only used for tracking -> Not actually sent anywhere
    'PS-medical':    'skip',
    'PS-UCY':        'skip',
    
    // not in list, as of 28 May 2017
    'Food vouchers': 'MoveFail',
    'Education':     'MoveFail'
  };
  
  return map[service];
}

// ----------------------------------- main function ----------------------------------- 
function mainFn() {
  // get active spreadsheet & sheet
  var mainSS = SpreadsheetApp.getActiveSpreadsheet();
  var thisS = mainSS.getActiveSheet();
  
  // get data range of sheet
  var sheetVs = thisS.getDataRange().getValues();
  
  var headerRow = sheetVs[getHeaderRowRow()];
  
  var startRowIndex = getStartingRowIndex(sheetVs);
  
  moveRows(sheetVs, startRowIndex, headerRow);
//  var dataToSend = getData(headerRow, sheetVs[startRowIndex]);
//  Logger.log(data);
  
  var lastRow = sheetVs.length;
  
  // set starting row to end of data range
  thisS.getRange(getStartRowRow() + 1, getStartRowCol() + 1).setValue( lastRow + 1 );
}

// -------------------------------- helper functions -----------------------------------

// function gets the client data starting on row @param1
function moveRows(sheetVs, startRowIndex, headerRow) {
  
  // loop through all rows of data in spreadsheet, after the "next row to send"
  for (var i = startRowIndex; i < sheetVs.length; i++) {
    var row = sheetVs[i];
    
    var rowData = getRowData(headerRow, row);
    
    if (rowData.length < 1) break;
    else {
      moveRow(rowData, i);
    }
  }
}

// function moves client data to new spreadsheets
// if "Issue" is not to be moved to a sheet, should be called 'skip' in object
function moveRow(row, rowIndex) {
  var service = row[ getServiceColIndex() ];
  var newSheetName = transformRowLocation(service);
  
  if (newSheetName == undefined) {
    Logger.log('couldn\'t find Issue: ' + service);
    Logger.log('make sure getServiceColIndex() is set to return the correct column number!');
    // TODO: maybe throw data into an 'unknown' sheet?
    
    // error, so color service red
    SpreadsheetApp.getActiveSheet().getRange(rowIndex+1, getServiceColIndex()+1).setBackground('red');
  } else if (newSheetName === 'skip') {
    // skip this service [service is only used for tracking purposes, not to be sent on call]
    Logger.log('skipping service: ' + service);
    return;
  } else {
    var newSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(newSheetName);
    
    var rowNum = getMoveRow( newSheet.getDataRange().getValues() );
    
    newSheet.getRange(rowNum, getTargetColIndex()+1, 1, row.length).setValues([row]);
    
//    Logger.log(newSheet.getName());
  }
}

// get row to move data to -> start at bottom of getDataRange, look for empty target column
// dataRange is 2D array, I think
function getMoveRow(dataRange) {
  for (var rowNum = dataRange.length; rowNum > 0; rowNum--) {

    var row = dataRange[rowNum - 1];
    
    var colIndex = getTargetColIndex();
    var cell = row[colIndex];
    
    if (cell !== '') {
      return rowNum + 1;
      break;
    } else {
      continue;
    }
  }
  
  return -2;
}

function getRowData(headerRow, dataRow) {
  if (dataRow[0] === '') return [];
  var dataToSend = [];
  
  // loop through all data in rows
  for (var i = 0; i < headerRow.length; i++) {
//    Logger.log(dataToSend);
    
    var cellHeader = headerRow[i];
    var cellData = dataRow[i];
    
    if (cellHeader == '') // break once an empty column is found
      break;
    
    var cellDataArr = manipulate(cellHeader, cellData, i);
//    Logger.log(cellDataArr);
    
    dataToSend = joinArr(dataToSend, cellDataArr);
  }
  
  return dataToSend;
}

// i think this is like [arr].push(), but can push arrays onto the end of arrays in a flat manner
function joinArr(orig, addition) {
  for (var i = 0; i < addition.length; i++) {
    orig.push(addition[i]);
  }
  
  return orig;
}

// function gets the row index to start grabbing client data for move
function getStartingRowIndex(sheetVs) {
  var row = getStartRowRow();
  var col = getStartRowCol();
  
  // substract 1 because cell is row #, not row index
  var startRowIndex = sheetVs[row][col] - 1;
 
  if (startRowIndex < 1 || startRowIndex === undefined)
    return -1;
  else
    return startRowIndex;
}

function strHas(orig, str) {
  var found = orig.indexOf(str);
//  Logger.log(found);
  if ( found === -1)
    return false;
  else
    return true;
}

// ---------------------------------- the manipulator --------------------------------------
function manipulate(key, data, headerRowIndex) {
  var returnArr = [];
  
  returnArr.push(data);
  
  // big, generic cases: -> Commented out because we will probably just standardize the look of every tab!
  // if row index matches where time column should go, automatically add time :D
  if ( headerRowIndex === getTimeColIndex() ) {
    // TODO: calculate time here, and push into array here
    var currentTimeArr = (new Date()).toLocaleTimeString().split(':');
    
    returnArr = joinArr([currentTimeArr[0] + ":" + currentTimeArr[1]], returnArr);
  }
  /*if ( strHas(key, 'Issue') && (strHas(data, "RST") || strHas(data, "RSD")) ) {
    // 
//    Logger.log('Issue contains RST / RST -> add 2 entries to array to bypass "Legal Officer" and "Time"');
//    Logger.log('pushing [][]');
    returnArr = joinArr(returnArr, ['','']);
  }
  */
  
  // individual, specific cases [use switch function?]
//  Logger.log('pushing ' + data);
  
  
  return returnArr;
}

// ----------------------------------------------------------------------------------------------------
//                                     END - SIO Copy & Move script
// ----------------------------------------------------------------------------------------------------
