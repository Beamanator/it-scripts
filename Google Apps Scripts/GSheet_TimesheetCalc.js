// PLEASE keep code up to date with source repo: https://github.com/Beamanator/stars-scripts
/*
TODOs:
1) allow multiple groups of times
ex: '8:00am - 2:23pm, 4:00pm - 5:15pm' // returns '7hrs 38mins'
*/

/**
 * Function calculates number of hours and minutes between two times. Must include 'am'
 * or 'pm' after every time, times are separated with a dash '-', and you can add location
 * specifiers like (home) or (away) or (mogamma) if you'd like.
 * @param {"9:00am - 5:32pm (location)"} timeRange  Start - End time for when you
 *                                                  worked that day, with optional
 *                                                  location specifier in parenthesis
 * @return total amount of time in '#hrs #mins' format
 * @customfunction
 */
function GET_H_MM(totalTime) {
  totalTime = totalTime.toLowerCase().trim();

  // treat plan dash '-' as 0 hours
  if (totalTime === "-") return "0hrs";

  // error if time range is undefined / falsy but not an empty string
  if (!totalTime)
    throw Error("Please pass a cell reference with a time, or '-' for 0hrs.");

  // ask user (politely) to remove her / his time calculation
  if (totalTime.indexOf("=") !== -1)
    throw Error("Remove your calculation (=) please.");

  // remove parenthesis groups from timeRange
  totalTime = removeParenthesisGroupsFromString(totalTime);

  // f no special characters, treat as special cell like annual leave, other leave, etc
  if (!/[-(),:]/.test(totalTime)) {
    var expectedSpecialCells = [
      "annual leave",
      "al",
      "other leave",
      "ol",
      "compassionate leave",
      "cl",
      "toil"
    ];

    var specialCellFound = false;

    for (var i = 0; i < expectedSpecialCells.length; i++) {
      var specialCell = expectedSpecialCells[i];

      // first, check if text IS EXACTLY one of the expected values
      if (specialCell === totalTime) {
        specialCellFound = true;
        break;
      }
    }

    if (specialCellFound) {
      return "0hrs";
    }
    // otherwise, throw error
    else {
      throw Error("Cannot determine what you entered!");
    }
  }

  // here, split into multiple calcuations, 1 + 1 extra calc for every comma
  var timeRanges = totalTime.split(",");

  var totalHrs = 0;
  var totalMins = 0;

  for (var i = 0; i < timeRanges.length; i++) {
    var timeRange = timeRanges[i];

    // make sure time format looks right
    var firstSplit = timeRange.split("-");
    if (firstSplit.length !== 2)
      throw Error('Time "' + timeRange + '" should have exactly 1 dash "-".');

    // get start & end times
    var startTime = firstSplit[0].trim();
    var endTime = firstSplit[1].trim();

    // get total number of hours & minutes in each time
    var startHrs = getHoursFromTime(startTime);
    var startMins = getMinutesFromTime(startTime);
    var endHrs = getHoursFromTime(endTime);
    var endMins = getMinutesFromTime(endTime);

    // calculate totals
    var sumMins = endMins - startMins;
    var sumHrs = endHrs - startHrs;
    if (sumMins < 0) {
      sumHrs--;
      sumMins += 60;
    }

    // handle end time BEFORE start time
    if (sumHrs < 0)
      throw Error("Start time must be BEFORE End time: " + timeRange);

    totalHrs += sumHrs;
    totalMins += sumMins;
  }

  // TODO: handle mins > 60 (convert to extra hour)

  // output totals
  return totalHrs + "hrs " + totalMins + "mins";
}

function getHoursFromTime(time) {
  time = time.toLowerCase();

  var isPm;

  // determine if time in 'am' or 'pm' or invalid
  if (time.indexOf("am") !== -1) isPm = false;
  else if (time.indexOf("pm") !== -1) isPm = true;
  else throw Error("Couldn't find 'am' or 'pm' in '" + time + "'");

  // calculate number of hours + if valid
  if (time.indexOf(":") === -1)
    throw Error("Couldn't find time / minute separator ':' in '" + time + "'");
  var numHours = parseInt(time.split(":")[0]);
  if (numHours > 12 || numHours < 1)
    throw Error(
      "Number of hours should be between 1 and 12 (inclusive) in '" + time + "'"
    );

  // convert hours to 0 - 23, instead of 0-11 for am + pm
  if (isPm) {
    // no need to change if 12pm
    if (numHours === 12) numHours = numHours;
    // else add 12
    else numHours += 12;
  }
  if (!isPm) {
    // handle possible 12am case (midnight) -> hour 0
    if (numHours === 12) numHours = 0;
  }

  return numHours;
}

function getMinutesFromTime(time) {
  time = time.toLowerCase();

  // determine if time is valid
  if (time.indexOf("am") === -1 && time.indexOf("pm") === -1)
    throw Error("Couldn't find 'am' or 'pm' in '" + time + "'");
  if (time.indexOf(":") === -1)
    throw Error("Couldn't find time / minute separator ':' in '" + time + "'");

  // remove 'am' and / or 'pm'
  var numMinutesString = time
    .replace("am", "")
    .replace("pm", "")
    .split(":")[1]
    .trim();

  // determine if numMinutesString is valid
  if (numMinutesString.length !== 2)
    throw Error("Minutes in '" + time + "' should be exactly 2 digits");

  var numMinutes = parseInt(numMinutesString, 10);

  // determine if numMinutes is valid
  if (isNaN(numMinutes) || numMinutes < 0 || numMinutes > 59)
    throw Error(
      "Number of minutes in '" +
        time +
        "' should be between 0 and 59 (inclusive)"
    );

  return numMinutes;
}

function removeParenthesisGroupsFromString(string) {
  // --> search timeRange for groups of (...) text, and remove them <--
  //   match any letter, number, or ! and ? and ' ' inside a set of parenthesis ()
  var validCharsInParenthesis = "[a-z0-9!? ]";
  var parenthesisRegex = RegExp(
    "\\(" + validCharsInParenthesis + "{0,}\\)",
    "gi"
  );
  var regexMatch,
    parenthesisTextMatches = [];
  while ((regexMatch = parenthesisRegex.exec(string)) !== null) {
    // push string that was matched into matchArray
    parenthesisTextMatches.push(regexMatch[0]);
  }
  // delete parenthesis group from timeRange text
  for (var i = 0; i < parenthesisTextMatches.length; i++) {
    string = string.replace(parenthesisTextMatches[i], "");
  }

  // search for other left or right parens, throw errors if they exist!
  if (string.indexOf("(") !== -1 || string.indexOf(")") !== -1)
    throw Error(
      "Err in (): Nested OR unmatched open / closing OR invalid text " +
        validCharsInParenthesis
    );

  return string;
}
