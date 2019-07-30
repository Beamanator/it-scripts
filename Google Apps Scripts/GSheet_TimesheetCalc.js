// PLEASE keep code up to date with source repo: https://github.com/Beamanator/stars-scripts
/*
TODOs:
1) allow multiple groups of times
ex: '8:00am - 2:23pm, 4:00pm - 5:15pm' // returns '7hrs 38mins'
2) allow blank ('') / 0 / 0hrs input
ex: '0hrs' // returns '0hrs'
ex: '' // returns '0hrs'
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
function GET_H_MM(timeRange) {
  if (!timeRange)
    return "Err: Please pass a cell reference with a time to the function.";

  // --> search timeRange for groups of (...) text, and remove them <--
  //   match any letter, number, or ! and ? and ' ' inside a set of parenthesis ()
  var validCharsInParenthesis = "[a-z0-9!? ]";
  var parenthesisRegex = RegExp(
    "\\(" + validCharsInParenthesis + "{0,}\\)",
    "gi"
  );
  var regexMatch,
    parenthesisTextMatches = [];
  while ((regexMatch = parenthesisRegex.exec(timeRange)) !== null) {
    // push string that was matched into matchArray
    parenthesisTextMatches.push(regexMatch[0]);
  }
  // delete parenthesis group from timeRange text
  for (var i = 0; i < parenthesisTextMatches.length; i++) {
    timeRange = timeRange.replace(parenthesisTextMatches[i], "");
  }

  // search for other left or right parens, throw errors if they exist!
  if (timeRange.indexOf("(") !== -1 || timeRange.indexOf(")") !== -1)
    return (
      "Err in (): Nested OR unmatched open / closing OR invalid text " +
      validCharsInParenthesis
    );

  // make sure time format looks right
  if (timeRange.indexOf(",") !== -1)
    return "Err: Found ',' but function doesn't handle multiple time ranges yet";
  var firstSplit = timeRange.split("-");
  if (firstSplit.length !== 2)
    return 'Err: Time should have 1 "-" character ONLY.';
  if (timeRange.indexOf("=") !== -1)
    return "Err: Remove your calculation (=) please.";

  // get start & end times
  var startTime = firstSplit[0].trim();
  var endTime = firstSplit[1].trim();

  // get total number of hours & minutes in each time
  var startHrs = getHoursFromTime(startTime);
  var startMins = getMinutesFromTime(startTime);
  var endHrs = getHoursFromTime(endTime);
  var endMins = getMinutesFromTime(endTime);

  // display errors if there are any so far
  if (startHrs.err) return startHrs.err;
  if (startMins.err) return startMins.err;
  if (endHrs.err) return endHrs.err;
  if (endMins.err) return endMins.err;

  // calculate totals
  var totalMins = endMins - startMins;
  var totalHrs = endHrs - startHrs;
  if (endMins < startMins) {
    totalHrs--;
    totalMins += 60;
  }

  // handle end time BEFORE start time
  if (totalHrs < 0) return "Err: Start time must be BEFORE End time";

  // output totals
  return totalHrs + "hrs " + totalMins + "mins";
}

function getHoursFromTime(time) {
  time = time.toLowerCase();

  var isPm;

  // determine if time in 'am' or 'pm' or invalid
  if (time.indexOf("am") !== -1) isPm = false;
  else if (time.indexOf("pm") !== -1) isPm = true;
  else return { err: "Err: Couldn't find 'am' or 'pm' in '" + time + "'" };

  // calculate number of hours + if valid
  if (time.indexOf(":") === -1)
    return {
      err: "Err: Couldn't find time / minute separator ':' in '" + time + "'"
    };
  var numHours = parseInt(time.split(":")[0]);
  if (numHours > 12 || numHours < 1)
    return {
      err:
        "Err: Number of hours should be between 1 and 12 (inclusive) in '" +
        time +
        "'"
    };

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
    return { err: "Err: Couldn't find 'am' or 'pm' in '" + time + "'" };
  if (time.indexOf(":") === -1)
    return {
      err: "Err: Couldn't find time / minute separator ':' in '" + time + "'"
    };

  // remove 'am' and / or 'pm'
  var numMinutesString = time
    .replace("am", "")
    .replace("pm", "")
    .split(":")[1]
    .trim();

  // determine if numMinutesString is valid
  if (numMinutesString.length !== 2)
    return { err: "Err: Minutes in '" + time + "' should be exactly 2 digits" };

  var numMinutes = parseInt(numMinutesString, 10);

  // determine if numMinutes is valid
  if (isNaN(numMinutes) || numMinutes < 0 || numMinutes > 59)
    return {
      err:
        "Err: Number of minutes in '" +
        time +
        "' should be between 0 and 59 (inclusive)"
    };

  return numMinutes;
}
