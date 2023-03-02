// todo: 
//    done: get rows of month and year
//    done:   get Z; AB; AC; AD (Z=tech name; AB=date installed; AC=INSTALLATION STATUS; AD=remarks)
//            Z=25/26, AB=27/28; AC=28/29; AD=29/30;
//    done: cleanup team names
//    done: count INSTALLATION STATUS per tech name
//    done (7 oct 2022): classify and count remarks per team
//    done (11 oct 2022) check if addRemarksToSpread() works. >> it works
//    done (11 oct 2022) kijken of ik addInstallStatToSpread() en addRemarksToSpread() kan samenvoegen
//    done (11 oct 2022) meerdere remarks in 1 record scheiden en goed opsommen
//    done (11 oct 2022) no pole attachments samenvoegen
//    done (11 oct 2022) ?? reason cancelation filtered maken
//    done (12 oct 2022) reason cancelation unfiltered maken
//    dropped (12 oct 2022) reason cancelation filtered toevoegen
//    done (12 oct 2022) remarks unfiltered toevoegen
//    done (12 oct 2022) reason cancelation unfiltered toevoegen
//    done (18 oct 2022) make code more efficient search for: ">>> SEE #THIS#"
//    done (18 oct 2022) make dynamic headers
//    done (18 oct 2022) maak de super headers mooier
//    done (18 oct 2022) haal de eerste letter weg uit de namen
//    done (19 oct 2022) add count data to sheet
//    done (19 oct 2022) delete old sheet data before filling the sheet
//    done (19 oct 2022) fix the key names for the dictionaries.
//    done (19 oct 2022) clean the script - makeHeaders() is now more efficient
//    done (24 oct 2022) check of remarks kloppen met ruwe remarks
//    done (25 oct 2022) maak kleurtjes
//    pending (25 oct 2022) maak een (optioneel) dag overzicht (nog controleren)
//    pending (25 oct 2022) maak een (optioneel) jaar overzicht (nog controleren)
//    done (30 oct 2022) average daily SO amken.
//    done (30 oct 2022) translate comments to english
//    done (29 nov 2022) Corrected the sums by taking duplicate days into account.
//    done (30 nov 2022) can now use external files
//    done (5 dec 2022) check if count data is correct
//    done (5 dec 2022) remove sums with duplicate days
//    done (5 dec 2022) why are some days missing for the test set for the 3 man teams??
//    - get day overview working // not working
//    - get year overview working
//    - do THE BIG THINK


// PROBLEM !!: Only year works. Year and month; Year, month, day can give problems.
// PROBLEEM !!:   Day data can be wrong.
//                Some rows are not unique.
//                Maybe data is overwitten somewhere?
//                The culprit is most possibly that there is a problem with time zones.
//                The timezone is an Azian timezone, but my locale is Dutch timezone.
//                2-Dec-2022 in date attended is displayed as 2-nov-2022
// https://developers.google.com/apps-script/add-ons/how-tos/access-user-locale

// input: 11 Aug 2022
// output: 12 Aug 2022


///////////////////
// THE BIG THINK //
///////////////////
// ^^^^^^^^^^^^^^^^
// is it usefull to make a team summary?
// it is more about the individual people, than the team.
// I can add a team summary
// add team name next to the current result
// take numbers from the sheet
// recalculate days for the teams, summ all the non day stuff, recalculate thing/day stuff
// probably good to add the team key to the spreadsheet to calculate the overlapping days


const STARTROW = 11;


function generateSummary() {
  /**
   * The main function that gets you the counts and sums of all the teams.
   * The program only looks at specific columns of the sheet, by index.
   */

  // get parameters
  const params = getParameters(parameter_sheet_name="getInfo");

  // clear sheet
  clearSheet(params);

  // get the raw data
  var rawdatasheet = getSheet(params.spreadsheet, params.sheetname);

  // filter year and month
  var rawdata_filtered = filterByDate(rawdatasheet, params);
  

  // loop through data
  // CAUTION!! THE VARIABLE COUNTS ALSO CONTAINS A KEY 'TEAMS', WHICH IS JUST A LIST OF THE TEAMS.
  // CAUTION!! THIS NEEDS TO BE SEPERATED WHEN THE COUNTS ARE FILLED IN.
  var counts = loopData(rawdata_filtered);

  var counts = calcPerDay(counts);
  var number_of_days = counts["all_days"];
  delete counts["all_days"];

  
  // would be better if this was a global variable and already made in loopData
  // Getting the keys/headers
  // keys for counts (and keys) are: install_status remark remark_unf cancel_reason teams
  var keys = getKeys(counts);

  // get the sum of every column
  var counts_sum = sumCounts(counts, keys);

  // // Browser.msgBox(JSON.stringify(rawdata_filtered));
  // // Browser.msgBox(JSON.stringify(keys));
  // // Browser.msgBox(JSON.stringify(counts_sum));
  // // Browser.msgBox(JSON.stringify(counts));
  // // Browser.msgBox(JSON.stringify(counts));

  // // display all the counts.
  fillCounts(counts, params, keys, counts_sum, number_of_days);
}


function getNames(filtered_data) {
  // THIS IS PROBABLY OVERKILL, AND EASIER SOLUTION CAN BE MADE, ALTHOUGH LESS AUTOMATED
  // THE CURRENT INFORMATION IS USEFULL, SO NO NEED TO DESTROY IT.
  // USE THE CURRENT OUTPUT TO LET THE USER FILL IN TEAM NAMES, THEN GENERATE A TEAM SUMMARY UNDER EVERYTHING.

  // precies op dezelfde manier de teamnnamen verkrijgen
  // de namen niet joinen, maar toevoegen aan een nieuwe lijst, die meteen uniek gemaakt wordt

  // click button1: 
  // 1) get unique list of teams as list1: [['ALEJANDRO','ATIENZA','GOMEZ'],['ALEJANDRO','ASTROLOGO','GOMEZ'], etc]
  // 1.1) while making this list1, add every unique name to list2: [all unique names]
  // 1.2) use linklist12 to keep track where every list1 index connects to list2
  // 2) send list1, list2, linklist12 to a new tab
  // 3) in the list2 tab the user creates table3 and linklist23 by adding teamnames to table2

  // if user clicks on button1 again:
  // 4) get list1 again
  // 5) get list1_backup from tab
  //    compare list1 with list1_backup
  //    if different:
  //        add new stuff to list1
  //        generate linklist12
  //        add new elements from linklist12 to linklist12_tab
  //        ----> STILL NEED TO THINK
  // 6) check if there are new names in tab for list1
  // 7) add new names to tab
  // 8) user adds teamname for new members
  // 9) MAKE A POPUP BOX IF USER FORGOT TO GIVE NEW MEMBER A TEAM

  // if user clicks the summary button:
  // 10) loopData() starts
  // 11) teamname as list gets calculated again
  // 12) behind teamname output you can say who are in the team.


  var counts = {};
  // loop through rows of the filtered data set.
  for (var r=0; r < filtered_data.length; r++) {
    // cleanup names
    var team = filtered_data[r][0];
    team = team.toUpperCase();

    // remove the word team, remove :, replace separators
    team = team.replace("TEAM", "").replaceAll(":", "").replaceAll(",", "/").replace(/\d+/g, '');
    team = team.replaceAll("\\s+", " ");
    var splitteam = team.trim();

    // split the names into a list
    if (team.indexOf("/") > -1) {
      splitteam = splitteam.split("/");
    } else {
      splitteam = splitteam.split(" ");
    }

    // remove initials and empty names
    for (var i=0; i < splitteam.length; i++) {
      if (splitteam[i].indexOf(".") > -1 && splitteam[i].split(".")[0].trim().length==1) {
        splitteam[i] = splitteam[i].split(".")[1].trim();
      }
      // needed to remove empty names
      splitteam[i] = "@"+splitteam[i].trim()+"@";
      if (splitteam[i] === "@@") { 
        splitteam.splice(i, 1); 
      }
    }

    // restore names
    for( var i = 0; i < splitteam.length; i++){ 
      splitteam[i] = splitteam[i].replaceAll("@", "");
    }
    splitteam = splitteam.sort();
    var names = splitteam.join("/");
    filtered_data[r][0] = names;
    team = filtered_data[r][0];

    // add team if it's not there yet.
    if(!counts.hasOwnProperty(team)) {
      counts[team] = {};
    }
  }
}

function calcPerDay(counts) {
  /**
   * Calculates the RSO average and installed average.
   * This is done by counting the dates in the spreadsheet and counting the uniqe days in the spreadsheet.
   * The problem is that a team on a day can have wrongfully imputed the team name, so you will count the days double.
   * As long as the team names are not fixed, this statistic isn't reliable.
   * @params  {Object}  counts  counts[team][install_status/remark/remark_unf/cancel_reason] = dict with counts.
   * @return  {Object}  counts  counts[team][install_status/remark/remark_unf/cancel_reason] = dict with counts.
   */
    var all_days = [];
    let teams = Object.keys(counts);
    for(let ti=0; ti<teams.length; ti++) { // ti = team index
      let team = teams[ti];
      let unique_days = [...new Set(counts[team]["dates"])];
      let days_worked = unique_days.length;
      let installed = counts[team]["install_status"]["INSTALLED"];
      let rso = counts[team]["install_status"]["RSO"];
      if (rso==0 || rso===undefined) {
        counts[team]["install_status"]["RSO"] = 0;
        counts[team]["install_status"]["RSO/DAY"] = 0;
      }
      if (installed==0 || installed===undefined) {
        counts[team]["install_status"]["INSTALLED/DAY"] = 0;
        counts[team]["install_status"]["INSTALLED"] = 0;
      }
      if(days_worked>0) {
        if(installed>0) {
          counts[team]["install_status"]["INSTALLED/DAY"] = installed / days_worked;
        }
        if(rso>0) {
          counts[team]["install_status"]["RSO/DAY"] = rso / days_worked;
        }
      }
      counts[team]["install_status"]["DAYS"] = days_worked;
      delete counts[team]["dates"];
      all_days = all_days.concat(unique_days);
    }
    counts["all_days"] = [...new Set(all_days)].length;
    return counts;
}


function getParameters(parameter_sheet_name) {
  /**
   * Get the input data to generate the counts.
   * Month and Day variables are allowed to be empty.
   * Year only: year summary will be generated.
   * Year and Month: Month summary will be generated.
   * Year, month and day: Day summary will be genearted.
   * input: cell locations:
   * B1: spreadsheet link; B2: sheet name; B3: year; B4: month; B5: day}.
   * @return {Object} params  Needed data for the script.
   */
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('getInfo');
  var params = {
    spreadsheet : sheet.getRange('B1').getValue(),
    sheetname: sheet.getRange('B2').getValue(),
    year: sheet.getRange('B3').getValue(),
    month: sheet.getRange('B4').getValue(),
    day: sheet.getRange('B5').getValue(),
    parameter_spreadsheet: ss.getId(),
    parameter_sheetname: parameter_sheet_name
  }

  if(sheet.getRange('B4').isBlank()) {
    params.month = NaN;
  }
  if(sheet.getRange('B5').isBlank()) {
    params.day = NaN;
  }
  return params; 
}


function clearSheet(params) {
  /**
   * Cleans the getInfo tab from the active spreadsheet starting at
   * the row that is declared in global variable STARTROW.
   * @params  {Object}  Contains paramteters for file locations and filters.
   *                    parameters are created from getParameters().
   */
  var spreadsheet = SpreadsheetApp.openById(params.parameter_spreadsheet);
  var sheet = spreadsheet.getSheetByName(params.parameter_sheetname);
  var range = sheet.getRange(STARTROW, 1, sheet.getMaxRows(), sheet.getMaxColumns());
  range.clear();
}


function getSheet(spreadsheetid, sheetname) {
  /**
   * Fetches the spreadsheet that contains the data needed for the summation.
   * @params  {String}  spreadsheetid   Which spreadsheet needs to be activated.
   * @params  {String}  sheetname       Which sheet needs to be activated.
   * @return  {object}  raw spreadsheet with data from selected spreadsheet and sheet.
   */
  var spreadsheet = SpreadsheetApp.openById(spreadsheetid);
  var rawdatasheet = spreadsheet.getSheetByName(sheetname);
  return rawdatasheet;
}

// the real one
function filterByDate(rawdatasheet, params) {
  /**
   * Filter the spreadsheet by year, month and day.
   * Only the rows [TECH NAME, DATE ATTENDED, DATE INSTALLED, INSTALLATION STATUS, REMARKS]
   * are selected from the spreadsheet.
   * @param   {object}  rawdatasheet  Data from the spreadsheet; see: getSelectedSheet().
   * @param   {object}  params        Contains day,year,month for filtering; see: getParameters().
   * @return  {object}  result        Subset of the spreadsheet data filtered by year/month/day.
   */
  // getRange(row, column, numRows, numColumns)
  var range = rawdatasheet.getRange(2,26,rawdatasheet.getLastRow(),5);
  var data  = range.getValues();
  var result;

  if(Number.isNaN(params.month) && Number.isNaN(params.day)) {
    // get whole year
    result = data.filter(function(item) {
      let date = new Date(item[2]);
      let year = date.getFullYear();
      return year===params.year;
    });
  } else if(!Number.isNaN(params.month) && Number.isNaN(params.day)) {
    // get whole month
    result = data.filter(function(item) {
      let date = new Date(item[2]);
      let year = date.getFullYear();
      let month = date.toLocaleString("en-US", { month: "short" });
      return year===params.year && month===params.month;
    });
  } else if(!Number.isNaN(params.month) && !Number.isNaN(params.day)) {
    // get a day
    result = data.filter(function(item) {
      let date = new Date(item[2]);
      let year = date.getFullYear();
      let month = date.toLocaleString("en-US", { month: "short" });
      // let day = date.getDay(); // THIS GIVES THE WEEK DAY AS AN INTEGER!!!! MO=1;TU=2 ETC....
      let daydate = date.getDate().valueOf(); // STARTS AT 0 !
      // Browser.msgBox(item[2].toString() + " | "+daydate.toString()+" | "+year.toString()+" | "+month.toString()+" | "+item[4].toString());
      return year===params.year && month===params.month && daydate===(params.day-1);
    });
  }
  return result;
}

// // this one is for test
// function filterByDate(rawdatasheet, params) {
//   /**
//    * Filter the spreadsheet by year, month and day.
//    * WE USE DATE INSTALLED (index 2).
//    * Only the rows [TECH NAME, DATE ATTENDED, DATE INSTALLED, INSTALLATION STATUS, REMARKS]
//    * are selected from the spreadsheet.
//    * @param   {object}  rawdatasheet  Data from the spreadsheet; see: getSelectedSheet().
//    * @param   {object}  params        Contains day,year,month for filtering; see: getParameters().
//    * @return  {object}  result        Subset of the spreadsheet data filtered by year/month/day.
//    */
//   // getRange(row, column, numRows, numColumns)
//   var range = rawdatasheet.getRange(2,1,rawdatasheet.getLastRow(),rawdatasheet.getLastColumn());
//   var data  = range.getValues();
//   var result;

//   if(Number.isNaN(params.month) && Number.isNaN(params.day)) {
//     // get whole year
//     result = data.filter(function(item) {
//       let date = new Date(item[27]);
//       let year = date.getFullYear();
//       return year===params.year;
//     });
//   } else if(!Number.isNaN(params.month) && Number.isNaN(params.day)) {
//     // get whole month
//     result = data.filter(function(item) {
//       let date = new Date(item[27]);
//       let year = date.getFullYear();
//       let month = date.toLocaleString("en-US", { month: "short" });
//       return year===params.year && month===params.month;
//     });
//   } else if(!Number.isNaN(params.month) && !Number.isNaN(params.day)) {
//     // get a day
//     result = data.filter(function(item) {
//       let date = new Date(item[27]);
//       let year = date.getFullYear();
//       let month = date.toLocaleString("en-US", { month: "short" });
//       // let day = date.getDay(); // THIS GIVES THE WEEK DAY AS AN INTEGER!!!! MO=1;TU=2 ETC....
//       let daydate = date.getDate().valueOf(); // date is as an index....
//       // Browser.msgBox(item[2].toString() + " | "+daydate.toString()+" | "+year.toString()+" | "+month.toString()+" | "+item[4].toString()+" | "+params.day);
//       return year===params.year && month===params.month && daydate===(params.day-1);
//     });
//   }
//   return result;
// }

function loopData(filtered_data) {
  /**
   * Itterates through the rows of the filtered dataset and cleans up the team name;
   * counts the install statuses, remarks and cancel reasons by team.
   * @param {Object}  filtered_data contains the filtered spreadsheet from filterByDate().
   * @return {Object} counts        counts[team][install_status/remark/remark_unf/cancel_reason] = dict with counts.
   * NOTE:  The cleaning up of names didn't work when put into a function.
   *        Could not find the reason why.
   */
  var counts = {};
  // loop through rows of the filtered data set.
  for (var r=0; r < filtered_data.length; r++) {
    // cleanup names
    var team = filtered_data[r][0];
    team = team.toUpperCase();

    // remove the word team, remove :, replace separators
    team = team.replace("TEAM", "").replaceAll(":", "").replaceAll(",", "/").replace(/\d+/g, '');
    team = team.replaceAll("\\s+", " ");
    var splitteam = team.trim();

    // split the names into a list
    if (team.indexOf("/") > -1) {
      splitteam = splitteam.split("/");
    } else {
      splitteam = splitteam.split(" ");
    }

    // remove initials and empty names
    for (var i=0; i < splitteam.length; i++) {
      if (splitteam[i].indexOf(".") > -1 && splitteam[i].split(".")[0].trim().length==1) {
        splitteam[i] = splitteam[i].split(".")[1].trim();
      }
      // needed to remove empty names
      splitteam[i] = "@"+splitteam[i].trim()+"@";
      if (splitteam[i] === "@@") { 
        splitteam.splice(i, 1); 
      }
    }

    // restore names
    for( var i = 0; i < splitteam.length; i++){ 
      splitteam[i] = splitteam[i].replaceAll("@", "");
    }
    splitteam = splitteam.sort();
    var names = splitteam.join("/");
    filtered_data[r][0] = names;
    team = filtered_data[r][0];

    // add team if it's not there yet.
    if(!counts.hasOwnProperty(team)) {
      counts[team] = {};
    }

    // counts install status
    counts[team] = countInstalStatus(counts[team], filtered_data[r]); // new version

    // counts remarks and cancel reasons
    counts[team] = countRemarksCancelReas(counts[team], filtered_data[r]);
  }
  return counts;
}


function countInstalStatus(teamcounts, datarow) {
  /**
   * Adds a "install status" object to teamcounts, wherin every type of install status gets counted.
   * @param {Object}  teamcounts  Count data for the current team.
   * @param {array}   datarow     Contains data for the current team in the order: 
   *  [TECH NAME, DATE ATTENDED, DATE INSTALLED, INSTALLATION STATUS, REMARKS]
   * @return {Object} teamcounts  Updated count data for the current team.
   */
  // Add install_status if it isn't there yet.
  if (!teamcounts.hasOwnProperty("install_status")) {
    teamcounts["install_status"] = {};
  }
  // datarow = 0:TECH NAME; 1:DATE ATTENDED; 2:DATE INSTALLED; 3:INSTALLATION STATUS; 4:REMARKS)
  let install_status = datarow[3].trim();
  // Initialise a install status if it isn't there yet.
  if (!teamcounts["install_status"].hasOwnProperty(install_status)) {
    teamcounts["install_status"][install_status] = 0;
  }
  // Increase count for the instal status
  teamcounts["install_status"][install_status] += 1;

  // add date to list
  if(!teamcounts.hasOwnProperty("dates")) {
      teamcounts["dates"] = [];
  }
  let date = new Date(datarow[1]).toDateString();
  teamcounts["dates"].push(date);
  return teamcounts;
}


function countRemarksCancelReas(teamcounts, datarow) {
  /**
   * Counts remarks, cancel reasons, and raw remarks, and adds those to teamcounts,
   * a team specific slice of all the counts.
   * @param {Object}  teamcounts  Count data for the current team.
   * @param {array}   datarow     Contains data for the current team in the order: 
   *  [TECH NAME, DATE ATTENDED, DATE INSTALLED, INSTALLATION STATUS, REMARKS]
   * @return {Object} teamcounts  Updated count data for the current team.
   */
  // make the sub dictionaries/objects
  if (!teamcounts.hasOwnProperty("remark")) {
    teamcounts["remark"] = {};
  }
  if (!teamcounts.hasOwnProperty("remark_unf")) {
    teamcounts["remark_unf"] = {};
  }
  if (!teamcounts.hasOwnProperty("cancel_reason")) {
    teamcounts["cancel_reason"] = {};
  }
  
  var remark = datarow[4].trim().toLowerCase();
  // look for reason of cancelation
  if (remark.match(/cancel *by *subs?/)) {
    teamcounts = countCancelReason(remark, teamcounts);
  }

  // add filtered remark to the count table
  teamcounts = countFilteredRemark(remark, teamcounts);

  // count raw unfiltered remarks
    if (!teamcounts["remark_unf"].hasOwnProperty(remark)) {
      teamcounts["remark_unf"][remark] = 0;
    }
    teamcounts["remark_unf"][remark] += 1;
  return teamcounts;
}


function countCancelReason(remark, teamcounts) {
  /**
   * Counts cancel reasons if calcelled by subs, and adds those to teamcounts.
   * @param {String}  remark      The remark of the datarow which can contain the reaosn for cancelation.
   * @param {Object}  teamcounts  Count data for the current team.
   * @return {Object} teamcounts  Updated count data for the current team.
   */
    // Increase count of cancel by subs, and if needed, initialise the counter for the remark.
    remark_cancel = "CANCEL BY SUBS";
    if (!teamcounts["remark"].hasOwnProperty(remark_cancel)) {
      teamcounts["remark"][remark_cancel] = 0;
    }
    teamcounts["remark"][remark_cancel] += 1;

    // start looking for the reason of the cancelation
    reason = remark.split("cancel by subs")[1];
    reason = reason.replaceAll(":", "").replaceAll(",", "").replaceAll("/", "").trim();
    if (reason.match(/not in\w*ed/)) {
      reason = "NOT INTERESTED";
    } else if (reason.match(/ot\w* *provi/)) {
      reason = "OTHER PROVIDER";
    } else if (reason.match(/fin\w* * pro\w*/)) { 
      reason = "FINANCIAL PROBLEM";
    } else {
      reason = reason.trim();
    }

    // Increase count of cancel reason, and if needed, initialise the counter for the cancel reason.
    if(!teamcounts["cancel_reason"].hasOwnProperty(reason)) {
      teamcounts["cancel_reason"][reason] = 0;
    }
    teamcounts["cancel_reason"][reason] += 1;

    return teamcounts;
}


function countFilteredRemark(remark, teamcounts) {
  /**
   * Searches for a fixed set of remarks using regex, and counts them.
   * Other and empty remarks are handled on their own.
   * @param {String}  remark      The remark of the datarow which can contain the reason for cancelation.
   * @param {Object}  teamcounts  Count data for the current team.
   * @return {Object} teamcounts  Updated count data for the current team.
   */
  // sometimes there are multiple remarks in one sentece, so we just check for the occurance of some words
  // The dsp in and out looks a bit weird, but it is faster, because it skips a loop to check if one of them is true.
  var dsp = false; // = did_something_pass
  teamcounts["remark"], dsp = scanForRemark(teamcounts["remark"], remark.match(/^re\w*h\w*d/), "RESCHED", dsp);
  teamcounts["remark"], dsp = scanForRemark(teamcounts["remark"], remark.match(/^al\w*dy *in\w*d/), "ALREADY INSTALLED", dsp);
  teamcounts["remark"], dsp = scanForRemark(teamcounts["remark"], remark.match(/cant *locate/), "CANT LOCATE",dsp);
  teamcounts["remark"], dsp = scanForRemark(teamcounts["remark"], remark.match(/full? *nap/), "FULL NAP",dsp);
  teamcounts["remark"], dsp = scanForRemark(teamcounts["remark"], remark.match(/no *pole/), "NO POLE (ATTACHMENT)",dsp);
  teamcounts["remark"], dsp = scanForRemark(teamcounts["remark"], remark.match(/^in\w*led/), "INSTALLED",dsp);
  teamcounts["remark"], dsp = scanForRemark(teamcounts["remark"], remark.match(/^un\w*con\w*ta\w*ted/), "UNCONTACTED",dsp);
  
  // if nothing passed
  if (!dsp) {
    if (remark.trim()=="") {
        if (!teamcounts["remark"].hasOwnProperty("NO REMARK ENTERED")) {
          teamcounts["remark"]["NO REMARK ENTERED"] = 0;
        }
      teamcounts["remark"]["NO REMARK ENTERED"] += 1;
    } else {
      if (!teamcounts["remark"].hasOwnProperty("OTHER REMARK")) {
          teamcounts["remark"]["OTHER REMARK"] = 0;
        }
      teamcounts["remark"]["OTHER REMARK"] += 1;
    }
  }
  return teamcounts;
}


function scanForRemark(teamremarkcounts, match, remark, prev_pass) {
  /**
   * Takes the result of the regex on the remark.
   * If found, increase count or initiate the count.
   * Also turns prev_pass into true if it found a match.
   * @param   {Object}  teamremarkcounts    Keeps track the counts for the remarks for the team.
   * @param   {Bool}    remark              True when the regex matches
   * @param   {Bool}    prev_pass           Takes the pass flag from previous search and turns it on if a match is found.
   * @return  {Object}  teamremarkcounts    Updated counts for the remarks for the team.
   * @return  {Bool}    prev_pass           Updated flag for a regex hit.
   */
  if (match) {
    if (!teamremarkcounts.hasOwnProperty(remark)) {
      teamremarkcounts[remark] = 0;
    }
    teamremarkcounts[remark] += 1;
    prev_pass = true;
  }
  return teamremarkcounts, prev_pass;
}


function getKeys(counts) {
  /**
   * Searches for all the keys in the third dimension of the counts object, and sorts them.
   * @param {Object}  counts    Count data for all teams.
   * @return {Object} keys      2d object that contains the sorted keys for every key group.
   */
  // reminder: counts[team][install_status/remark/remark_unf/cancel_reason] = dict with counts
  keys = {};
  keys["install_status"] = [];
  keys["remark"] = [];
  keys["remark_unf"] = [];
  keys["cancel_reason"] = [];
  // We also add and sort the teams, while we are at it.
  keys["teams"] = Object.keys(counts).sort();  

  // get all keys from the objects
  for (const team in counts) {
    let is = Object.keys(counts[team]["install_status"]);
    let re = Object.keys(counts[team]["remark"]);
    let ru = Object.keys(counts[team]["remark_unf"]);
    let cr = Object.keys(counts[team]["cancel_reason"]);

    keys["install_status"] = mergeNoDuplicates(keys["install_status"], is);
    keys["remark"] = mergeNoDuplicates(keys["remark"], re);
    keys["remark_unf"] = mergeNoDuplicates(keys["remark_unf"], ru);
    keys["cancel_reason"] = mergeNoDuplicates(keys["cancel_reason"], cr);
  }
  keys["install_status"].sort();
  keys["remark"].sort();
  keys["remark_unf"].sort();
  keys["cancel_reason"].sort();
  return keys;
}


function sumCounts(counts, keys) {
  /**
   * Gets the sum of every subgroup/column over all the teams.
   * @param {Object}  counts      Count data for all teams.
   * @param {Object}  keys        2d object that contains the sorted keys for every key group.
   * @return {Object} counts_sum  Contains the same keys for group and subgroup as counts,
   *                              but the value is just the sum integer, not a new object.
   */
  var counts_sum = {};
  var dgroups = ["install_status", "remark", "remark_unf", "cancel_reason"];  // todo: make this a const global var.

  // loop through teams
  for (let ti=0; ti<keys["teams"].length; ti++) { // ti = team index
    // get the current team
    let cteam = keys["teams"][ti];
    // loop through groups of the team
    for (let gi=0; gi<dgroups.length; gi++) {
      let dgroup = dgroups[gi]; // data group
      let group_counts = counts[cteam][dgroup];
      let sub_groups = Object.keys(group_counts);
      if (!counts_sum.hasOwnProperty(dgroup)) {
        counts_sum[dgroup] = {};
      }
      // loop through subgroups of group for the team
      for (let sg=0; sg<sub_groups.length; sg++) {
        let sub_group = sub_groups[sg];
        if (!counts_sum[dgroup].hasOwnProperty(sub_group)) {
          counts_sum[dgroup][sub_group] = 0;
        }
        let ccount = group_counts[sub_group];
        counts_sum[dgroup][sub_group] += ccount;
      }
    }
  }
  return counts_sum;
}


function fillCounts(counts, params, keys, counts_sum, number_of_days) {
  /**
   * Fill the counts and sums into the spreadsheet.
   * @param {Object}  counts      Count data for all teams.
   * @param {Object}  params      contains the parameters for the target sheet.
   * @param {Object}  keys        2d object that contains the sorted keys for every key group.
   * @param {Object}  counts_sum  Contains the same keys for group and subgroup as counts,
   *                              but the value is just the sum integer, not a new object.
   * @param {int}     number_of_days  Amount of unique days worked.
   */
  // This function fills the sheet with the counts
  // counts[team][install_status/remark/remark_unf/cancel_reason] = dict met counts

  // Activating the spreadsheet for writing
  var spreadsheet = SpreadsheetApp.openById(params.parameter_spreadsheet);
  var sheet = spreadsheet.getSheetByName(params.parameter_sheetname);

  // Make headers
  var offsets = makeHeaders(sheet, keys, STARTROW);

  // Adding objects to spreadsheet
  for(let ti=0; ti<keys["teams"].length; ti++) { // ti = team index
    let team = keys["teams"][ti];
    
    // add team names
    addNameToSpreadsheet(sheet, ti, team);

    // add install status / remark / cancel_reason / remark_unf
    addObjToSpread(sheet, ti, offsets["install_status"], keys["install_status"], counts[team]["install_status"]);
    addObjToSpread(sheet, ti, offsets["remark"], keys["remark"], counts[team]["remark"]);
    addObjToSpread(sheet, ti, offsets["cancel_reason"], keys["cancel_reason"], counts[team]["cancel_reason"]);
    addObjToSpread(sheet, ti, offsets["remark_unf"], keys["remark_unf"], counts[team]["remark_unf"]);
  }

  // Make a header at the bottom for the sums
  var offsets = makeHeaders(sheet, keys, STARTROW+keys["teams"].length+3);

  // correct the day rates
  counts_sum["install_status_adjusted"] = {};
  counts_sum["install_status_adjusted"][""] = counts_sum["install_status"][""];
  counts_sum["install_status_adjusted"]["DAYS"] = number_of_days;
  counts_sum["install_status_adjusted"]["INSTALLED"] = counts_sum["install_status"]["INSTALLED"];
  counts_sum["install_status_adjusted"]["INSTALLED/DAY"] = counts_sum["install_status"]["INSTALLED"]/number_of_days;
  counts_sum["install_status_adjusted"]["RSO"] = counts_sum["install_status"]["RSO"];
  counts_sum["install_status_adjusted"]["RSO/DAY"] = counts_sum["install_status"]["RSO"]/number_of_days;

  // Add the sums to the spreadsheet
  addObjToSpread(sheet, keys["teams"].length+4, offsets["install_status"], keys["install_status"], counts_sum["install_status_adjusted"]);
  addObjToSpread(sheet, keys["teams"].length+4, offsets["remark"], keys["remark"], counts_sum["remark"]);
  addObjToSpread(sheet, keys["teams"].length+4, offsets["cancel_reason"], keys["cancel_reason"], counts_sum["cancel_reason"]);
  addObjToSpread(sheet, keys["teams"].length+4, offsets["remark_unf"], keys["remark_unf"], counts_sum["remark_unf"]);

  // add label to right amount of days
  sheet.getRange(keys["teams"].length+STARTROW+6, 1, 1, offsets["install_status"]-1).mergeAcross().setValue("Corrected sums");

  // old code one with duplicate days
  // addObjToSpread(sheet, keys["teams"].length+6, offsets["install_status"], keys["install_status"], counts_sum["install_status"]);
  // sheet.getRange(keys["teams"].length+STARTROW+8, 1, 1, offsets["install_status"]-1).mergeAcross().setValue("Sums with duplicate days");
  
}


function addNameToSpreadsheet(sheet, ti, team_name) {
  /**
   * Adds the name of the team to the spreadsheet for the data row.
   * @param {Object}  sheet       Sheet object, where the data will be written.
   * @param {integer} ti          Index of the current team that is handeld.
   * @param {String}  team_name   index of the current team that is handeld.
   */
  var team_members = team_name.split("/");
    for (var tm=0; tm<team_members.length; tm++) {
      // Browser.msgBox("#"+team_members + "#"+ team_members[tm] + "#");
      sheet.getRange(ti+STARTROW+2, tm+1).setValue(team_members[tm]);
    }
}


function addObjToSpread(sheet, ti, offset, keys, data) {
  /**
   * Used to put individual counts and sums on the spreadsheet.
   * @param {Object}  sheet   Sheet object, where the data will be written.
   * @param {integer} ti      Index of the current team that is handeld.
   * @para, {integer} offset  The first column where the data should be written.
   * @param {Object}  keys    2d object that contains the sorted keys for every key group.
   * @param {list}    data    contains the counts of the current sub groups for the current team to be written in the sheet.
   */
  for (let ki=0; ki<keys.length; ki++) {
    let title = keys[ki];
    // check if title excists
    if(data.hasOwnProperty(title)) {
      let count = data[title];
      let col_index = ki + offset;
      let row_index = ti + STARTROW+2;
      // add data
      sheet.getRange(row_index, col_index).setValue(count);
    }
  }
}


function makeHeaders(sheet, keys, start_row) {
  /**
   * Sets headers for the counts
   * @param {Object}  sheet       Where to write the headers to.
   * @param {Object}  keys        2d object that contains the sorted keys for every key group.
   * @param {integer} start_row   Where to place the header.
   */
  // sets all the headers and subheaders into the sheet.

  // constants
  const group_height = start_row;
  const header_height = start_row+1;
  // declaration of lengths
  var len_names = getMaxTeamSize(keys["teams"]);
  var len_inst_stat = keys["install_status"].length;
  var len_remarks = keys["remark"].length;
  var len_can_reas = keys["cancel_reason"].length;
  var len_remar_unf = keys["remark_unf"].length;

  // Initiate variables to keep track of the offsets for the headers, subheaders and location of vars
  var offsets = {};
  var start_group = 1;
  var start_subgroup = 1+len_names;

  // set teams group header
  sheet.getRange(group_height, start_group, 1, len_names).mergeAcross().setValue("teams");
  start_group += len_names;

  // loop through install status / remarks / cancel reasons / unfiltered remarks
  const group_keys = ["install_status", "remark", "cancel_reason", "remark_unf"];
  const group_titles = ["install status", "remarks", "cancel reasons", "unfiltered remarks"];
  const group_lengths  = [len_inst_stat, len_remarks, len_can_reas, len_remar_unf];
  const colors_header = ["#dcf3ff", "#a2d2df", "#d6d6d6", "#f4f4f4", "#ffffff"];
  const colors_sub = ["#e2ffe1", "#f7e7e7", "#fcffe7", "#e3fdfd", "#f3eeff"];
  // color palettes: https://www.color-hex.com/color-palettes/
  
  for (let x=0; x<group_titles.length; x++) {
    // set header
    if (group_lengths[x] >= 1) {
      sheet.getRange(group_height, start_group, 1, group_lengths[x]).mergeAcross().setValue(group_titles[x]).setBackground(colors_header[x % colors_header.length]);
      
      // loop through sub headers
      offsets[group_keys[x]] = start_subgroup; // to keep track of offsets for filling in values
      for (let y=0; y<keys[group_keys[x]].length; y++) {
        // set sub headers
        sheet.getRange(header_height, start_subgroup+y).setValue(keys[group_keys[x]][y]).setBackground(colors_sub[y % colors_sub.length]);;
      }
      // keep track of 
      start_group += group_lengths[x];
      start_subgroup += group_lengths[x];
    }
    // else {
    //   // need to do somewthing here
    //   Browser.msgBox("the following group length index 0: "+x);

    // }
  }
  return offsets;
}


function getMaxTeamSize(teams) {
  /**
   * Takes the list of teams and checks what the size is of the biggest team.
   * @param {list}  teams     Contains a list of the cleaned up names of the teams.
   * @return {list} max_size  The size of the biggest team.
   */
  var max_size = 0;
  for (x=0; x<teams.length; x++) {
    let size = teams[x].split("/").length;
    if (size > max_size) {
      max_size = size;
    }
  }
  return max_size;
}
