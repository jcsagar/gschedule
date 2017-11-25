/**
 * BCMRAD CALL SCHEDULE AUTOMATION FUNCTIONS
 * JCR -- (jcsagar@gmail.com)
 * ver 2.4b
 * 
 * - MIT License. Free to use.
 *
 */


// Date vars //

// Academic year
var acad_year1 = SpreadsheetApp.getActive().getSheetByName('R1-MASTER').getRange("N2").getValue(); 
var acad_year1b = acad_year1 + 1;

var acad_year2 = SpreadsheetApp.getActive().getSheetByName('R2-MASTER').getRange("N2").getValue(); 
var acad_year2b = acad_year2 + 1;

//var acad_year3 = SpreadsheetApp.getActive().getSheetByName('R3-MASTER').getRange("E1").getValue(); 
//var acad_year3b = acad_year2 + 1;

// R1 PFR
var r1_pfr_beginDate = new Date("7/1/"+acad_year1);
var r1_pfr_endDate = new Date("6/30/"+acad_year1b);

// R2 PFR
var r2_pfr_beginDate = new Date("7/1/"+acad_year2);
var r2_pfr_endDate = new Date("9/30/"+acad_year2);

// R2 NEURO
var r2_neuro_beginDate = new Date("7/1/"+acad_year2);
var r2_neuro_endDate = new Date("6/30/"+acad_year2b);

// R2 BODY
var r2_body_beginDate = new Date("3/1/"+acad_year2b);
var r2_body_endDate = new Date("6/30/"+acad_year2b);


// Google Calendar ID
//var calId = "1@group.calendar.google.com";

// On spreadsheet open, setup the menus...
function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  var entries = [
    {
    name : "AutoSchedule R1",
    functionName : "autoSchedule_R1"
    },    
    {
    name : "AutoSchedule R2",
    functionName : "autoSchedule_R2"
    }
//    ,
//    {
//    name : "Export to BCMRAD C/O 2019 Calendar",
//    functionName : "exportCallSchedule"
//    }
//    ,
//    {
//    name : "Clear the Calendar",
//    functionName : "clearCalendar"
//    }
//    ,
//    {
//    name : "Email Test",
//    functionName : "emailReminders"
//    },

                ];
  spreadsheet.addMenu("BCM RAD", entries);
}


//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// CALENDAR STUFF
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

// deprecated & deleted



///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Generate Calendar Index for use with auto-scheduling month requests
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
function CalIndxGen(academic_year, academic_year2){

  var acad_months = {"July":["7/1/"+academic_year, "7/31/"+academic_year], 
                     "August":["8/1/"+academic_year, "8/31/"+academic_year], 
                     "September":["9/1/"+academic_year, "9/30/"+academic_year], 
                     "October":["10/1/"+academic_year, "10/31/"+academic_year], 
                     "November":["11/1/"+academic_year, "11/30/"+academic_year], 
                     "December":["12/1/"+academic_year, "12/31/"+academic_year], 
                     "January":["1/1/"+academic_year2, "1/31/"+academic_year2], 
                     "February":["2/1/"+academic_year2, "2/28/"+academic_year2], 
                     "March":["3/1/"+academic_year2, "3/31/"+academic_year2], 
                     "April":["4/1/"+academic_year2, "4/30/"+academic_year2], 
                     "May":["5/1/"+academic_year2, "5/31/"+academic_year2],
                     "June":["6/1/"+academic_year2, "6/30/"+academic_year2],
                    }
  
  var retCalIndx = {};
  
  for (var month in acad_months) {
    if (acad_months.hasOwnProperty(month)) {
      //Logger.log(acad_months[month][0])
      retCalIndx[month] = [];
      retCalIndx[month].push(new Date(acad_months[month][0]) );
      retCalIndx[month].push(new Date(acad_months[month][1]) );
    }
  }
  
  //Logger.log(retCalIndx);
  return retCalIndx;

}


//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// COMMON FUNCTIONS 
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

function CheckToProceed(display_string) {
  // Confirmation dialog box warning
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(display_string, ui.ButtonSet.YES_NO);
  
  // Process response
  if (response == ui.Button.YES) return true
  else return false;

}

// create valid_dates[], which is a list of all valid dates    
function create_valid_dates (data, beginDate, endDate) {
  
  var valid_dates = [];
  
  // Loop through dates
  for (i=0; i < data.length; i++) {
    
    // CHECK DATE VALIDITY //
    r = data[i]; // load row
    if (r[0] == "") continue; // skip if row is a blank date
    var date = new Date(r[0]); // First column - Date
    if (date < beginDate || date > endDate ) continue; // skip if date is out of range
    
    valid_dates.push(date); // push if valid
    
  }
  
  return valid_dates;
  
}


// **** Create dateHash function, for use with autoscheduler ****
function create_dateHash (valid_dates, person_data) {
  //
  // Setup date hash - date : [person1, person2, etc.]
  var dateHash = {};
  
  var sA = shuffled_array(person_data.length); // shuffled array
  var defArr = []; for (i in person_data) defArr.push(person_data[ sA[i] ]); // default array, to contain all 12 people in a random fashion
  for (i=0; i < valid_dates.length; i++) dateHash[valid_dates[i]] = defArr.slice(); // copy to each date

  return dateHash;
  
}


// **** Trim dateHash function, for use with autoscheduler ****
function trim_dateHash (valid_dates, dateHash, vacdata, CalIndx, nb_months) {
  
  // Initialize pTally : a tracker of how many shifts each person worked
  var pTally = []; // pTally = [ [Jesse, 0], [CJ, 2], etc. ]

  // Loop through columns/persons
  for (x=0; x < vacdata[0].length; x++) {
    
    var person = vacdata[0][x]; // define person    
    pTally.push( [person, 0] ); // initialize count for each person at 0
    
    if (nb_months > 0){
      // TRIM OUT MONTH REQUESTS //
      // Loop through MONTHS section
      for (y=1; y <= nb_months; y++) {  
        
        // Continue if cell not blank
        if (vacdata[y][x] == "") continue; // skip if blank
        
        try {
          var begin_mo = new Date( CalIndx[vacdata[y][x]][0] );
          var end_mo = new Date( CalIndx[vacdata[y][x]][1] );
        }
        catch (e) {
          //
          Logger.log("invalid month");
          Logger.log(e);
          Logger.log(person);
          Logger.log(vacdata[y][x]);
          Logger.log("------------");
          var ui = SpreadsheetApp.getUi();
          ui.alert("Invalid month "+vacdata[y][x]+" for person "+person, ui.ButtonSet.OK);
          throw new Error("Execution failed due to error!");
        }
        
        begin_mo.setDate(begin_mo.getDate() - 1);
        end_mo.setDate(end_mo.getDate() + 1);
  
        Logger.log("begin " + begin_mo);
        Logger.log("end " + end_mo);
        
        for (var date_key in dateHash) {
          if (dateHash.hasOwnProperty(date_key)) {
            var tmpdate = new Date(date_key);
            if (begin_mo < tmpdate && tmpdate < end_mo){
              //Logger.log("HERE!!! "+date_key);
              var people_assigned = dateHash[date_key];
              
              index = dateHash[date_key].indexOf(person);
  
              // Remove person as "valid option" from the date
              if(index != -1){
                dateHash[date_key].splice(index, 1);
              }
              
            }
          }
        }
      }
    
    }
    
    
    // TRIM OUT DATE REQUESTS //
    // Initialize this var for this segment
    for (y=7; y < vacdata.length; y++) {
      
      if (vacdata[y][x] == "") continue; // skip if blank
      
      var vDate = new Date(vacdata[y][x]); // date-ify it  
      var index = -1;
      
      // 
      try {
        // Find index of person for that date
        index = dateHash[vDate].indexOf(person);
      }
      catch (e) {
        Logger.log("Invalid index request:");
        Logger.log(e);
        Logger.log(person);
        Logger.log(vacdata[y][x]);
        Logger.log("------------");
        var ui = SpreadsheetApp.getUi();
        ui.alert("Invalid date "+vacdata[y][x]+" for person "+person, ui.ButtonSet.OK);
        throw new Error("Execution failed due to error!");
      }
      
      // Remove person as "valid option" from the date
      if(index != -1){
        dateHash[vDate].splice(index, 1);
      }
      
    }
    
  }
  
  return [dateHash, pTally];
  
}


// Random integer generation
function getRandomInt(min, max) {
    return Math.floor(Math.random() * (max - min + 1)) + min;
}


// Update function of pTally
function update_pTally(pTally, name, incr) {
  for (j=0; j < pTally.length; j++) {
    if (pTally[j] == name) {
      pTally[j][1] = pTally[j][1] + incr;
      break;
    }
  }
  return pTally;
}


// http://stackoverflow.com/questions/962802#962890
function shuffled_array(maxElements) {
  
  // create ordered array : 0,1,2,3..maxElements
  for (var temArr = [], i = 0; i < maxElements; i++) {
    temArr[i] = i;
  }
  
  for (var finalArr = [maxElements], i = 0; i < maxElements; i++) {
    // remove random element from the temArr and push it into finalArrr
    finalArr[i] = temArr.splice(Math.floor(Math.random() * (maxElements - i)), 1)[0];
  }
  
  return finalArr;
} 


//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// AUTOSCHEDULE R1 PFR
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

function autoSchedule_R1() {
  
  if (!CheckToProceed('This will erase any existing data on the R1-MASTER sheet. Are you sure you want to continue?')) return;
  
  // Access spreadsheet
  var spreadsheet = SpreadsheetApp.getActive();
  
  // Access MASTER sheet
  var mastersheet = spreadsheet.getSheetByName('R1-MASTER');
   
  // Clear range - ERASE ALL THE THINGZ
  mastersheet.getRange("B2:B100").clearContent();
  
  var range = mastersheet.getRange("A2:C100");
  var data;
  
  var diff = 100;
  for (var run=0; run < 30; run++) {
    [data, diff] = autoSchedule_R1_PFR(spreadsheet, mastersheet);
    if (diff == -1) return;
    Logger.log("OPTIMZATION SCORE: " + diff + " | RUN #: " + run);
    Logger.log("=-----------------------------------=");
    if (diff < 1) break;
  }
  
  // ******* Record all event IDs to spreadsheet ********* //
  range.setValues(data);
  
}


/**
 * Automatically schedule R1 PFR people based on vacation preferences
 */
function autoSchedule_R1_PFR(spreadsheet, mastersheet) {
  
  // Initialize some vars
  var sheet_name = 'R1-REQUESTS';
  
  var index;
  var error_collection = []; // error checking to return
  var pTally; // Tally for each person
  
  // Get mastersheet range/data (with dates as first column)
  var range = mastersheet.getRange("A2:C100");
  var data = range.getValues();
  
  // get valid_dates, ["7/1/2016", "7/2/2016", etc.]
  var valid_dates = [];
  
  // Loop through dates
  for (i=0; i < data.length; i++) {
    
    // CHECK DATE VALIDITY //
    r = data[i]; // load row
    if (r[0] == "") continue; // skip if row is a blank date
    var date = new Date(r[0]); // First column - Date
    if (date < r1_pfr_beginDate || date > r1_pfr_endDate ) continue; // skip if date is out of range
    
    valid_dates.push(date); // push if valid
    
  }
  
  var perdata = spreadsheet.getSheetByName(sheet_name).getRange("B1:M1").getValues()[0]; // get persons
  var vacation_data = spreadsheet.getSheetByName(sheet_name).getRange("B1:M").getValues(); // get range/data from vacation sheet
  
  // ******* Initialize dateHash with "everyone for everyday" --------------------------
  var dateHash = create_dateHash(valid_dates, perdata);
  
  var Calendar_Index = CalIndxGen(acad_year1, acad_year1b); // Generate Calendear-Index
  //Logger.log("Date Test : " + new Date( CalIndx[vacdata[y][x]][0] ) );
  
  // ******* Go through vacation requests, and trim out people from dateHash
  [dateHash, pTally] = trim_dateHash(valid_dates, dateHash, vacation_data, Calendar_Index, 0);
  
  
  // ******* Iterate through dates again, and this time place people --------------------------
  
  var sA = shuffled_array(data.length); // create shuffled array of numbers, length is up to number of data rows
  
  // Loop through dates
  for (n=0; n < data.length; n++) {
    
    var i = n;//sA[n]; // use i as a "random data row index"
    
    // CHECK DATE VALIDITY //
    r = data[i]; // load row
    if (valid_dates.map(Number).indexOf(+r[0]) < 0) continue; // skip if invalid date
    var dateval = new Date(r[0]);

    // SCHEDULE PFR //
    pTally.sort( function(a,b) { return a[1] - b[1]; } ); // sort pTally [least -> most]
    
    // if person is available for the date, place them
    for (j=0; j < pTally.length; j++) {
      //var tmpval = dateHash[dateval];
      if (dateHash[dateval].indexOf(pTally[j][0]) != -1) {
        r[1] = pTally[j][0];
        pTally[j][1] = pTally[j][1] + 1;
        break;
      }
    }
    
    // if no one was placed, then record error
    if (r[1] == "") {
      //ui.alert('ERROR - No one available to work on : ' + dateval.toString(), ui.ButtonSet.OK);
      var tmp_s = " " + (dateval.getMonth()+1).toString() + "/" + dateval.getDate().toString() + "/" + dateval.getFullYear().toString();
      error_collection.push(tmp_s);
      //return [-1, -1]; // return error code
    }
  
  } // END INITIAL SCHEDULE SEGMENT //
  
  
  // ERROR DISPLAY
  spreadsheet.getSheetByName(sheet_name).getRange("S2:S").clearContent();
  if (error_collection.length > 0) {
    var ui = SpreadsheetApp.getUi();
    ui.alert('ERROR --- INSUFFICIENT PEOPLE AVAILABLE FOR THE FOLLOWING DATES : \n' + error_collection.toString(), ui.ButtonSet.OK);
    var error_range = spreadsheet.getSheetByName(sheet_name).getRange("S2:S");
    var error_data = error_range.getValues();
    for (var e=0; e < error_collection.length; e++) {
      error_data[e][0] = error_collection[e];
    }
    error_range.setValues(error_data);
    return [-1, -1];
  }
  
  // ******* Calculate optimization of randomness, and return data ********* //
  pTally.sort( function(a,b) { return a[1] - b[1]; } );
  var optimization_score = pTally[pTally.length-1][1] - pTally[0][1];
  //return [data, optimization_score];
  
  // ******* Iterate through dates one more time, and this time swap people --------------------------
  
  if (pTally[pTally.length-1][1] - pTally[0][1] > 1){
    
    Logger.log(pTally);
        
    var overPerson = pTally[pTally.length-1][0];
    var underPerson = pTally[0][0];
    
    var overPINDX = 1;
    var count = 0;
    
    while ( (pTally[pTally.length-1][1] - pTally[0][1]) > 1 ){
      
      var swap_happened = -1; // keep track of whether a swap happened (for scenario when same people remain as over/under)
      
      for (i=0; i < valid_dates.length; i++) {
        
        // CHECK DATE VALIDITY //
        r = data[i]; // load row
        if (valid_dates.map(Number).indexOf(+r[0]) < 0) continue; // skip if invalid date
        var dateval = new Date(r[0]);
        
        if (r[1] != overPerson) continue; // skip the date if overperson is not working

        if (dateHash[dateval].indexOf(underPerson) != -1) {
          
          r[1] = underPerson;
          
          pTally[0][1] = pTally[0][1] + 1;
          pTally[pTally.length - overPINDX][1] = pTally[pTally.length - overPINDX][1] - 1;
          swap_happened = 1;
          
          Logger.log(dateval);
          Logger.log("oP:" + overPerson + " | uP:" + underPerson);
          
          break;
        }
      }
      
      pTally.sort( function(a,b) { return a[1] - b[1]; } );
      
      // CHECK SWAPS, AND TRY LOWER PERSON INDICES //
      if (swap_happened == 1){
        overPINDX = 1;
        overPerson = pTally[pTally.length - overPINDX][0]; underPerson = pTally[0][0];
      } else {
        overPINDX = overPINDX + 1; 
        if ( (pTally[pTally.length - overPINDX][1] - pTally[0][1]) > 1){
          overPerson = pTally[pTally.length - overPINDX][0]; underPerson = pTally[0][0];  
        } else {
          //ui.alert('Imbalanced Schedule', ui.ButtonSet.OK);
          break;
        }
      }
      
    }
    
  }
  // END SWAP SEGMENT
  
  //*/
  
  // ******* Calculate optimization of randomness, and return data ********* //
  pTally.sort( function(a,b) { return a[1] - b[1]; } );
  var optimization_score = pTally[pTally.length-1][1] - pTally[0][1];
  return [data, optimization_score];
  
}


///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// AUTOSCHEDULE R2 (MAIN FUNCTION)
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
function autoSchedule_R2() {
  
  if (!CheckToProceed('This will erase any existing data on the R2-MASTER sheet. Are you sure you want to continue?')) return;
  
  // Access spreadsheet
  var spreadsheet = SpreadsheetApp.getActive();
  
  // Access MASTER sheet
  var mastersheet = spreadsheet.getSheetByName('R2-MASTER');
   
  // Clear range
  mastersheet.getRange("B2:B300").clearContent();
  mastersheet.getRange("D2:D300").clearContent();
  mastersheet.getRange("F2:F300").clearContent();
   
  // Get R2 PFR data/range
  var range = mastersheet.getRange("A2:G300");
  var data = range.getValues();
  
  var diff = 100;  
  for (var run=0; run < 30; run++) {
    [data, diff] = autoSchedule_R2_PFRNEUROBODY(spreadsheet, mastersheet, data);
    if (diff == -1) return;
    Logger.log("OPTIMZATION SCORE: " + diff + " | RUN #: " + run);
    Logger.log("=-----------------------------------=");
    if (diff < 1) break;
  }
  
  // ******* Record all event IDs to spreadsheet ********* //
  range.setValues(data);
  
}


///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// AUTOSCHEDULE R2 PFR - NEURO - BODY
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
function autoSchedule_R2_PFRNEUROBODY(spreadsheet, mastersheet, data) {
  
  // Assign value for pfr weight
  var pfr_value = spreadsheet.getSheetByName('R2-MASTER').getRange("Q2").getValue(); 

  // get valid_dates, ["7/1/2016", "7/2/2016", ..., "6/30/2017"]
  var valid_dates = create_valid_dates (data, r2_neuro_beginDate, r2_neuro_endDate);
    
  
  // ******* Initialize dateHash with "everyone for everyday" --------------------------
  var perdata = spreadsheet.getSheetByName('R2-REQUESTS').getRange("B1:M1").getValues()[0]; // get persons
  var dateHash = create_dateHash(valid_dates, perdata);
  
  
  // ******* Go through vacation requests, and trim out people from dateHash
  var pTally;
  var vacation_data = spreadsheet.getSheetByName('R2-REQUESTS').getRange("B1:M").getValues(); // get range/data from vacation sheet
  var Calendar_Index = CalIndxGen(acad_year2, acad_year2b); // Generate Calendear-Index
  [dateHash, pTally] = trim_dateHash(valid_dates, dateHash, vacation_data, Calendar_Index, 6);
  
  
  // ******* Iterate through dates, and place people *******
  
  // Initialize some standard vars, for reuse later as well
  var dateval;
  var r;
  var index;
  
  // Error checking to return
  var error_collection = [];
  
  // create shuffled array of numbers, length is up to number of datevalvals
  var sA = shuffled_array(data.length);
  
  // Loop through dates
  for (x=0; x < data.length; x++) {
    
    var i = sA[x];

    // CHECK DATE VALIDITY //
    r = data[i]; // load row
    if (r[0] == "") continue; // skip if row is a blank date
    dateval = new Date(r[0]); // First column - Date
    if (dateval < r2_pfr_beginDate || dateval > r2_pfr_endDate ) continue; // skip if date is out of range
    
    // SCHEDULE PFR //
    pTally.sort(
      function(a,b) { return a[1] - b[1]; }
    );
    for (j=0; j < pTally.length; j++) {
      index = dateHash[dateval].indexOf(pTally[j][0]);
      if (index != -1) {
        r[1] = pTally[j][0];
        pTally[j][1] = pTally[j][1] + pfr_value; // default pfr_value = 0.5
        break;
      }
    }
    if (r[1] == "") {
      //ui.alert('ERROR - No one available to work on : ' + date.toString(), ui.ButtonSet.OK);
      var tmp_s = " " + (dateval.getMonth()+1).toString() + "/" + dateval.getDate().toString() + "/" + dateval.getFullYear().toString();
      error_collection.push(tmp_s);
      //return [-1, -1]; // return error code
    }
  
  }
  
  
  // SCHEDULE NEURO AND BODY //
  // create re-shuffled array of numbers, length is up to number of dates
  //var sA = shuffled_array(data.length); // RE-SHUFFLE DATES?
  
  // Loop through dates
  for (x=0; x < data.length; x++) {
    
    var i = sA[x];

    // CHECK DATE VALIDITY //
    r = data[i]; // load row
    if (r[0] == "") continue; // skip if row is a blank date
    dateval = new Date(r[0]); // First column - Date
    if (dateval < r2_neuro_beginDate || dateval > r2_neuro_endDate ) continue; // skip if date is out of range

    // SCHEDULE NEURO SHIFT //
    pTally.sort(
      function(a,b) { return a[1] - b[1]; }
    );
    for (var z=0; z < pTally.length; z++) {  
      index = dateHash[dateval].indexOf(pTally[z][0]);
      if (index != -1 && r[1] != pTally[z][0]) {
        r[3] = pTally[z][0];
        pTally[z][1] = pTally[z][1] + 1;
        break;
      }
    }
    if (r[3] == "") {
      //ui.alert('ERROR - No one available to work NEURO on : ' + dateval.toString(), ui.ButtonSet.OK);
      var tmp_s = " " + (dateval.getMonth()+1).toString() + "/" + dateval.getDate().toString() + "/" + dateval.getFullYear().toString();
      error_collection.push(tmp_s);
    }
    Logger.log(dateval); Logger.log(r2_body_beginDate);
    // SCHEDULE BODY SHIFT, IF VALID //
    if ( dateval > r2_body_beginDate || r[6] != "" ) {
      pTally.sort(
        function(a,b) { return a[1] - b[1]; }
      );
      for (var z=0; z < pTally.length; z++) {
        index = dateHash[dateval].indexOf(pTally[z][0]);
        if (index != -1 && r[1] != pTally[z][0] && r[3] != pTally[z][0] ) {
          r[5] = pTally[z][0];
          pTally[z][1] = pTally[z][1] + 1;
          break;
        }
      }
      if (r[5] == "") {
        //ui.alert('ERROR - No one available to work BODY on : ' + date.toString(), ui.ButtonSet.OK);
        var tmp_s = " " + (dateval.getMonth()+1).toString() + "/" + dateval.getDate().toString() + "/" + dateval.getFullYear().toString();
        error_collection.push(tmp_s);
      }
    }
    

  } // END INITIAL SCHEDULE SEGMENT //
  spreadsheet.getSheetByName('R2-REQUESTS').getRange("S2:S").clearContent();
  if (error_collection.length > 0) {
    var ui = SpreadsheetApp.getUi();
    ui.alert('ERROR --- INSUFFICIENT PEOPLE AVAILABLE FOR THE FOLLOWING DATES : \n' + error_collection.toString(), ui.ButtonSet.OK);
    var error_range = spreadsheet.getSheetByName('R2-REQUESTS').getRange("S2:S");
    var error_data = error_range.getValues();
    for (var e=0; e < error_collection.length; e++) {
      error_data[e][0] = error_collection[e];
    }
    error_range.setValues(error_data);
    return [-1, -1];
  }
  
  
  // ******* BEGIN SWAP SEGMENT - Iterate through dates one more time, and this time swap people
  pTally.sort(
      function(a,b) { return a[1] - b[1]; }
    );
  
  if (pTally[pTally.length-1][1] - pTally[0][1] > 1){
    
    Logger.log(pTally);
        
    var overPerson = pTally[pTally.length-1][0];
    var underPerson = pTally[0][0];
    
    var overPINDX = 1;
    var count = 0;
    
    while ( (pTally[pTally.length-1][1] - pTally[0][1]) > 1 ){
      
      var swap_happened = -1; // keep track of whether a swap happened (for scenario when same people remain as over/under)
      
      for (i=0; i < data.length; i++) {
        
        // CHECK DATE VALIDITY //
        r = data[i]; // load row
        if (r[0] == "") continue; // skip if row is a blank date
        dateval = new Date(r[0]); // First column - Date
        if (dateval < r2_neuro_beginDate || dateval > r2_neuro_endDate ) continue; // skip if date is out of range
        
        // skip the date if overperson is not working
        if (r[3] != overPerson && r[5] != overPerson) continue;
        
        // skip the date if underperson is already working
        if (r[1] == underPerson || r[3] == underPerson || r[5] == underPerson) continue;

        index = dateHash[dateval].indexOf(underPerson);
        
        if (index != -1) {
          if (r[3] == overPerson) {
            r[3] = underPerson;
          } else {
            r[5] = underPerson;
          }          
          pTally[0][1] = pTally[0][1] + 1;
          pTally[pTally.length - overPINDX][1] = pTally[pTally.length - overPINDX][1] - 1;
          swap_happened = 1;
          
          Logger.log(dateval);
          Logger.log("oP:" + overPerson + " | uP:" + underPerson);
          
          break;
        }
      }
      
      pTally.sort(
        function(a,b) { return a[1] - b[1]; }
      );
      
      // CHECK SWAPS, AND TRY LOWER PERSON INDICES //
      if (swap_happened == 1){
        overPINDX = 1;
        overPerson = pTally[pTally.length - overPINDX][0]; underPerson = pTally[0][0];
      } else {
        overPINDX = overPINDX + 1; 
        if ( (pTally[pTally.length - overPINDX][1] - pTally[0][1]) > 1){
          overPerson = pTally[pTally.length - overPINDX][0]; underPerson = pTally[0][0];  
        } else {
          //ui.alert('Imbalanced Schedule', ui.ButtonSet.OK);
          break;
        }
      }
      
    }
    
  }
  // END SWAP SEGMENT
  
  
  // ******* BEGIN BALANCE SEGMENT - Iterate through dates one more time, and this time swap within the same date to balance neuro and body
  
  // initialize bodyTally (which is to keep track of neuro vs body shifts, used in the balance segment)
  var bodyTally = {};
  for (b=0; b < pTally.length; b++) { 
    bodyTally[pTally[b][0]] = 0; 
  }
  
  // Iterate through dates and tally body shifts
  for (i=0; i < data.length; i++) {
    // CHECK DATE VALIDITY //
    r = data[i]; // load row
    if (r[0] == "" || r[5] == "") continue; // skip if row is a blank date or not assigned to body
        
    bodyTally[r[5]] = bodyTally[r[5]] + 1;
    
  }
  
  Logger.log(bodyTally);
  
  // Iterate through dates switch if criteria met
  var tmp_name;
  for (k=0; k < 12; k++) {    
    var swap_happened = -1;
    for (i=0; i < data.length; i++) {
      // CHECK DATE VALIDITY //
      r = data[i]; // load row
      if (r[0] == "" || r[5] == "") continue; // skip if row is a blank date or not assigned to body
      dateval = new Date(r[0]); // First column - Date
      
      // check and swap to balance body shifts
      if ( (bodyTally[r[5]] - bodyTally[r[3]]) > 1 ) {
        Logger.log(dateval); Logger.log("NEURO:" + r[3] + " | BODY:" + r[5]);
        bodyTally[r[5]] = bodyTally[r[5]] - 1;
        bodyTally[r[3]] = bodyTally[r[3]] + 1;
        tmp_name = r[3]; r[3] = r[5]; r[5] = tmp_name;
        swap_happened = 1;
      }
      
    }
    if (swap_happened == -1) break;  
  }


  // ******* Calculate optimization of randomness, and return data ********* //
  pTally.sort(
      function(a,b) { return a[1] - b[1]; }
    );
  
  var optimization_score = pTally[pTally.length-1][1] - pTally[0][1];
  
  return [data, optimization_score];
  

}
