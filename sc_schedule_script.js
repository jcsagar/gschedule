/**
 * Smith ML Automated Scheduler
 * JCR -- (jcsagar@gmail.com)
 * ver 1.51b
 * MIT License
 * 
 * Basic scheduling with Google spreadsheet and Google forms. Can be extended as necessary.
 */


function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  var entries = [
    {
    name : "Schedule",
    functionName : "schedule"
    },{
    name : "Export to Calendar",
    functionName : "exportToCalendar"
    },{
    name : "Clear the Calendar",
    functionName : "clearCalendar"
    }
  ];
  spreadsheet.addMenu("BCM Rad", entries);
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

// Take input responses from form, and output to schedule
function schedule() {
  
  var spreadsheet = SpreadsheetApp.getActive();
  var formsheet = spreadsheet.getSheetByName('FormResponses');
  var schedulesheet = spreadsheet.getSheetByName('Schedule');
  var adminsheet = spreadsheet.getSheetByName('Admin');
  
  var currentmonth = adminsheet.getRange("M1").getValue();
  
  var monthname = {
    'January' : '1',
    'February' : '2',
    'March' : '3',
    'April' : '4',
    'May' : '5',
    'June' : '6',
    'July' : '7',
    'August' : '8',
    'September' : '9',
    'October' : '10',
    'November' : '11',
    'December' : '12'
  }
  
  Logger.log(currentmonth);
  
  
  /* Create date hash */
  var scheduleRange = schedulesheet.getRange("A2:C400")
  var scheduleData = scheduleRange.getValues();
  
  var dateHash = {};
  
  // Create hash table with values representing who is available for each date
  for (x=0; x < scheduleData.length; x++) {

    // Set row
    var r = scheduleData[x];
      
    // Skip conditions
    if (r[0] != "" || r[1] == "") continue; // skip if date is invalid or blank
    
    // Check validity of date/month
    var tempdate = new Date(r[1]); tempdate.setHours(tempdate.getHours() + 8); // Daylight savings hotfix
    if (tempdate.getMonth()+1 != monthname[currentmonth]) continue;
    
    Logger.log("Date: " + tempdate.toString());
    
    dateHash[tempdate.getDate()] = [];
    
  }
  Logger.log("Datehash: " + dateHash.toString());
  
  
  /* Add to date hash based on responses */
  var responseData = formsheet.getRange("A2:E1000").getValues();
  var pool_of_people = [];
  
  var len_responseData = responseData.length;
  for (x=len_responseData-1; x >= 0 ; x--) {
    
    var r = responseData[x];
    
    if ( r[0] == "" || r[2] != currentmonth || pool_of_people.indexOf(r[1]) != -1 ) continue;
    
    pool_of_people.push(r[1]);
    
    var temparr = r[3];
    temparr = temparr.split(' ').join('').split(',');
    
    //Logger.log(temparr);
    
    for (i=0; i < temparr.length; i++) {
      //
      //Logger.log(i);
      if(dateHash.hasOwnProperty(temparr[i]))
        dateHash[temparr[i]].push(r[1]);
    }
  }
  
  Logger.log(dateHash);    
  Logger.log("--");
  Logger.log(dateHash[3]);
  
  /* Get people info, and hash it */
  var peopleData = adminsheet.getRange("A2:D50").getValues();
  //Logger.log(peopleData)
  //var pTally = {5:[], 4:[], 3:[], 2:[]};
  var pTally = [];
  for (i=1; i<peopleData.length; i++) {
    var eRow = peopleData[i];
    if (eRow[0] == "" || pool_of_people.indexOf(eRow[0]) == -1) continue;
    //pTally[eRow[1]].push( [eRow[0],eRow[2]] );
    pTally.push( [eRow[0],eRow[2]] );
  }
  
  pTally.sort( function(a,b) { return a[1] - b[1]; } );  
  Logger.log(pTally);
  //return;
  
  //Logger.log(peopleHash);
  
  var sA = shuffled_array(scheduleData.length);
  
  /* Place people */
  for (i=0; i < scheduleData.length; i++) {
    
    var x = sA[i];
    
    // Set row
    var r = scheduleData[x];
    
    // Skip conditions
    if (r[0] != "" || r[1] == "") continue; // skip if date is invalid or blank
    
    // Check validity of date/month
    var tempdate = new Date(r[1]); tempdate.setHours(tempdate.getHours() + 8); // Daylight savings hotfix
    if (tempdate.getMonth()+1 != monthname[currentmonth]) continue;
    
    Logger.log("Acting on date: " + r[1]);

    Logger.log(dateHash[tempdate.getDate()]);
    //Logger.log("--");


    // Sort tally list
    pTally.sort( function(a,b) { return a[1] - b[1]; } );  
    //Logger.log(pTally);
    
    // Place someone
    for (p=0; p < pTally.length; p++) {
      index = dateHash[tempdate.getDate()].indexOf(pTally[p][0]);
      //Logger.log(index);
      var tempvar = pTally[p][0];
      //Logger.log(tempvar);
      if (index != -1) {
        r[2] = pTally[p][0];
        pTally[p][1] = pTally[p][1] + 1;
        break;
      }
    }
    Logger.log("Placed: " + r[2] + " on " + r[1]);
    Logger.log("---------------");  
  }
  
  
  /* BEGIN SWAP SEGMENT - Iterate through dates one more time, and this time swap people */
  pTally.sort( function(a,b) { return a[1] - b[1]; } );  
  Logger.log(pTally);
  
  if (pTally[pTally.length-1][1] - pTally[0][1] > 1){
    
    Logger.log("Attempt swap");
        
    var overPerson = pTally[pTally.length-1][0];
    var underPerson = pTally[0][0];
    
    var overPINDX = 1;
    var count = 0;
    
    while ( (pTally[pTally.length-1][1] - pTally[0][1]) > 1 ){
      
      var swap_happened = -1; // keep track of whether a swap happened (for scenario when same people remain as over/under)
      
      for (x=0; x < scheduleData.length; x++) {
        
        // Set row
        var r = scheduleData[x];
        
        // Skip conditions
        if (r[0] != "" || r[1] == "") continue; // skip if date is invalid or blank
        
        // Check validity of date/month
        var tempdate = new Date(r[1]); tempdate.setHours(tempdate.getHours() + 8); // Daylight savings hotfix
        if (tempdate.getMonth()+1 != monthname[currentmonth]) continue;
        
        // skip the date if overperson is not working
        if (r[2] != overPerson) continue;
        
        // skip the date if underperson is already working
        if (r[2] == underPerson) continue;

        index = dateHash[tempdate.getDate()].indexOf(underPerson);
        
        if (index != -1) {
          r[2] = underPerson;
          
          pTally[0][1] = pTally[0][1] + 1;
          pTally[pTally.length - overPINDX][1] = pTally[pTally.length - overPINDX][1] - 1;
          swap_happened = 1;
          
          Logger.log("For date : " + tempdate.toString());
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
  
  /* Set schedule */
  scheduleRange.setValues(scheduleData);

}

var calId = "?????@group.calendar.google.com";


/**
 * Export events from spreadsheet to calendar
 */
function exportToCalendar() { 
  var spreadsheet = SpreadsheetApp.getActive();
  
  var schedulesheet = spreadsheet.getSheetByName('Schedule');
  var calsheet = spreadsheet.getSheetByName('calData');
  
  // Check for change in mastersheet data, and if so, update the calsheet
  var tempRange1 = schedulesheet.getRange("A:E").getValues();
  var tempD2 = calsheet.getRange("A:E");
  var tempRange2 = tempD2.getValues();
  
  var changed = 0;
  for (a=0; a<tempRange1.length; a++) {
    // load row
    var row1 = tempRange1[a];
    var row2 = tempRange2[a];
    if (row1[1] == "") continue; // skip if row blank
    for (b=0; b < row1.length; b++) {
      if (row1[b] == row2[b]) continue;
      else {
        if (b == 1) {
          var D1 = new Date(row2[b]); D1.setHours(D1.getHours() + 8); // Daylight savings hotfix
          var D2 = new Date(row1[b]); D2.setHours(D2.getHours() + 8); // Daylight savings hotfix
          //Logger.log(D1.getMonth() + " " + D1.getDate() + " " + D1.getYear());
          if (D1.getMonth() == D2.getMonth() && D1.getDate() == D2.getDate() && D1.getYear() == D2.getYear()) continue;
        }
        changed = 1;
        row2[b] = row1[b];
        Logger.log("Changes noted");
        break;
      }
    }
  }
  
  if (changed == 0) {
    Logger.log("No changes noted. Ending.");
    return;
  }
  
  // Update calsheet only if change in mastersheet 
  tempD2.setValues(tempRange1);
  
  //return;
  
  var headerRows = 2;  // Number of rows of header info (to skip)
  var range = calsheet.getRange("A1:G400");
  var data = range.getValues();
  
  // Access Calendar
  var cal = CalendarApp.getCalendarById(calId);
  
  var pauser = 0;  
  for (i=headerRows; i<data.length; i++) {

    // load row
    var row = data[i];
    if (row[1] == "") continue; // skip if row is a blank date
    
    // Define Date
    var date_value = new Date(row[1]); date_value.setHours(date_value.getHours() + 8); // Daylight savings hotfix
    
    
    // Make/Update Event -------------------------------------------------

    // CalendarEventID
    var eventID = row[6];
    Logger.log('Date: ' + date_value);
    var name = row[2];
    
    // First see if the spreadsheet has any event ID
    if (eventID == "") {
      // If blank name, then just skip it
      if (row[2] == "") continue;
      
      // Otherwise, a name is defined... so create an event with name
      var eventObj = cal.createAllDayEvent(name, date_value);
      row[6] = eventObj.getId();
      Logger.log('Event ID: ' + row[6]);
      
      // Need to pause once in a while because Google limits actions/second
      pauser = pauser + 1;
      if (pauser > 10) {
        Utilities.sleep(3000); 
        pauser = 0;
      }
      
    } else {
      //
      try {
        var eventObj = cal.getEventSeriesById(eventID);
        if (eventObj.getTitle() != name)
          eventObj.setTitle(name);
      }
      catch (e) {
        // do nothing - just avoiding stupid exception
      }
      
      // Need to pause once in a while because Google limits actions/second
      pauser = pauser + 1;
      if (pauser > 10) {
        Utilities.sleep(3000); 
        pauser = 0;
      }
      
    }
  }
          
  // Record all event IDs to spreadsheet
  range.setValues(data);
  
}


/**
 * Invoke to clear all events from the calendar
 */
function clearCalendar(){
  // Access Calendar
  var cal = CalendarApp.getCalendarById(calId);
  
  // Clear Calendar
  var fromDate = new Date(2016,0,1,0,0,0);
  var toDate = new Date(2030,0,1,0,0,0);
  var events = cal.getEvents(fromDate, toDate);
  var pauser = 0;
  for(var i=0; i<events.length;i++){
    pauser = pauser + 1;
    if (pauser > 10) {
      Utilities.sleep(3000); 
      pauser = 0;
    }
    var ev = events[i];
    //Logger.log(ev.getTitle()); // show event name in log
    ev.deleteEvent();
  }
  
  var spreadsheet = SpreadsheetApp.getActive();  
  var calsheet = spreadsheet.getSheetByName('calData');
  
  calsheet.clearContents();
  
}
