function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Sync to Calendar')
      .addItem('Create Events', 'createEvents')
      .addSeparator()
      .addSubMenu(ui.createMenu('Delete')
          .addItem('Clear Events', 'clearCalendar'))
      .addToUi();
}


/**
 * check if a JavaScript value is a Date object
 *
 * See https://stackoverflow.com/a/44198641/297797
 */
function isValidDate(date) {
  return date && Object.prototype.toString.call(date) === "[object Date]" && !isNaN(date);
}

/**
 *
 * get the value defined by a named range in the active spreadsheet
 *
 * Parameters:
 *   name (string) - the name of the range
 * 
 * Returns:
 *   - the content of the (first cell of the) named range.  The value may 
 *     be of type Number, Boolean, Date, or String depending on the value
 *     of the cell. Empty cells return an empty string.
 *
 *   - a null value if the range isn't found.
 *
 */
function getNamedRangeValue(name) {
  var namedRanges = SpreadsheetApp.getActive().getNamedRanges();
  var range;
  for (var i = 0; i < namedRanges.length; i++) {
    range = namedRanges[i];
    if (range.getName() == name) {
      break;
    }
  }
  value = range.getRange().getValue();
  Logger.log(value)
  return value;
}

/**
 * Get the calendar associated to this spreadsheet
 */
function getCalendar() {  
  var spreadsheet = SpreadsheetApp.getActive();
  var calendarId = getNamedRangeValue('calendarID');
  var calendar = CalendarApp.getCalendarById(calendarId);
  return calendar;
}

/**
 * Get a list of event records (not Event objects)
 *
 * Looks in a sheet called 'events'
 **/
function getEventRecords() {
  var sheet = SpreadsheetApp.getActive().getSheetByName('events');
  var data = sheet.getDataRange().getValues()
  var events = []
  
  // This loop parses the data from the spreadsheet and adds it to an array
  // Row 0 (Row 1 on the spreadsheet) is the header row
  for (var i = 1; i < data.length; i++) {
    // Skips blank rows
    if (data[i][0] == "") {
      break
    }
    var event = {};
    // Column A is the event's title, a string
    event.title = data[i][0];
    // Column B is the event's description, a string
    event.description = data[i][1];
    // Column C is the day of the event, and Column D is the time.
    // Both get converted by AppScript into JavaScript Date objects (that include time),
    // except the first has the wrong time (00:00:00) 
    // and the second has the wrong date (1970-01-01 00:00:00 UTC)
    var date = data[i][2];
    var start = data[i][3];
    const MILLISECONDS_PER_HOUR = 60*60*1000;
    const MILLISECONDS_PER_MINUTE = 60*1000;
    event.startDate = new Date(date.getTime() 
                               + start.getHours()*MILLISECONDS_PER_HOUR
                               + start.getMinutes()*MILLISECONDS_PER_MINUTE);
    // Column E has the duration, either another JS date or a number of minutes
    var dur = data[i][4];
    var duration = 0; // actual duration in milliseconds
    if (isValidDate(dur)) {
      duration = dur.getHours()* MILLISECONDS_PER_HOUR
                 + dur.getMinutes() * MILLISECONDS_PER_MINUTE;
    } else {
      duration = dur * MILLISECONDS_PER_MINUTE; 
    }
    event.endDate = new Date(event.startDate.getTime() + duration);
    // Column F has the 'type' (used by Sakai to set the icon);
    event.type = data[i][5];
    // Column G has the location
    event.location = data[i][6];
    Logger.log(event);
    events.push(event);  
  }
  return events;
}

function createEvents() {
  // get the calendar.  It's defined in the spreadsheet as a named range.
  var calendar = getCalendar();
  var recs = getEventRecords();
  for (var i = 0; i < recs.length; i++) {
    rec = recs[i]
    calendar.createEvent(rec.title,rec.startDate,rec.endDate,
                           {location: rec.location, description: rec.description});
    // slow down the script to keep away an error message about creating too many events quickly                       
    if (i % 10 == 9) {
      Utilities.sleep(3000);
    }
  }
}

function clearCalendar() {
  // This one removes all of the shifts from the event
  // Calendar, so I put it in a sub-menu to make it
  // difficult to click by accident!
  var calendar = getCalendar();
  events = calendar.getEvents(new Date(1999,0), new Date(2099,0));
  for (var i = 0; i < events.length; i++) {
    events[i].deleteEvent();
    // slow down the script to keep away an error message about deleting too many events quickly                       
    if (i % 10 == 9) {
      Utilities.sleep(3000);
    }
  }
}