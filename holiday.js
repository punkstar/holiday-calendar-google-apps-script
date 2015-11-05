/**
 * Pull all holidays out of the Meanbee Holiday Calendar and insert it into a raw data spreadsheet.
 */
function fetch_holidays(){
  var meanbee_holiday_calendar = 'calendarname@group.calendar.google.com';
  var meanbee_holiday_spreadsheet_id = 'your_spreadsheet_id';
  var meanbee_holiday_sheet_name = 'Raw Holiday Data';
  
  var cal = CalendarApp.getCalendarById(meanbee_holiday_calendar);
  var sheet = SpreadsheetApp.openById(meanbee_holiday_spreadsheet_id).getSheetByName(meanbee_holiday_sheet_name);
   
  var events = cal.getEvents(new Date("January 1, 2000"), new Date("December 31, 3000"));

  // 1 indexed, need the header row for ourselves.  
  var row_number = 1;
  
  sheet.clearContents();
  
  var header_data = ['Person', 'Description', 'Working Days', 'Event Start Date', 'Event End Date', 'Is Half Day'];
  var header_range = sheet.getRange(row_number++, 1, 1, header_data.length);
  
  header_range.setValues([header_data]);
  
  for (var i = 0; i < events.length; i++) {
    
    var event_title         = events[i].getTitle();
    var event_start_day     = events[i].getStartTime();
    var event_end_day       = events[i].getEndTime();
    
    // Call the accessors again to get a fresh object
    var event_duration_working_days = working_days_in_range(events[i].getStartTime(), events[i].getEndTime());
    
    var staff_member = event_title.split(/:/)[0];
    var is_half_day  = !!event_title.match(/half day/i);
    var is_holiday   = !!event_title.match(/holiday/i);
    
    
    if (!is_holiday) {
      continue; 
    }
    
    if (is_half_day) {
      event_duration_working_days = 0.5;
    }
    
    var row_data = [
      staff_member,
      event_title,
      event_duration_working_days,
      event_start_day,
      event_end_day,
      is_half_day
    ];
    
    var range = sheet.getRange(row_number++, 1, 1, row_data.length);
    
    range.setValues([row_data]);
  }
}

function working_days_in_range(start_range, end_range) {
  var day_in_ms = 24 * 60 * 60 * 1000;
  var current_day = start_range;
  var working_days = 0;
  
  while (current_day < end_range) {
    var day_of_week = current_day.getDay();
    
    if (day_of_week >= 1 && day_of_week <= 4) {
      working_days += 1.0;
    } else if (day_of_week == 5) {
      working_days += 0.5; 
    }
    
    current_day.setDate(current_day.getDate() + 1);
  }
  
  return working_days;
}
