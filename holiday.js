/**
 * Pull all holidays out of the Holiday Calendar and insert it into a raw data spreadsheet.
 */
function fetch_holidays(){
  var meanbee_holiday_calendar = 'yourcalendarurl@group.calendar.google.com';
  var meanbee_holiday_spreadsheet_id = 'yourspreadsheetid';
  var meanbee_holiday_sheet_name = 'Raw Holiday Data';
  
  var cal = CalendarApp.getCalendarById(meanbee_holiday_calendar);
  var sheet = SpreadsheetApp.openById(meanbee_holiday_spreadsheet_id).getSheetByName(meanbee_holiday_sheet_name);
   
  var events = cal.getEvents(new Date("January 1, 2000"), new Date("December 31, 3000"));
  
  for (var i = 0; i < events.length; i++) {
    
    var event_title         = events[i].getTitle();
    var event_start_day     = events[i].getStartTime();
    var event_end_day       = events[i].getEndTime();
    
    var row_number = i + 2; // Because we want the header row for ourselves
    
    var staff_member = event_title.split(/:/)[0];
    var is_half_day  = !!event_title.match(/half day/i);
    var event_duration_working_days = working_days_in_range(event_start_day, event_end_day);
    
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
    
    var range = sheet.getRange(row_number, 1, 1, row_data.length);
    
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
