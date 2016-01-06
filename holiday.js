/**
 * Pull all holidays out of the Meanbee Holiday Calendar and insert it into a raw data spreadsheet.
 */
function fetch_holidays_v2() {
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
    
    var event = events[i];
    
    var event_title         = event.getTitle();
    var event_start_day     = event.getStartTime();
    var event_end_day       = event.getEndTime();
    
    var date_current_day    = event_start_day;
    
    /* Loop through each day in the event and create a new row in the spreadsheet for it */ 
    while (date_current_day < event_end_day) {
      var staff_member = event_title.split(/:/)[0];
      var is_half_day  = is_event_half_day(event, date_current_day);
      var is_working_day = is_date_working_day(date_current_day);
      var is_holiday   = is_event_holiday(event);
      
      if (is_holiday && is_working_day) {
        var event_duration_working_days = (is_half_day) ? 0.5 : 1;
        
        var row_data = [
          staff_member,
          event_title,
          event_duration_working_days,
          date_current_day,
          date_current_day,
          is_half_day
        ];
        
        var range = sheet.getRange(row_number++, 1, 1, row_data.length);
        
        range.setValues([row_data]);
      }
      
      date_current_day.setDate(date_current_day.getDate() + 1);
    }
  }
}

function is_event_holiday(event) {
  return !!event.getTitle().match(/holiday/i);
}

function is_event_half_day(event, event_start_time) {
  var half_day_in_title = !!event.getTitle().match(/half day/i);
  var is_friday         = event_start_time.getDay() == 5;
  
  return half_day_in_title || is_friday;
}

function is_date_working_day(date) {
  var day = date.getDay();
  
  return day >= 1 && day <= 5;
}
