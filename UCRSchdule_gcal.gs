function gcalToSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var cal = CalendarApp.getCalendarById("shuyama240@gmail.com");
  var events = cal.getEvents(new Date("8/1/2018 12:00 AM"), new Date("8/30/2018 11:59 PM"));
  var titles = ["first name", "last name", "date1", "day1", "time1_start", "time1_end", "date2", "day2", "time1_start", "time2_end"];
  for (var i = 0; i < titles.length; i++){
    ss.getRange(1, i + 1).setValue(titles[i]);
  }
  for (var i = 0; i < events.length; i++){
    var name = events[i].getTitle();
    var first_last = name.split(" ");
    var firstName = first_last[0];
    var lastName = first_last[1];
    var start = events[i].getStartTime();
    var end = events[i].getEndTime();
    var day = start.getDate();
    ss.getRange(i + 2, 1).setValue(firstName);
    ss.getRange(i + 2, 2).setValue(lastName);
    ss.getRange(i + 2, 3).setValue(start);
    ss.getRange(i + 2, 3).setNumberFormat("mm/dd/yyyy");
    ss.getRange(i + 2, 4).setValue(day);
    ss.getRange(i + 2, 5).setValue(start);
    ss.getRange(i + 2, 5).setNumberFormat("h:mm AM/PM");
    ss.getRange(i + 2, 6).setValue(end);
    ss.getRange(i + 2, 6).setNumberFormat("h:mm AM/PM");
  
  }

}