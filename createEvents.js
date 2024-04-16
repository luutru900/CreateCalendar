function createCalendarEvents() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var calendar = CalendarApp.getDefaultCalendar();
  var schedule = sheet.getRange("A2:C" + sheet.getLastRow()).getValues();
  
  // Định nghĩa ngày bắt đầu và kết thúc cho phạm vi tạo sự kiện
  var startDate = new Date(2024, 3, 17); // Ngày bắt đầu
  var endDate = new Date(2024, 7, 20);   // Ngày kết thúc

  // Tạo sự kiện cho mỗi ngày từ startDate đến endDate
  for (var day = startDate; day <= endDate; day.setDate(day.getDate() + 1)) {
    for (var i = 1; i < schedule.length; i++) { // Bắt đầu từ 1 để bỏ qua hàng tiêu đề
      var row = schedule[i];
      var eventTitle = row[2]; // Tên sự kiện
      
      // Lấy giờ bắt đầu và kết thúc từ bảng tính và thêm vào ngày hiện tại
      var eventStartTime = parseTime(row[0], day);
      var eventEndTime = parseTime(row[1], day);

      // Tạo sự kiện trên lịch nếu thời gian kết thúc sau thời gian bắt đầu
      if (eventEndTime > eventStartTime) {
        calendar.createEvent(eventTitle, eventStartTime, eventEndTime);
      }
    }
  }
}

// Hàm parseTime để lấy thời gian từ chuỗi và thêm vào ngày cụ thể
function parseTime(timeStr, date) {
  var timeParts = timeStr.split(":");
  var hours = parseInt(timeParts[0], 10);
  var minutes = parseInt(timeParts[1], 10);
  
  // Tạo một bản sao mới của ngày để tránh thay đổi ngày gốc
  var newDate = new Date(date);
  newDate.setHours(hours);
  newDate.setMinutes(minutes);
  newDate.setSeconds(0);
  newDate.setMilliseconds(0);
  return newDate;
}
