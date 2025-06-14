function createNextMonthSchedule() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();
  const now = new Date();
  const nextMonth = new Date(now.getFullYear(), now.getMonth() + 1, 1);
  const yearMonth = Utilities.formatDate(nextMonth, tz, "yyyy-MM");
  const sheetName = `משמרות ${yearMonth}`;

  if (ss.getSheetByName(sheetName)) {
    SpreadsheetApp.getUi().alert(`הגיליון "${sheetName}" כבר קיים.`);
    return;
  }

  const empSheet = ss.getSheetByName("עובדים");
  const shiftTypes = empSheet.getRange(1, 2, 1, empSheet.getLastColumn() - 1).getValues()[0];

  const dayMap = {
    "Sunday": "ראשון",
    "Monday": "שני",
    "Tuesday": "שלישי",
    "Wednesday": "רביעי",
    "Thursday": "חמישי",
    "Friday": "שישי",
    "Saturday": "שבת"
  };
  const dayLetterMap = {
    "ראשון": "א",
    "שני": "ב",
    "שלישי": "ג",
    "רביעי": "ד",
    "חמישי": "ה",
    "שישי": "ו",
    "שבת": "ש"
  };

  const needSheet = ss.getSheetByName("כמות נדרשת");
  const needData = needSheet ? needSheet.getDataRange().getValues() : [];
  const needHeaders = needData[0].slice(1); // א–ש
  const needMap = Object.fromEntries(
    needData.slice(1).map(row => {
      const shift = row[0];
      const needs = {};
      row.slice(1).forEach((val, idx) => {
        const dayLetter = needHeaders[idx];
        if (val !== "") needs[dayLetter] = val;
      });
      return [shift, needs];
    })
  );

  const holidaySheet = ss.getSheetByName("חגים");
  const restDaySet = holidaySheet
    ? new Set(
        holidaySheet.getRange("A2:C" + holidaySheet.getLastRow()).getValues()
          .filter(r => r[2] === "חופש")
          .map(r => new Date(r[0]).setHours(0, 0, 0, 0))
      )
    : new Set();

  const days = [];
  const lastDay = new Date(nextMonth.getFullYear(), nextMonth.getMonth() + 1, 0).getDate();
  for (let d = 1; d <= lastDay; d++) {
    const date = new Date(nextMonth.getFullYear(), nextMonth.getMonth(), d);
    const timestamp = date.setHours(0, 0, 0, 0);
    const formatted = Utilities.formatDate(new Date(timestamp), tz, "yyyy-MM-dd");
    const engDay = Utilities.formatDate(new Date(timestamp), tz, "EEEE");
    const hebDay = dayMap[engDay];
    const dayLetter = dayLetterMap[hebDay];
    const isHoliday = restDaySet.has(timestamp);
    const holidayEves = holidaySheet
    ? new Set(
        holidaySheet.getRange("A2:C" + holidaySheet.getLastRow()).getValues()
          .filter(r => r[2] === "מידע" && isHolidayEve(r[1]))
          .map(r => new Date(r[0]).setHours(0, 0, 0, 0))
      )
    : new Set();

    const isHolidayEveDate = holidayEves.has(timestamp);

    const holidayDayLetter = isHoliday ? "ש" : dayLetter;

    for (const shift of shiftTypes) {
      let required = "";

      const isSaturday = new Date(timestamp).getDay() === 6;
      const treatAsShabbat = isHoliday || (isHolidayEveDate && isSaturday);

      if (treatAsShabbat) {
        const shabbatNeed = needMap[shift]?.["ש"];
        required = shabbatNeed !== undefined ? shabbatNeed : 0;
      } else if (isHolidayEveDate) {
        const fridayNeed = needMap[shift]?.["ו"];
        required = fridayNeed !== undefined ? fridayNeed : 0;
      } else {
        const regularNeed = needMap[shift]?.[dayLetter];
        required = regularNeed !== undefined ? regularNeed : 0;
      }

      days.push([formatted, hebDay, shift, "", "", required]);
    }

  }

  const newSheet = ss.insertSheet(sheetName);
  newSheet.setRightToLeft(true);
  newSheet.getRange(1, 1, 1, 6).setValues([
    ["תאריך", "יום", "סוג משמרת", "שם", "זכאים", "כמות נדרשת"]
  ]);
  newSheet.getRange(2, 1, days.length, 6).setValues(days);

  // 🟨🟦🟫 Apply coloring + add חג to יום
  if (holidaySheet) {
    const holidayData = holidaySheet.getRange("A2:C" + holidaySheet.getLastRow()).getValues();
    const restDates = holidayData.filter(r => r[2] === "חופש").map(r => [new Date(r[0]).toDateString(), r[1]]);
    const infoDates = holidayData.filter(r => r[2] === "מידע").map(r => [new Date(r[0]).toDateString(), r[1]]);

    const scheduleRange = newSheet.getRange(2, 1, newSheet.getLastRow() - 1, 2); // תאריך, יום
    const scheduleData = scheduleRange.getValues();

    for (let i = 0; i < scheduleData.length; i++) {
      const row = i + 2;
      const cellDate = new Date(scheduleData[i][0]);
      const dayStr = scheduleData[i][1];
      const dateStr = cellDate.toDateString();

      let color = null;
      let holidayName = "";

      const matchRest = restDates.find(([d]) => d === dateStr);
      const matchInfo = infoDates.find(([d]) => d === dateStr);

      if (matchRest) {
        color = "#FFF59D"; // Yellow for שבת or חופש
        holidayName = matchRest[1];
      } else if (matchInfo) {
        color = "#DCEFFF"; // Light blue for מידע
        holidayName = matchInfo[1];
      } else if (matchInfo && isHolidayEve(matchInfo[1])) {
        color = "#FFE0B2"; // Treat ערב חג like Friday
      } else if (cellDate.getDay() === 5) {
        color = "#FFE0B2"; // Light orange for שישי
      } else if (cellDate.getDay() === 6) {
        color = "#FFF59D"; // שבת = Yellow (same as חופש)
      }

      if (color) {
        newSheet.getRange(row, 1, 1, newSheet.getLastColumn()).setBackground(color);
      }

      if (holidayName) {
        newSheet.getRange(row, 2).setValue(`${dayStr} - ${holidayName}`);
      }
    }
  }

  Logger.log(`המשמרות לחודש ${yearMonth} נוצרו בהצלחה!`);
}
