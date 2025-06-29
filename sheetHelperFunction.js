function isHoliday(date) {
  const holidaysSheet = SpreadsheetApp.getActive().getSheetByName("חגים");
  if (!holidaysSheet) return false;

  const data = holidaysSheet.getDataRange().getValues();
  const tz = Session.getScriptTimeZone();

  const holidayDates = new Set(
    data
      .slice(1)
      .filter(row => row[2] === "חופש")  // Only rows marked חופש
      .map(row =>
        Utilities.formatDate(new Date(row[0]), tz, "yyyy-MM-dd")
      )
  );

  const dateStr = Utilities.formatDate(date, tz, "yyyy-MM-dd");
  return holidayDates.has(dateStr);
}

function assignAttendingsForMonth(scheduleSheet, tz) {
  const ss = SpreadsheetApp.getActive();
  const attendingSheet = ss.getSheetByName("אטנדינג");
  if (!attendingSheet) throw new Error("לא נמצאה לשונית 'אטנדינג'");

  const attData = attendingSheet.getDataRange().getValues();
  const headers = attData[0];
  const rows = attData.slice(1);

  const monthIndex = headers.indexOf("חודש");
  const att1Index = headers.indexOf("אטנדינג 1");
  const att2Index = headers.indexOf("אטנדינג 2");
  if (monthIndex === -1 || att1Index === -1 || att2Index === -1)
    throw new Error("שמות עמודות לא תקינים בלשונית 'אטנדינג'");

  const schedule = scheduleSheet.getDataRange().getValues();
  const header = schedule[0];
  const rowsData = schedule.slice(1);

  const dateCol = header.indexOf("תאריך");
  const shiftCol = header.indexOf("סוג משמרת");
  const nameCol = header.indexOf("שם");

  if ([dateCol, shiftCol, nameCol].includes(-1))
    throw new Error("שמות עמודות לא תקינים בגיליון המשמרות");

  if (!rowsData.length) return;

  const firstDate = new Date(rowsData[0][dateCol]);
  const monthStr = Utilities.formatDate(firstDate, tz, "yyyy-MM");

  const match = rows.find(row => row[monthIndex] === monthStr);
  if (!match) throw new Error(`אין אטנדינג עבור חודש ${monthStr}`);
  const [att1, att2] = [match[att1Index], match[att2Index]];

  for (let i = 0; i < rowsData.length; i++) {
    const row = rowsData[i];
    const date = new Date(row[dateCol]);
    const day = date.getDate();
    const weekday = getHebrewDayName(date);
    const shift = row[shiftCol];

    if (shift !== "אטנדינג") continue;
    if (["שישי", "שבת"].includes(weekday)) continue;
    if (isHoliday(date)) continue;

    const attending = day <= 15 ? att1 : att2;
    if (attending) {
      scheduleSheet.getRange(i + 2, nameCol + 1).setValue(attending);
    }
  }
}


function appendAttendingsToFixedAssignments(scheduleSheet, tz) {
  const ss = SpreadsheetApp.getActive();
  const attSheet = ss.getSheetByName("אטנדינג");
  const fixedSheet = ss.getSheetByName("שיבוצים קבועים") || ss.insertSheet("שיבוצים קבועים");

  const attData = attSheet.getDataRange().getValues();
  const attHeaders = attData[0];
  const attRows = attData.slice(1);

  const schedule = scheduleSheet.getDataRange().getValues();
  const headers = schedule[0];
  const rows = schedule.slice(1);

  const dateCol = headers.indexOf("תאריך");
  const shiftCol = headers.indexOf("סוג משמרת");

  const monthIndex = attHeaders.indexOf("חודש");
  const att1Index = attHeaders.indexOf("אטנדינג 1");
  const att2Index = attHeaders.indexOf("אטנדינג 2");

  if ([monthIndex, att1Index, att2Index].includes(-1)) throw new Error("שמות עמודות חסרים בטאב אטנדינג");

  const attAssignments = [];

  for (const row of rows) {
    const date = new Date(row[dateCol]);
    const shift = row[shiftCol];
    if (shift !== "אטנדינג") continue;

    const day = date.getDate();
    const monthStr = Utilities.formatDate(date, tz, "yyyy-MM");

    const match = attRows.find(r => {
      const val = r[monthIndex];
      return Utilities.formatDate(new Date(val), tz, "yyyy-MM") === monthStr;
    });

    if (!match) continue;
    const name = (day <= 15) ? match[att1Index] : match[att2Index];
    if (!name) continue;

    attAssignments.push({ name, date, monthStr });
  }

  const grouped = {};
  for (const { name, date, monthStr } of attAssignments) {
    if (!grouped[monthStr]) grouped[monthStr] = {};
    if (!grouped[monthStr][name]) grouped[monthStr][name] = [];
    grouped[monthStr][name].push(date);
  }

  const result = [];

  for (const monthStr in grouped) {
    for (const name in grouped[monthStr]) {
      const dates = grouped[monthStr][name].sort((a, b) => a - b);
      const start = Utilities.formatDate(dates[0], tz, "yyyy-MM-dd");
      const end = Utilities.formatDate(dates[dates.length - 1], tz, "yyyy-MM-dd");
      result.push(["אטנדינג", start, end, name]);
    }
  }

  // 🧼 Load existing assignments to prevent duplicates
  const existing = fixedSheet.getDataRange().getValues().map(r => r.slice(0, 4).join("|"));
  const unique = result.filter(r => !existing.includes(r.join("|")));

  if (unique.length) {
    fixedSheet.getRange(fixedSheet.getLastRow() + 1, 1, unique.length, 4).setValues(unique);
    Logger.log(`✅ נוספו ${unique.length} שיבוצים חדשים לאטנדינג (נמנעו כפילויות)`);
  } else {
    Logger.log("ℹ️ לא נוספו שיבוצים — כל השורות כבר קיימות.");
  }
}
