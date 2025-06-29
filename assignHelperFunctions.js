function adjustColumnWidthsByContent(sheet) {
  const range = sheet.getDataRange();
  const values = range.getDisplayValues();
  const colCount = values[0].length;

  for (let col = 0; col < colCount; col++) {
    const maxLen = values.reduce(
      (max, row) => Math.max(max, (row[col] + "").length),
      0
    );
    const width = Math.min(400, maxLen * 7); // Limit to 400px max
    sheet.setColumnWidth(col + 1, width);
  }
}


function getHebrewDayName(date) {
  const days = ["ראשון", "שני", "שלישי", "רביעי", "חמישי", "שישי", "שבת"];
  return days[new Date(date).getDay()];
}

function formatDate(d) {
  return Utilities.formatDate(new Date(d), Session.getScriptTimeZone(), "yyyy-MM-dd");
}

function appendRequestedFixedAssignments() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();

  const formSheet = ss.getSheetByName("תגובות לטופס זמינות");
  const fixedSheet = ss.getSheetByName("שיבוצים קבועים") || ss.insertSheet("שיבוצים קבועים");
  const fixedData = fixedSheet.getDataRange().getValues();
  const fixedRows = fixedData.slice(1).map(r => ({
    shift: r[0], date: formatDate(r[1], tz), name: r[3]
  }));

  const data = formSheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);

  const toranIndex = headers.indexOf("תאריך רצוי לתורנות מיון");
  const conanIndex = headers.indexOf("תאריך רצוי לכוננות מיון");
  const nameIndex = headers.indexOf("בחר את שמך");

  if (toranIndex === -1 || conanIndex === -1 || nameIndex === -1)
    throw new Error("חסרים עמודות בטופס");

  const newAssignments = [];

  for (const row of rows) {
    const name = row[nameIndex];
    const toranDates = parseMultiDates(row[toranIndex], tz);
    const conanDates = parseMultiDates(row[conanIndex], tz);

    for (const d of toranDates) {
      if (!fixedRows.some(r => r.shift === "תורן מיון" && r.date === d))
        newAssignments.push(["תורן מיון", new Date(d), new Date(d), name]);
    }

    for (const d of conanDates) {
      if (!fixedRows.some(r => r.shift === "כונן מיון" && r.date === d))
        newAssignments.push(["כונן מיון", new Date(d), new Date(d), name]);
    }
  }

  if (newAssignments.length) {
    fixedSheet.getRange(fixedData.length + 1, 1, newAssignments.length, 4)
      .setValues(newAssignments);
  }
}

function parseMultiDates(cellValue, tz) {
  if (!cellValue) return [];
  return cellValue.toString().split(/[,;\n]+/)
    .map(s => new Date(s.trim()))
    .filter(d => !isNaN(d))
    .map(d => formatDate(d, tz));
}

function formatDate(d, tz) {
  return Utilities.formatDate(new Date(d), tz, "yyyy-MM-dd");
}

function appendAvailabilityRow(sheet, name, date, type, source) {
  const hebrewWeekdays = {
    0: "ראשון",
    1: "שני",
    2: "שלישי",
    3: "רביעי",
    4: "חמישי",
    5: "שישי",
    6: "שבת"
  };

  let dateStr = "";
  let dayName = "";
  if (date) {
    dateStr = Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd");
    dayName = hebrewWeekdays[date.getDay()];
  }
  sheet.appendRow([name, dateStr, dayName, type, source]);
}

function getAllowedPostClinicDates() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("פוסט אשפוז");
  if (!sheet) throw new Error("לא נמצאה לשונית 'פוסט אשפוז'");

  const data = sheet.getDataRange().getValues().slice(1); // skip header
  const tz = ss.getSpreadsheetTimeZone();
  const allowed = new Set();

  data.forEach(row => {
    const date = new Date(row[0]);
    if (!isNaN(date)) {
      const dateStr = Utilities.formatDate(date, tz, "yyyy-MM-dd");
      allowed.add(dateStr);
    }
  });

  return allowed;
}
