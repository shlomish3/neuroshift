function extractLatestResponses(rows, nameIndex, timestampIndex) {
  const map = new Map();
  rows.forEach(row => {
    const name = row[nameIndex];
    const ts = new Date(row[timestampIndex]);
    if (!name || isNaN(ts)) return;
    if (!map.has(name) || ts > map.get(name).timestamp) {
      map.set(name, { timestamp: ts, row });
    }
  });
  return map;
}

function prepareOutputSheets(ss) {
  const outName = "זמינות מפורקת";
  const fixedName = "שיבוצים קבועים";

  if (ss.getSheetByName(outName)) ss.deleteSheet(ss.getSheetByName(outName));
  const outSheet = ss.insertSheet(outName);
  outSheet.appendRow(["שם", "תאריך", "יום בשבוע", "סוג חסימה/זמינות", "מקור"]);

  let fixedSheet = ss.getSheetByName(fixedName);
  if (!fixedSheet) {
    fixedSheet = ss.insertSheet(fixedName);
    fixedSheet.appendRow(["סוג משמרת", "תאריך התחלה", "תאריך סיום", "שם"]);
  }

  return { outSheet, fixedSheet };
}

function getColumnIndices(headers) {
  const idxOf = (pattern) =>
    headers.reduce((arr, h, i) => (h && h.includes(pattern) ? [...arr, i] : arr), []);

  const partTime = headers.reduce((map, h, i) => {
    const m = h?.match(/מי שבמשרה חלקית[^[]*\[([^]+)\]/);
    if (m) map[m[1]] = i;
    return map;
  }, {});

  return {
    rotationStartCols: idxOf("תאריך התחלה של הסבב"),
    rotationEndCols: idxOf("תאריך סיום של הסבב"),
    exceptionCols: idxOf("תאריך חריג"),
    singleBlockCols: idxOf("תאריך חסימה ספציפי"),
    blockRangeStartCols: idxOf("תאריך התחלה לחסימה"),
    blockRangeEndCols: idxOf("תאריך סיום לחסימה"),
    nightSingleCols: idxOf("תאריך חסימה ספציפי לתורנות"),
    nightRangeStartCols: idxOf("תאריך התחלה לחסימת תורנות"),
    nightRangeEndCols: idxOf("תאריך סיום לחסימת תורנות"),
    desiredShiftCols: idxOf("תאריך רצוי לתורנות"),
    recurringClinicCols: idxOf("מרפאה קבועה"),
    recurringOutCols: idxOf("עבודה קבועה אחה\"צ"),
    notesCols: idxOf("הערות"),
    partTimeCols: partTime,
  };
}

function expandDateRange(startRaw, endRaw) {
  const start = new Date(startRaw), end = new Date(endRaw);
  if (isNaN(start) || isNaN(end) || end < start) return [];
  const days = [];
  for (let d = new Date(start); d <= end; d.setDate(d.getDate() + 1)) {
    days.push(new Date(d));
  }
  return days;
}

function appendAvailabilityRow(sheet, name, dateObj, type, source) {
  let dateStr = "", weekday = "";
  if (dateObj instanceof Date && !isNaN(dateObj)) {
    dateStr = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "yyyy-MM-dd");
    const weekdays = ["ראשון", "שני", "שלישי", "רביעי", "חמישי", "שישי", "שבת"];
    weekday = weekdays[dateObj.getDay()];
  }
  sheet.appendRow([name, dateStr, weekday, type, source]);
}

function alreadyInFixedRowSet(existingFixed, name, dateStr, shiftType) {
  const key = `${dateStr}|${name}|${shiftType}`;
  return existingFixed.has(key);
}

function buildExistingFixedSet(fixedSheet) {
  const existingFixed = new Set();
  const values = fixedSheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    const [shiftType, start, , name] = values[i];
    if (shiftType && start && name) {
      const dateStr = Utilities.formatDate(new Date(start), Session.getScriptTimeZone(), "yyyy-MM-dd");
      const key = `${dateStr}|${name}|${shiftType}`;
      existingFixed.add(key);
    }
  }
  return existingFixed;
}


function processResponseRow(name, row, headers, outSheet) {
  const tz = Session.getScriptTimeZone();
  const today = new Date();
  const nextMonth = new Date(today.getFullYear(), today.getMonth() + 1, 1);
  const daysInMonth = new Date(nextMonth.getFullYear(), nextMonth.getMonth() + 1, 0).getDate();

  const hebrewDays = ["ראשון", "שני", "שלישי", "רביעי", "חמישי", "שישי", "שבת"];
  const getDayName = d => hebrewDays[d.getDay()];

  const formatDate = d => Utilities.formatDate(d, tz, "yyyy-MM-dd");

    // חסימה לפי סבב / מבחן
  const rotStatus = row[headers.indexOf("האם ברוטציה / לפני מבחן (משמש לחסימה של כל החודש)?")];
  if (rotStatus && rotStatus.toLowerCase().includes("כן")) {
    const startIdx = headers.indexOf("תאריך התחלה של הסבב (אם יש)");
    const endIdx = headers.indexOf("תאריך סיום של הסבב (אם יש)");
    const start = new Date(row[startIdx]);
    const end = new Date(row[endIdx]);

    if (start instanceof Date && end instanceof Date && !isNaN(start) && !isNaN(end)) {
      for (let d = new Date(start); d <= end; d.setDate(d.getDate() + 1)) {
        appendAvailabilityRow(new Date(d), "לא זמין", "סבב / לפני מבחן");
      }
    }
  }


  function appendAvailabilityRow(date, type, source) {
    if (!(date instanceof Date) || isNaN(date)) return;
    outSheet.appendRow([name, formatDate(date), getDayName(date), type, source]);
  }

  // מרפאות קבועות
  const clinicDaysStr = row[headers.indexOf("באילו ימים יש לך מרפאה קבועה בבית החולים? (אם אין - לדלג על השאלה)")];
  if (clinicDaysStr) {
    const days = clinicDaysStr.split(",").map(s => s.trim());
    for (let d = 1; d <= daysInMonth; d++) {
      const date = new Date(nextMonth.getFullYear(), nextMonth.getMonth(), d);
      if (days.includes(getDayName(date))) {
        appendAvailabilityRow(date, "לא זמין", "מרפאה קבועה");
      }
    }
  }

  // עבודה אחה"צ
  const afternoonWork = row[headers.indexOf("באילו ימים יש לך עבודה קבועה אחה\"צ שלא מאפשרת כוננות? (אם אין - לדלג על השאלה)")];
  if (afternoonWork) {
    const days = afternoonWork.split(",").map(s => s.trim());
    for (let d = 1; d <= daysInMonth; d++) {
      const date = new Date(nextMonth.getFullYear(), nextMonth.getMonth(), d);
      if (days.includes(getDayName(date))) {
        appendAvailabilityRow(date, "לא זמין לתורנות", "עבודה קבועה אחה\"צ");
      }
    }
  }

  // משרה חלקית
  const partialHeaders = {
    ראשון: "מי שבמשרה חלקית - מהם ימי העבודה הקבועים בביה\"ח שמיר? [ראשון]",
    שני: "מי שבמשרה חלקית - מהם ימי העבודה הקבועים בביה\"ח שמיר? [שני]",
    שלישי: "מי שבמשרה חלקית - מהם ימי העבודה הקבועים בביה\"ח שמיר? [שלישי]",
    רביעי: "מי שבמשרה חלקית - מהם ימי העבודה הקבועים בביה\"ח שמיר? [רביעי]",
    חמישי: "מי שבמשרה חלקית - מהם ימי העבודה הקבועים בביה\"ח שמיר? [חמישי]"
  };
  for (const [dayName, colName] of Object.entries(partialHeaders)) {
    const idx = headers.indexOf(colName);
    if (row[idx]) {
      for (let d = 1; d <= daysInMonth; d++) {
        const date = new Date(nextMonth.getFullYear(), nextMonth.getMonth(), d);
        if (getDayName(date) === dayName) {
          appendAvailabilityRow(date, "לא זמין", "משרה חלקית");
        }
      }
    }
  }

  // חסימות יומיות
  for (let i = 1; i <= 6; i++) {
    const idx = headers.indexOf(`תאריך חסימה ספציפי ${i}?`);
    const date = new Date(row[idx]);
    appendAvailabilityRow(date, "לא זמין", `חסימה יומית - בלוק ${i}`);
  }

  // חסימות טווח תאריכים
  for (let i = 1; i <= 6; i++) {
    const start = new Date(row[headers.indexOf(`תאריך התחלה לחסימה (טווח תאריכים ${i})`)]);
    const end = new Date(row[headers.indexOf(`תאריך סיום לחסימה (טווח תאריכים ${i})`)]);
    for (let d = new Date(start); d <= end; d.setDate(d.getDate() + 1)) {
      appendAvailabilityRow(new Date(d), "לא זמין", `חסימה יומית - טווח ${i}`);
    }
  }

  // חסימות תורנות
  for (let i = 1; i <= 6; i++) {
    const oncall = new Date(row[headers.indexOf(`תאריך חסימה ספציפי לתורנות/כוננות ${i}?`)]);
    appendAvailabilityRow(oncall, "לא זמין לתורנות", `חסימת תורנות - בלוק ${i}`);
  }

  // חסימות טווח תורנות
  for (let i = 1; i <= 6; i++) {
    const start = new Date(row[headers.indexOf(`תאריך התחלה לחסימת תורנות/כוננות (טווח תאריכים ${i})`)]);
    const end = new Date(row[headers.indexOf(`תאריך סיום לחסימת תורנות/כוננות (טווח תאריכים ${i})`)]);
    for (let d = new Date(start); d <= end; d.setDate(d.getDate() + 1)) {
      appendAvailabilityRow(new Date(d), "לא זמין לתורנות", `חסימת תורנות - טווח ${i}`);
    }
  }

  // בקשות לתורנות / כוננות
  for (let i = 1; i <= 4; i++) {
    const idx = headers.indexOf(`תאריך רצוי לתורנות/כוננות ${i}`);
    const date = new Date(row[idx]);
    appendAvailabilityRow(date, "מבקש תורנות", "בקשת תורנות");
  }

  // תאריכים חריגים (מגיעים למרות החסם)
  for (let i = 1; i <= 6; i++) {
    const idx = headers.indexOf(`תאריך חריג ${i} בו אגיע לעבוד בביה\"ח`);
    const date = new Date(row[idx]);
    appendAvailabilityRow(date, "זמין חריג", `תאריך חריג ${i}`);
  }

  // הערה חופשית
  const notesIdx = headers.findIndex(h => h.includes("הערות נוספות"));
  const note = row[notesIdx];
  if (note) {
    outSheet.appendRow([name, "", "", "הערה חופשית", note]);
  }
}


function handleRequestedShifts(row, indices, record, fixedSheet, name, blockedDates, insertedShifts, existingFixed) {
  indices.desiredShiftCols.forEach((colIdx) => {
    const val = row[colIdx];
    if (!val) return;

    const d = new Date(val);
    const dateStr = Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
    const key = `${dateStr}|${name}|תורן מיון`;
    const conflict = blockedDates.get(dateStr);

    Logger.log(`▶ שם: ${name}, תאריך רצוי: ${dateStr}`);
    Logger.log(`🔎 KEY = ${key}`);
    Logger.log(`🔍 existingFixed.has(key)? ${existingFixed.has(key)}`);
    Logger.log(`🆕 insertedShifts.has(key)? ${insertedShifts.has(key)}`);

    if (!conflict) {
      if (!existingFixed.has(key) && !insertedShifts.has(key)) {
        insertedShifts.add(key);
        record(d, "מבקש תורנות לילה", "בקשת תורנות");
        Logger.log(`✅ מוסיף לשיבוצים קבועים: תורן מיון | ${dateStr} | ${name}`);
        fixedSheet.appendRow(["תורן מיון", dateStr, dateStr, name]);
      } else {
        Logger.log(`🔁 הבקשה כבר קיימת — לא נוסף שוב`);
      }
    } else {
      Logger.log(`⛔ לא נוסף לשיבוצים — חסימה: ${conflict}`);
      fixedSheet.appendRow(["⛔ בקשה עם חסימה", dateStr, dateStr, `${name} (קונפליקט: ${conflict})`]);
    }
  });
}


function handleRotationBlocks(row, indices, record) {
  const startVal = row[indices.rotationStartCols[0]];
  const endVal = row[indices.rotationEndCols[0]];
  if (startVal && endVal) {
    expandDateRange(startVal, endVal).forEach(d =>
      record(d, "לא זמין", "סבב חיצוני"));
  }
}

function handleSingleBlocks(row, indices, record) {
  indices.singleBlockCols.forEach((colIdx, i) => {
    const val = row[colIdx];
    if (val) record(new Date(val), "לא זמין", `חסימה יומית - בלוק ${i + 1}`);
  });
}

function handleRangeBlocks(row, indices, record) {
  indices.blockRangeStartCols.forEach((startIdx, i) => {
    const endIdx = indices.blockRangeEndCols[i] || -1;
    const startVal = row[startIdx], endVal = row[endIdx];
    if (startVal && endVal) {
      expandDateRange(startVal, endVal).forEach(d =>
        record(d, "לא זמין", `חסימה יומית - טווח ${i + 1}`));
    }
  });
}

function handleNightBlocks(row, indices, record) {
  indices.nightSingleCols.forEach((colIdx, i) => {
    const val = row[colIdx];
    if (val) record(new Date(val), "לא זמין לתורנות", `חסימת תורנות - בלוק ${i + 1}`);
  });
}

function handleNightRanges(row, indices, record) {
  indices.nightRangeStartCols.forEach((startIdx, i) => {
    const endIdx = indices.nightRangeEndCols[i] || -1;
    const startVal = row[startIdx];
    const endVal = row[endIdx];

    if (startVal) {
      const range = expandDateRange(startVal, endVal || startVal); // Treat as single-day if no end
      range.forEach(d =>
        record(d, "לא זמין לתורנות", `חסימת תורנות - טווח ${i + 1}`));
    }
  });
}


function handleRecurringClinics(row, indices, record, daysInNextMonth, nextMonth, hebrewDayIndex) {
  indices.recurringClinicCols.forEach((colIdx) => {
    const val = row[colIdx];
    if (!val) return;
    val.split(/[,;]+/).map(s => s.trim()).forEach((dayName) => {
      const idx = hebrewDayIndex[dayName];
      if (idx === undefined) return;
      for (let d = 1; d <= daysInNextMonth; d++) {
        const cur = new Date(nextMonth.getFullYear(), nextMonth.getMonth(), d);
        if (cur.getDay() === idx) {
          record(cur, "לא זמין", "מרפאה קבועה");
        }
      }
    });
  });
}

function handleRecurringAfternoons(row, indices, record, daysInNextMonth, nextMonth, hebrewDayIndex) {
  indices.recurringOutCols.forEach((colIdx) => {
    const val = row[colIdx];
    if (!val) return;
    val.split(/[,;]+/).map(s => s.trim().replace(/ אחה"?צ/, "")).forEach((dayName) => {
      const idx = hebrewDayIndex[dayName];
      if (idx === undefined) return;
      for (let d = 1; d <= daysInNextMonth; d++) {
        const cur = new Date(nextMonth.getFullYear(), nextMonth.getMonth(), d);
        if (cur.getDay() === idx) {
          record(cur, "לא זמין לתורנות", "עבודה אחה\"צ קבועה");
        }
      }
    });
  });
}

function handlePartTimeDays(row, indices, record, daysInNextMonth, nextMonth, hebrewDayIndex) {
  Object.entries(indices.partTimeCols).forEach(([dayName, colIdx]) => {
    const val = row[colIdx];
    if (!val || !String(val).toLowerCase().includes("כן")) return;
    const idx = hebrewDayIndex[dayName];
    if (idx === undefined) return;
    for (let d = 1; d <= daysInNextMonth; d++) {
      const cur = new Date(nextMonth.getFullYear(), nextMonth.getMonth(), d);
      if (cur.getDay() === idx) {
        record(cur, "יום עבודה קבוע", `עובד קבוע ביום ${dayName}`);
      }
    }
  });
}

function handlePositiveExceptions(row, indices, record) {
  indices.exceptionCols.forEach((colIdx) => {
    const val = row[colIdx];
    if (val) record(new Date(val), "זמין חריג", "תאריך חריג");
  });
}

function handleNotes(row, noteCols, outSheet, name) {
  noteCols.forEach((colIdx) => {
    const val = row[colIdx];
    if (val && String(val).trim()) {
      appendAvailabilityRow(outSheet, name, null, "הערה חופשית", String(val).trim());
    }
  });
}
