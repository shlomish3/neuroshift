function parseAvailabilityResponses() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const formSheet = ss.getSheetByName("תגובות לטופס זמינות");
  const outputSheetName = "זמינות מפורקת";
  const fixedSheetName = "שיבוצים קבועים";

  if (!formSheet) throw new Error("לא נמצאה לשונית 'תגובות לטופס זמינות'");

  // קבלת כל שורות הגיליון
  const data = formSheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);

  // מציאת אינדקסים דינמיים לפי מילות מפתח בכותרות
  const idxOf = (pattern) => {
    return headers.reduce((arr, h, i) => {
      if (h && h.toString().includes(pattern)) arr.push(i);
      return arr;
    }, []);
  };

  // אינדקס חותמת זמן ושם
  const tsCols = idxOf("חותמת זמן");
  const nameCols = idxOf("בחר את שמך");
  if (!tsCols.length || !nameCols.length) {
    throw new Error("שדות 'חותמת זמן' או 'בחר את שמך' חסרים");
  }
  const timestampIndex = tsCols[0];
  const nameIndex = nameCols[0];

  // בניית מפת התשובה האחרונה לכל שם
  const latestMap = new Map();
  rows.forEach((row) => {
    const name = row[nameIndex];
    const tsRaw = row[timestampIndex];
    const ts = new Date(tsRaw);
    if (!name || isNaN(ts)) return;
    if (!latestMap.has(name) || ts > latestMap.get(name).timestamp) {
      latestMap.set(name, { timestamp: ts, row });
    }
  });
  if (!latestMap.size) {
    throw new Error("לא נמצאו שורות תקינות עם תאריך ושם תקינים");
  }

  // מחיקת גיליון פלט קיים ויצירה מחדש
  let outSheet = ss.getSheetByName(outputSheetName);
  if (outSheet) ss.deleteSheet(outSheet);
  outSheet = ss.insertSheet(outputSheetName);
  outSheet.appendRow(["שם", "תאריך", "יום בשבוע", "סוג חסימה/זמינות", "מקור"]);

  // וידוא גיליון שיבוצים קבועים
  let fixedSheet = ss.getSheetByName(fixedSheetName);
  if (!fixedSheet) {
    fixedSheet = ss.insertSheet(fixedSheetName);
    fixedSheet.appendRow(["סוג משמרת", "תאריך התחלה", "תאריך סיום", "שם"]);
  }

  // עזרה להוספת שורה לגיליון הפלט
  function appendAvailabilityRow(sheet, name, dateObj, type, source) {
    let dateStr = "";
    let weekday = "";
    if (dateObj instanceof Date && !isNaN(dateObj)) {
      dateStr = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "yyyy-MM-dd");
      const day = dateObj.getDay(); // 0=ראשון ... 6=שבת
      const weekdayMap = ["ראשון", "שני", "שלישי", "רביעי", "חמישי", "שישי", "שבת"];
      weekday = weekdayMap[day];
    }
    sheet.appendRow([name, dateStr, weekday, type, source]);
  }

  // מרחיב טווח תאריכים (כולל קצוות)
  function expandDateRange(startRaw, endRaw) {
    const start = new Date(startRaw);
    const end = new Date(endRaw);
    if (isNaN(start) || isNaN(end) || end < start) return [];
    const arr = [];
    for (let d = new Date(start); d <= end; d.setDate(d.getDate() + 1)) {
      arr.push(new Date(d));
    }
    return arr;
  }

  // זיהוי אינדקסים לפי תבניות
  const rotationStartCols = idxOf("תאריך התחלה של הסבב");
  const rotationEndCols = idxOf("תאריך סיום של הסבב");

  const exceptionCols = idxOf("תאריך חריג"); // "תאריך חריג N בו אגיע"
  const singleBlockCols = idxOf("תאריך חסימה ספציפי"); // תאריך חסימה ספציפי N?
  const blockRangeStartCols = idxOf("תאריך התחלה לחסימה (טווח"); // תאריך התחלה לחסימה (טווח תאריכים N)
  const blockRangeEndCols = idxOf("תאריך סיום לחסימה (טווח"); // תאריך סיום לחסימה (טווח תאריכים N)

  const nightSingleCols = idxOf("תאריך חסימה ספציפי לתורנות"); // תאריך חסימה ספציפי לתורנות/כוננות N?
  const nightRangeStartCols = idxOf("תאריך התחלה לחסימת תורנות"); // תאריך התחלה לחסימת תורנות (טווח תאריכים N)
  const nightRangeEndCols = idxOf("תאריך סיום לחסימת תורנות"); // תאריך סיום לחסימת תורנות (טווח תאריכים N)

  const desiredShiftCols = idxOf("תאריך רצוי לתורנות"); // תאריך רצוי לתורנות/כוננות N

  const recurringClinicCols = idxOf("באילו ימים יש לך מרפאה קבועה");
  const recurringOutCols = idxOf("באילו ימים יש לך עבודה קבועה אחה\"צ");

  // חלק-זמן: "מי שבמשרה חלקית - ... [יום]"
  const partTimeCols = headers.reduce((map, h, i) => {
    const m = h && h.match(/מי שבמשרה חלקית[^[]*\[([^]+)\]/);
    if (m) {
      // m[1] זה היום העברי, למשל "ראשון", "שני" וכו'
      map[m[1]] = i;
    }
    return map;
  }, {}); // eg: { "ראשון": colIndex, "שני": colIndex, ... }

  const notesCols = idxOf("הערות נוספות");

  // קביעת חודשי מרחיב לפי היום הנוכחי
  const today = new Date();
  const nextMonth = new Date(today.getFullYear(), today.getMonth() + 1, 1);
  const daysInNextMonth =
    new Date(nextMonth.getFullYear(), nextMonth.getMonth() + 1, 0).getDate();
  const hebrewDayIndex = { ראשון: 0, שני: 1, שלישי: 2, רביעי: 3, חמישי: 4, שישי: 5, שבת: 6 };

  // עיבוד כל שורת תשובה אחרונה לכל שם
  latestMap.forEach(({ row }, name) => {
    const blockedDates = new Set();

    // פונקציה לרישום חסימה/זמינות בפלט
    function recordEntry(dateObj, type, source) {
      if (!(dateObj instanceof Date) || isNaN(dateObj)) return;
      const key = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "yyyy-MM-dd");
      // כדי שלא לרשום כפול
      if (!blockedDates.has(key)) {
      appendAvailabilityRow(outSheet, name, dateObj, type, source);
      blockedDates.add(key);  // <-- Only the date string
    }

    }

    // 1. חסימות לפי "סבב" חודשי (אם קיימים שני עמודות התחלה וסיום עבורם)
    if (rotationStartCols.length && rotationEndCols.length) {
      const startVal = row[rotationStartCols[0]];
      const endVal = row[rotationEndCols[0]];
      if (startVal && endVal) {
        expandDateRange(startVal, endVal).forEach((d) => {
          recordEntry(d, "לא זמין", "סבב חיצוני");
        });
      }
    }

    // 2. תאריכי חריג (זמינות חיובית)
    exceptionCols.forEach((colIdx) => {
      const val = row[colIdx];
      if (val) {
        const d = new Date(val);
        recordEntry(d, "זמין חריג", headers[colIdx]);
      }
    });

    // 3. חסימות חד-יומיות רגילות (בלוקים)
    singleBlockCols.forEach((colIdx, i) => {
      const val = row[colIdx];
      if (val) {
        const d = new Date(val);
        recordEntry(d, "לא זמין", `חסימה יומית - בלוק ${i + 1}`);
      }
    });

    // 4. טווחי חסימות רגילות (כלומר, התחלה/סיום)
    blockRangeStartCols.forEach((startIdx, i) => {
      const endIdx = blockRangeEndCols[i] || -1;
      const startVal = row[startIdx];
      const endVal = endIdx >= 0 ? row[endIdx] : null;
      if (startVal && endVal) {
        expandDateRange(startVal, endVal).forEach((d) => {
          recordEntry(d, "לא זמין", `חסימה יומית - טווח ${i + 1}`);
        });
      }
    });

    // 5. חסימות לתורנות/כוננות חד-יומיות (לילה)
    nightSingleCols.forEach((colIdx, i) => {
      const val = row[colIdx];
      if (val) {
        const d = new Date(val);
        recordEntry(d, "לא זמין לתורנות", `חסימת תורנות - בלוק ${i + 1}`);
      }
    });

    // 6. טווחי חסימות לתורנות/כוננות
    nightRangeStartCols.forEach((startIdx, i) => {
      const endIdx = nightRangeEndCols[i] || -1;
      const startVal = row[startIdx];
      const endVal = endIdx >= 0 ? row[endIdx] : null;
      if (startVal && endVal) {
        expandDateRange(startVal, endVal).forEach((d) => {
          recordEntry(d, "לא זמין לתורנות", `חסימת תורנות - טווח ${i + 1}`);
        });
      }
    });

    // 7. בקשות תאריך רצוי לתורנות/כוננות (זמינות חיובית + כתיבה ל"שיבוצים קבועים")
    desiredShiftCols.forEach((colIdx) => {
  const val = row[colIdx];
  if (val) {
    const d = new Date(val);
    recordEntry(d, "מבקש תורנות לילה", "בקשת תורנות");

    const dateStr = Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
    const isBlocked = blockedDates.has(dateStr);


    Logger.log(`▶ שם: ${name}, תאריך רצוי: ${dateStr}, חסום? ${isBlocked}`);

    if (!isBlocked) {
      Logger.log(`✅ מוסיף לשיבוצים קבועים: תורן מיון | ${dateStr} | ${name}`);
      fixedSheet.appendRow(["תורן מיון", dateStr, dateStr, name]);
    } else {
      Logger.log(`⛔ לא נוסף לשיבוצים — התאריך כנראה מופיע גם כחסימה`);
    }
  }
});


    // 8. ימים קבועים - מרפאה בבית החולים (חזרה שבועית בחודש הבא, חוסם)
    recurringClinicCols.forEach((colIdx) => {
      const val = row[colIdx];
      if (val) {
        const days = String(val)
          .split(/[,;]+/)
          .map((s) => s.trim())
          .filter((s) => s);
        days.forEach((dayName) => {
          const idxDay = hebrewDayIndex[dayName];
          if (idxDay !== undefined) {
            for (let d = 1; d <= daysInNextMonth; d++) {
              const cur = new Date(nextMonth.getFullYear(), nextMonth.getMonth(), d);
              if (cur.getDay() === idxDay) {
                recordEntry(cur, "לא זמין", "מרפאה קבועה");
              }
            }
          }
        });
      }
    });

    // 9. ימים קבועים - עבודה אחה"צ (חוסם לתורנות לילה)
      recurringOutCols.forEach((colIdx) => {
      const val = row[colIdx];
      if (val) {
        const days = String(val)
          .split(/[,;]+/)
          .map((s) => s.trim().replace(/ אחה"?צ/, "")) // remove " אחה״צ"
          .filter((s) => s);
        days.forEach((cleanDay) => {
          const idxDay = hebrewDayIndex[cleanDay];
          if (idxDay !== undefined) {
            for (let d = 1; d <= daysInNextMonth; d++) {
              const cur = new Date(nextMonth.getFullYear(), nextMonth.getMonth(), d);
              if (cur.getDay() === idxDay) {
                recordEntry(cur, "לא זמין לתורנות", "עבודה אחה\"צ קבועה");
              }
            }
          }
        });
      }
    });


    // 10. ימי עבודה קבועים של משרה חלקית (זמינות חיובית)
    Object.entries(partTimeCols).forEach(([dayName, colIdx]) => {
      const val = row[colIdx];
      if (val && String(val).toLowerCase().includes("כן")) {
        const idxDay = hebrewDayIndex[dayName];
        if (idxDay !== undefined) {
          for (let d = 1; d <= daysInNextMonth; d++) {
            const cur = new Date(nextMonth.getFullYear(), nextMonth.getMonth(), d);
            if (cur.getDay() === idxDay) {
              recordEntry(cur, "יום עבודה קבוע", `עובד קבוע ביום ${dayName}`);
            }
          }
        }
      }
    });

    // 11. הערות נוספות (אם קיימות) — רישום חופשי בלי תאריך
    notesCols.forEach((colIdx) => {
      const text = row[colIdx];
      if (text && String(text).trim()) {
        appendAvailabilityRow(outSheet, name, null, "הערה חופשית", String(text).trim());
      }
    });
  });
}
