function fillEligibility() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();

  // Find the latest "משמרות" sheet
  const scheduleSheet = ss.getSheets()
    .filter(s => s.getName().startsWith("משמרות "))
    .sort((a, b) => b.getName().localeCompare(a.getName()))[0];

  const data = scheduleSheet.getDataRange().getValues();
  const rows = data.slice(1); // skip headers

  const empSheet = ss.getSheetByName("עובדים");
  const empData = empSheet.getDataRange().getValues();
  const empHeaders = empData[0].slice(1); // shift types
  const empList = empData.slice(1);       // names + shift eligibility

  // ✅ Use זמינות מפורקת instead of זמינות
  const parsedSheet = ss.getSheetByName("זמינות מפורקת");
  if (!parsedSheet) throw new Error("לא נמצא גיליון בשם 'זמינות מפורקת'");
  const parsedData = parsedSheet.getDataRange().getValues();
  const parsedRows = parsedData.slice(1); // skip header

  // Build a map: name -> Set of unavailable date strings (yyyy-MM-dd)
  const unavailabilityMap = {};
  parsedRows.forEach(([name, dateStr, , type]) => {
    if (!name || !dateStr) return;
    const normalized = Utilities.formatDate(new Date(dateStr), tz, "yyyy-MM-dd");
    const blockType = String(type || "");
    // Only mark actual blocking types
    const isBlocking = blockType.includes("לא זמין");
    if (isBlocking) {
      if (!unavailabilityMap[name]) unavailabilityMap[name] = new Set();
      unavailabilityMap[name].add(normalized);
    }
  });

  const fixedSheet = ss.getSheetByName("שיבוצים קבועים");
  const fixedRules = fixedSheet ? fixedSheet.getDataRange().getValues().slice(1) : [];

  rows.forEach((row, i) => {
    const [dateStr, , shiftType] = row;
    const rowIndex = i + 2;
    const dateFormatted = Utilities.formatDate(new Date(dateStr), tz, "yyyy-MM-dd");

    const demand = row[5]; // כמות נדרשת
    const needed = parseInt(demand, 10);
    if (isNaN(needed) || needed === 0) return;

    const shiftColIndex = empHeaders.indexOf(shiftType);
    const eligible = (shiftColIndex === -1) ? [] : empList
      .filter(emp => (emp[shiftColIndex + 1] === 1 || emp[shiftColIndex + 1] === "1"))
      .map(emp => emp[0])
      .filter(name => !(unavailabilityMap[name] || new Set()).has(dateFormatted));

    const fixed = fixedRules.find(r => {
      if (!r[0] || !r[1] || !r[2] || !r[3]) return false;
      const [shift, start, end, name] = r;
      const startStr = Utilities.formatDate(new Date(start), tz, "yyyy-MM-dd");
      const endStr = Utilities.formatDate(new Date(end), tz, "yyyy-MM-dd");
      return shift === shiftType && dateFormatted >= startStr && dateFormatted <= endStr;
    });

    if (fixed) {
      const fixedName = fixed[3];
      const blocked = unavailabilityMap[fixedName] || new Set();

      if (blocked.has(dateFormatted)) {
        Logger.log(`❌ שיבוץ קבוע עבור ${fixedName} ב-${dateFormatted} נדחה (חסם זמינות)`);
        return;
      }

      // 🟢 Write fixed name visibly
      scheduleSheet.getRange(rowIndex, 4).setValue(`${fixedName} (קבוע)`);

      // 🟢 Add fixedName to eligibility list if not already present
      const fullEligibility = Array.from(new Set([fixedName, ...eligible]));
      scheduleSheet.getRange(rowIndex, 5).setValue(fullEligibility.join(", "));

      return;
    }

    // No fixed assignment — just normal eligibility
    scheduleSheet.getRange(rowIndex, 5).setValue(eligible.join(", "));
  });

  Logger.log("השלמת חישוב זכאים לפי זמינות מפורקת ושיבוצים קבועים.");
}
