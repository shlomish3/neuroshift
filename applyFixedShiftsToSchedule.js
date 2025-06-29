//Probably unnecessary
function applyFixedShiftsToSchedule() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();
  const scheduleSheet = getScheduleSheet(ss);
  const fixedSheet = ss.getSheetByName("שיבוצים קבועים");
  const parsedSheet = ss.getSheetByName("זמינות מפורקת");
  if (!scheduleSheet || !fixedSheet || !parsedSheet) throw new Error("חסרה לשונית");

  const fixedRows = fixedSheet.getDataRange().getValues().slice(1);
  const parsedRows = parsedSheet.getDataRange().getValues().slice(1);
  const schedule = scheduleSheet.getDataRange().getValues();
  const shiftRows = schedule.slice(1); // skip header

  const nameCol = 3, shiftCol = 2, dateCol = 0;

  const blocked = buildBlockedMap(parsedRows, tz);

  const fixedAssignments = fixedRows.filter(([shift]) => ["תורן מיון", "תורן חצי"].includes(shift));

  for (const [shift, start, end, name] of fixedAssignments) {
    const startStr = Utilities.formatDate(new Date(start), tz, "yyyy-MM-dd");
    const endStr = Utilities.formatDate(new Date(end), tz, "yyyy-MM-dd");

    for (const d of getDateRange(new Date(startStr), new Date(endStr))) {
      if ((blocked[name] || new Set()).has(d)) {
        Logger.log(`🚫 דילוג על ${shift} ל-${name} בתאריך ${d} עקב חסימת זמינות`);
        continue;
      }

      const idx = shiftRows.findIndex(r =>
        Utilities.formatDate(new Date(r[dateCol]), tz, "yyyy-MM-dd") === d &&
        r[shiftCol] === shift &&
        !r[nameCol] // empty
      );

      if (idx !== -1) {
        const rowIndex = idx + 2; // account for header
        scheduleSheet.getRange(rowIndex, nameCol + 1).setValue(name);
        Logger.log(`✅ שיבוץ ${name} ל-${shift} בתאריך ${d}`);
      }
    }
  }

  adjustColumnWidthsByContent(scheduleSheet);
}
