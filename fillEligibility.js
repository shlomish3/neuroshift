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
    if (shiftColIndex === -1) return;

    const eligible = empList
      .filter(emp => {
        const name = emp[0];
        const isEligible = emp[shiftColIndex + 1] === 1 || emp[shiftColIndex + 1] === "1";
        if (!isEligible) return false;

        // Find all blocking types for this user/date
        const blockTypes = parsedRows
          .filter(([n, d]) => n === name && Utilities.formatDate(new Date(d), tz, "yyyy-MM-dd") === dateFormatted)
          .map(([, , , type]) => String(type || ""));

        const blocksToran = blockTypes.includes("לא זמין לתורנות");
        const blocksGeneral = blockTypes.includes("לא זמין");

        if (blocksToran && ["תורן מיון", "תורן חצי", "כונן מיון"].includes(shiftType)) {
          return false;
        }

        if (blocksGeneral && !["תורן מיון", "תורן חצי", "כונן מיון"].includes(shiftType)) {
          return false;
        }

        return true;
      })
      .map(emp => emp[0]);

    const fixed = fixedRules.find(r => {
      if (!r[0] || !r[1] || !r[2] || !r[3]) return false;
      const [shift, start, end, name] = r;
      const startStr = Utilities.formatDate(new Date(start), tz, "yyyy-MM-dd");
      const endStr = Utilities.formatDate(new Date(end), tz, "yyyy-MM-dd");
      return shift === shiftType && dateFormatted >= startStr && dateFormatted <= endStr;
    });

    if (fixed) {
      const fixedName = fixed[3];

      // Apply the same block logic to fixed names
      const blockTypes = parsedRows
        .filter(([n, d]) => n === fixedName && Utilities.formatDate(new Date(d), tz, "yyyy-MM-dd") === dateFormatted)
        .map(([, , , type]) => String(type || ""));

      const blocksToran = blockTypes.includes("לא זמין לתורנות");
      const blocksGeneral = blockTypes.includes("לא זמין");

      if (
        (blocksToran && ["תורן מיון", "תורן חצי", "כונן מיון"].includes(shiftType)) ||
        (blocksGeneral && !["תורן מיון", "תורן חצי", "כונן מיון"].includes(shiftType))
      ) {
        Logger.log(`❌ שיבוץ קבוע עבור ${fixedName} ב-${dateFormatted} נדחה (חסם זמינות)`);
        return;
      }

      const markAsFixed = !["תורן מיון", "כונן מיון", "תורן חצי"].includes(shiftType);
      scheduleSheet.getRange(rowIndex, 4).setValue(markAsFixed ? `${fixedName} (קבוע)` : fixedName);

      const fullEligibility = Array.from(new Set([fixedName, ...eligible]));
      scheduleSheet.getRange(rowIndex, 5).setValue(fullEligibility.join(", "));
      return;
    }

    // No fixed assignment — just normal eligibility
    scheduleSheet.getRange(rowIndex, 5).setValue(eligible.join(", "));
  });

  Logger.log("השלמת חישוב זכאים לפי זמינות מפורקת ושיבוצים קבועים.");
}
