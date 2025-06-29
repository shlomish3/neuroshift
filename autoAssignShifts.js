function autoAssignShifts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const scheduleSheet = getScheduleSheet(ss);
  if (!scheduleSheet) return;

  const allowedPostClinicDates = getAllowedPostClinicDates();
  const { rows, nameCol, eligibleCol, neededCol } = getScheduleData(scheduleSheet);

  const fairShifts = ["תורן מיון", "מיון", "בכיר מיון", "תורן חצי", "כונן מיון", "ייעוצים מובילים"];
  const doubleAllowed = new Set(["תורן מיון", "בכיר מיון", "כונן מיון", "תורן חצי"]);
  const dynamicCaps = computeDynamicCaps(rows, fairShifts);
  const dailyAssignments = {};
  // 🔹 Build blockTypesMap from "זמינות מפורקת"
  const availabilitySheet = ss.getSheetByName("זמינות מפורקת");
  const tz = ss.getSpreadsheetTimeZone();
  const blockTypesMap = {};
  const empSheet = ss.getSheetByName("עובדים");
  const empData = empSheet.getDataRange().getValues();
  const empHeaders = empData[0].slice(1); // shift types
  const empList = empData.slice(1);       // name + shift eligibility

  if (availabilitySheet) {
    const data = availabilitySheet.getDataRange().getValues();
    const rows = data.slice(1); // skip header
    rows.forEach(([name, dateStr, , type]) => {
      if (!name || !dateStr || !type) return;
      const normalized = Utilities.formatDate(new Date(dateStr), tz, "yyyy-MM-dd");
      if (!blockTypesMap[name]) blockTypesMap[name] = {};
      if (!blockTypesMap[name][normalized]) blockTypesMap[name][normalized] = new Set();
      blockTypesMap[name][normalized].add(String(type));
    });
}


  const {
    assignmentCounts, manualCounts, blockedDates,
    sameDayAssignments, dateMap, toranMionLog
  } = initializeAssignmentState(rows, nameCol, eligibleCol, fairShifts);
  cleanAssignmentCounts(assignmentCounts, fairShifts);
  const pastCounts = mergePastAndManualCounts(countPastShiftLoads(fairShifts), manualCounts);
  rows.sort((a, b) => new Date(a[0]) - new Date(b[0]));  // ✅ Sort by date

  // 🔹 Phase 1: Only assign תורן מיון and תורן חצי
  const toranRows = rows
    .map((row, i) => ({ row, i }))
    .filter(({ row }) => ["תורן מיון", "תורן חצי"].includes(row[2]));

  toranRows.forEach(({ row, i }) => {
    const rowIndex = i + 2;
    const shiftType = row[2];
    const dateStr = row[0];
    const nameCell = scheduleSheet.getRange(rowIndex, nameCol + 1);
    const current = nameCell.getValue().toString().trim();

    if (!current) {
      assignShiftToRow({
        row, i, rows, scheduleSheet, nameCol, eligibleCol, neededCol,
        fairShifts, doubleAllowed, pastCounts, dynamicCaps,
        assignmentCounts, blockedDates, sameDayAssignments,
        dateMap, toranMionLog, dailyAssignments, blockTypesMap, empList, empHeaders
      });

      const assigned = nameCell.getValue().toString().trim();
      if (!assigned) {
        Logger.log(`❌ לא הצלחנו לשבץ ל-${shiftType} בתאריך ${dateStr}`);
      }
    }
  });

  // 🔹 Phase 2: Continue with all other shifts
  const phaseGroups = { מרפאות: [], תורן_כונן: [], שישי_שבת: [], ייעוצים: [], מחלקה: [] };
  rows.forEach((row, i) => {
    const [dateStr, , shiftType] = row;
    const day = new Date(dateStr).getDay();
    if (shiftType.startsWith("מרפאת")) phaseGroups.מרפאות.push({ row, i });
    else if (["תורן מיון", "כונן מיון"].includes(shiftType)) phaseGroups.תורן_כונן.push({ row, i });
    else if (day >= 5) phaseGroups.שישי_שבת.push({ row, i });
    else if (["ייעוצים מובילים", "EEG", "EMG"].includes(shiftType)) phaseGroups.ייעוצים.push({ row, i });
    else if (shiftType === "מחלקה") phaseGroups.מחלקה.push({ row, i });
  });

  for (const phase of ["מרפאות", "תורן_כונן", "שישי_שבת", "ייעוצים", "מחלקה"]) {
    phaseGroups[phase].forEach(({ row, i }) => {
      const [dateStr, , shiftType] = row;
      const date = new Date(dateStr);
      const day = date.getDay();
      const rowIndex = i + 2;
      const nameCell = scheduleSheet.getRange(rowIndex, nameCol + 1);

      // Skip if already filled
      if (nameCell.getValue().toString().trim()) return;

      if (shiftType === "מרפאת פוסט אשפוז") {
        const formatted = Utilities.formatDate(date, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");
        if (!allowedPostClinicDates.has(formatted)) {
          Logger.log(`🛑 מדלג על מרפאת פוסט אשפוז ב-${formatted}`);
          nameCell.setValue("אין מרפאה");
          return;
        }
      }

      assignShiftToRow({
        row, i, rows, scheduleSheet, nameCol, eligibleCol, neededCol,
        fairShifts, doubleAllowed, pastCounts, dynamicCaps,
        assignmentCounts, blockedDates, sameDayAssignments,
        dateMap, toranMionLog, dailyAssignments, blockTypesMap, empList, empHeaders
      });

      const getAssigned = () => nameCell.getValue().toString().split(" ")[0];

      // 🔁 Chain logic
      if (day === 5 && shiftType === "תורן מיון") {
        const idx = rows.findIndex(r => r[0] === dateStr && r[2] === "מיון");
        if (idx !== -1 && !rows[idx][3]) {
          const name = getAssigned();
          rows[idx][3] = name;
          scheduleSheet.getRange(idx + 2, nameCol + 1).setValue(name);
          updateTracking(name, "מיון", dateStr, assignmentCounts, sameDayAssignments, dailyAssignments);
        }
      }

      if (day === 5 && shiftType === "בכיר מיון") {
        const idx = rows.findIndex(r => r[0] === dateStr && r[2] === "כונן מיון");
        if (idx !== -1 && !rows[idx][3]) {
          const name = getAssigned();
          rows[idx][3] = name;
          scheduleSheet.getRange(idx + 2, nameCol + 1).setValue(name);
          updateTracking(name, "כונן מיון", dateStr, assignmentCounts, sameDayAssignments, dailyAssignments);
        }
      }

      if (day === 6 && shiftType === "תורן מיון") {
        const friday = new Date(date); friday.setDate(friday.getDate() - 1);
        const fridayStr = Utilities.formatDate(friday, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");
        const idx = rows.findIndex(r => r[0] === fridayStr && r[2] === "מחלקה");
        if (idx !== -1) {
          const name = getAssigned();
          const prev = rows[idx][3] || "";
          if (!prev.includes(name)) {
            const updated = prev ? `${prev}, ${name}` : name;
            rows[idx][3] = updated;
            scheduleSheet.getRange(idx + 2, nameCol + 1).setValue(updated);
            updateTracking(name, "מחלקה", fridayStr, assignmentCounts, sameDayAssignments, dailyAssignments);
          }
        }
      }
    });
  }

  // 📊 Show assignment summary
  const summary = Object.entries(assignmentCounts).map(
    ([name, counts]) => `${name}: ${Object.entries(counts).map(([s, c]) => `${s}=${c}`).join(", ")}`
  ).sort().join("\n");

  try {
    SpreadsheetApp.getUi().alert("השיבוץ האוטומטי הושלם:\n\n" + summary);
  } catch (e) {
    Logger.log(summary);
  }

  adjustColumnWidthsByContent(scheduleSheet);
}
