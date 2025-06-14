function autoAssignShifts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const scheduleSheet = getScheduleSheet(ss);
  if (!scheduleSheet) return;

  const { rows, nameCol, eligibleCol, neededCol } = getScheduleData(scheduleSheet);
  const fairShifts = ["תורן מיון", "מיון", "בכיר מיון", "תורן חצי", "כונן מיון", "ייעוצים מובילים"];
  const doubleAllowed = new Set(["תורן מיון", "בכיר מיון", "כונן מיון", "תורן חצי"]);
  const dynamicCaps = computeDynamicCaps(rows, fairShifts);
  const dailyAssignments = {};

  // Initialize assignment trackers
  const {
    assignmentCounts, manualCounts, blockedDates, sameDayAssignments, dateMap, toranMionLog
  } = initializeAssignmentState(rows, nameCol, eligibleCol, fairShifts);

  cleanAssignmentCounts(assignmentCounts, fairShifts);
  const pastCounts = mergePastAndManualCounts(countPastShiftLoads(fairShifts), manualCounts);

  // === Modular Phased Assignment ===
  const phaseGroups = {
    מרפאות: [],
    תורן_כונן: [],
    שישי_שבת: [],
    ייעוצים: [],
    מחלקה: []
  };

  rows.forEach((row, i) => {
    const [dateStr, , shiftType] = row;
    const date = new Date(dateStr);
    const day = date.getDay(); // 0=Sunday, ..., 6=Saturday

    if (shiftType.startsWith("מרפאת")) {
      phaseGroups.מרפאות.push({ row, i });
    } else if (["תורן מיון", "כונן מיון"].includes(shiftType)) {
      phaseGroups.תורן_כונן.push({ row, i });
    } else if (day === 5 || day === 6) {
      phaseGroups.שישי_שבת.push({ row, i });
    } else if (["ייעוצים מובילים", "EEG", "EMG"].includes(shiftType)) {
      phaseGroups.ייעוצים.push({ row, i });
    } else if (shiftType === "מחלקה") {
      phaseGroups.מחלקה.push({ row, i });
    }
  });

  const phaseOrder = ["מרפאות", "תורן_כונן", "שישי_שבת", "ייעוצים", "מחלקה"];
  phaseGroups[phase].forEach(({ row, i }) => {
  const [dateStr, , shiftType] = row;
  const date = new Date(dateStr);
  const day = date.getDay(); // 5 = Friday, 6 = Saturday

  assignShiftToRow({
    row, i, rows, scheduleSheet, nameCol, eligibleCol, neededCol,
    fairShifts, doubleAllowed, pastCounts, dynamicCaps,
    assignmentCounts, blockedDates, sameDayAssignments,
    dateMap, toranMionLog, dailyAssignments
  });

  const rowIndex = i + 2;

  // 1. Friday תורן מיון → also assign to מיון
  if (day === 5 && shiftType === "תורן מיון") {
    const miunRowIndex = rows.findIndex(r => r[0] === dateStr && r[2] === "מיון");
    if (miunRowIndex !== -1) {
      const miunRow = rows[miunRowIndex];
      const toran = scheduleSheet.getRange(rowIndex, nameCol + 1).getValue().toString().split(" ")[0]; // remove markers
      if (toran && !miunRow[3]) {
        rows[miunRowIndex][3] = toran;
        scheduleSheet.getRange(miunRowIndex + 2, nameCol + 1).setValue(toran);
        updateTracking(toran, "מיון", dateStr, assignmentCounts, sameDayAssignments, dailyAssignments);
      }
    }
  }

  // 2. Friday בכיר מיון → also assign to כונן מיון
  if (day === 5 && shiftType === "בכיר מיון") {
    const conanRowIndex = rows.findIndex(r => r[0] === dateStr && r[2] === "כונן מיון");
    if (conanRowIndex !== -1) {
      const conanRow = rows[conanRowIndex];
      const bakhir = scheduleSheet.getRange(rowIndex, nameCol + 1).getValue().toString().split(" ")[0];
      if (bakhir && !conanRow[3]) {
        rows[conanRowIndex][3] = bakhir;
        scheduleSheet.getRange(conanRowIndex + 2, nameCol + 1).setValue(bakhir);
        updateTracking(bakhir, "כונן מיון", dateStr, assignmentCounts, sameDayAssignments, dailyAssignments);
      }
    }
  }

  // 3. Saturday תורן מיון → also assign to Friday מחלקה
  if (day === 6 && shiftType === "תורן מיון") {
    const friday = new Date(date);
    friday.setDate(date.getDate() - 1);
    const fridayStr = Utilities.formatDate(friday, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");

    const machlakaRowIndex = rows.findIndex(r => r[0] === fridayStr && r[2] === "מחלקה");
    if (machlakaRowIndex !== -1) {
      const machlakaRow = rows[machlakaRowIndex];
      const toran = scheduleSheet.getRange(rowIndex, nameCol + 1).getValue().toString().split(" ")[0];
      if (toran && (!machlakaRow[3] || !machlakaRow[3].includes(toran))) {
        const prev = machlakaRow[3] || "";
        const updated = prev ? `${prev}, ${toran}` : toran;
        rows[machlakaRowIndex][3] = updated;
        scheduleSheet.getRange(machlakaRowIndex + 2, nameCol + 1).setValue(updated);
        updateTracking(toran, "מחלקה", fridayStr, assignmentCounts, sameDayAssignments, dailyAssignments);
      }
    }
  }

});


  // Mark extra day off eligibility based on תורן מיון
  const extraDayOffEligibleMap = markExtraDayOffEligible(toranMionLog);
  Logger.log("EXTRA-DAY-OFF MAP: " + JSON.stringify(Object.fromEntries(
    Array.from(extraDayOffEligibleMap).map(([k, v]) => [k, Array.from(v)])
  )));

  // Annotate תורן מיון rows with this information
  annotateEligibleNames({
    rows, scheduleSheet, nameCol, eligibleCol, extraDayOffEligibleMap
  });

  // Summary of assignments
  const summary = Object.entries(assignmentCounts)
    .map(([name, shiftObj]) => {
      const shiftsSummary = Object.entries(shiftObj)
        .map(([shift, count]) => `${shift}=${count}`)
        .join(", ");
      return `${name}: ${shiftsSummary}`;
    })
    .sort()
    .join("\n");

  try {
    SpreadsheetApp.getUi().alert("השיבוץ האוטומטי הושלם:\n\n" + summary);
  } catch (e) {
    Logger.log(summary);
  }

  adjustColumnWidthsByContent(scheduleSheet);
}
