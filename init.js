function getScheduleSheet(ss) {
  return ss.getSheets()
    .filter(s => s.getName().startsWith("משמרות "))
    .sort((a, b) => b.getName().localeCompare(a.getName()))[0];
}

function getScheduleData(sheet) {
  const data = sheet.getDataRange().getValues();
  return {
    rows: data.slice(1),
    nameCol: 3,
    eligibleCol: 4,
    neededCol: 5
  };
}

function initializeAssignmentState(rows, nameCol, eligibleCol, fairShifts) {
  const assignmentCounts = {};
  const manualCounts = {};
  const blockedDates = {};
  const sameDayAssignments = {};
  const dateMap = {};
  const toranMionLog = {};

  rows.forEach((row, i) => {
    const date = row[0];
    const shiftType = row[2];
    const nameStr = row[nameCol];
    const eligibleStr = row[eligibleCol];
    dateMap[i] = date;

    // Process assigned names
    if (nameStr) {
      nameStr.split(",").map(n => n.trim()).forEach(name => {
        if (!assignmentCounts[name]) assignmentCounts[name] = {};
        if (!manualCounts[name]) manualCounts[name] = {};
        fairShifts.forEach(shift => {
          if (assignmentCounts[name][shift] === undefined) assignmentCounts[name][shift] = 0;
          if (manualCounts[name][shift] === undefined) manualCounts[name][shift] = 0;
        });
        assignmentCounts[name][shiftType] += 1;
        manualCounts[name][shiftType] += 1;

        if (!sameDayAssignments[date]) sameDayAssignments[date] = new Set();
        sameDayAssignments[date].add(name);
      });
    }

    // Pre-initialize eligible names
    if (eligibleStr) {
      eligibleStr.split(",").map(n => n.trim()).forEach(name => {
        if (!assignmentCounts[name]) assignmentCounts[name] = {};
        fairShifts.forEach(shift => {
          if (assignmentCounts[name][shift] === undefined) {
            assignmentCounts[name][shift] = 0;
          }
        });
      });
    }
  });

  return {
    assignmentCounts,
    manualCounts,
    blockedDates,
    sameDayAssignments,
    dateMap,
    toranMionLog
  };
}

// ✅ Utility function to detect ערב חג (used in multiple files)
function isHolidayEve(title) {
  return typeof title === "string" && title.startsWith("ערב ");
}

function formatAssignmentShortWarning(assigned, totalNeeded) {
  return `❗ רק ${assigned}/${totalNeeded}`;
}