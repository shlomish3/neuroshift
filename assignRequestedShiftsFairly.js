function assignRequestedShiftsFairly() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const requestSheet = ss.getSheetByName("זמינות מפורקת");
  const fixedSheet = ss.getSheetByName("שיבוצים קבועים") || ss.insertSheet("שיבוצים קבועים");
  const workersSheet = ss.getSheetByName("עובדים");

  const shiftTypeMap = buildWorkerShiftMap(workersSheet);
  const fixedMap = getFixedAssignmentsMap(fixedSheet);

  const reqData = requestSheet.getDataRange().getValues().slice(1);
  const requestsMap = new Map();

  for (const row of reqData) {
    const [name, date, , type] = row;
    if (type === "מבקש תורנות" && date) {
      if (!requestsMap.has(date)) requestsMap.set(date, []);
      requestsMap.get(date).push(name);
    }
  }

  const assigned = new Set();
  for (const dateStr of [...requestsMap.keys()].sort()) {
    const requesters = requestsMap.get(dateStr);
    const conflict = requesters.filter(n => fixedMap.has(dateStr) && fixedMap.get(dateStr).includes(n));
    const eligible = requesters.filter(n => !assigned.has(n) && !conflict.includes(n));
    const pool = eligible.length ? eligible : requesters.filter(n => !conflict.includes(n));
    if (!pool.length) {
      logConflict(dateStr, requesters, null, conflict);
      continue;
    }

    const chosen = chooseFairAssignee(pool);
    const shift = shiftTypeMap.get(chosen) || "תורן מיון";
    assigned.add(chosen);

    // Prevent duplicate assignment
    const alreadyAssigned = fixedMap.has(dateStr) && fixedMap.get(dateStr).includes(chosen);
    if (!alreadyAssigned) {
      fixedSheet.appendRow([shift, dateStr, dateStr, chosen]);
    }

    logConflict(dateStr, requesters, chosen, conflict);
  }
}
