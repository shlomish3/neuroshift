
//Helper functions:
function buildWorkerShiftMap(sheet) {
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);

  const options = ["תורן מיון", "תורן חצי", "כונן מיון"];
  const indices = options.map(s => headers.indexOf(s));

  const map = new Map();
  for (const row of rows) {
    const name = row[0];
    for (let i = 0; i < options.length; i++) {
      if (row[indices[i]] === 1) {
        map.set(name, options[i]);
        break;
      }
    }
  }
  return map;
}

function getFixedAssignmentsMap(fixedSheet) {
  const map = new Map();
  const rows = fixedSheet.getDataRange().getValues().slice(1); // skip header

  for (const [shiftType, startRaw, endRaw, name] of rows) {
    Logger.log("📄 Raw row: shiftType=%s | start=%s | end=%s | name=%s", shiftType, startRaw, endRaw, name);

    // Defensive conversion
    const start = (startRaw instanceof Date) ? startRaw : new Date(startRaw);
    const end = (endRaw instanceof Date) ? endRaw : new Date(endRaw);

    const startStr = formatDateSafe(start);
    const endStr = formatDateSafe(end);

    if (startStr === "INVALID-DATE" || endStr === "INVALID-DATE") {
      Logger.log("⚠️ Skipping row due to invalid date: start=%s | end=%s | name=%s", startRaw, endRaw, name);
      continue;
    }

    const range = getDateRange(start, end);
    for (const d of range) {
      if (!map.has(d)) map.set(d, []);
      map.get(d).push(name);
    }
  }

  Logger.log("✅ Finished building fixedMap with %s dates", map.size);
  return map;
}


function chooseFairAssignee(names) {
  return names[Math.floor(Math.random() * names.length)];
}

function logConflict(dateStr, requesters, chosen, fixedConflict) {
  Logger.log(`📅 %s | Requests: %s`, dateStr, requesters.join(", "));
  if (chosen) {
    const losers = requesters.filter(n => n !== chosen);
    Logger.log(`🎯 Assigned: %s`, chosen);
    if (losers.length) Logger.log(`❌ Not Assigned: %s`, losers.join(", "));
  } else {
    Logger.log(`🚫 No eligible assignee.`);
  }
  if (fixedConflict.length) {
    Logger.log(`⚠️ Conflict with שיבוצים קבועים: %s`, fixedConflict.join(", "));
  }
}

function formatDateSafe(date) {
  let d = date;

  if (typeof d === "string") d = new Date(d);

  if (!(d instanceof Date) || isNaN(d.getTime())) {
    Logger.log("❌ formatDateSafe received invalid date: %s", date);
    return "INVALID-DATE";
  }

  return Utilities.formatDate(d, "Asia/Jerusalem", "yyyy-MM-dd");
}


function getDateRange(start, end) {
  const dates = [];
  for (
    let d = new Date(start);
    d <= end;
    d.setDate(d.getDate() + 1)
  ) {
    const copy = new Date(d.getTime());
    const formatted = formatDateSafe(copy);
    if (formatted !== "INVALID-DATE") dates.push(formatted);
  }
  return dates;
}


