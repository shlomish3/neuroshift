function markExtraDayOffEligible(toranMionLog) {
  const eligibleMap = new Map();

  for (const [name, dateStrs] of Object.entries(toranMionLog)) {
    const sortedDates = dateStrs.map(d => new Date(d)).sort((a, b) => a - b);
    const times = sortedDates.map(d => d.getTime());

    for (let i = 0; i < sortedDates.length; i++) {
      const date = sortedDates[i];
      const day = date.getDay(); // 0 = Sun ... 6 = Sat

      const prev = times[i - 1] ?? 0;
      const curr = times[i];
      const next = times[i + 1] ?? 0;

      const diffPrev = (curr - prev) / 86400000;
      const diffNext = (next - curr) / 86400000;

      if (
        day === 5 ||                          // Friday alone
        (day === 4 && diffNext === 2) ||      // Thursday + Saturday
        (day === 5 && diffNext === 1)         // Friday + Saturday
      ) {
        const triggeringDate = sortedDates[i].toISOString().split('T')[0];
        if (!eligibleMap.has(name)) eligibleMap.set(name, new Set());
        eligibleMap.get(name).add(triggeringDate);
      }
    }
  }

  return eligibleMap;
}


function getWeekStart(dateStr) {
  const date = new Date(dateStr);
  const day = date.getDay(); // 0 = Sunday
  const sunday = new Date(date);
  sunday.setDate(date.getDate() - day);
  sunday.setHours(0, 0, 0, 0);
  return sunday;
}

function annotateEligibleNames({
  rows, scheduleSheet, nameCol, eligibleCol, extraDayOffEligibleMap
}) {
  rows.forEach((row, i) => {
    const rowIndex = i + 2;
    const [dateStr, , shiftType, nameStr, eligiblesStr] = row;

    // Only mark rows that are תורן מיון
    if (shiftType !== "תורן מיון") return;

    for (const [name, dateSet] of extraDayOffEligibleMap.entries()) {
      if (nameStr.includes(name) || eligiblesStr.includes(name)) {
        const rowDate = new Date(dateStr).toISOString().split("T")[0];
        if (!dateSet.has(rowDate)) continue;

        const matchPattern = new RegExp(`(^|,\\s*)${name}(?=,|$)`);
        let newNameVal = nameStr;
        let newEligiblesVal = eligiblesStr;

        if (!newNameVal.includes(`${name} 😴`)) {
          newNameVal = newNameVal.replace(matchPattern, `$1${name} 😴`);
        }
        if (!newEligiblesVal.includes(`${name} 😴`)) {
          newEligiblesVal = newEligiblesVal.replace(matchPattern, `$1${name} 😴`);
        }

        if (newNameVal !== nameStr)
          scheduleSheet.getRange(rowIndex, nameCol + 1).setValue(newNameVal);
        if (newEligiblesVal !== eligiblesStr)
          scheduleSheet.getRange(rowIndex, eligibleCol + 1).setValue(newEligiblesVal);
      }
    }
  });
}

