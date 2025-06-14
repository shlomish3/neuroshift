function countPastShiftLoads(shiftTypes) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const counts = {};

  ss.getSheets()
    .filter(s => s.getName().startsWith("משמרות "))
    .forEach(sheet => {
      const data = sheet.getDataRange().getValues();
      data.slice(1).forEach(row => {
        const [ , , shift, name] = row;
        if (!shiftTypes.includes(shift)) return;
        if (!name) return;

        name.split(",").map(n => n.trim()).forEach(n => {
          if (!counts[n]) counts[n] = {};
          counts[n][shift] = (counts[n][shift] || 0) + 1;
        });
      });
    });

  return counts;
}

function mergePastAndManualCounts(pastCounts, manualCounts) {
  for (const [name, shiftMap] of Object.entries(manualCounts)) {
    if (!pastCounts[name]) pastCounts[name] = {};
    for (const [shift, count] of Object.entries(shiftMap)) {
      pastCounts[name][shift] = (pastCounts[name][shift] || 0) + count;
    }
  }
  return pastCounts;
}
