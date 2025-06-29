//Probably unnecessary
function buildBlockedMap(parsedRows, tz) {
  const map = {};
  for (const [name, dateStr, , type] of parsedRows) {
    if (!name || !dateStr) continue;
    const normalized = Utilities.formatDate(new Date(dateStr), tz, "yyyy-MM-dd");
    const blockType = String(type || "");
    if (!blockType.includes("לא זמין")) continue;
    if (!map[name]) map[name] = new Set();
    map[name].add(normalized);
  }
  return map;
}
