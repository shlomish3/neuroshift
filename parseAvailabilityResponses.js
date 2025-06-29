function parseAvailabilityResponses() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const formSheet = ss.getSheetByName("תגובות לטופס זמינות");
  const outputSheet = ss.getSheetByName("זמינות מפורקת") || ss.insertSheet("זמינות מפורקת");
  outputSheet.clear().appendRow(["שם", "תאריך", "יום בשבוע", "סוג חסימה/זמינות", "מקור"]);

  const data = formSheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);

  const nameIndex = headers.indexOf("בחר את שמך");
  const timestampIndex = headers.indexOf("חותמת זמן");
  if (nameIndex === -1 || timestampIndex === -1) throw new Error("Missing required fields");

  const latestResponses = new Map();
  for (const row of rows) {
    const name = row[nameIndex];
    const timestamp = new Date(row[timestampIndex]);
    if (!latestResponses.has(name) || timestamp > latestResponses.get(name).timestamp) {
      latestResponses.set(name, { row, timestamp });
    }
  }

  for (const [name, { row }] of latestResponses.entries()) {
    processResponseRow(name, row, headers, outputSheet);
  }
}
