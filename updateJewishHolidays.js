function updateJewishHolidays() {
  const year = new Date().getFullYear();
  const url = `https://www.hebcal.com/hebcal?v=1&year=${year}&cfg=json&maj=on&mod=on&ss=on&mf=on&c=on&geo=IL&m=50&s=on`;
  const response = UrlFetchApp.fetch(url);
  const data = JSON.parse(response.getContentText());

  let sheet = SpreadsheetApp.getActive().getSheetByName("חגים");
  if (!sheet) sheet = SpreadsheetApp.getActive().insertSheet("חגים");
  else sheet.clear();

  sheet.appendRow(["תאריך", "חג", "סוג"]);
  sheet.getRange("A1:C1").setFontWeight("bold").setHorizontalAlignment("right");

  const translations = {
    "Rosh Hashana 5786": "ראש השנה",
    "Rosh Hashana II": "ראש השנה ב",
    "Yom Kippur": "יום כיפור",
    "Sukkot I": "סוכות א",
    "Shmini Atzeret": "שמיני עצרת",
    "Simchat Torah": "שמחת תורה",
    "Pesach I": "פסח א",
    "Pesach VII": "שביעי של פסח",
    "Shavuot I": "שבועות א",
    "Yom HaAtzma’ut": "יום העצמאות",
    "Chanukah: 1 Candle": "חנוכה",
    "Chanukah: 2 Candles": "חנוכה",
    "Chanukah: 3 Candles": "חנוכה",
    "Chanukah: 4 Candles": "חנוכה",
    "Chanukah: 5 Candles": "חנוכה",
    "Chanukah: 6 Candles": "חנוכה",
    "Chanukah: 7 Candles": "חנוכה",
    "Chanukah: 8 Candles": "חנוכה",
    "Asara B’Tevet": "עשרה בטבת",
    "Ta’anit Esther": "תענית אסתר",
    "Tzom Tammuz": "צום י״ז בתמוז",
    "Erev Tish’a B’Av": "ערב תשעה באב",
    "Tish’a B’Av": "תשעה באב",
    "Tzom Gedaliah": "צום גדליה",
    "Yom HaShoah": "יום השואה",
    "Yom HaZikaron": "יום הזיכרון"
  };

  const restDays = [
    "ראש השנה", "ראש השנה ב", "יום כיפור",
    "סוכות א", "שמיני עצרת", "שמחת תורה",
    "פסח א", "שביעי של פסח",
    "שבועות א", "יום העצמאות"
  ];

  const holidays = data.items.filter(e => e.category === "holiday");

  holidays.forEach(item => {
    const date = new Date(item.date);
    const title = translations[item.title] || item.title;
    const type = restDays.includes(title) ? "חופש" : "מידע";
    sheet.appendRow([date, title, type]);
  });

  sheet.getRange(`A2:C${sheet.getLastRow()}`).setHorizontalAlignment("right");
}

