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

  const titleToHebrew = {
    "Rosh Hashana 5786": "ראש השנה",
    "Rosh Hashana II": "ראש השנה ב",
    "Erev Rosh Hashana": "ערב ראש השנה",
    "Erev Yom Kippur": "ערב יום כיפור",
    "Yom Kippur": "יום כיפור",
    "Erev Sukkot": "ערב סוכות",
    "Sukkot I": "סוכות יום א",
    "Sukkot II": "סוכות יום ב (חול המועד)",
    "Sukkot III (CH’’M)": "סוכות יום ג (חול המועד)",
    "Sukkot IV (CH’’M)": "סוכות יום ד (חול המועד)",
    "Sukkot V (CH’’M)": "סוכות יום ה (חול המועד)",
    "Sukkot VI (CH’’M)": "סוכות יום ו (חול המועד)",
    "Sukkot VII (Hoshana Raba)": "הושענא רבה",
    "Shmini Atzeret": "שמיני עצרת",
    "Simchat Torah": "שמחת תורה",
    "Erev Pesach": "ערב פסח",
    "Pesach I": "פסח יום א",
    "Pesach II": "פסח יום ב (חול המועד)",
    "Pesach III (CH’’M)": "פסח יום ג (חול המועד)",
    "Pesach IV (CH’’M)": "פסח יום ד (חול המועד)",
    "Pesach V (CH’’M)": "פסח יום ה (חול המועד)",
    "Pesach VI (CH’’M)": "פסח יום ו (חול המועד)",
    "Pesach VII": "שביעי של פסח",
    "Erev Shavuot": "ערב שבועות",
    "Shavuot I": "שבועות",
    "Erev Purim": "ערב פורים",
    "Purim": "פורים",
    "Ta’anit Esther": "תענית אסתר",
    "Asara B’Tevet": "עשרה בטבת",
    "Tzom Gedaliah": "צום גדליה",
    "Tzom Tammuz": "צום י״ז בתמוז",
    "Erev Tish’a B’Av": "ערב תשעה באב",
    "Tish’a B’Av": "תשעה באב",
    "Chanukah: 1 Candle": "חנוכה נר 1",
    "Chanukah: 2 Candles": "חנוכה נר 2",
    "Chanukah: 3 Candles": "חנוכה נר 3",
    "Chanukah: 4 Candles": "חנוכה נר 4",
    "Chanukah: 5 Candles": "חנוכה נר 5",
    "Chanukah: 6 Candles": "חנוכה נר 6",
    "Chanukah: 7 Candles": "חנוכה נר 7",
    "Chanukah: 8 Candles": "חנוכה נר 8",
    "Yom HaAtzma’ut": "יום העצמאות",
    "Yom HaZikaron": "יום הזיכרון",
    "Yom HaShoah": "יום השואה",
    "Sigd": "סיגד",
    "Yom Yerushalayim": "יום ירושלים",
  };

  const restDays = [
    "ראש השנה", "ראש השנה ב", "ערב יום כיפור", "יום כיפור", "ערב ראש השנה",
    "ערב סוכות", "סוכות יום א", "שמיני עצרת", "שמחת תורה",
    "ערב פסח", "פסח יום א", "שביעי של פסח",
    "ערב שבועות", "שבועות",
    "יום העצמאות"
  ];

  const holidays = data.items.filter(e => e.category === "holiday");

  holidays.forEach(item => {
    const date = new Date(item.date);
    const hebrew = titleToHebrew[item.title];
    if (!hebrew) return;

    const type = restDays.includes(hebrew) ? "חופש" : "מידע";
    sheet.appendRow([date, hebrew, type]);
  });

  sheet.getRange(`A2:C${sheet.getLastRow()}`).setHorizontalAlignment("right");
}