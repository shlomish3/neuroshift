function onOpen() {
  SpreadsheetApp.getUi().createMenu("שיבוץ")
    .addItem("צור גיליון לחודש הבא", "createNextMonthSchedule")  
    .addItem("חשב זכאים", "fillEligibility")
    .addItem("שבץ אוטומטית", "autoAssignShifts")
    .addSeparator()
    .addItem("עדכן חגים", "updateJewishHolidays")
    .addToUi();
}
