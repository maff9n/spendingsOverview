const creditLeft = {
  cornerStone: "A1",
  display: function(sheet){
    const cornerStone = sheet.getRange(this.cornerStone);

    const incomeRange = sheet.getRange(income.cornerStone).offset(2,1).getA1Notation() + ':' + sheet.getRange(income.cornerStone).offset(2+income.listLength,1).getA1Notation();

    const spendingsRange = sheet.getRange(spendings.cornerStone).offset(2,1).getA1Notation() + ':' + sheet.getRange(spendings.cornerStone).offset(2+spendings.listLength,1).getA1Notation();

    cornerStone.setValue("Restliches Geld");
    cornerStone.offset(1,0).setFormula("=SUM(" + incomeRange + ")-SUM(" + spendingsRange + ")");
  }
}

const spendings = {
  cornerStone: "A4",
  listLength : 100,
  headers: ["Beschreibung", "Betrag", "Wiederkehrend", "Start Monat"],
  recurringTypes: ['Ausstehend', 'Monatlich', 'Zweimonatlich', 'Vierteljährlich', 'Halbjährlich', 'Jährlich'],
  months: ['Januar', 'Februar', 'März', 'April', 'Mai', 'Juni', 'Juli', 'August', 'September', 'Oktober', 'November', 'Dezember'],
  display: function(sheet) {
    const titleCell = sheet.getRange(this.cornerStone).setValue("Ausgaben");
    
    // Set headers
    this.headers.forEach((header, index) => {
      titleCell.offset(1, index).setValue(header);
    });
    
    // Apply data validation
    this.applyDataValidation(sheet);
  },
  applyDataValidation: function(sheet) {
    const cornerStone = sheet.getRange(this.cornerStone);
    const typeRange = sheet.getRange(cornerStone.offset(2, 2).getA1Notation() + ':' + cornerStone.offset(2+this.listLength, 2).getA1Notation());
    const monthRange = sheet.getRange(cornerStone.offset(2, 3).getA1Notation() + ':' + cornerStone.offset(2+this.listLength, 3).getA1Notation());

    
    typeRange.setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(this.recurringTypes)
        .setAllowInvalid(false)
        .build()
    );
    
    monthRange.setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(this.months)
        .setAllowInvalid(false)
        .build()
    );
  },
  parse: function(sheet, previousSheet){
    const prevCornerStone = previousSheet.getRange(this.cornerStone);
    const cornerStone = sheet.getRange(this.cornerStone);
    let prevEntries = previousSheet.getRange(prevCornerStone.offset(2,0).getA1Notation() + ':' + prevCornerStone.offset(2+this.listLength, 3).getA1Notation()).getValues();

    prevEntries = prevEntries.filter(obj => obj[2] !== '' && obj[2] !== null)
                    .sort((first, second) => first[1] - second[1])
                      .reverse()
    
    sheet.getRange(cornerStone.offset(2, 0).getA1Notation() + ':' + cornerStone.offset(2+prevEntries.length-1, 3).getA1Notation()).setValues(prevEntries);
  }
}

const income = {
  cornerStone: "F1",
  listLength: 4,
  headers: ["Beschreibung", "Betrag", "Wiederkehrend", "Start Monat"],
  recurringTypes: ['Ausstehend', 'Monatlich', 'Zweimonatlich', 'Vierteljährlich', 'Halbjährlich', 'Jährlich'],
  months: ['Januar', 'Februar', 'März', 'April', 'Mai', 'Juni', 'Juli', 'August', 'September', 'Oktober', 'November', 'Dezember'],
  display: function(sheet) {
    const titleCell = sheet.getRange(this.cornerStone).setValue("Einkommen");
    
    // Set headers
    this.headers.forEach((header, index) => {
      titleCell.offset(1, index).setValue(header);
    });
    
    // Apply data validation
    this.applyDataValidation(sheet);
  },
  applyDataValidation: function(sheet) {
    const cornerStone = sheet.getRange(this.cornerStone);
    const typeRange = sheet.getRange(cornerStone.offset(2, 2).getA1Notation() + ':' + cornerStone.offset(2+this.listLength, 2).getA1Notation());
    const monthRange = sheet.getRange(cornerStone.offset(2, 3).getA1Notation() + ':' + cornerStone.offset(2+this.listLength, 3).getA1Notation());
    
    typeRange.setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(this.recurringTypes)
        .setAllowInvalid(false)
        .build()
    );
    
    monthRange.setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(this.months)
        .setAllowInvalid(false)
        .build()
    );
  },
  parse: function(sheet, previousSheet){
    const prevCornerStone = previousSheet.getRange(this.cornerStone);
    const cornerStone = sheet.getRange(this.cornerStone);
    let prevEntries = previousSheet.getRange(prevCornerStone.offset(2,0).getA1Notation() + ':' + prevCornerStone.offset(2+this.listLength, 3).getA1Notation()).getValues();

    prevEntries = prevEntries.filter(obj => obj[2] !== '' && obj[2] !== null)
                    .sort((first, second) => first[1] - second[1])
                      .reverse()
    
    sheet.getRange(cornerStone.offset(2, 0).getA1Notation() + ':' + cornerStone.offset(2+prevEntries.length-1, 3).getA1Notation()).setValues(prevEntries);
  }
}

const motivation = {
  cornerStone: "F9",
  headers: ["Beschreibung", "Betrag"],
  display: function(sheet) {
    const titleCell = sheet.getRange(this.cornerStone).setValue("Ziele / Wünsche");
    
    // Set headers
    this.headers.forEach((header, index) => {
      titleCell.offset(1, index).setValue(header);
    });
  },
}

function createSheetName(situation){
  const dateObj = new Date();
  if (situation == "test"){
    const sheetName = `${dateObj.getUTCFullYear()}/${dateObj.getUTCMonth() + 1}/${dateObj.getUTCDate()}/${dateObj.getUTCMinutes()}/${dateObj.getUTCSeconds()}`;
    Logger.log("createSheetName() returns the following string: " + sheetName);
    return sheetName
  } else {
    const sheetName = `${dateObj.getUTCFullYear()}/${dateObj.getUTCMonth() + 1}`;
    Logger.log("createSheetName() returns the following string: " + sheetName);
    return sheetName
  }
}

function newMonthNewTab(){

  const previousSheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const newSheet = SpreadsheetApp.getActiveSpreadsheet();
  newSheet.insertSheet(createSheetName("test"));
  newSheet.moveActiveSheet(0);

  creditLeft.display(newSheet)
  spendings.display(newSheet)
  income.display(newSheet)
  motivation.display(newSheet)

  income.parse(newSheet, previousSheet)
  spendings.parse(newSheet, previousSheet)

}