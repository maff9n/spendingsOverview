const creditLeft = {
  cornerStone: "A1",
  display: function(sheet){
    const cellCreditTitle = sheet.getRange(this.cornerStone);
    const cellCreditValue = cellCreditTitle.offset(1,0);

    cellCreditTitle.setBackground("red").setValue("Restliches Geld");
    cellCreditValue.setBackground("red").setFormula("=SUM(G3:G6)-SUM(B6:B33)");
  }
}

const spendings = {
  cornerStone: "A4",
  headers: ["Beschreibung", "Betrag", "Wiederkehrend?", "Start Monat"],
  typesList: { 
    range: ['Ausstehend', 'Monatlich', 'Zweimonatlich', 'Vierteljährlich', 'Halbjährlich', 'Jährlich'],
    start: "C6",
    end: "C10"
  },
  monthsList: { 
    range: ['Januar', 'Februar', 'März', 'April', 'Mai', 'Juni', 'Juli', 'August', 'September', 'Oktober', 'November', 'Dezember'],
    start: "D6",
    end: "D10"
  },
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
    const typeRange = sheet.getRange(`${this.typesList.start}:${this.typesList.end}`);
    const monthRange = sheet.getRange(`${this.monthsList.start}:${this.monthsList.end}`);
    
    typeRange.setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(this.typesList.range)
        .setAllowInvalid(false)
        .build()
    );
    
    monthRange.setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(this.monthsList.range)
        .setAllowInvalid(false)
        .build()
    );
  },
  parse: function(sheet, previousSheet){
    const titleCell = sheet.getRange(this.cornerStone);
    previousSheet.getRange(`${this.monthsList.start}:${this.monthsList.end}`).getValues().forEach(
    (_a, index) => {
      const entryName = previousSheet.getRange(this.monthsList.start).offset(index, -3)
      const entryAmount = previousSheet.getRange(this.monthsList.start).offset(index, -2)
      const entryRecurring = previousSheet.getRange(this.monthsList.start).offset(index, -1)
      const entryMonth = previousSheet.getRange(this.monthsList.start).offset(index, 0)
      if (entryName.getValue() !== "" && entryName.getValue() !== null) {
        if (entryAmount.getValue() !== "" && entryAmount.getValue() !== null) {
          if (entryRecurring.getValue() !== "" && entryRecurring.getValue() !== null) {
            sheet.getRange(entryName.getA1Notation()).setValue(entryName.getValue())
            sheet.getRange(entryAmount.getA1Notation()).setValue(entryAmount.getValue())
            sheet.getRange(entryRecurring.getA1Notation()).setValue(entryRecurring.getValue())
            sheet.getRange(entryMonth.getA1Notation()).setValue(entryMonth.getValue())
          }
        } 
      }
    })
  }
}

const income = {
  cornerStone: "F1",
  headers: ["Beschreibung", "Betrag", "Wiederkehrend?", "Start Monat"],
  typesList: { 
    range: ['Ausstehend', 'Monatlich', 'Zweimonatlich', 'Vierteljährlich', 'Halbjährlich', 'Jährlich'],
    start: "H3",
    end: "H6"
  },
  monthsList: { 
    range: ['Januar', 'Februar', 'März', 'April', 'Mai', 'Juni', 'Juli', 'August', 'September', 'Oktober', 'November', 'Dezember'],
    start: "I3",
    end: "I6"
  },
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
    const typeRange = sheet.getRange(`${this.typesList.start}:${this.typesList.end}`);
    const monthRange = sheet.getRange(`${this.monthsList.start}:${this.monthsList.end}`);
    
    typeRange.setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(this.typesList.range)
        .setAllowInvalid(false)
        .build()
    );
    
    monthRange.setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(this.monthsList.range)
        .setAllowInvalid(false)
        .build()
    );
  },
  parse: function(sheet, previousSheet){
    const titleCell = sheet.getRange(this.cornerStone);
    previousSheet.getRange(`${this.monthsList.start}:${this.monthsList.end}`).getValues().forEach(
    (_a, index) => {
      const entryName = previousSheet.getRange(this.monthsList.start).offset(index, -3)
      const entryAmount = previousSheet.getRange(this.monthsList.start).offset(index, -2)
      const entryRecurring = previousSheet.getRange(this.monthsList.start).offset(index, -1)
      const entryMonth = previousSheet.getRange(this.monthsList.start).offset(index, 0)
      if (entryName.getValue() !== "" && entryName.getValue() !== null) {
        if (entryAmount.getValue() !== "" && entryAmount.getValue() !== null) {
          if (entryRecurring.getValue() !== "" && entryRecurring.getValue() !== null) {
            sheet.getRange(entryName.getA1Notation()).setValue(entryName.getValue())
            sheet.getRange(entryAmount.getA1Notation()).setValue(entryAmount.getValue())
            sheet.getRange(entryRecurring.getA1Notation()).setValue(entryRecurring.getValue())
            sheet.getRange(entryMonth.getA1Notation()).setValue(entryMonth.getValue())
          }
        } 
      }
    })
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

function createSheetName(){
  const dateObj = new Date();
  return `${dateObj.getUTCFullYear()}/${dateObj.getUTCMonth() + 1}/${dateObj.getUTCMinutes()}/${dateObj.getUTCSeconds()}`;
}

function newMonthNewTab(){

  const previousSheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];

  const currentSheet = SpreadsheetApp.getActiveSpreadsheet();
  currentSheet.insertSheet(createSheetName());
  currentSheet.moveActiveSheet(0);

  creditLeft.display(currentSheet)
  spendings.display(currentSheet)
  spendings.parse(currentSheet, previousSheet)
  income.display(currentSheet)
  income.parse(currentSheet, previousSheet)
  motivation.display(currentSheet)
  
}
