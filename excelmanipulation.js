const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');

class ExcelReader {
  constructor(filename) {
    this.filename = filename;
    this.workbook = xlsx.readFile(filename);
  }

  queryRow(rowNumber) {
    const sheet = this.workbook.Sheets[this.workbook.SheetNames[0]];
    const range = xlsx.utils.decode_range(sheet['!ref']);
    if (rowNumber < range.s.r || rowNumber > range.e.r) {
      return null;
    }
    const row = rowNumber;
    // get the column names from the first row of the sheet
    const header = xlsx.utils.sheet_to_json(sheet, { range: 0 })[0];
    // convert the header object to an array of keys
    const columns = Object.keys(header);
    // use the columns array as the header option
    const rowData = xlsx.utils.sheet_to_json(sheet, { range: row, header: columns });
    return rowData.length > 0 ? rowData[0] : null;
  }

  emptyRowJson() {
    const sheet = this.workbook.Sheets[this.workbook.SheetNames[0]];
    // get the column names from the first row of the sheet
    const header = xlsx.utils.sheet_to_json(sheet, { range: 0 })[0];
    // create an empty row JSON
    let emptyRow = {};
    for (let key in header) {
      emptyRow[key] = null;
    }
    return emptyRow;
  }

  close() {
    // No specific action required to close the Excel reader
  }
}

class ExcelWriter {
  constructor(filename) {
    this.filename = filename;
    this.workbook = xlsx.utils.book_new();
    this.sheetName = 'Sheet1';
    xlsx.utils.book_append_sheet(this.workbook, xlsx.utils.aoa_to_sheet([]), this.sheetName);
  }

  append(data) {
    const sheet = this.workbook.Sheets[this.sheetName];
    const range = xlsx.utils.decode_range(sheet['!ref']);
    const rowIndex = range.e.r + 1;
    data.forEach((value, colIndex) => {
      const cellAddress = xlsx.utils.encode_cell({ r: rowIndex, c: colIndex });
      sheet[cellAddress] = { t: 's', v: value };
    });
  }

  save() {
    xlsx.writeFile(this.workbook, this.filename);
  }

  close() {
    // No specific action required to close the Excel writer
  }
}

class ExcelController {
  constructor(filename) {
    this.filename = filename;
    this.excelReader = new ExcelReader(filename);
    this.excelWriter = new ExcelWriter(filename);
  }

  readRow(rowNumber) {
    return new Promise((resolve, reject) => {
      const rowData = this.excelReader.queryRow(rowNumber);
      if (rowData !== null) {
        resolve(rowData);
      } else {
        reject(new Error(`Row ${rowNumber} does not exist in the file.`));
      }
    });
  }

  appendRow(data) {
    this.excelWriter.append(data);
    this.excelWriter.save();
  }

  readEmptyRow() {
    return this.excelReader.emptyRowJson();
  }

  close() {
    this.excelReader.close();
    this.excelWriter.close();
  }
}

module.exports = ExcelController;
