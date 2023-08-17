// Imports
var Excel = require("exceljs");
var path = require("path");

// Require route modules.
var excelJSBridge = require("./ExcelJS-Bridge");

// Setting a workbook
var workbook = new Excel.Workbook();
async function setWorkBook(excelName) {
  await workbook.xlsx.readFile(excelName);
}

exports.readBase = function () {};

exports.readWeek = async function (weekNum, Cycle) {
  //await downloadLuz(Cycle);
  let xlpath = path.join(__dirname, "/../../test.xlsx");

  await setWorkBook(xlpath);

  // Const Definition
  const DATE_BASE_NAME = "DATE";

  // Get workSheet by the name of the weekNum.
  let worksheet = workbook.getWorksheet(weekNum);

  // Schedule (Array of Arrays)
  let schedule = [];

  let cText;
  let cFontBold;
  let cFontSize;
  let cFontColor;
  let cFGColor;
  let cType;

  // Going Over Each Row and Each Cell
  worksheet.eachRow(function (row, rowNumber) {
    // Creating An Row Array
    let rowArray = [];

    row.eachCell(function (currCell, colNumber) {
      // Get Text Field
      cText = excelJSBridge.getCellText(currCell);

      // Get Is Bold
      cFontBold = excelJSBridge.getCellFontBold(currCell);

      // Get Font Size
      cFontSize = excelJSBridge.getCellFontSize(currCell);

      // Get Font Color
      cFontColor = excelJSBridge.getCellFontColor(currCell);

      // Get foreground Color
      cFGColor = excelJSBridge.getCellFGColor(currCell);

      // Get Type
      cType = excelJSBridge.getCellType(workbook, rowNumber, colNumber);

      // Format Date
      if (cType === DATE_BASE_NAME) {
        let dateObj = new Date(cText);

        // is the date is valid
        if (dateObj instanceof Date && !isNaN(dateObj)) {
          cText = formatTime(dateObj);
        }
      }

      // Create The Cell
      let newCell = {
        text: cText,
        font: {
          bold: cFontBold,
          size: cFontSize,
          color: cFontColor,
        },
        fgColor: cFGColor,
        type: cType,
        rowSpanNumber: 1,
      };

      // Push the cell to the array
      rowArray.push(newCell);
    });

    schedule.push(rowArray);
  });

  return schedule;
};

exports.getNumOfWeeks = async function (Cycle) {
  let xlpath = path.join(__dirname, "/../../test.xlsx");

  await setWorkBook(xlpath);

  let numOfWeeks = 0;

  while (workbook.getWorksheet((numOfWeeks + 1).toString()) !== undefined) {
    numOfWeeks++;
  }

  return numOfWeeks;
};

function formatTime(date) {
  return date.getUTCDate() + "/" + (date.getUTCMonth() + 1) + "/" + date.getUTCFullYear();
}
