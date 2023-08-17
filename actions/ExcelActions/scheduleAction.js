// Imports
var Excel = require("exceljs");
var path = require("path");
const { google } = require("googleapis");
var fs = require("fs");

// var cron = require("node-cron");

// cron.schedule("* * * * *", async () => {
//   await downloadLuz();
//   console.log("Update Luz");
// });

setInterval(function () {
  var date = new Date();
  if (date.getSeconds() === 0) {
    console.log("Downloading Luz...");
    downloadLuz();
  }
}, 1000);

// Require route modules.
var excelJSBridge = require("./ExcelJS-Bridge");

// Setting a workbook
var workbook = new Excel.Workbook();
async function setWorkBook(excelName) {
  await workbook.xlsx.readFile(excelName);
}
async function downloadLuz() {
  const drive = google.drive({ version: "v3", auth: global.oauth2Client });
  const fileId = "1x6PjiaUTt8u5E3NoN6b8ogYcVpwyEnAS";
  const dest = fs.createWriteStream("LUZ2.xlsx");
  try {
    const file = await drive.files.get(
      {
        fileId: fileId,
        alt: "media",
      },
      { responseType: "arraybuffer" }
    );
    // file.data.on("end", () => console.log("onCompleted"));
    //file.data.pipe(dest);
    fs.writeFileSync("LUZ2.xlsx", Buffer.from(file.data));
    //return file.status;
  } catch (err) {
    // TODO(developer) - Handle error
    throw err;
  }
  console.log("Success text file");
}

exports.readBase = function () {};

exports.readWeek = async function (weekNum, Cycle) {
  let xlpath = path.join(__dirname, "/../../LUZ2.xlsx");

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
  let xlpath = path.join(__dirname, "/../../LUZ2.xlsx");

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
