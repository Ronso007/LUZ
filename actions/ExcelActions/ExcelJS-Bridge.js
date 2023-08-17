// Imports
var Excel = require("exceljs");
const DEFAULT_COLOR_FONT = "000000";
const DEFAULT_COLOR_FG = "FFFFFF";

exports.getCellText = function (cell) {
  return cell.text;
};

exports.getCellFontBold = function (cell) {
  if ("font" in cell.style) {
    if ("bold" in cell.style.font) {
      return cell.style.font.bold;
    }
  }

  return false;
};

exports.getCellFontSize = function (cell) {
  if ("font" in cell.style) {
    if ("size" in cell.style.font) {
      return cell.style.font.size + "px";
    }
  }

  return 12 + "px";
};

exports.getCellFontColor = function (cell) {
  // Const Definition

  // Excel Is Pretty wierd, everyone can have his own theme so if we use any theme color,
  // It won't come as color but as the theme number.
  // We want first try to get the color if it isn't default.

  let cellFontColor;

  if ("font" in cell.style && "color" in cell.style.font) {
    if ("argb" in cell.style.font.color) {
      cellFontColor = convertARGBtoRGB(cell.style.font.color.argb);
    } else {
      let cFontColorTheme = cell.style.font.color.theme;
      let cFontColorTint = 0;

      // If there is tint - tint is a kind of alpha value.
      if ("tint" in cell.style.font.color) {
        cFontColorTint = cell.style.font.color.tint;
      }

      cellFontColor = convertThemeToColorHex(cFontColorTheme, cFontColorTint);
    }
  } else {
    cellFontColor = DEFAULT_COLOR_FONT;
  }

  cellFontColor = "#" + cellFontColor;

  return cellFontColor;
};

exports.getCellFGColor = function (cell) {
  // Const Definition

  let cellFGColor;

  // Get foreground Color
  if ("fill" in cell.style) {
    if ("argb" in cell.style.fill.fgColor) {
      cellFGColor = convertARGBtoRGB(cell.style.fill.fgColor.argb);
    } else if ("theme" in cell.style.fill.fgColor) {
      let cFillColorTheme = cell.style.fill.fgColor.theme;
      let cFillColorTint = 0;

      // If there is tint - tint is a kind of alpha value.
      if ("tint" in cell.style.fill.fgColor) {
        cFillColorTint = cell.style.fill.fgColor.tint;
      }

      cellFGColor = convertThemeToColorHex(cFillColorTheme, cFillColorTint);
    }
  } else {
    cellFGColor = DEFAULT_COLOR_FG;
  }

  cellFGColor = "#" + cellFGColor;

  return cellFGColor;
};

exports.getCellType = function (workbook, rowNumber, colNumber) {
  // Const Definition
  const WORKBOOK_BASE = "Base";

  let workSheetBase = workbook.getWorksheet(WORKBOOK_BASE);

  return workSheetBase.getRow(rowNumber).getCell(colNumber).text;
};

function convertARGBtoRGB(argb) {
  return argb.substring(2);
}

// Convert theme and tint to HEX color RGB
function convertThemeToColorHex(theme, tint) {
  // Get the Theme Hex
  let colorHex = getHexBaseByTheme(theme);

  // If the tint is 0, no need to convert.
  if (tint != 0) {
    // Convert To RGB
    let objRGB = hexToRGB(colorHex);

    // Get the calculate tint RGB
    objRGB = calculateTint(objRGB, tint);

    // Convert back to HEX
    colorHex = rgbToHex(objRGB);
  }

  return colorHex;
}

function calculateTint(objRGB, tint) {
  // Check if need to Shade Or Tint And calculate for R, G and B
  if (tint >= 0) {
    Object.keys(objRGB).forEach((key) => {
      objRGB[key] = Math.round(objRGB[key] + (255 - objRGB[key]) * tint);
    });
  } else {
    Object.keys(objRGB).forEach((key) => {
      objRGB[key] = Math.round(objRGB[key] * (1 + tint));
    });
  }

  return objRGB;
}

function hexToRGB(colorHex) {
  let bigint = parseInt(colorHex, 16);
  let r = (bigint >> 16) & 255;
  let g = (bigint >> 8) & 255;
  let b = bigint & 255;

  let objRGB = {
    r: r,
    g: g,
    b: b,
  };

  return objRGB;
}

function componentToHex(c) {
  let hex = c.toString(16);
  return hex.length == 1 ? "0" + hex : hex;
}

function rgbToHex(objRGB) {
  return "" + componentToHex(objRGB.r) + componentToHex(objRGB.r) + componentToHex(objRGB.b);
}

function getHexBaseByTheme(theme) {
  const colorTheme = [
    "FFFFFF",
    "262626",
    "E6E6E6",
    "445468",
    "4472C6",
    "ED7D31",
    "A5A5A5",
    "FFC000",
    "5B9BD5",
    "70AD47",
  ];

  return colorTheme[theme];
}
