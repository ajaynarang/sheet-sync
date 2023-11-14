const fs = require("fs");
const xlsx = require("xlsx");
const chalk = require("chalk");

function syncSheet(
  baseSheetPath,
  sourceSheetPath,
  columnNames,
  overrideFlag,
  newSheetName
) {
  try {
    // Load base sheet
    const baseWorkbook = xlsx.readFile(baseSheetPath, {
      cellStyles: true,
    });
    // Load source sheet
    const sourceWorkbook = xlsx.readFile(sourceSheetPath);

    // const newWorkbook = xlsx.utils.book_new();
    const newWorkbook = { ...baseWorkbook };

    for (
      let sheetIndex = 0;
      sheetIndex < baseWorkbook.SheetNames.length;
      sheetIndex++
    ) {
      const baseSheet =
        baseWorkbook.Sheets[baseWorkbook.SheetNames[sheetIndex]];
      const sourceSheet =
        sourceWorkbook.Sheets[sourceWorkbook.SheetNames[sheetIndex]];

      // Clone base sheet to create new sheet
      const newSheet = newWorkbook.Sheets[newWorkbook.SheetNames[sheetIndex]];

      console.log("sourceSheet", sourceSheet);

      if (sourceSheet && newSheet) {
        //for each provided columns
        columnNames.forEach((column) => {
          for (
            let row = 1;
            row <= xlsx.utils.decode_range(newSheet["!ref"]).e.r + 1;
            row++
          ) {
            const cellRef = column + row;
            if (sourceSheet.hasOwnProperty(cellRef)) {
              const cellValue = overrideFlag
                ? sourceSheet[cellRef].v
                : newSheet[cellRef].v + sourceSheet[cellRef].v;

              newSheet[cellRef] = cellValue;
            }
          }
          //xlsx.utils.book_append_sheet(newWorkbook, newSheet);
        });
      }
    }

    console.log("newWorkbook", newWorkbook.Sheets["Sheet 1"]["A6"]);
    xlsx.writeFile(newWorkbook, `${newSheetName}`);

    console.log("New sheet generated successfully.");
  } catch (error) {
    console.error("Error:", error.message);
  }
}

const baseSheetPath = "sheets/base.xlsx";
const sourceSheetPath = "sheets/source.xlsx";
const columnNames = ["B", "C"];
const overrideFlag = true;
const newSheetName = "sheets/output.xlsx";

syncSheet(
  baseSheetPath,
  sourceSheetPath,
  columnNames,
  overrideFlag,
  newSheetName
);
