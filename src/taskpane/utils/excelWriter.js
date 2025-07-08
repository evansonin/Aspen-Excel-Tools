import { baseUrl } from "../taskpane.js";
// Date formatter used to generate sheet names in the format "Month Year" (e.g., "June 2025") based on dates from the CSV.
const formatter = new Intl.DateTimeFormat("en-US", {
  year: "numeric",
  month: "long",
});

/**
 * Writes parsed CSV data to the Excel spreadsheet.
 * @param {Array<Array<any>>} data The parsed CSV data.
 * @param {string} bank The type of bank ("bok" or "anb").
 * @param {function} showErrorDialog The function to display error dialogs.
 */
export async function writeToSpreadsheet(data, bank = "bok", showErrorDialog) {
  let correctSheet;
  let monthNumber = 0;
  const bokClearedCheckbox = document.getElementById("bok-markCleared");
  const anbClearedCheckbox = document.getElementById("anb-markCleared");
  let dataMonth;

  // How many rows to skip in the open spreadsheet when checking where the last entry is.
  // This only needs to be changed if accounting changes formatting.
  const ignoreUpToRowConstant = 7;

  // Define column indices for the Excel sheet
  const COL_DATE = 0;
  const COL_DESCRIPTION = 1;
  const COL_CLEARED = 2;
  const COL_CHECK_NUMBER = 3;
  const COL_IN = 4;
  const COL_OUT = 5;
  const COL_BALANCE = 6; 

  // Define CSV data indices for BOK and ANB
  const BOK_CSV = {
    DATE: 3,
    TYPE: 4,
    AMOUNT: 5,
    DESCRIPTION: 12,
  };

  const ANB_CSV = {
    DATE: 1,
    DESCRIPTION: 2,
    OUT_AMOUNT: 4,
    IN_AMOUNT: 5,
  };

  // Date-retrieving logic
  try {
    switch (bank) {
      case "bok":
        dataMonth = new Date(data[0][BOK_CSV.DATE]); // Retrieves data date from first entry in CSV
        break;
      case "anb":
        dataMonth = new Date(data[0][ANB_CSV.DATE]);
        break;
      default: // This should theoretically never happen
        throw new Error("Invalid bank type specified.");
    }
    monthNumber = (dataMonth.getMonth() + 1).toString().padStart(2, "0");
    correctSheet = formatter.format(dataMonth).toString();
  } catch (error) {
    showErrorDialog(
      "generic",
      'Error extracting data from the given CSV. Right click the task pane, click "Inspect", and go to the console for more information. Ensure you are using the respective CSV for the tool used.'
    );
    if (error.message === "Invalid time value") { // Use strict equality
      console.error(
        "Error finding a date in the CSV provided. This program currently searches for a date in cell D2 for BOK and B5 for ANB."
      );
    } else {
      console.error(`Error opening the given CSV: ${error.message}`);
    }
    return; // Exit function if date extraction fails
  }

  try {
    // This async function finds (and switches to) the correct sheet given the
    // information in the CSV, and then returns the index of the first empty
    // row (after row 7) in that sheet
    await Excel.run(async (context) => {
      const workbook = context.workbook;
      let sheet = workbook.worksheets.getActiveWorksheet();
      const sheets = workbook.worksheets;
      try {
        sheets.load("items/name");
        sheet.load("name");
        await context.sync(); // Request data from Excel

        // Sheet-locating logic
        let match = sheet.name.trim() === correctSheet; // Boolean stating if the current sheet has the same name as the date retrieved from the CSV
        if (!match) {
          // If the current sheet is not the one found in the CSV
          let sheetNames = []; // Get list of all sheets in the workbook
          sheets.items.forEach(function (currentSheetInLoop) {
            sheetNames.push(currentSheetInLoop.name);
          });
          let correctSheetExists = sheetNames.includes(correctSheet);
          if (correctSheetExists) {
            // Switch to correct sheet if it's found
            workbook.worksheets.getItem(correctSheet).activate();
          } else {
            // If it's not found, it may be because it has leading/following spaces in its name
            let correctedSheetNamesObj = {}; // An object to store corrected sheet names plus their original names
            sheetNames.forEach((element) => {
              correctedSheetNamesObj[element.trim()] = element;
            });
            correctSheetExists = Object.keys(correctedSheetNamesObj).includes(correctSheet);
            if (correctSheetExists) {
              workbook.worksheets.getItem(correctedSheetNamesObj[correctSheet]).activate(); 
              // If the correct worksheet (albeit with leading/following spaces) exists, switch to it
            } else {
              // If a matching sheet still cannot be found (this program will not create a new sheet itself)
              showErrorDialog("missingSheet", undefined, correctSheet, "ok", baseUrl);
              throw new Error("The desired sheet was not found");
            }
          }
        }
        await context.sync();
      } catch (error) {
        throw new Error(error.message);
      }

      // Data-copying logic
      try {
        sheet = workbook.worksheets.getActiveWorksheet();
        await context.sync();
        const firstEmptyRowIndex = await getFirstEmptyRow(ignoreUpToRowConstant);

        const firstEmptyRowInCheckCol_D = await getFirstEmptyRow(ignoreUpToRowConstant, "D");
        if (firstEmptyRowInCheckCol_D === null) {
          showErrorDialog(
            "generic",
            "Could not determine the last check number. Column D might be empty or structured unexpectedly."
          );
          throw new Error("Could not find first empty row in Check # column (D).");
        }

        let numericLastCheckNumber = 0; // Default if no previous check numbers or if cell is not a number

        // If the first empty row in Col D is > 7, it means there's data in or above row 7.
        if (firstEmptyRowInCheckCol_D > 7) {
          const cellWithLastCheck = sheet.getCell(firstEmptyRowInCheckCol_D - 1, COL_CHECK_NUMBER);
          cellWithLastCheck.load("values");
          await context.sync(); // Sync 3: Load the value of the last check number

          // .values is a 2D array, e.g., [[value]]. Also check if it's truly a number.
          if (
            cellWithLastCheck.values &&
            cellWithLastCheck.values[0][0] !== "" &&
            !isNaN(Number(cellWithLastCheck.values[0][0]))
          ) {
            numericLastCheckNumber = parseInt(cellWithLastCheck.values[0][0]);
          } else if (cellWithLastCheck.values && cellWithLastCheck.values[0][0] !== "") {
            console.warn(
              `Last check number in cell D${firstEmptyRowInCheckCol_D - 1} ("${cellWithLastCheck.values[0][0]}") was not a parsable number. Starting new checks from 1.`
            );
            // numericLastCheckNumber remains 0, so first new check will be 1
          }
          // If cell was empty, numericLastCheckNumber remains 0, so new checks start from 1.
        } else {
          // If the current month's check column is empty, try to get the last check number from the previous month.
          let newCheckNumber = await getPreviousMonthsCheckNumber(correctSheet);
          if(newCheckNumber) {
            numericLastCheckNumber = newCheckNumber;
          }
        }
        // Loop to iterate through each item on the daily list
        /*      
        Column 0: Date
        Column 1: Description
        Column 2: Cleared
        Column 3: Check # 
        Column 4: In 
        Column 5: Out 
        Column 6 (Preset formula): Balance */
        for (let i = 0; i < data.length; i++) {
          if (bank === "bok"){ // Use strict equality
            // Date
            sheet.getCell(firstEmptyRowIndex + i, COL_DATE).values = data[i][BOK_CSV.DATE];
            if (data[i][BOK_CSV.AMOUNT] > 0) {
              // If positive
              sheet.getCell(firstEmptyRowIndex + i, COL_IN).values = data[i][BOK_CSV.AMOUNT];
              if (bokClearedCheckbox.checked) {
                // Only cleared transactions are colored
                sheet.getRangeByIndexes(firstEmptyRowIndex + i, COL_DATE, 1, COL_BALANCE + 1).format.font.color = // Use COL_BALANCE + 1 for range width
                  "#008000"; // Green
              }
            } else {
              // If negative, out, absolute value (outflows are not kept as negative amounts in spreadsheet)
              sheet.getCell(firstEmptyRowIndex + i, COL_OUT).values = Math.abs(data[i][BOK_CSV.AMOUNT]);
              if (bokClearedCheckbox.checked) {
                // Only cleared transactions are colored
                sheet.getRangeByIndexes(firstEmptyRowIndex + i, COL_DATE, 1, COL_BALANCE + 1).format.font.color = // Use COL_BALANCE + 1 for range width
                  "#FF0000"; // Red
              }
            }
            if (data[i][BOK_CSV.TYPE].endsWith("DEBIT")) {
              sheet.getCell(firstEmptyRowIndex + i, COL_DESCRIPTION).values = `ACH Debit: ${data[i][BOK_CSV.DESCRIPTION]}`; // e.g. ACH Debit: EWALLET
            } else if (data[i][BOK_CSV.TYPE].endsWith("CREDIT")) {
              sheet.getCell(firstEmptyRowIndex + i, COL_DESCRIPTION).values = `ACH Deposit: ${data[i][BOK_CSV.DESCRIPTION]}`; // e.g. ACH Deposit: NBOA
            } else {
              sheet.getCell(firstEmptyRowIndex + i, COL_DESCRIPTION).values = `${data[i][BOK_CSV.TYPE]}: ${data[i][BOK_CSV.DESCRIPTION]}`;
            }
            if (bokClearedCheckbox.checked) {
              sheet.getCell(firstEmptyRowIndex + i, COL_CLEARED).values = `X${monthNumber}`; // e.g. X06 for June
            } else {
              // For transactions that are not marked as cleared
              let checkNumberCell = sheet.getCell(firstEmptyRowIndex + i, COL_CHECK_NUMBER);
              checkNumberCell.values = numericLastCheckNumber + (i + 1);
              checkNumberCell.format.horizontalAlignment = Excel.HorizontalAlignment.center;
              // This should be redundant, but just in case the text is formatted differently for some reason
              sheet.getRangeByIndexes(firstEmptyRowIndex + i, COL_DATE, 1, COL_BALANCE + 1).format.font.color = "#000000"; // Use COL_BALANCE + 1 for range width
            }
          } else if (bank === "anb") { // Use strict equality
            sheet.getCell(firstEmptyRowIndex + i, COL_DATE).values = data[i][ANB_CSV.DATE];
            if (data[i][ANB_CSV.DESCRIPTION].startsWith("Square Inc")) {
              if (data[i][ANB_CSV.IN_AMOUNT]) {
                sheet.getCell(firstEmptyRowIndex + i, COL_DESCRIPTION).values = "ACH Deposit: Square, BSE Sales";
                sheet.getCell(firstEmptyRowIndex + i, COL_IN).values = data[i][ANB_CSV.IN_AMOUNT];
              } else if (data[i][ANB_CSV.OUT_AMOUNT]) {
                sheet.getCell(firstEmptyRowIndex + i, COL_DESCRIPTION).values = data[i][ANB_CSV.DESCRIPTION];
                sheet.getCell(firstEmptyRowIndex + i, COL_OUT).values = data[i][ANB_CSV.OUT_AMOUNT];
              }
            } else {
              sheet.getCell(firstEmptyRowIndex + i, COL_DESCRIPTION).values = data[i][ANB_CSV.DESCRIPTION];
              if (data[i][ANB_CSV.OUT_AMOUNT]){
                sheet.getCell(firstEmptyRowIndex + i, COL_OUT).values = data[i][ANB_CSV.OUT_AMOUNT];
              }
              else if (data[i][ANB_CSV.IN_AMOUNT]) {
                sheet.getCell(firstEmptyRowIndex + i, COL_IN).values = data[i][ANB_CSV.IN_AMOUNT];
              }
            }
            if (anbClearedCheckbox.checked) {
              sheet.getCell(firstEmptyRowIndex + i, COL_CLEARED).values = `X${monthNumber}`; // e.g. X06 for June
              if (data[i][ANB_CSV.IN_AMOUNT]) {
                sheet.getRangeByIndexes(firstEmptyRowIndex + i, COL_DATE, 1, COL_BALANCE + 1).format.font.color = "#008000"; // Use COL_BALANCE + 1 for range width
              } else if (data[i][ANB_CSV.OUT_AMOUNT]) {
                sheet.getRangeByIndexes(firstEmptyRowIndex + i, COL_DATE, 1, COL_BALANCE + 1).format.font.color = "#FF0000"; // Use COL_BALANCE + 1 for range width
              }
            } else {
              // For transactions that are not marked as cleared
              let checkNumberCell = sheet.getCell(firstEmptyRowIndex + i, COL_CHECK_NUMBER);
              checkNumberCell.values = numericLastCheckNumber + (i + 1);
              checkNumberCell.format.horizontalAlignment = Excel.HorizontalAlignment.center;
              // This should be redundant, but just in case the text is formatted differently for some reason
              sheet.getRangeByIndexes(firstEmptyRowIndex + i, COL_DATE, 1, COL_BALANCE + 1).format.font.color = "#000000"; // Use COL_BALANCE + 1 for range width
            }
          }
        }  
        await context.sync();
      } catch (error) {
        throw new Error(error);
      }
    });
  } catch (error) {
    console.error(error);
  }
}

/**
 * Finds the first empty row in a specified column of an Excel sheet.
 * @param {number} ignoreUpToRow Rows before this index are ignored.
 * @param {string} [columnLetter="A"] The column letter to check (e.g., "A", "D").
 * @param {string} [sheetName=null] The name of the sheet to check. If null, uses the active sheet.
 * @returns {Promise<number|null>} The 0-indexed row number of the first empty row, or null if an error occurs.
 */
export async function getFirstEmptyRow(ignoreUpToRow = 0, columnLetter = "A", sheetName = null) {
  const rowIndex = await Excel.run(async (context) => {
    let sheet;
    if (!sheetName) {
      sheet = context.workbook.worksheets.getActiveWorksheet();
    } else {
      sheet = context.workbook.worksheets.getItem(sheetName);
    }
    
    const column = sheet.getRange(`${columnLetter}:${columnLetter}`); // e.g. A:A

    let targetRowIndex;
    // Retrieve first empty cell after row ignoreUpToRow
    try {
      const lastUsedCell = column.getUsedRange(true).getLastCell();
      lastUsedCell.load("rowIndex");

      await context.sync();

      const firstVacantRowIndex = lastUsedCell.rowIndex + 1;

      targetRowIndex = Math.max(ignoreUpToRow, firstVacantRowIndex);
    } catch (error) {
      console.error(`Column  ${columnLetter} is empty or another error occured: ${error.message}`);
      return null;
    }
    // 0 Indexed
    return targetRowIndex;
  });
  return rowIndex;
}

/**
 * Retrieves the last check number from the previous month's sheet.
 * @param {string} dateString A date string that can be parsed into a Date object.
 * @returns {Promise<number>} The last check number from the previous month, or 0 if not found/error.
 */
export async function getPreviousMonthsCheckNumber(dateString) {
  const date = new Date(dateString);
  date.setMonth(date.getMonth() - 1);

  const newMonth = date.toLocaleDateString("en-US", {
    month: "long",
    year: "numeric"
  });

  let lastMonthsCheckNumberRow;
  try {
    lastMonthsCheckNumberRow = await getFirstEmptyRow(ignoreUpToRowConstant, "D", newMonth);
  } catch (error) {
    console.error(`Failed to get first empty row for previous month (${newMonth}). Assuming check number is 0. Error: ${error}`);
    return 0; // Default to 0 if previous month's sheet or row cannot be found
  }

  if (lastMonthsCheckNumberRow === null || lastMonthsCheckNumberRow <= 7) {
    // If the previous month's sheet is empty or doesn't exist, start checks from 0.
    return 0;
  }

  try {
    const lastCheckNumber = await Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getItem(newMonth);
      const cellWithLastCheck = worksheet.getCell(lastMonthsCheckNumberRow - 1, 3); // Row index, Column index (D is 3)

      cellWithLastCheck.load("values");
      await context.sync();

      const cellValue = cellWithLastCheck.values[0][0];
      const parsedValue = parseInt(cellValue);

      if (!isNaN(parsedValue)) {
        return parsedValue;
      } else {
        console.warn(`Last check number in previous month's sheet (${newMonth}) cell D${lastMonthsCheckNumberRow} ("${cellValue}") was not a parsable number. Starting new checks from 0.`);
        return 0;
      }
    });
    return lastCheckNumber;
  } catch (error) {
    console.error(`Failed to get check number from previous month (${newMonth}): ${error}`);
    return 0; // Default to 0 if Excel.run fails
  }
}
