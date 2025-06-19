// src/taskpane/taskpane.js

/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */
import Papa from "papaparse";
let dialog = null;

// Date formatter used to generate sheet names in the format "Month Year" (e.g., "June 2025") based on dates from the CSV.
const formatter = new Intl.DateTimeFormat("en-US", {
  year: "numeric",
  month: "long",
});

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // Bank login logic

    const openBokSiteButton = document.getElementById("openBokSiteButton");
    openBokSiteButton.addEventListener("click", () => {
      const url = "https://www.bokfinancial.com/business";
      Office.context.ui.displayDialogAsync(
        url,
        { height: 90, width: 80, displayInIframe: true },
        (asyncResult) => {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            console.error("Dialog could not be opened: " + asyncResult.error.message);
          } else {
            dialog = asyncResult.value;
            dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
              console.log("Message received from dialog: " + arg.message);
              dialog.close();
            });
            dialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
              console.log("Dialog event received: " + arg.error);
            });
          }
        }
      );
    });
    const openAnbSiteButton = document.getElementById("openAnbSiteButton");
    openAnbSiteButton.addEventListener("click", () => {
      const url = "https://www.anbbank.com";
      Office.context.ui.displayDialogAsync(
        url,
        { height: 90, width: 80, displayInIframe: true },
        (asyncResult) => {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            console.error("Dialog could not be opened: " + asyncResult.error.message);
          } else {
            dialog = asyncResult.value;
            dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
              console.log("Message received from dialog: " + arg.message);
              dialog.close();
            });
            dialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
              console.log("Dialog event received: " + arg.error);
            });
          }
        }
      );
    });

    // CSV import logic

    const importBokCsvButton = document.getElementById("import-bok-csv-button");
    const importAnbCsvButton = document.getElementById("import-anb-csv-button");
    const bokFileInput = document.getElementById("bok-csv-file-input");
    const anbFileInput = document.getElementById("anb-csv-file-input");

    importBokCsvButton.addEventListener("click", () => {
      const file = bokFileInput.files[0];
      if (file) {
        const reader = new FileReader();
        reader.onload = function (event) {
          let parsedData;
          const csvData = event.target.result;

          // Use Papa Parse to convert CSV to a 2D array
          parsedData = Papa.parse(csvData, { skipEmptyLines: true, skipFirstNLines: 1 }).data;

          if (parsedData && parsedData.length > 0) {
            // Call our function to write the data to Excel
            writeToSpreadsheet(parsedData, "bok");
          } else {
            console.error("Could not read data from CSV.");
            showErrorDialog("generic", "Could not parse the CSV. Ensure you are using the CSV from ANB for this tool.");
          }
        };

        reader.onerror = function (event) {
          console.error("File could not be read! " + event.target.error);
        };

        reader.readAsText(file);
      } else {
        showErrorDialog("generic", "Please select a file.");
      }
    });
    importAnbCsvButton.addEventListener("click", () => {
      const file = anbFileInput.files[0];
      if (file) {
        const reader = new FileReader();
        reader.onload = function (event) {
          let parsedData;
          const csvData = event.target.result;

          // Use Papa Parse to convert CSV to a 2D array
          parsedData = Papa.parse(csvData, { skipEmptyLines: true, skipFirstNLines: 4 }).data;

          if (parsedData && parsedData.length > 0) {
            // Call our function to write the data to Excel
            parsedData.sort((a, b) => new Date(a[1]) - new Date(b[1])); // Sort array of arrays by date, since the export doesn't seem to put them in order
            console.log(parsedData);
            writeToSpreadsheet(parsedData, "anb");
          } else {
            console.error("Could not read data from CSV.");
            showErrorDialog("generic", "Could not parse the CSV. Ensure you are using the CSV from ANB for this tool.");
          }
        };

        reader.onerror = function (event) {
          console.error("File could not be read! " + event.target.error);
        };

        reader.readAsText(file);
      } else {
        showErrorDialog("generic", "Please select a file.");
      }
    });
  }
});

async function writeToSpreadsheet(data, bank = "bok") {
  let correctSheet;
  let monthNumber = 0;
  const bokClearedCheckbox = document.getElementById("bok-markCleared");
  const anbClearedCheckbox = document.getElementById("anb-markCleared");
  let firstEmptyRowIndex;
  let dataMonth;

  // Date-retrieving logic

  try {
    switch (bank) {
      case "bok":
        dataMonth = new Date(data[0][3]); // Retrieves data date from first entry in CSV
        break;
      case "anb":
        dataMonth = new Date(data[0][1]);
        console.log(dataMonth);
        break;
      case "default":
        throw new Error();
    }
    monthNumber = (dataMonth.getMonth() + 1).toString().padStart(2, "0");
    console.log(monthNumber);
    correctSheet = formatter.format(dataMonth).toString();
    console.log(`fjweq ${correctSheet}`);
  } catch (error) {
    showErrorDialog(
      "generic",
      'Error extracting data from the given CSV. Right click the task pane, click "Inspect", and go to the console for more information. Ensure you are using the respective CSV for the tool used.'
    );
    if (error.message == "Invalid time value") {
      console.error(
        "Error finding a date in the CSV provided. This program currently searches for a date in cell D2 for BOK and B5 for ANB."
      );
    } else {
      console.error(`Error opening the given CSV: ${error.message}`);
    }
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

        let match = sheet.name.trim() == correctSheet; // Boolean stating if the current sheet has the same name as the date retrieved from the CSV
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
              workbook.worksheets.getItem(correctedSheetNamesObj[correctSheet]).activate(); // If the correct worksheet (albeit with leading/following spaces) exists, swtich to it
            } else {
              // If all else fails
              showErrorDialog("missingSheet", undefined, correctSheet);
              console.log(correctSheet);
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
        const firstEmptyRowIndex = await getFirstEmptyRow(7);

        const firstEmptyRowInCheckCol_D = await getFirstEmptyRow(7, "D");
        if (firstEmptyRowInCheckCol_D === null) {
          showErrorDialog(
            "generic",
            "Could not determine the last check number. Column D might be empty or structured unexpectedly."
          );
          throw new Error("Could not find first empty row in Check # column (D).");
        }

        let numericLastCheckNumber = 0; // Default if no previous check numbers or if cell is not a number

        // If the first empty row in Col D is > 7, it means there's data in or above row 7.
        // The actual last check number is in the row *above* this first empty row.
        if (firstEmptyRowInCheckCol_D > 7) {
          const cellWithLastCheck = sheet.getCell(firstEmptyRowInCheckCol_D - 1, 3);
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
              `Last check number in cell E${firstEmptyRowInCheckCol_D - 1} ("${cellWithLastCheck.values[0][0]}") was not a parsable number. Starting new checks from 1.`
            );
            // numericLastCheckNumber remains 0, so first new check will be 1
          }
          // If cell was empty, numericLastCheckNumber remains 0, so new checks start from 1.
        }
        // Loop to iterate through each item on the daily list
        /*      Column 0: Date
        Column 1: Description
        Column 2: Cleared
        Column 3: Check # 
        Column 4: In 
        Column 5: Out 
        Column 6 (Preset formula): Balance */
        for (let i = 0; i < data.length; i++) {
          if (bank == "bok"){
            // Date
            sheet.getCell(firstEmptyRowIndex + i, 0).values = data[i][3];
            if (data[i][5] > 0) {
              // If positive
              sheet.getCell(firstEmptyRowIndex + i, 4).values = data[i][5];
              if (bokClearedCheckbox.checked) {
                // Only cleared transactions are colored
                sheet.getRangeByIndexes(firstEmptyRowIndex + i, 0, 1, 6).format.font.color =
                  "#008000"; // Green
              }
            } else {
              // If negative, out, absolute value (outflows are not kept as negative amounts in spreadsheet)
              sheet.getCell(firstEmptyRowIndex + i, 5).values = Math.abs(data[i][5]);
              if (bokClearedCheckbox.checked) {
                // Only cleared transactions are colored
                sheet.getRangeByIndexes(firstEmptyRowIndex + i, 0, 1, 6).format.font.color =
                  "#FF0000"; // Red
              }
            }
            if (data[i][4].endsWith("DEBIT")) {
              sheet.getCell(firstEmptyRowIndex + i, 1).values = `ACH Debit: ${data[i][12]}`; // e.g. ACH Debit: EWALLET
            } else if (data[i][4].endsWith("CREDIT")) {
              sheet.getCell(firstEmptyRowIndex + i, 1).values = `ACH Deposit: ${data[i][12]}`; // e.g. ACH Deposit: NBOA
            } else {
              sheet.getCell(firstEmptyRowIndex + i, 1).values = `${data[i][4]}: ${data[i][12]}`;
            }
            if (bokClearedCheckbox.checked) {
              sheet.getCell(firstEmptyRowIndex + i, 2).values = `X${monthNumber}`; // e.g. X06 for June
            } else {
              // For transactions that are not marked as cleared
              let checkNumberCell = sheet.getCell(firstEmptyRowIndex + i, 3);
              checkNumberCell.values = numericLastCheckNumber + (i + 1);
              checkNumberCell.format.horizontalAlignment = Excel.HorizontalAlignment.center;
              // This should be redundant, but just in case the text is formatted differently for some reason
              sheet.getRangeByIndexes(firstEmptyRowIndex + i, 0, 1, 6).format.font.color = "#000000";
            }
          } else if (bank == "anb") {
            sheet.getCell(firstEmptyRowIndex + i, 0).values = data[i][1];
            console.log(data);
            if (data[i][2].startsWith("Square Inc")) {
              if (data[i][5]) {
                sheet.getCell(firstEmptyRowIndex + i, 1).values = "ACH Deposit: Square, BSE Sales";
                sheet.getCell(firstEmptyRowIndex + i, 4).values = data[i][5];
              } else if (data[i][4]) {
                sheet.getCell(firstEmptyRowIndex + i, 1).values = data[i][2];
                sheet.getCell(firstEmptyRowIndex + i, 5).values = data[i][4];
              }
            } else {
              sheet.getCell(firstEmptyRowIndex + i, 1).values = data[i][2];
              if (data[i][4]){
                sheet.getCell(firstEmptyRowIndex + i, 5).values = data[i][4];
              }
              else if (data[i][5]) {
                sheet.getCell(firstEmptyRowIndex + i, 4).values = data[i][5];
              }
            }
            if (anbClearedCheckbox.checked) {
              sheet.getCell(firstEmptyRowIndex + i, 2).values = `X${monthNumber}`; // e.g. X06 for June
              if (data[i][5]) {
                sheet.getRangeByIndexes(firstEmptyRowIndex + i, 0, 1, 6).format.font.color = "#008000";
              } else if (data[i][4]) {
                sheet.getRangeByIndexes(firstEmptyRowIndex + i, 0, 1, 6).format.font.color = "#FF0000";
              }
            } else {
              // For transactions that are not marked as cleared
              let checkNumberCell = sheet.getCell(firstEmptyRowIndex + i, 3);
              checkNumberCell.values = numericLastCheckNumber + (i + 1);
              checkNumberCell.format.horizontalAlignment = Excel.HorizontalAlignment.center;
              // This should be redundant, but just in case the text is formatted differently for some reason
              sheet.getRangeByIndexes(firstEmptyRowIndex + i, 0, 1, 6).format.font.color = "#000000";
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



async function getFirstEmptyRow(ignoreUpToRow = 0, columnLetter = "A") {
  const rowIndex = await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
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

function showErrorDialog(errorType = null, errorText = "Generic error", sheetName = null) {
  const correctSheet = sheetName || "the specified month"; // Fallback text
  const urlData = `?type=${encodeURIComponent(errorType)}&text=${encodeURIComponent(errorText)}&month=${encodeURIComponent(correctSheet)}`;
  const dialogUrl = `${window.location.origin}/errorDialog.html${urlData}`;
  // A variable to hold the dialog object
  let dialog;

  Office.context.ui.displayDialogAsync(
    dialogUrl,
    { height: 25, width: 25, displayInIframe: true },
    (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        // Handle error opening dialog
        console.error(`Error opening dialog: ${asyncResult.error.message}`);
        return;
      }

      // If successful, we get a dialog object.
      dialog = asyncResult.value;

      // Add an event handler to listen for messages from the dialog.
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
        // Check if the message is the one we expect
        if (arg.message === "close-dialog") {
          // Close the dialog
          dialog.close();
        }
      });
    }
  );
}

export async function run() {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.format.fill.color = "yellow";
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}
