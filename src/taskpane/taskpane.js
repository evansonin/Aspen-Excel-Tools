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

let fullPath;
let openFileName;
// For making the error dialog work both in bebugging and in the published version
let baseUrl;

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // File name and error dialog definitions
    fullPath = Office.context.document.url;
    openFileName = fullPath.substring(fullPath.lastIndexOf('/')+1);
    // For making the error dialog work both in production and during debugging
    baseUrl = window.location.href.substring(0, window.location.href.lastIndexOf('/') + 1);
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

    const SETTINGS_KEY = 'excelSettings';
    const defaultSettings = {
      checkFileName: true
    };
    const currentSettings = loadSettings();
    applySettingsToUI(currentSettings);
    let saveSettingsButton = document.getElementById('save-button')

    function loadSettings() {
      const storedSettingsString = getFromLocalStorage(SETTINGS_KEY);

      // Handle the case where nothing is stored or the API returns "null" as a string
      if (!storedSettingsString || storedSettingsString === "null") {
          return defaultSettings;
      }

      try {
          const loadedSettings = JSON.parse(storedSettingsString);
          
          // Merge with defaults. This is crucial for future-proofing.
          // It ensures that if we add a new property to `defaultSettings`,
          // users with old saved settings will still get that new property.
          // The properties from `loadedSettings` will overwrite the `defaultSettings`.
          const finalSettings = { ...defaultSettings, ...loadedSettings };
          
          return finalSettings;

      } catch (error) {
          console.error("Error parsing stored settings. Falling back to defaults.", error);
          // If the stored JSON is corrupted, fall back to the default settings.
          return defaultSettings;
      }
    }
    function saveSettings(settingsObject) {
      try {
          const settingsString = JSON.stringify(settingsObject);
          setInLocalStorage(SETTINGS_KEY, settingsString);
      } catch (error) {
          console.error("Could not save settings.", error);
      }
    }
    function getSettingsFromUI() {
      const checkFilenameEl = document.getElementById('check-filename-checkbox');

      // Add other UI elements here in the future
      // const authorNameEl = document.getElementById('author-name-input');
      // const maxRowCountEl = document.getElementById('max-row-count-input');

      return {
          checkFileName: checkFilenameEl.checked,
          // authorName: authorNameEl.value,
          // maxRowCount: parseInt(maxRowCountEl.value, 10) || 0 // Parse integers
      };
    }
    function applySettingsToUI(settings) {
      const checkFilenameEl = document.getElementById('check-filename-checkbox');
      if (checkFilenameEl) {
          checkFilenameEl.checked = settings.checkFileName;
      }

      // Add other UI elements here in the future
      // const authorNameEl = document.getElementById('author-name-input');
      // if (authorNameEl) {
      //     authorNameEl.value = settings.authorName;
      // }
    }
    saveSettingsButton.addEventListener("click", () => {
      const settingsToSave = getSettingsFromUI();
      saveSettings(settingsToSave);
    })


    // CSV import logic

    const importBokCsvButton = document.getElementById("import-bok-csv-button");
    const importAnbCsvButton = document.getElementById("import-anb-csv-button");
    const bokFileInput = document.getElementById("bok-csv-file-input");
    const anbFileInput = document.getElementById("anb-csv-file-input");

    importBokCsvButton.addEventListener("click", async () => { // <-- Make the function async
      const file = bokFileInput.files[0];
      if (file) {
        const reader = new FileReader();
        reader.onload = async function (event) { // <-- Make this function async too
          let parsedData;
          const csvData = event.target.result;

          parsedData = Papa.parse(csvData, { skipEmptyLines: true, skipFirstNLines: 1 }).data;

          if (parsedData && parsedData.length > 0) {
            if (correctWorkbookOpen("bok")) {
              writeToSpreadsheet(parsedData, "bok");
            } else {
              let proceed = false;
              // Await the result of the dialog
              if (document.getElementById('check-filename-checkbox').checked) {
                proceed = await showErrorDialog("generic", "\"BOK\" not found in the currently open workbook's name. Proceed anyway?", null, "yes-no");
              } else {
                proceed = true;
              }
              
              if (proceed) {
                writeToSpreadsheet(parsedData, "bok");
              }
            }
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
    // MODIFIED importAnbCsvButton event listener
    importAnbCsvButton.addEventListener("click", async () => { // <-- Make the function async
      const file = anbFileInput.files[0];
      if (file) {
        const reader = new FileReader();
        reader.onload = async function (event) { // <-- Make this function async too
          let parsedData;
          const csvData = event.target.result;

          parsedData = Papa.parse(csvData, { skipEmptyLines: true, skipFirstNLines: 4 }).data;

          if (parsedData && parsedData.length > 0) {
            parsedData.sort((a, b) => new Date(a[1]) - new Date(b[1]));
            if (correctWorkbookOpen("anb")) {
              writeToSpreadsheet(parsedData, "anb");
            } else {
              let proceed = false;
              // Await the result of the dialog
              if (document.getElementById('check-filename-checkbox').checked) {
                proceed = await showErrorDialog("generic", "\"ANB\" not found in the currently opened workbook's name. Proceed anyway?", null, "yes-no");
              } else {
                proceed = true;
              }
              if (proceed) {
                writeToSpreadsheet(parsedData, "anb");
              }
            }
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
        break;
      case "default":
        throw new Error();
    }
    monthNumber = (dataMonth.getMonth() + 1).toString().padStart(2, "0");
    correctSheet = formatter.format(dataMonth).toString();
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

// MODIFIED showErrorDialog function
function showErrorDialog(errorType = null, errorText = "Generic error", sheetName = null, button = "ok") {
  // Return a new Promise
  return new Promise((resolve) => {
    const correctSheet = sheetName || "the specified month"; // Fallback text
    const urlData = `?type=${encodeURIComponent(errorType)}&text=${encodeURIComponent(errorText)}&month=${encodeURIComponent(correctSheet)}&button=${button}`;
    const dialogUrl = `${baseUrl}errorDialog.html${urlData}`;
    let dialog;

    Office.context.ui.displayDialogAsync(
      dialogUrl,
      { height: 25, width: 25, displayInIframe: true },
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.error(`Error opening dialog: ${asyncResult.error.message}`);
          resolve(false); // Resolve with false if the dialog fails to open
          return;
        }

        dialog = asyncResult.value;

        // Event handler for messages from the dialog.
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
          dialog.close(); // Close the dialog on any message
          if (arg.message === "yes") {
            resolve(true); // Resolve the promise with true for "yes"
          } else {
            // This handles "close-dialog" or any other message as a "no"
            resolve(false); // Resolve the promise with false
          }
        });

        // Event handler for when the dialog is closed by the user (e.g., clicking the 'X')
        dialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
            console.log("Dialog event received: " + arg.error);
            resolve(false); // Also resolve with false if the dialog is just closed
        });
      }
    );
  });
}

function correctWorkbookOpen(bank) {
  if (openFileName.toLowerCase().includes(bank)) {
    return true;
  } else {
    return false;
  }

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

function setInLocalStorage(key, value) {
  const myPartitionKey = Office.context.partitionKey;

  // Check if local storage is partitioned. 
  // If so, use the partition to ensure the data is only accessible by your add-in.
  if (myPartitionKey) {
    localStorage.setItem(myPartitionKey + key, value);
  } else {
    localStorage.setItem(key, value);
  }
}

function getFromLocalStorage(key) {
  const myPartitionKey = Office.context.partitionKey;

  // Check if local storage is partitioned.
  if (myPartitionKey) {
    return localStorage.getItem(myPartitionKey + key);
  } else {
    return localStorage.getItem(key);
  }
}