// src/taskpane/taskpane.js

/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */
import { subDays, format } from 'date-fns';
import Papa from "papaparse";
import { writeToSpreadsheet } from "./utils/excelWriter";
import { 
  SETTINGS_KEY, 
  defaultSettings, 
  loadSettings, 
  saveSettings, 
  getSettingsFromUI, 
  applySettingsToUI,
  resetToDefaultSettings, 
  setInLocalStorage, 
  getFromLocalStorage 
} from "./utils/settingsManager";
import { 
  getBillData,
  manualDivvyWrite,
  writeToDivvySpreadsheet
} from "./utils/divvyTool";
import { showErrorDialog } from "./utils/errorDialogManager";
let dialog = null;


let fullPath;
let openFileName;
export let baseUrl; // for pop-ups

/**
 * Initializes global variables related to the workbook and add-in base URL.
 */
function initializeGlobalVariables() {
  fullPath = Office.context.document.url;
  openFileName = fullPath.substring(fullPath.lastIndexOf('/') + 1);
  baseUrl = window.location.href.substring(0, window.location.href.lastIndexOf('/') + 1);
}

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    initializeGlobalVariables();
    setupBankLoginButtons();
    setupSettings();
    setupCsvImportButtons();
    billThingy();
    setupDateSelection();
  }
  console.log("Aspen Excel Tools: Successfully loaded.");
});

/**
 * Attach a click event listener to a button to fetch Bill.com transactions.
 */
function billThingy() {
  const billButton = document.getElementById("billSubmit");
  const betterBillButton = document.getElementById("billSubmitBetter");

  billButton.addEventListener("click", async () => {
    const startDate = document.getElementById("startDate").value;
    const endDate = document.getElementById("endDate").value;
    billButton.innerHTML = "Please wait..."
    // startDate/endDate example: "2025-06-01"
    if (new Date(startDate) > new Date(endDate)) {
      showErrorDialog("generic", "Start date cannot be later than end date.", null, "ok", baseUrl);
      billButton.innerHTML = "Submit";
      return;
    }
    if (!startDate || !endDate) {
      showErrorDialog("generic", "Please select start and ending dates.", null, "ok", baseUrl);
      billButton.innerHTML = "Submit";
      return;
    }
    

    try {
      let billData = await getBillData(startDate, endDate);

      const users = billData['users']['results'];
      const transactions = billData['transactions']['results'];
      const employees = billData['employees'];
      if(transactions.length == 0) {
        showErrorDialog("generic", "No transactions were found for the given time period.", null, "ok", baseUrl);
      }
      else {
        writeToDivvySpreadsheet(users, transactions, employees, startDate, endDate);
      }
    } catch (error) {
      console.error(error);
      showErrorDialog("generic", "Failed to retrieve information from Divvy.", null, "ok", baseUrl);
    }
    billButton.innerHTML = "Submit";
  });

}

/**
 * Sets up event listeners for the bank login buttons.
 */
function setupBankLoginButtons() {
  const openBokSiteButton = document.getElementById("openBokSiteButton");
  openBokSiteButton.addEventListener("click", () => {
    const url = "https://exchange.bokfinancial.com/cdp/login/";
    Office.context.ui.displayDialogAsync(
      url,
      { height: 90, width: 80, displayInIframe: true },
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.error("Dialog could not be opened: " + asyncResult.error.message);
        } else {
          dialog = asyncResult.value;
          dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
            dialog.close();
          });
          dialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
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
            dialog.close();
          });
          dialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
          });
        }
      }
    );
  });
  const openDivvySiteButton = document.getElementById("openDivvySiteButton");
  openDivvySiteButton.addEventListener("click", () => {
    const url = "https://login.us.bill.com/neo/login";
    Office.context.ui.displayDialogAsync(
      url,
      { height: 90, width: 80, displayInIframe: true },
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.error("Dialog could not be opened: " + asyncResult.error.message);
        } else {
          dialog = asyncResult.value;
          dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
            dialog.close();
          });
          dialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
          });
        }
      }
    );
  });
}

/**
 * Sets up the settings UI and save functionality.
 */
function setupSettings() {
  const currentSettings = loadSettings();
  applySettingsToUI(currentSettings);
  let saveSettingsButton = document.getElementById('save-button');
  saveSettingsButton.addEventListener("click", () => {
    const settingsToSave = getSettingsFromUI();
    saveSettings(settingsToSave);
  });
  let defaultSettingsButton = document.getElementById('resetToDefaults');
  defaultSettingsButton.addEventListener("click", () => {
    resetToDefaultSettings();
  });
}

/**
 * Handles the import process
 * @param {File} file The CSV file to import.
 * @param {string} bankType The type of bank ("bok" or "anb").
 * @param {number} skipLines The number of lines to skip at the beginning of the CSV.
 * @returns {Promise<Array | undefined>} A promise that resolves with the parsed data for 'divvy' bank type,
 *                                       or undefined for other bank types or on error.
 */
async function handleCsvImport(file, bankType, skipLines) {
  return new Promise((resolve, reject) => {
    if (!file) {
      showErrorDialog("generic", "Please select a file.", null, "ok", baseUrl);
      resolve(undefined); // Resolve with undefined if no file is selected
      return;
    }

    const reader = new FileReader();
    reader.onload = async function (event) {
      let parsedData;
      const csvData = event.target.result;

      parsedData = Papa.parse(csvData, { skipEmptyLines: true, skipFirstNLines: skipLines }).data;

      if (parsedData && parsedData.length > 0) {
        if (bankType === "divvy") {
          resolve(parsedData); // Resolve the promise with the parsed data for 'divvy'
          return; // Exit the onload function
        }

        // Logic for non-divvy banks (bok, anb)
        if (bankType === "anb") {
          parsedData.sort((a, b) => new Date(a[1]) - new Date(b[1]));
        }

        if (correctWorkbookOpen(bankType)) {
          writeToSpreadsheet(parsedData, bankType, showErrorDialog);
          resolve(undefined); // Resolve with undefined as no data needs to be returned to caller for these types
        } else {
          let proceed = false;
          if (document.getElementById('check-filename-checkbox').checked) {
            proceed = await showErrorDialog("generic", `"${bankType.toUpperCase()}" not found in the currently open workbook's name. Proceed anyway?`, null, "yes-no", baseUrl);
          } else {
            proceed = true;
          }

          if (proceed) {
            writeToSpreadsheet(parsedData, bankType, showErrorDialog);
            resolve(undefined); // Resolve with undefined
          } else {
            resolve(undefined); // Resolve with undefined if user cancels
          }
        }
      } else {
        console.error("Could not read data from CSV.");
        showErrorDialog("generic", `Could not parse the CSV. Ensure you are using the CSV from ${bankType.toUpperCase()} for this tool.`, null, "ok", baseUrl);
        resolve(undefined); // Resolve with undefined on parsing error
      }
    };

    reader.onerror = function (event) {
      console.error("File could not be read! " + event.target.error);
      reject(event.target.error); // Reject the promise on file read error
    };

    reader.readAsText(file);
  });
}

/**
 * Sets up event listeners for the CSV import buttons.
 */
function setupCsvImportButtons() {
  const bokFileInput = document.getElementById("bok-csv-file-input");
  const anbFileInput = document.getElementById("anb-csv-file-input");
  const divvyFileInput = document.getElementById("divvy-csv-file-input");
  const employeeFileInput = document.getElementById("employees-csv-file-input");

  document.getElementById("import-bok-csv-button").addEventListener("click", async () => {
    await handleCsvImport(bokFileInput.files[0], "bok", 1);
  });

  document.getElementById("import-anb-csv-button").addEventListener("click", async () => {
    await handleCsvImport(anbFileInput.files[0], "anb", 4);
  });

  const importDivvyButton = document.getElementById("import-divvy-csv-button");
  importDivvyButton.addEventListener("click", async () => {
    const transactionData = await handleCsvImport(divvyFileInput.files[0], "divvy", 1);
    const employeeData = await handleCsvImport(employeeFileInput.files[0], "divvy", 0);
    if (transactionData && employeeData) {
      importDivvyButton.innerHTML = "Please wait..."; 
      // Change the button to indicate loading since this can take a while
      await manualDivvyWrite(transactionData, employeeData);
      importDivvyButton.innerHTML = "Import to Excel";

    }
  });
}

function setupDateSelection() {
  const currentDate = new Date();
  const oneWeekAgo = subDays(currentDate, 7);
  const formattedCurrentDate = format(currentDate, 'yyyy-MM-dd');
  const formattedDateOneWeekAgo = format(oneWeekAgo, 'yyyy-MM-dd');
  document.getElementById('startDate').value = formattedDateOneWeekAgo;
  document.getElementById('endDate').value = formattedCurrentDate;
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