import { showErrorDialog } from "./errorDialogManager.js";
import { baseUrl  } from "../taskpane.js";

/**
 * Retrieves a list of transactions by calling our secure backend proxy.
 * This version fetches transactions within a specified date range.
 *
 * @param {string} startDate - The start date for filtering transactions (YYYY-MM-DD).
 * @param {string} endDate - The end date for filtering transactions (YYYY-MM-DD).
 * @returns {Promise<Array>} A promise that resolves to an array of transaction objects.
 */
export async function getBillData(startDate, endDate) {
  const divvyAddress = document.getElementById('divvyProxyAddress').value;
  const divvyPort = document.getElementById('divvyProxyPort').value;
  
  // Construct the proxy URL with date parameters
  const proxyUrl = `${divvyAddress}:${divvyPort}/api/transactions?startDate=${startDate}&endDate=${endDate}`;


  try {
    const response = await fetch(proxyUrl);
    if (!response.ok) {
      // This block runs if the status is 4xx or 5xx
      console.error('3. [Client] Response was NOT OK. Status:', response.status);
      const errorData = await response.json();
      throw new Error(`API Error: ${errorData.message || response.statusText}`);
    }
    const responseData = await response.json();
    return responseData;

  } catch (error) {
    console.error('!!! [Client] An error occurred in getBillTransactions:', error);
    return []; // Return empty array on failure
  }
}


/**
 * Writes Divvy transaction data to the currently active Excel spreadsheet.
 *
 * @param {Object} divvyData - An array of transaction objects to be written.
 */
export async function writeToDivvySpreadsheet(users, transactions, startDate, endDate) {

  // Sheet name logic
  const formatDateForSheetName = (dateString) => {
    // Parse the YYYY-MM-DD string directly to avoid timezone issues
    const [year, month, day] = dateString.split('-');
    const shortYear = String(parseInt(year) % 100).padStart(2, '0');
    return `${parseInt(month)}.${day}.${shortYear}`;
  };
  const dateRangeForSheetName = `${formatDateForSheetName(startDate)} - ${formatDateForSheetName(endDate)} Transactions`;

  // Data sorting 
  let sortedTransactions = {};
  for (let transactionKey in transactions) {
    const currentTransaction = transactions[transactionKey]; // Get the transaction object
    // Name
    const userId = currentTransaction.userId; // Get the userId from the transaction

    // Find the user in the 'users' array based on userId
    const user = users.find(u => u.id === userId);

    if (user) {
      // If user is found, format the name as "LastName, FirstName"
      const fullName = `${user.lastName}, ${user.firstName}`;
      sortedTransactions[transactionKey] = { name: fullName }; // Store the full name associated with the transaction's key/index
    } else {
      // If user is not found, log a warning and assign a default value
      console.warn(`User with ID ${userId} not found for transaction at key/index: ${transactionKey}`);
      sortedTransactions[transactionKey] = { name: "Unknown Cardholder" };
    }
    // Date
    const occurredTime = currentTransaction.occurredTime;
    const formattedDate = new Date(occurredTime).toISOString().split('T')[0]; // Formats to "YYYY-MM-DD"
    sortedTransactions[transactionKey].date = formattedDate;
    // Amount
    sortedTransactions[transactionKey].amount = (currentTransaction.amount);
    // Transaction Description
    sortedTransactions[transactionKey].merchantName = (currentTransaction.merchantName);
  }
  console.log(sortedTransactions);
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      // Set the name defined above
      sheet.name = dateRangeForSheetName;
      // TODO: Implement logic to write divvyData to the sheet
      // Set the height of the header row
      const headerRow = sheet.getRange("1:1"); 
      // Write column headers and apply formatting
      for (const column of columns) {
        const headerRange = sheet.getRange(column.letter + "1");
        headerRange.values = [[column.title]];
        headerRange.format.fill.color = headerColor;
        headerRange.format.font.bold = true;
        headerRange.format.verticalAlignment = "Bottom";
        headerRange.format.horizontalAlignment = "Center";

        // Set column width
        sheet.getRange(column.letter + ":" + column.letter).format.columnWidth = column.width/2; // Not sure why it has to divded by two but it does
        headerRange.format.wrapText = true; // Enable text wrapping
      }

      // Write transaction data
      let rowIndex = 2; // Start writing data from the second row (after headers)
      for (const transactionKey in sortedTransactions) {
        const transactionData = sortedTransactions[transactionKey];
        
        // Write date to the Date column (A)
        const dateRange = sheet.getRange(dateColumn.letter + rowIndex);
        dateRange.values = [[transactionData.date]];

        // Write cardholder name to the Cardholder column (J)
        const cardholderRange = sheet.getRange(cardholderColumn.letter + rowIndex);
        cardholderRange.values = [[transactionData.name]];

        // Write amount to Amount column (H)
        const amountRange = sheet.getRange(importAmountColumn.letter + rowIndex);
        amountRange.values = [[transactionData.amount]];
        amountRange.numberFormat = "0.00";


        // Write merchant name to Transaction Description (B)
        const transactionDescriptionRange = sheet.getRange(transactionDescriptionColumn.letter + rowIndex);
        transactionDescriptionRange.values = [[transactionData.merchantName]];

        rowIndex++;
      }

      await context.sync();
    });
  } catch (error) {
    console.error("!!! [Client] An error occurred in writeToDivvySpreadsheet:", error);
    showErrorDialog("generic", "Error copying Divvy/Bill export data to Excel. If you recently requested data, please wait a few minutes or change the date range.", null, "ok", baseUrl);
  }
}


// COLUMN OBJECTS 
const dateColumn = {
  title: "Date",
  index: 0,
  letter: "A",
  width: 111
};
const transactionDescriptionColumn = {
  title: "TRANSACTION DESCRIPTION",
  index: 1,
  letter: "B",
  width: 385
};
const accountingCodeColumn = {
  title: "ACCOUNTING CODE",
  index: 2,
  letter: "C",
  width: 278
};
const firstProgramColumn = {
  title: "Program",
  index: 3,
  letter: "D",
  width: 336
};
const departmentColumn = {
  title: "Department",
  index: 4,
  letter: "E",
  width: 302
};
const feAccountColumn = {
  title: "FE Account",
  index: 5,
  letter: "F",
  width: 126
};
const secondProgramColumn = {
  title: "Program",
  index: 6,
  letter: "G",
  width: 93
};
const importAmountColumn = {
  title: "Import Amount",
  index: 7,
  letter: "H", // Changed 'h' to 'H' for consistency
  width: 127
};
const expenseDescriptionColumn = {
  title: "EXPENSE DESCRIPTION",
  index: 8,
  letter: "I",
  width: 605
};
const cardholderColumn = {
  title: "Cardholder",
  index: 9,
  letter: "J",
  width: 217
};
const typeColumn = {
  title: "Type",
  index: 10,
  letter: "K",
  width: 88
};
const fixedAssetColumn = {
  title: "Fixed Asset",
  index: 11,
  letter: "L",
  width: 135
};
const strategicInitiativesColumn = {
  title: "Stategic Initiatives",
  index: 12,
  letter: "M",
  width: 163
};
const jeDescriptionColumn = {
  title: "JE Description",
  index: 13,
  letter: "N",
  width: 865
};
const columns = 
  [dateColumn,
  transactionDescriptionColumn,
  accountingCodeColumn,
  firstProgramColumn,
  departmentColumn,
  feAccountColumn,
  secondProgramColumn,
  importAmountColumn,
  expenseDescriptionColumn,
  cardholderColumn,
  typeColumn,
  fixedAssetColumn,
  strategicInitiativesColumn,
  jeDescriptionColumn
  ];

const headerColor = "DAF2D0";
const columnHeight = 29; // Corrected variable name from 'columnHeight' to 'const columnHeight'