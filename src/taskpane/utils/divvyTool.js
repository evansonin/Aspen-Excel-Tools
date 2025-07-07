import { showErrorDialog } from "./errorDialogManager.js";
import { baseUrl } from "../taskpane.js";



/**
 * Retrieves a list of transactions by calling the backend proxy.
 * Fetches transactions within a specified date range.
 *
 * @param {string} startDate - The start date for filtering transactions (YYYY-MM-DD).
 * @param {string} endDate - The end date for filtering transactions (YYYY-MM-DD).
 * @returns {Promise<Array>} A promise that resolves to an array of transaction objects.
 */
export async function getBillData(startDate, endDate) {
  const divvyAddress = document.getElementById('divvyProxyAddress').value;
  const divvyPort = document.getElementById('divvyProxyPort').value;
  const divvyPassword = document.getElementById('divvyPassword').value;
  
  // Construct the proxy URL with date parameters, using HTTPS
  const proxyUrl = `https://${divvyAddress}:${divvyPort}/api/transactions?startDate=${startDate}&endDate=${endDate}`;


  try {
    const response = await fetch(proxyUrl, {
      headers: {
        'x-passcode': divvyPassword // Add the passcode to the request headers
      }
    });
    if (!response.ok) {
      // If the status is 4xx or 5xx
      console.error('[Client] Response was NOT OK. Status:', response.status);
      const errorData = await response.json();
      throw new Error(`API Error: ${errorData.message || response.statusText}`);
    }
    const responseData = await response.json();
    return responseData;

  } catch (error) {
    console.error('[Client] An error occurred in getBillTransactions:', error);
    if(error.message.includes("Invalid or missing passcode")) {
      showErrorDialog("generic", "Incorrect password. Please check you have entered the correct password and hit \"Save Settings\".", null, "ok", baseUrl);
    } else if(error.message.includes("Failed to fetch")) {
      showErrorDialog("generic", "Failed to retrieve Divvy transaction data; a call was made to the server but no response was received. Make sure the server address and port are correct.", null, "ok", baseUrl);
    }
    return []; // Return empty array on failure
  }
}




/**
 * Writes Divvy transaction data to the currently active Excel spreadsheet.
 *
 * @param {Object} divvyData - An array of transaction objects to be written.
 */
export async function writeToDivvySpreadsheet(users, transactions, employees, startDate, endDate) {

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

      // Find the position type from the employees array
      let positionType = "Unknown";
      const foundEmployee = employees.find(employee => employee[0] === fullName);
      if (foundEmployee) {
        positionType = foundEmployee[1]; // Get the position string
      }
      sortedTransactions[transactionKey].positionType = positionType;

    } else {
      // If user is not found, log a warning and assign a default value
      console.warn(`User with ID ${userId} not found for transaction at key/index: ${transactionKey}`);
      sortedTransactions[transactionKey] = { name: "Unknown Cardholder", positionType: "Unknown" };
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
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      // Set the name defined above
      sheet.name = dateRangeForSheetName;

      await setupSpreadsheet(context, sheet); // Call the spreadsheet setup function

      // Write transaction data
      let rowIndex = 2; // Start writing data from the second row (after headers)
      await writeTransactionDataToSheet(context, sheet, sortedTransactions, rowIndex);

      await context.sync();
    });
  } catch (error) {
    console.error("[Client] An error occurred in writeToDivvySpreadsheet:", error);
    showErrorDialog("generic", "Error copying Divvy/Bill export data to Excel. If you recently requested data, please wait a few minutes or change the date range.", null, "ok", baseUrl);
  }
}

/**
 * Writes transaction data to the specified Excel sheet.
 * @param {Excel.RequestContext} context - The Excel request context.
 * @param {Excel.Worksheet} sheet - The active worksheet.
 * @param {Object} sortedTransactions - The processed transaction data.
 * @param {number} startRowIndex - The row index to start writing data from.
 */
async function writeTransactionDataToSheet(context, sheet, sortedTransactions, startRowIndex) {
  let rowIndex = startRowIndex;

  // Define the mapping of transaction data keys to Excel columns and their formatting
  const columnMappings = [
    { column: dateColumn, dataKey: 'date' },
    { column: accountingCodeColumn, dataKey: 'accountingCode'},
    { column: firstProgramColumn, dataKey: 'firstProgram'},
    { column: departmentColumn, dataKey: 'department'},
    { column: feAccountColumn, dataKey: 'feAccount'},
    { column: secondProgramColumn, dataKey: 'secondProgram', numberFormat: '@'}, 
    { column: importAmountColumn, dataKey: 'amount', numberFormat: '0.00' },
    { column: expenseDescriptionColumn, dataKey: 'expenseDescription'},
    { column: cardholderColumn, dataKey: 'name' },
    { column: transactionDescriptionColumn, dataKey: 'merchantName' },
    { column: typeColumn, dataKey: 'positionType' },
    { column: jeDescriptionColumn, dataKey: 'jeDescription' }
  ];

  for (const transactionKey in sortedTransactions) {
    const transactionData = sortedTransactions[transactionKey];

    for (const mapping of columnMappings) {
      const range = sheet.getRange(mapping.column.letter + rowIndex);

      // Apply specific formatting if defined in the mapping *before* setting values
      if (mapping.numberFormat) {
        range.numberFormat = mapping.numberFormat;
      }
      
      range.values = [[transactionData[mapping.dataKey]]];
      range.format.horizontalAlignment = "General";
    }
    rowIndex++;
  }
}

export async function manualDivvyWrite(transactions, employees) {
  let sortedTransactions = {}
  const CSV_LAYOUT = {
    transactionId: 0,
    splitFrom: 1,
    dateIndex: 2,
    clearedTime: 3,
    firstName: 4,
    lastName: 5,
    merchant: 6,
    merchantNameIndex: 7,
    amount: 8,
    cardName: 9,
    cardType: 10,
    cardLastFour: 11,
    cardExpDate: 12,
    cardProgram: 13,
    reviews: 14,
    budgetId: 16,
    cardId: 17,
    userId: 18,
    authorizedAt: 19,
    cardHolderEmail: 20,
    localAmount: 21,
    foreignExchangeFee: 22,
    merchantAddress: 23,
    status: 24,
    budget: 29,
    receipt: 31,
    mmc: 32,
    description: 33,
    accountCode: 34,
    department: 35,
    program: 36
  };

  let earliestDate = null; // Initialize to null
  let latestDate = null;  // Initialize to null

  for (let transactionKey in transactions) {
    // Date
    const occurredTime = transactions[transactionKey][CSV_LAYOUT.dateIndex];
    const formattedDate = new Date(occurredTime).toISOString().split('T')[0];
    sortedTransactions[transactionKey] = { date: formattedDate};
    
    // Update earliest and latest dates
    if (!earliestDate || formattedDate < earliestDate) {
      earliestDate = formattedDate;
    }
    if (!latestDate || formattedDate > latestDate) {
      latestDate = formattedDate;
    }
    
    // TRANSACTION DESCRIPTION
    const transactionDescription = transactions[transactionKey][CSV_LAYOUT.merchantNameIndex];
    sortedTransactions[transactionKey].merchantName = transactionDescription;

    // ACCOUNTING CODE
    const accountingCode = transactions[transactionKey][CSV_LAYOUT.accountCode];
    sortedTransactions[transactionKey].accountingCode = accountingCode

    // Program (1st)
    const firstProgram = transactions[transactionKey][CSV_LAYOUT.program];
    sortedTransactions[transactionKey].firstProgram = firstProgram;

    // Department
    const department = transactions[transactionKey][CSV_LAYOUT.department];
    sortedTransactions[transactionKey].department = department;

    // FE Account
    // TBA

    // Program (2nd) - Extract the leading part before the first space, or use the whole string if no space
    let secondProgramValue = firstProgram;
    const spaceIndex = firstProgram.indexOf(' ');
    if (spaceIndex !== -1) {
      secondProgramValue = firstProgram.substring(0, spaceIndex);
    }
    sortedTransactions[transactionKey].secondProgram = secondProgramValue;

    // Amount
    const amountString = transactions[transactionKey][CSV_LAYOUT.amount];
    // Remove '$' and spaces, then parse as a float and invert the sign
    const amount = -parseFloat(amountString.replace(/[\$\s]/g, ''));
    sortedTransactions[transactionKey].amount = amount;

    // Description of Purchase
    const expenseDescription = transactions[transactionKey][CSV_LAYOUT.description];
    sortedTransactions[transactionKey].expenseDescription = expenseDescription ? expenseDescription.replace(/\r?\n|\r/g, ' ') : '';

    // Cardholder
    const firstName = transactions[transactionKey][CSV_LAYOUT.firstName];
    const lastName = transactions[transactionKey][CSV_LAYOUT.lastName];
    let cardholderName;
    if (firstName && lastName){
      cardholderName = lastName + ", " + firstName;
    } else if (firstName && !lastName) {
      cardholderName = firstName;
    } else if (lastName && !firstName) {
      cardholderName = lastName;
    } else {
      cardholderName = "None";
    }
    sortedTransactions[transactionKey].name = cardholderName;

    // Position (from selected CSV)
    let positionType = "Unknown (Not found)"; // Initialize with default

    if (cardholderName === "Divvy") {
      positionType = "Divvy";
    } else {
      // Find the position type from the employees array, removing spaces for comparison
      const foundEmployee = employees.find(employee => employee[0].replaceAll(' ','') === cardholderName.replaceAll(' ',''));
      if (foundEmployee) {
        positionType = foundEmployee[1]; // Get the position string
      }
    }
    sortedTransactions[transactionKey].positionType = positionType; // Assign the determined position type

    // JE Description
    let jeDescription;
    if (expenseDescription) {
      jeDescription = transactionDescription + ' - ' + expenseDescription;
    } else {
      jeDescription = transactionDescription;
    }
    
    if (jeDescription.length > 80) {
      jeDescription = jeDescription.slice(0, 77) + "...";
    } 
    sortedTransactions[transactionKey].jeDescription = jeDescription;
  }   
  
  // Convert the sortedTransactions object into an array of its values
  let transactionsArray = Object.values(sortedTransactions);

  // Sort the array: "Divvy" comes first, then sort alphabetically by name
  transactionsArray.sort((a, b) => {
    if (a.name === "Divvy" && b.name !== "Divvy") {
      return -1; // "Divvy" comes before other names
    }
    if (a.name !== "Divvy" && b.name === "Divvy") {
      return 1; // Other names come after "Divvy"
    }
    // If neither are "Divvy", sort alphabetically by name
    return a.name.localeCompare(b.name);
  });

  // Helper function to format dates for the sheet name
  const formatDateForSheetName = (dateString) => {
    // Parse the YYYY-MM-DD string directly to avoid timezone issues
    const [year, month, day] = dateString.split('-');
    const shortYear = String(parseInt(year) % 100).padStart(2, '0');
    return `${parseInt(month)}.${day}.${shortYear}`;
  };

  // Construct the sheet name based on the earliest and latest dates
  let dateRangeForSheetName = "Transactions"; // Default name
  if (earliestDate && latestDate) {
    const formattedEarliest = formatDateForSheetName(earliestDate);
    const formattedLatest = formatDateForSheetName(latestDate);
    dateRangeForSheetName = `${formattedEarliest} - ${formattedLatest} Transactions`;
  }

  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      // Set the sheet name
      sheet.name = dateRangeForSheetName;
      await setupSpreadsheet(context, sheet); // Call the new setup function
      await context.sync();

      let rowIndex = 2;
      // Pass the sorted array to the writing function
      await writeTransactionDataToSheet(context, sheet, transactionsArray, rowIndex);

    });
  } catch (error) {
    console.error("[Client] An error occurred in manualDivvyWrite:", error);
    showErrorDialog("generic", "Error setting up spreadsheet for manual Divvy write.", null, "ok", baseUrl);
  }
}

/**
 * Sets up the Excel spreadsheet with predefined headers and formatting.
 * @param {Excel.RequestContext} context - The Excel request context.
 * @param {Excel.Worksheet} sheet - The active worksheet.
 */
const headerColor = "DAF2D0";

async function setupSpreadsheet(context, sheet) {
  // Write column headers and apply formatting
  for (const column of columns) {
    const headerRange = sheet.getRange(column.letter + "1");
    headerRange.values = [[column.title]];
    headerRange.format.fill.color = headerColor;
    headerRange.format.font.bold = true;
    headerRange.format.verticalAlignment = "Bottom";
    headerRange.format.horizontalAlignment = "Center"; 
    headerRange.format.wrapText = true;

    // Set column width
    sheet.getRange(column.letter + ":" + column.letter).format.columnWidth = column.width / 2; // Not sure why it has to divded by two but it does
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
  letter: "H",
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
  title: "Stategic\nInitiatives",
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

