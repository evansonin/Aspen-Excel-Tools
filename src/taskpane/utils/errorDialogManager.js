/* global Office */

/**
 * Displays a custom error dialog to the user.
 * @param {string} errorType The type of error (e.g., "missingSheet", "generic").
 * @param {string} [errorText="Generic error"] The main error message to display.
 * @param {string} [sheetName=null] The name of the sheet, if relevant to the error.
 * @param {string} [button="ok"] The type of buttons to display ("ok" or "yes-no").
 * @param {string} baseUrl The base URL for the add-in's taskpane.
 * @returns {Promise<boolean>} Resolves to true if "yes" is clicked, false otherwise (for "yes-no" button type).
 */
export function showErrorDialog(errorType = null, errorText = "Generic error", sheetName = null, button = "ok", baseUrl) {
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
