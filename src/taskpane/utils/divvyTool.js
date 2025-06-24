/**
 * Retrieves a list of transactions by calling our secure backend proxy.
 * This version fetches all available transactions without a date filter.
 *
 * @returns {Promise<Array>} A promise that resolves to an array of transaction objects.
 */
export async function getBillData() {
  const divvyAddress = document.getElementById('divvyProxyAddress').value;
  const divvyPort = document.getElementById('divvyProxyPort').value;
  const proxyUrl = `${divvyAddress}:${divvyPort}/api/transactions`;

  try {
    console.log('1. [Client] Sending request to proxy:', proxyUrl);

    const response = await fetch(proxyUrl);

    console.log('2. [Client] Received raw response from proxy:', response);

    if (!response.ok) {
      // This block runs if the status is 4xx or 5xx
      console.error('3. [Client] Response was NOT OK. Status:', response.status);
      const errorData = await response.json();
      throw new Error(`API Error: ${errorData.message || response.statusText}`);
    }

    console.log('3. [Client] Response was OK. Parsing JSON...');
    const responseData = await response.json();

    console.log('4. [Client] Parsed JSON data:', responseData);

    return responseData;

  } catch (error) {
    console.error('!!! [Client] An error occurred in getBillTransactions:', error);
    return []; // Return empty array on failure
  }
}