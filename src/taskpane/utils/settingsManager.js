/* global Office, document */

export const SETTINGS_KEY = 'excelSettings';
export const defaultSettings = {
  checkFileName: true,
  divvyProxyAddress: "localhost",
  divvyProxyPort: 3001,
  divvyPassword: ""
};

/**
 * Loads settings from local storage, merging with default settings.
 * @returns {object} The loaded settings object.
 */
export function loadSettings() {
  const storedSettingsString = getFromLocalStorage(SETTINGS_KEY);

  // Handle the case where nothing is stored or the API returns "null" as a string
  if (!storedSettingsString || storedSettingsString === "null") {
      return defaultSettings;
  }

  try {
      const loadedSettings = JSON.parse(storedSettingsString);
      
      // Merge with defaults
      const finalSettings = { ...defaultSettings, ...loadedSettings };
      
      return finalSettings;

  } catch (error) {
      console.error("Error parsing stored settings. Falling back to defaults.", error);
      // If the stored JSON is corrupted, fall back to the default settings.
      return defaultSettings;
  }
}

/**
 * Saves settings to local storage.
 * @param {object} settingsObject 
 */
export function saveSettings(settingsObject) {
  try {
      const settingsString = JSON.stringify(settingsObject);
      setInLocalStorage(SETTINGS_KEY, settingsString);
  } catch (error) {
      console.error("Could not save settings.", error);
  }
}

/**
 * Retrieves current settings from the UI elements.
 * @returns {object} An object containing the current settings from the UI.
 */
export function getSettingsFromUI() {
  const checkFilenameEl = document.getElementById('check-filename-checkbox');
  const divvyProxyAddressEl = document.getElementById('divvyProxyAddress');
  const divvyProxyPortEl = document.getElementById('divvyProxyPort');
  const divvyPasswordEl = document.getElementById('divvyPassword');

  return {
    checkFileName: checkFilenameEl.checked,
    divvyProxyAddress: divvyProxyAddressEl.value,
    divvyProxyPort: divvyProxyPortEl.value,
    divvyPassword: divvyPasswordEl.value
  };
}

/**
 * Applies the given settings to the UI elements.
 * @param {object} settings The settings object to apply.
 */
export function applySettingsToUI(settings) {
  const checkFilenameEl = document.getElementById('check-filename-checkbox');
  const divvyProxyAddressEl = document.getElementById('divvyProxyAddress');
  const divvyProxyPortEl = document.getElementById('divvyProxyPort');
  const divvyPasswordEl = document.getElementById('divvyPassword');
  if (checkFilenameEl) {
      checkFilenameEl.checked = settings.checkFileName;
  }
  if (divvyProxyAddressEl) {
    divvyProxyAddressEl.value = settings.divvyProxyAddress;
  }
  if (divvyProxyPortEl) {
    divvyProxyPortEl.value = settings.divvyProxyPort;
  }
  if (divvyPasswordEl) {
    divvyPasswordEl.value = settings.divvyPassword;
  }
}

/**
 * Resets all settings to their default values and applies them to the UI.
 */
export function resetToDefaultSettings() {
  saveSettings(defaultSettings);
  applySettingsToUI(defaultSettings);
}

/**
 * Stores a key-value pair in local storage, respecting Office.context.partitionKey.
 * @param {string} key 
 * @param {string} value 
 */
export function setInLocalStorage(key, value) {
  const myPartitionKey = Office.context.partitionKey;

  // Check if local storage is partitioned, and if it is, use it
  if (myPartitionKey) {
    localStorage.setItem(myPartitionKey + key, value);
  } else {
    localStorage.setItem(key, value);
  }
}

/**
 * Retrieves a value from local storage, respecting Office.context.partitionKey.
 * @param {string} key 
 * @returns {string|null} 
 */
export function getFromLocalStorage(key) {
  const myPartitionKey = Office.context.partitionKey;

  // Check if local storage is partitioned.
  if (myPartitionKey) {
    return localStorage.getItem(myPartitionKey + key);
  } else {
    return localStorage.getItem(key);
  }
}
