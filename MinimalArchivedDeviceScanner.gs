/**
 * Minimal Archived Device Scanner - Improved Version
 * Clean, efficient Google Apps Script to remove mobile devices from archived accounts
 * 
 * @author Your Organization
 * @version 3.0.1
 * @license MIT
 */

//================== CONFIGURATION ==================
const CONFIG = {
  ACCOUNTS_SHEET_NAME: 'Archived Accounts',
  LOG_SHEET_NAME: 'Device Log',
  DEBUG_SHEET_NAME: 'Debug Log',
  MAX_EXECUTION_TIME: 4.5 * 60 * 1000, // 4.5 minutes
  STATE_KEY: 'ArchivedDeviceScannerState',
  BATCH_SIZE: 100, // Devices per API call
  RETRY_ATTEMPTS: 3,
  RETRY_DELAY: 1000 // Base delay in ms
};

//================== MENU & SETUP ==================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📱 Device Scanner')
    .addItem('1️⃣ Setup Sheets', 'setupSheets')
    .addSeparator()
    .addItem('🚀 Start Scanning', 'startScan')
    .addItem('🔍 Test Single Account', 'testAccount')
    .addSeparator()
    .addItem('📊 Check Status', 'checkStatus')
    .addItem('🔄 Reset Scanner', 'resetScan')
    .addToUi();
}

function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Accounts sheet
  let accountsSheet = ss.getSheetByName(CONFIG.ACCOUNTS_SHEET_NAME);
  if (!accountsSheet) accountsSheet = ss.insertSheet(CONFIG.ACCOUNTS_SHEET_NAME, 0);
  
  accountsSheet.clear();
  accountsSheet.getRange('A1:B1')
    .setValues([['Email Address', 'Notes']])
    .setFontWeight('bold');
  accountsSheet.getRange('A2:B2')
    .setValues([['user@example.com', 'Add emails here, then run Start Scanning']]);
  accountsSheet.setFrozenRows(1);
  
  // Log sheet
  let logSheet = ss.getSheetByName(CONFIG.LOG_SHEET_NAME);
  if (!logSheet) logSheet = ss.insertSheet(CONFIG.LOG_SHEET_NAME, 1);
  
  logSheet.clear();
  logSheet.getRange('A1:G1')
    .setValues([['Timestamp', 'Email', 'Status', 'Device Model', 'Device Type', 'Device ID', 'Notes']])
    .setFontWeight('bold');
  logSheet.setFrozenRows(1);
  
  // Debug sheet
  let debugSheet = ss.getSheetByName(CONFIG.DEBUG_SHEET_NAME);
  if (!debugSheet) debugSheet = ss.insertSheet(CONFIG.DEBUG_SHEET_NAME, 2);
  
  debugSheet.clear();
  debugSheet.getRange('A1:B1')
    .setValues([['Timestamp', 'Message']])
    .setFontWeight('bold');
  debugSheet.setFrozenRows(1);
  
  SpreadsheetApp.getUi().alert('✅ Setup Complete', 'Sheets created. Add emails to "Archived Accounts" sheet.', SpreadsheetApp.getUi().ButtonSet.OK);
}

//================== MAIN SCAN LOGIC ==================
function startScan() {
  const startTime = Date.now();
  
  try {
    let state = loadState();
    
    if (!state) {
      const accounts = getAccounts();
      if (!accounts.length) {
        SpreadsheetApp.getUi().alert('No accounts found. Please add emails to the Archived Accounts sheet.');
        return;
      }
      
      state = initializeState(accounts);
      debugLog('Starting new scan for ' + accounts.length + ' accounts');
    } else {
      debugLog('Resuming scan at index ' + state.currentIndex);
    }
    
    // Main processing loop
    while (state.currentIndex < state.accounts.length) {
      if (Date.now() - startTime > CONFIG.MAX_EXECUTION_TIME) {
        saveState(state);
        scheduleContinuation();
        debugLog('Time limit reached, continuing in 60 seconds');
        return;
      }
      
      const email = state.accounts[state.currentIndex];
      if (isValidEmail(email)) {
        const result = processAccount(email);
        state.devicesRemoved += result.devicesRemoved;
        state.errorsCount += result.errors;
      }
      
      state.currentIndex++;
      
      // Save progress every 10 accounts
      if (state.currentIndex % 10 === 0) {
        saveState(state);
        SpreadsheetApp.getActiveSpreadsheet()
          .toast(`Progress: ${state.currentIndex}/${state.accounts.length} accounts`);
      }
    }
    
    completeScan(state);
    
  } catch (error) {
    debugLog('CRITICAL ERROR: ' + error.toString());
    throw error;
  }
}

function processAccount(email) {
  let devicesRemoved = 0;
  let errors = 0;
  
  try {
    debugLog('Processing: ' + email);
    
    // Get and validate user
    const user = retryApiCall(() => AdminDirectory.Users.get(email));
    
    if (!user.archived && !user.suspended) {
      logResult(email, 'ACTIVE_USER', null, null, null, 'User is active - skipped');
      return { devicesRemoved: 0, errors: 0 };
    }
    
    // Get user devices efficiently
    const devices = getUserDevices(email);
    
    if (!devices.length) {
      logResult(email, 'NO_DEVICES', null, null, null, 'No mobile devices found');
      return { devicesRemoved: 0, errors: 0 };
    }
    
    // Remove devices
    for (const device of devices) {
      try {
        retryApiCall(() => AdminDirectory.Mobiledevices.remove('my_customer', device.resourceId));
        logResult(email, 'REMOVED', device.model, device.type, device.resourceId, 'Device removed');
        devicesRemoved++;
      } catch (error) {
        logResult(email, 'FAILED', device.model, device.type, device.resourceId, 'Remove failed: ' + error.message);
        errors++;
      }
    }
    
  } catch (error) {
    if (error.message.includes('Resource Not Found')) {
      logResult(email, 'NOT_FOUND', null, null, null, 'User not found');
    } else {
      logResult(email, 'ERROR', null, null, null, error.message);
      errors++;
    }
  }
  
  return { devicesRemoved, errors };
}

//================== DEVICE FETCHING (OPTIMIZED) ==================
function getUserDevices(targetEmail) {
  const devices = [];
  let pageToken = null;
  
  do {
    try {
      // This query makes the API call more specific and efficient
      // by only asking for devices matching the user's email.
      const response = retryApiCall(() => AdminDirectory.Mobiledevices.list('my_customer', {
        maxResults: CONFIG.BATCH_SIZE,
        pageToken: pageToken,
        query: 'email:' + targetEmail,
        orderBy: 'EMAIL',
        projection: 'FULL'
      }));
      
      // The API response now only contains devices for the target user
      if (response.mobiledevices) {
        devices.push(...response.mobiledevices);
      }
      
      pageToken = response.nextPageToken;
      
    } catch (error) {
      // This ensures that if the API call fails, the main function will know about it
      // and log it correctly as an error.
      debugLog('Error fetching devices for ' + targetEmail + ': ' + error.message);
      throw new Error(error.message); 
    }
  } while (pageToken);
  
  return devices;
}

//================== API RETRY LOGIC ==================
function retryApiCall(apiCall) {
  let lastError;
  
  for (let attempt = 1; attempt <= CONFIG.RETRY_ATTEMPTS; attempt++) {
    try {
      return apiCall();
    } catch (error) {
      lastError = error;
      
      if (attempt === CONFIG.RETRY_ATTEMPTS) break;
      
      // Check if retry is appropriate
      if (error.message.includes('Rate Limit') || 
          error.message.includes('Backend Error') ||
          error.message.includes('Internal Error')) {
        
        const delay = CONFIG.RETRY_DELAY * Math.pow(2, attempt - 1);
        debugLog(`API error (attempt ${attempt}), retrying in ${delay}ms: ${error.message}`);
        Utilities.sleep(delay);
      } else {
        break; // Don't retry for other errors
      }
    }
  }
  
  throw lastError;
}

//================== STATE MANAGEMENT ==================
function loadState() {
  const stateJson = PropertiesService.getScriptProperties().getProperty(CONFIG.STATE_KEY);
  return stateJson ? JSON.parse(stateJson) : null;
}

function saveState(state) {
  PropertiesService.getScriptProperties().setProperty(CONFIG.STATE_KEY, JSON.stringify(state));
}

function initializeState(accounts) {
  const state = {
    accounts: accounts,
    currentIndex: 0,
    totalAccounts: accounts.length,
    devicesRemoved: 0,
    errorsCount: 0,
    startTime: new Date().toISOString()
  };
  saveState(state);
  return state;
}

function resetScan() {
  const ui = SpreadsheetApp.getUi();
  if (ui.alert('Reset Scanner', 'Clear all progress and start over?', ui.ButtonSet.YES_NO) === ui.Button.YES) {
    PropertiesService.getScriptProperties().deleteProperty(CONFIG.STATE_KEY);
    deleteAllTriggers();
    debugLog('Scanner reset by user');
    ui.alert('Reset complete');
  }
}

//================== CONTINUATION ==================
function scheduleContinuation() {
  deleteAllTriggers();
  ScriptApp.newTrigger('startScan')
    .timeBased()
    .after(60 * 1000)
    .create();
}

function completeScan(state) {
  deleteAllTriggers();
  PropertiesService.getScriptProperties().deleteProperty(CONFIG.STATE_KEY);
  
  const message = `Scan complete!\n\nAccounts: ${state.totalAccounts}\nDevices removed: ${state.devicesRemoved}\nErrors: ${state.errorsCount}`;
  debugLog('Scan completed: ' + message.replace(/\n/g, ' '));
  SpreadsheetApp.getUi().alert('✅ Scan Complete', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

//================== UTILITIES ==================
function getAccounts() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.ACCOUNTS_SHEET_NAME);
  if (!sheet) return [];
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  
  return sheet.getRange(2, 1, lastRow - 1, 1)
    .getValues()
    .map(row => row[0].toString().trim())
    .filter(email => email && isValidEmail(email));
}

function deleteAllTriggers() {
  ScriptApp.getProjectTriggers().forEach(trigger => ScriptApp.deleteTrigger(trigger));
}

function isValidEmail(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

//================== LOGGING ==================
function debugLog(message) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.DEBUG_SHEET_NAME);
    if (sheet) sheet.appendRow([new Date(), message]);
    console.log(message);
  } catch (e) {
    console.error('Debug log failed:', e);
  }
}

function logResult(email, status, model, type, resourceId, notes) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.LOG_SHEET_NAME);
  sheet.appendRow([new Date(), email, status, model || '', type || '', resourceId || '', notes || '']);
}

//================== STATUS & TESTING ==================
function checkStatus() {
  const state = loadState();
  if (!state) {
    SpreadsheetApp.getUi().alert('No active scan');
    return;
  }
  
  const percentage = Math.round((state.currentIndex / state.totalAccounts) * 100);
  const message = `Progress: ${state.currentIndex}/${state.totalAccounts} (${percentage}%)\n` +
                  `Devices removed: ${state.devicesRemoved}\n` +
                  `Errors: ${state.errorsCount}\n` +
                  `Started: ${new Date(state.startTime).toLocaleString()}`;
  
  SpreadsheetApp.getUi().alert('📊 Scan Status', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

function testAccount() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Test Account', 'Enter email address:', ui.ButtonSet.OK_CANCEL);
  
  if (response.getSelectedButton() === ui.Button.OK) {
    const email = response.getResponseText().trim();
    if (isValidEmail(email)) {
      debugLog('Testing account: ' + email);
      processAccount(email);
      ui.alert('Test complete. Check Device Log for results.');
    } else {
      ui.alert('Invalid email address');
    }
  }
}