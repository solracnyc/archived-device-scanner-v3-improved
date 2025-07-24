/**
 * Minimal Archived Device Scanner - Improved Version
 * Clean, efficient Google Apps Script to remove mobile devices from archived accounts
 * 
 * @author Your Organization
 * @version 3.3.0
 * @license MIT
 */

//================== CONFIGURATION ==================
const CONFIG = {
  ACCOUNTS_SHEET_NAME: 'Archived Accounts',
  LOG_SHEET_NAME: 'Device Log',
  DEBUG_SHEET_NAME: 'Debug Log',
  MAX_EXECUTION_TIME: 4.5 * 60 * 1000, // 4.5 minutes
  STATE_KEY: 'ArchivedDeviceScannerState',
  PREVIEW_STATE_KEY: 'ArchivedDeviceScannerPreviewState',
  BATCH_SIZE: 100, // Devices per API call
  RETRY_ATTEMPTS: 3,
  RETRY_DELAY: 1000, // Base delay in ms
  
  // Parallel processing configuration
  CONCURRENT_REQUESTS: 5, // Number of users to process in parallel (1 = sequential, max 20)
  CACHE_TTL_HOURS: 24, // Skip users scanned within this timeframe (0 = disabled)
  CACHE_KEY_PREFIX: 'last_scan_' // Prefix for cache keys in PropertiesService
};

//================== MENU & SETUP ==================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📱 Device Scanner')
    .addItem('1️⃣ Setup Sheets', 'setupSheets')
    .addSeparator()
    .addItem('🚀 Start Scanning', 'startScan')
    .addItem('🔍 Test Single Account', 'testAccount')
    .addItem('🔍 Scan Device Preview', 'scanDevicePreview')
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
  accountsSheet.getRange('A1:D1')
    .setValues([['Email Address', 'Device Count', 'Device Models', 'Last Scanned']])
    .setFontWeight('bold');
  accountsSheet.getRange('A2')
    .setValue('user@example.com');
  accountsSheet.getRange('B2')
    .setValue('Add emails here, then run Scan Device Preview');
  accountsSheet.setFrozenRows(1);
  
  // Set column widths for better visibility
  accountsSheet.setColumnWidth(1, 200); // Email
  accountsSheet.setColumnWidth(2, 100); // Device Count
  accountsSheet.setColumnWidth(3, 300); // Device Models
  accountsSheet.setColumnWidth(4, 150); // Last Scanned
  
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

//================== PREVIEW STATE MANAGEMENT ==================
function loadPreviewState() {
  const stateJson = PropertiesService.getScriptProperties().getProperty(CONFIG.PREVIEW_STATE_KEY);
  return stateJson ? JSON.parse(stateJson) : null;
}

function savePreviewState(state) {
  PropertiesService.getScriptProperties().setProperty(CONFIG.PREVIEW_STATE_KEY, JSON.stringify(state));
}

function initializePreviewState(emails) {
  const state = {
    emails: emails,
    currentIndex: 0,
    processed: 0,
    errors: 0,
    startTime: new Date().toISOString()
  };
  savePreviewState(state);
  return state;
}

function resetScan() {
  const ui = SpreadsheetApp.getUi();
  if (ui.alert('Reset Scanner', 'Clear all progress and start over?', ui.ButtonSet.YES_NO) === ui.Button.YES) {
    PropertiesService.getScriptProperties().deleteProperty(CONFIG.STATE_KEY);
    PropertiesService.getScriptProperties().deleteProperty(CONFIG.PREVIEW_STATE_KEY);
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

//================== PREVIEW CONTINUATION ==================
function schedulePreviewContinuation() {
  deleteAllTriggers();
  ScriptApp.newTrigger('scanDevicePreview')
    .timeBased()
    .after(60 * 1000)
    .create();
}

function completePreview(state) {
  deleteAllTriggers();
  PropertiesService.getScriptProperties().deleteProperty(CONFIG.PREVIEW_STATE_KEY);
  
  const duration = Math.round((new Date() - new Date(state.startTime)) / 1000);
  const message = `Preview complete! Accounts scanned: ${state.processed}, Errors: ${state.errors}, Duration: ${duration} seconds`;
  
  // Try to show UI alert if we're in a UI context
  try {
    SpreadsheetApp.getUi().alert('✅ Preview Complete', message.replace(/,/g, '\n'), SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    // If UI is not available (running from trigger), use toast notification
    try {
      SpreadsheetApp.getActiveSpreadsheet().toast(message, '✅ Preview Complete', 10);
    } catch (e2) {
      // If even toast fails, just log it
    }
  }
  
  debugLog(`Device preview completed: ${message}`);
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
  const previewState = loadPreviewState();
  
  if (!state && !previewState) {
    SpreadsheetApp.getUi().alert('No active scan or preview');
    return;
  }
  
  let message = '';
  
  if (state) {
    const percentage = Math.round((state.currentIndex / state.totalAccounts) * 100);
    message += `📊 DEVICE REMOVAL SCAN:\n` +
               `Progress: ${state.currentIndex}/${state.totalAccounts} (${percentage}%)\n` +
               `Devices removed: ${state.devicesRemoved}\n` +
               `Errors: ${state.errorsCount}\n` +
               `Started: ${new Date(state.startTime).toLocaleString()}\n\n`;
  }
  
  if (previewState) {
    const percentage = Math.round((previewState.currentIndex / previewState.emails.length) * 100);
    message += `🔍 PREVIEW SCAN:\n` +
               `Progress: ${previewState.currentIndex}/${previewState.emails.length} (${percentage}%)\n` +
               `Processed: ${previewState.processed}\n` +
               `Errors: ${previewState.errors}\n` +
               `Started: ${new Date(previewState.startTime).toLocaleString()}`;
  }
  
  SpreadsheetApp.getUi().alert('📊 Status', message, SpreadsheetApp.getUi().ButtonSet.OK);
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

//================== DEVICE PREVIEW FUNCTION ==================
/**
 * Scan emails and preview device information without deleting anything
 */
function scanDevicePreview() {
  const startTime = Date.now();
  
  try {
    let state = loadPreviewState();
    
    if (!state) {
      // Initialize preview
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const accountsSheet = ss.getSheetByName(CONFIG.ACCOUNTS_SHEET_NAME);
      
      if (!accountsSheet) {
        SpreadsheetApp.getUi().alert('Archived Accounts sheet not found. Run Setup Sheets first.');
        return;
      }
      
      const lastRow = accountsSheet.getLastRow();
      if (lastRow <= 1) {
        SpreadsheetApp.getUi().alert('No accounts found. Please add emails to the Archived Accounts sheet.');
        return;
      }
      
      const ui = SpreadsheetApp.getUi();
      const response = ui.alert(
        '🔍 Scan Device Preview',
        `This will scan ${lastRow - 1} accounts to preview their devices.\nNo devices will be deleted.\n\nContinue?`,
        ui.ButtonSet.YES_NO
      );
      
      if (response !== ui.Button.YES) return;
      
      // Get all emails
      const emails = accountsSheet.getRange(`A2:A${lastRow}`).getValues().flat().filter(email => email);
      
      if (emails.length === 0) {
        ui.alert('No valid email addresses found.');
        return;
      }
      
      state = initializePreviewState(emails);
      debugLog('Starting device preview for ' + emails.length + ' accounts');
    } else {
      debugLog('Resuming preview at index ' + state.currentIndex);
    }
    
    const accountsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.ACCOUNTS_SHEET_NAME);
    
    // Main processing loop with time checking
    while (state.currentIndex < state.emails.length) {
      if (Date.now() - startTime > CONFIG.MAX_EXECUTION_TIME) {
        savePreviewState(state);
        schedulePreviewContinuation();
        debugLog('Preview time limit reached, continuing in 60 seconds');
        
        // Try to show pause notification if possible
        try {
          SpreadsheetApp.getActiveSpreadsheet().toast(
            `Preview paused at ${state.currentIndex}/${state.emails.length}. Will resume automatically...`,
            '⏸️ Pausing'
          );
        } catch (e) {
          // Silent fail if UI context not available
        }
        
        return;
      }
      
      const email = state.emails[state.currentIndex];
      const rowNumber = state.currentIndex + 2; // +2 because we start from row 2
      
      if (!isValidEmail(email)) {
        state.processed++;
        state.currentIndex++;
        continue;
      }
      
      try {
        const devices = getUserDevices(email);
        
        if (devices.length > 0) {
          // Extract device models
          const deviceModels = devices.map(device => {
            const model = device.model || device.type || 'Unknown';
            const os = device.os ? ` (${device.os})` : '';
            return model + os;
          }).join(', ');
          
          // Update the row with device information
          accountsSheet.getRange(rowNumber, 2).setValue(devices.length); // Device Count
          accountsSheet.getRange(rowNumber, 3).setValue(deviceModels); // Device Models
          accountsSheet.getRange(rowNumber, 4).setValue(new Date().toLocaleString()); // Last Scanned
          
          debugLog(`Preview: ${email} has ${devices.length} devices: ${deviceModels}`);
        } else {
          // No devices found
          accountsSheet.getRange(rowNumber, 2).setValue(0); // Device Count
          accountsSheet.getRange(rowNumber, 3).setValue('No devices'); // Device Models
          accountsSheet.getRange(rowNumber, 4).setValue(new Date().toLocaleString()); // Last Scanned
          
          debugLog(`Preview: ${email} has no devices`);
        }
        
      } catch (error) {
        accountsSheet.getRange(rowNumber, 2).setValue('Error'); // Device Count
        accountsSheet.getRange(rowNumber, 3).setValue(error.message); // Error message
        accountsSheet.getRange(rowNumber, 4).setValue(new Date().toLocaleString()); // Last Scanned
        
        debugLog(`Error previewing devices for ${email}: ${error.message}`);
        state.errors++;
      }
      
      state.processed++;
      state.currentIndex++;
      
      // Save progress every 10 accounts (without toast notification)
      if (state.processed % 10 === 0) {
        savePreviewState(state);
      }
    }
    
    // Preview complete
    completePreview(state);
    
  } catch (error) {
    debugLog('CRITICAL ERROR in preview: ' + error.toString());
    throw error;
  }
}