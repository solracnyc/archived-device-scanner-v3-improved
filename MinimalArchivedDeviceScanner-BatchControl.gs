/**
 * Minimal Archived Device Scanner - Batch Control Version
 * Clean, efficient Google Apps Script with advanced batch processing controls
 * 
 * @author Your Organization
 * @version 3.1.0-beta
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
  RETRY_DELAY: 1000, // Base delay in ms
  
  // ===== BATCH CONTROL FEATURES =====
  ENABLE_BATCH_CONTROL: false, // Set to true to enable batch features
  ACCOUNTS_PER_BATCH: 50, // How many accounts to process per batch
  ENABLE_AUTO_CONTINUE: true, // Auto-continue to next batch
  STATUS_COLUMN: 'D' // Column for account processing status
};

//================== MENU & SETUP ==================
function onOpen() {
  const menu = SpreadsheetApp.getUi()
    .createMenu('📱 Device Scanner')
    .addItem('1️⃣ Setup Sheets', 'setupSheets')
    .addSeparator()
    .addItem('🚀 Start Scanning', 'startScan')
    .addItem('🔍 Test Single Account', 'testAccount')
    .addSeparator();
  
  // Add batch control menu items if enabled
  if (CONFIG.ENABLE_BATCH_CONTROL) {
    menu.addSubMenu(SpreadsheetApp.getUi()
      .createMenu('🎯 Batch Control')
      .addItem(`📦 Process Next ${CONFIG.ACCOUNTS_PER_BATCH}`, 'processNextBatch')
      .addItem('📝 Mark Range for Processing', 'markRangeForProcessing')
      .addItem('🔄 Reprocess Failed Only', 'reprocessFailed')
      .addItem('📊 Batch Status Report', 'batchStatusReport')
      .addSeparator()
      .addItem('⚙️ Setup Status Column', 'setupStatusColumn'))
    .addSeparator();
  }
  
  menu.addItem('📊 Check Status', 'checkStatus')
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

//================== BATCH CONTROL FUNCTIONS ==================

/**
 * Setup status column for batch processing
 */
function setupStatusColumn() {
  if (!CONFIG.ENABLE_BATCH_CONTROL) {
    SpreadsheetApp.getUi().alert('Batch control is disabled. Enable it in CONFIG first.');
    return;
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const accountsSheet = ss.getSheetByName(CONFIG.ACCOUNTS_SHEET_NAME);
  
  if (!accountsSheet) {
    SpreadsheetApp.getUi().alert('Archived Accounts sheet not found. Run Setup Sheets first.');
    return;
  }
  
  // Add status column header
  accountsSheet.getRange('D1').setValue('Status').setFontWeight('bold');
  
  // Set all existing accounts to "Pending" if they don't have a status
  const lastRow = accountsSheet.getLastRow();
  if (lastRow > 1) {
    const statusRange = accountsSheet.getRange(`D2:D${lastRow}`);
    const values = statusRange.getValues();
    
    for (let i = 0; i < values.length; i++) {
      if (!values[i][0]) {
        values[i][0] = 'Pending';
      }
    }
    statusRange.setValues(values);
  }
  
  SpreadsheetApp.getUi().alert('✅ Status Column Setup Complete', 'Status column added. Existing accounts marked as "Pending".', SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Process next batch of pending accounts
 */
function processNextBatch() {
  if (!CONFIG.ENABLE_BATCH_CONTROL) {
    SpreadsheetApp.getUi().alert('Batch control is disabled. Enable it in CONFIG first.');
    return;
  }
  
  const accounts = getPendingAccounts(CONFIG.ACCOUNTS_PER_BATCH);
  if (accounts.length === 0) {
    SpreadsheetApp.getUi().alert('No pending accounts found to process.');
    return;
  }
  
  SpreadsheetApp.getUi().alert(`🚀 Starting Batch`, `Processing ${accounts.length} pending accounts.`, SpreadsheetApp.getUi().ButtonSet.OK);
  
  // Process the batch
  processBatchAccounts(accounts);
}

/**
 * Get pending accounts for batch processing
 */
function getPendingAccounts(limit = CONFIG.ACCOUNTS_PER_BATCH) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const accountsSheet = ss.getSheetByName(CONFIG.ACCOUNTS_SHEET_NAME);
  
  if (!accountsSheet) return [];
  
  const lastRow = accountsSheet.getLastRow();
  if (lastRow <= 1) return [];
  
  const data = accountsSheet.getRange(`A2:D${lastRow}`).getValues();
  const pendingAccounts = [];
  
  for (let i = 0; i < data.length && pendingAccounts.length < limit; i++) {
    const email = data[i][0];
    const status = data[i][3] || 'Pending'; // Default to Pending if empty
    
    if (email && status === 'Pending') {
      pendingAccounts.push({
        email: email,
        row: i + 2, // +2 because we start from row 2 and array is 0-based
        status: status
      });
    }
  }
  
  return pendingAccounts;
}

/**
 * Process a batch of accounts with status tracking
 */
function processBatchAccounts(accounts) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const accountsSheet = ss.getSheetByName(CONFIG.ACCOUNTS_SHEET_NAME);
  let processed = 0;
  let errors = 0;
  
  debugLog(`Starting batch processing of ${accounts.length} accounts`);
  
  for (const account of accounts) {
    try {
      // Mark as processing
      accountsSheet.getRange(account.row, 4).setValue('Processing');
      debugLog(`Processing: ${account.email}`);
      
      // Process the account (reuse existing logic)
      const result = processAccount(account.email);
      
      // Update status based on result
      if (result.errors > 0) {
        accountsSheet.getRange(account.row, 4).setValue('Error');
        errors++;
      } else {
        accountsSheet.getRange(account.row, 4).setValue('Completed');
        processed++;
      }
      
    } catch (error) {
      accountsSheet.getRange(account.row, 4).setValue('Error');
      debugLog(`Error processing ${account.email}: ${error.message}`);
      errors++;
    }
  }
  
  debugLog(`Batch complete: ${processed} processed, ${errors} errors`);
  SpreadsheetApp.getUi().alert('🎯 Batch Complete', `Processed: ${processed}\nErrors: ${errors}`, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Show batch status report
 */
function batchStatusReport() {
  if (!CONFIG.ENABLE_BATCH_CONTROL) {
    SpreadsheetApp.getUi().alert('Batch control is disabled. Enable it in CONFIG first.');
    return;
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const accountsSheet = ss.getSheetByName(CONFIG.ACCOUNTS_SHEET_NAME);
  
  if (!accountsSheet) {
    SpreadsheetApp.getUi().alert('Archived Accounts sheet not found.');
    return;
  }
  
  const lastRow = accountsSheet.getLastRow();
  if (lastRow <= 1) {
    SpreadsheetApp.getUi().alert('No accounts found.');
    return;
  }
  
  const statusData = accountsSheet.getRange(`D2:D${lastRow}`).getValues();
  const statusCount = {
    'Pending': 0,
    'Processing': 0,
    'Completed': 0,
    'Error': 0,
    'Empty': 0
  };
  
  statusData.forEach(row => {
    const status = row[0] || 'Empty';
    if (statusCount.hasOwnProperty(status)) {
      statusCount[status]++;
    } else {
      statusCount['Empty']++;
    }
  });
  
  const total = lastRow - 1;
  const report = `📊 Batch Status Report
  
Total Accounts: ${total}
✅ Completed: ${statusCount.Completed}
⏳ Pending: ${statusCount.Pending}
🔄 Processing: ${statusCount.Processing}
❌ Errors: ${statusCount.Error}
⚪ No Status: ${statusCount.Empty}

Progress: ${Math.round((statusCount.Completed / total) * 100)}%`;

  SpreadsheetApp.getUi().alert('📊 Status Report', report, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Mark selected range for processing
 */
function markRangeForProcessing() {
  if (!CONFIG.ENABLE_BATCH_CONTROL) {
    SpreadsheetApp.getUi().alert('Batch control is disabled. Enable it in CONFIG first.');
    return;
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const selection = ss.getSelection();
  const range = selection.getActiveRange();
  
  if (!range) {
    SpreadsheetApp.getUi().alert('Please select a range of accounts first.');
    return;
  }
  
  const accountsSheet = ss.getSheetByName(CONFIG.ACCOUNTS_SHEET_NAME);
  if (range.getSheet().getName() !== CONFIG.ACCOUNTS_SHEET_NAME) {
    SpreadsheetApp.getUi().alert('Please select a range in the Archived Accounts sheet.');
    return;
  }
  
  const startRow = range.getRow();
  const numRows = range.getNumRows();
  
  if (startRow < 2) {
    SpreadsheetApp.getUi().alert('Please select account rows (not the header).');
    return;
  }
  
  // Mark selected accounts as Pending
  const statusRange = accountsSheet.getRange(startRow, 4, numRows, 1);
  const values = Array(numRows).fill(['Pending']);
  statusRange.setValues(values);
  
  SpreadsheetApp.getUi().alert('✅ Range Marked', `${numRows} accounts marked as "Pending" for processing.`, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Reprocess only failed accounts
 */
function reprocessFailed() {
  if (!CONFIG.ENABLE_BATCH_CONTROL) {
    SpreadsheetApp.getUi().alert('Batch control is disabled. Enable it in CONFIG first.');
    return;
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const accountsSheet = ss.getSheetByName(CONFIG.ACCOUNTS_SHEET_NAME);
  
  if (!accountsSheet) {
    SpreadsheetApp.getUi().alert('Archived Accounts sheet not found.');
    return;
  }
  
  const lastRow = accountsSheet.getLastRow();
  if (lastRow <= 1) {
    SpreadsheetApp.getUi().alert('No accounts found.');
    return;
  }
  
  // Mark all "Error" status as "Pending"
  const data = accountsSheet.getRange(`A2:D${lastRow}`).getValues();
  let errorCount = 0;
  
  for (let i = 0; i < data.length; i++) {
    if (data[i][3] === 'Error') {
      accountsSheet.getRange(i + 2, 4).setValue('Pending');
      errorCount++;
    }
  }
  
  if (errorCount === 0) {
    SpreadsheetApp.getUi().alert('No failed accounts found to reprocess.');
  } else {
    SpreadsheetApp.getUi().alert('🔄 Ready to Reprocess', `${errorCount} failed accounts marked as "Pending".`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}