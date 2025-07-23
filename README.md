# Archived Device Scanner v3.2.0 (Improved)

The best version of the Google Workspace Archived Device Scanner - a hybrid combining the robust features from v2 original with the improved UI/UX from the Gemini version.

## Why This Version?

This v3.0.0 "Improved" version is the recommended version because it:
- ✅ Combines the best features from both previous versions
- ✅ Has the most efficient code (smallest file size, fewest lines)
- ✅ Includes robust retry logic with exponential backoff
- ✅ Features an improved UI with setup wizard
- ✅ Maintains all critical functionality in a cleaner implementation

## Quick Start

1. **Create a new Google Apps Script project**
2. **Copy the code** from `MinimalArchivedDeviceScanner.gs`
3. **Add the Admin SDK API service**:
   - Click Services → Find "Admin SDK API" → Add
   - Ensure identifier is `AdminDirectory`
4. **Run the setup**:
   - Run → Select `onOpen` → Review permissions
   - Device Scanner → Setup Sheets
5. **Start scanning**:
   - Device Scanner → Start Scan

## Features

- 🚀 **Optimized Performance**: Processes accounts in batches of 100
- 🔄 **Smart Retry Logic**: Automatic retries with exponential backoff
- 📊 **Real-time Progress**: Live updates in status sheet
- 🎯 **Flexible Testing**: Test individual accounts before full scan
- 📝 **Debug Logging**: Optional debug sheet for troubleshooting
- ⏸️ **Resumable**: Pause and resume scans anytime
- 🔍 **Device Preview**: Check what devices accounts have before deletion
- ⏱️ **No Timeouts**: Both scan and preview handle large datasets with automatic pause/resume

## Menu Options

| Option | Description |
|--------|-------------|
| Setup Sheets | Creates required sheets with headers |
| Start Scan | Begins scanning all archived accounts and removes devices (auto-resumes if timeout) |
| Test Single Account | Tests one account before full scan |
| Scan Device Preview | Shows device count and models without deleting (auto-resumes if timeout) |
| Check Status | Shows current scan progress for both removal and preview operations |
| Reset Scan | Clears all data and state for both scans |

## Sheet Structure

### Archived Accounts Sheet
| Column | Description |
|--------|-------------|
| Email Address | User's email address |
| Device Count | Number of devices found (populated by preview) |
| Device Models | List of device models (populated by preview) |
| Last Scanned | Timestamp of last preview scan |

### Device Log Sheet
| Column | Description |
|--------|-------------|
| Timestamp | When the action occurred |
| Email | User's email address |
| Status | REMOVED/FAILED/NOT_FOUND/ERROR |
| Device Model | Device model name |
| Device Type | Type of device |
| Device ID | Unique device identifier |
| Notes | Additional information or error messages |

## Requirements

- Google Workspace Admin access
- Admin SDK API enabled
- Appropriate permissions to view users and devices

## Version History

- **v3.2.0** - Enhanced Preview with Timeout Protection
  - Added automatic pause/resume for preview function to handle thousands of accounts
  - Preview now saves state and continues from where it left off after timeouts
  - Updated Check Status to show progress for both scan types
  - Removed unused columns (Notes, Reserved, Status) from Archived Accounts sheet
- **v3.1.0** - Added Device Preview Feature
  - New "Scan Device Preview" function to check devices before deletion
  - Shows device count, models, and last scanned time
  - Non-destructive preview for planning purposes
- **v3.0.1** - Fixed API parameter syntax and alert dialog errors
  - Corrected Mobiledevices.list() to use proper parameter order
  - Fixed Ui.alert() calls with missing ButtonSet parameter
  - Improved API query efficiency with email filtering
- **v3.0.0** - Initial improved hybrid version
- **v2.0.0** - Previous versions (Gemini variant and original)
- **v1.0.0** - Initial release

## Recent Fixes (v3.0.1)

1. **API Parameter Fix**: The script now correctly passes `'my_customer'` as the first parameter to `AdminDirectory.Mobiledevices.list()`
2. **Query Optimization**: Added email-based filtering to reduce API load
3. **UI Alerts**: Fixed all alert dialogs to include required ButtonSet parameter

## License

This project is licensed under the MIT License - see the LICENSE file for details.