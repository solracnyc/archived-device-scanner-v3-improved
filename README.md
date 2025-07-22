# Archived Device Scanner v3.0.0 (Improved)

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

## Menu Options

| Option | Description |
|--------|-------------|
| Setup Sheets | Creates required sheets with headers |
| Start Scan | Begins scanning all archived accounts |
| Test Single Account | Tests one account before full scan |
| Check Status | Shows current scan progress |
| Reset Scan | Clears all data and state |

## Output Columns

| Column | Description |
|--------|-------------|
| Email | Archived user's email |
| Status | Found/Not Found/Error |
| Device Count | Number of devices found |
| Model | Device model names |
| Serial Number | Device serial numbers |
| Last Sync | Last sync times |
| OS Version | Chrome OS versions |
| Status | Device status |
| Error | Error messages (if any) |

## Requirements

- Google Workspace Admin access
- Admin SDK API enabled
- Appropriate permissions to view users and devices

## Version History

- **v3.0.0** - Current improved hybrid version (recommended)
- **v2.0.0** - Previous versions (Gemini variant and original)
- **v1.0.0** - Initial release

## License

This project is licensed under the MIT License - see the LICENSE file for details.