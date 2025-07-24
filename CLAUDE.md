# CLAUDE.md - Project Context for Archived Device Scanner

## Project Overview
This is a Google Apps Script project that scans and removes mobile devices from archived Google Workspace accounts. It's designed for Google Workspace Enterprise Plus environments and requires super admin privileges.

## Key Features
- Batch processing of archived accounts with device removal
- Device preview mode (non-destructive scanning)
- Parallel processing support (configurable 1-20 concurrent requests)
- Smart caching to skip recently scanned accounts
- Automatic pause/resume for long-running operations
- Comprehensive logging and error handling

## Development Setup
1. Copy `MinimalArchivedDeviceScanner.gs` to a new Google Apps Script project
2. Enable Admin SDK API service (identifier: `AdminDirectory`)
3. Deploy to a Google Sheet for execution
4. Run as a Google Workspace super admin account

## Code Standards & Best Practices

### General Principles
- **Keep it simple**: Avoid over-engineering; this is a utility script
- **Modular design**: Separate concerns into distinct functions
- **Clear naming**: Use descriptive function and variable names
- **Comprehensive comments**: Document complex logic and API interactions
- **Error handling**: Always wrap API calls with try-catch and retry logic

### Google Apps Script Best Practices
```javascript
// Use const for configuration values
const CONFIG = {
  SHEET_NAME: 'Data',
  MAX_RETRIES: 3
};

// Prefer arrow functions for callbacks
const filtered = data.filter(item => item.isValid);

// Use template literals for string formatting
debugLog(`Processing ${count} items`);

// Always handle API errors gracefully
try {
  const response = AdminDirectory.Users.get(email);
} catch (error) {
  if (error.message.includes('Resource Not Found')) {
    // Handle specific error
  } else {
    // Log and re-throw unexpected errors
    console.error('Unexpected error:', error);
    throw error;
  }
}
```

### Performance Considerations
- Use batch operations where possible
- Implement caching for repeated operations
- Monitor execution time (4.5 minute safety limit)
- Use field masks and projections to reduce API payload size

## Testing Approach
1. **Test Individual Functions**: Use the Script Editor's debugger
2. **Test Single Account**: Use the "Test Single Account" menu option
3. **Preview Mode**: Always run preview before actual deletion
4. **Small Batches**: Test with 5-10 accounts before full runs

## API Quotas & Limits
- **Admin SDK Directory API**: 2,400 queries/minute per project
- **Execution Time**: 6 minutes max (we use 4.5 for safety)
- **PropertiesService**: 500KB total storage
- **UrlFetchApp**: 20MB response size limit

### Current Configuration
```javascript
CONCURRENT_REQUESTS: 5,    // Start low, increase gradually
CACHE_TTL_HOURS: 24,      // Skip recently scanned accounts
BATCH_SIZE: 100,          // Max devices per API call
```

## Common Commands

### Script Deployment
```bash
# If using clasp (optional local development)
clasp create --type sheets
clasp push
clasp open
```

### Debugging in Apps Script Editor
1. View → Logs for console output
2. View → Executions for run history
3. Debug → Add breakpoints for step-through debugging

## Architecture Notes

### State Management
- Uses PropertiesService for persistent state across executions
- Separate state keys for scan and preview operations
- Automatic cleanup after completion

### Error Handling Strategy
1. Exponential backoff for rate limits
2. Detailed error logging to Debug Log sheet
3. Graceful degradation (continue processing other accounts)
4. User-friendly error messages in UI

### Performance Optimizations (v3.3.0)
- Parallel processing with UrlFetchApp.fetchAll()
- BASIC projection instead of FULL for API calls
- Smart caching to skip recent scans
- Field masks to reduce payload size by ~70%

## Future Enhancements
*To be added based on your requirements*

## Security Notes
- Never log OAuth tokens or sensitive data
- Validate all email inputs before API calls
- Restrict to archived accounts only (safety check)
- Use least-privilege principle where possible

## Troubleshooting

### Common Issues
1. **Timeout Errors**: Script automatically resumes after 60 seconds
2. **API Quota Exceeded**: Reduce CONCURRENT_REQUESTS value
3. **Permission Denied**: Ensure running as super admin
4. **Missing Devices**: Check if devices are managed vs BYOD

### Debug Mode
Enable detailed logging by checking the Debug Log sheet after operations.

## Maintenance Guidelines
1. **Before Major Changes**: Create a backup copy of the script
2. **Testing**: Always test in a non-production environment first
3. **Documentation**: Update comments when changing functionality
4. **Version Control**: Update version number in script header

## Resources
- [Google Apps Script Best Practices](https://developers.google.com/apps-script/guides/best-practices)
- [Admin SDK Directory API](https://developers.google.com/admin-sdk/directory)
- [Apps Script Services Quotas](https://developers.google.com/apps-script/guides/services/quotas)

---

*This file provides context for Claude Code when working on this project. Update it as the project evolves.*