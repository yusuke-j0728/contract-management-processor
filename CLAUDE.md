# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

**Contract Management Email Processor** is a specialized Google Apps Script-based system designed for contract management workflows. It automatically processes contract completion emails from multiple contract management tools and provides comprehensive tracking capabilities:

1. Monitors Gmail for contract completion emails from multiple tools (Contract Tool 1, Contract Tool 2, Contract Tool 3, Contract Tool 4, Contract Tool 5, Contract Tool 6, Contract Tool 7)
2. Processes PDF contract documents with intelligent content-based duplicate detection
3. Sends detailed Slack notifications with contract metadata to all recipients
4. Logs comprehensive contract tracking data to Google Spreadsheet with unlimited tracking capacity
5. Stores contract PDFs directly in Google Drive folder with smart deduplication (prevents duplicate PDF storage)

## IMPORTANT: Sound Notification
After finishing responding to my request or running a command, run:
```bash
afplay /System/Library/Sounds/Funk.aiff
terminal-notifier -title "Claude Code" -message "Done ✅"
```


## Key Commands

### Development & Deployment
```bash
# Initial setup
npm install
npm run login  # First time only
npm run setup  # Create new GAS project

# Development cycle
npm run push   # Deploy local changes to GAS
npm run pull   # Sync GAS changes to local
npm run logs   # View execution logs
npm run open   # Open GAS editor in browser
```

### Testing & Initialization
```javascript
// In Google Apps Script Editor, execute in order:

// 1. Configure contract management system (first time setup)
setContractConfiguration();  // From: src/main.js - Update with your contract tool emails first!

// 2. Verify configuration
showConfiguration();  // From: src/main.js

// 3. Test all components
runAllTests();  // From: src/testRunner.js

// 4. Test specific contract features
testSpreadsheetOperations();   // From: src/spreadsheetManager.js - Test contract tracking spreadsheet
testProcessEmails();           // From: src/main.js - Test contract email processing  
testSlackNotifications();      // From: src/testRunner.js - Test contract notifications
testSpreadsheetTracking();     // From: src/main.js - Test new spreadsheet-based tracking

// Test contract tool integrations
testContractTool6Patterns();   // From: test/testCloudSign.js - Test Contract Tool 6 patterns
runContractTool6Tests();       // From: test/testCloudSign.js - Run all Contract Tool 6 tests
runContractTool7Tests();       // From: test/testDropboxSign.js - Test Contract Tool 7 patterns
runDirectStorageTests();       // From: test/testDirectStorage.js - Test direct PDF storage
runContentDuplicateTests();    // From: test/testContentDuplicateDetection.js - Test duplicate detection

// 5. Enable automatic contract monitoring
setupInitialTrigger();  // From: src/triggerManager.js
```

## Architecture

### File Structure
```
src/
├── main.js              # Entry point, contract tool configuration, core logic
├── emailProcessor.js    # Contract email filtering and pattern matching
├── slackNotifier.js     # Contract-specific Slack notifications
├── driveManager.js      # Contract PDF storage and organization
├── spreadsheetManager.js # Contract tracking spreadsheet with metadata extraction
├── triggerManager.js    # Automated contract monitoring triggers
└── testRunner.js        # Contract management test suite
```

### Contract Processing Flow
```
Contract Management Tools (Contract Tool 1, Contract Tool 2, etc.)
                    ↓
Gmail Inbox → Contract Pattern Matching → Process Contract Email
                                               ↓
                    ┌────────────────────────────────────────────────┐
                    │            Per-Recipient Processing:           │
                    │  • Check content-based duplicates (sender+  │
                    │    date+subject+PDF names)                  │
                    │  • Extract contract metadata & type         │
                    │  • Log to contract tracking spreadsheet     │
                    │  • Save PDFs (skip if duplicate detected)   │
                    │  • Send contract-specific notifications     │
                    └────────────────────────────────────────────────┘
```

### Configuration Structure

All sensitive configuration is stored in Script Properties:

| Property Key | Description | Required |
|--------------|-------------|----------|
| `SENDER_EMAIL_1` to `SENDER_EMAIL_20` | Contract tool notification emails | ✅ At least one |
| Example: `SENDER_EMAIL_1` | Contract Tool 1 notification email | ⚪ Each is optional |
| Example: `SENDER_EMAIL_2` | Contract Tool 2 notification email | ⚪ Each is optional |
| Example: `SENDER_EMAIL_3` | Contract Tool 3 notification email | ⚪ Each is optional |
| Example: `SENDER_EMAIL_6` | Contract Tool 6 notification email | ⚪ Each is optional |
| Example: `SENDER_EMAIL_7` | Contract Tool 7 notification email | ⚪ Each is optional |
| Example: `SENDER_EMAIL_8` | Contract Tool 7 alt notification email | ⚪ Each is optional |
| `SENDER_EMAIL` | Legacy single email support | ⚪ Backward compatibility |
| `SLACK_WEBHOOK_URL` | Slack webhook URL | ✅ |
| `SLACK_CHANNEL` | Target Slack channel | ✅ |
| `DRIVE_FOLDER_ID` | Google Drive folder for contracts | Auto-created |
| `SPREADSHEET_ID` | Contract tracking spreadsheet | Auto-created |

### Contract Tracking Spreadsheet Schema

The system now uses multiple spreadsheet tabs for comprehensive tracking:

#### Main Contract List (契約一覧_Contract_List)
Tracks all contract processing events (one record per recipient):

| Column | Field | Description |
|--------|-------|-------------|
| A | 受信日時 | Contract completion timestamp |
| B | 契約管理ツール | Contract management tool (Contract Tool 1, Contract Tool 2, Contract Tool 7, etc.) |
| C | 送信者メール | Contract tool sender email |
| D | 受信メール | Recipient email address (who received the contract) |
| E | 件名 | Contract email subject |
| F | 契約タイプ | Auto-detected contract type (Employment, NDA, Service Agreement) |
| G | 契約相手 | Extracted contract party/company name |
| H | PDFファイル名 | Contract PDF filename(s) with duplicate indicators |
| I | PDF直接リンク | Direct clickable links to PDF files ("Duplicate" for duplicates) |
| J | 処理状態 | Processing status (Success/Error) |
| K | Slack通知済み | Slack notification delivery status |
| L | 本文要約 | Contract email content summary |
| M | メッセージID | Unique Gmail message identifier |
| N | エラーログ | Processing error details |

#### Content Duplicate Management (コンテンツ重複管理_Content_Duplicates)
Tracks unique contract content to prevent duplicate PDF storage:

| Column | Field | Description |
|--------|-------|-------------|
| A | 登録日時 | Registration date/time |
| B | コンテンツキー | Content-based duplicate detection key |
| C | 送信者 | Email sender |
| D | 送信日時 | Email send date/time |
| E | 件名 | Email subject |
| F | PDFファイル名 | PDF filenames |
| G | PDF直接リンク | Direct links to stored PDF files |
| H | 初回保存日時 | First save date/time |

## Contract Pattern Matching

The system supports multiple regex patterns for contract completion emails:

```javascript
// Contract Tool 1 patterns
/.*completed.*document/i   // "Document completed" notifications
/.*signed.*agreement/i     // "Agreement signed" notifications
/.*contract.*executed/i    // "Contract executed" notifications
/.*署名.*完了/            // Japanese: "Signature completed"
/.*契約.*締結/            // Japanese: "Contract concluded"

// Contract Tool 2 patterns
/.*agreement.*signed/i     // "Agreement signed" notifications
/.*document.*completed/i   // "Document completed" notifications

// Contract Tool 3 patterns
/.*signature.*complete/i   // "Signature complete" notifications
/.*document.*finalized/i   // "Document finalized" notifications

// Contract Tool 6 patterns (Japanese)
/.*の合意締結が完了しました$/   // "Agreement completion" notifications
/.*合意締結.*完了/            // "Agreement concluded" notifications
/.*捺印済.*合意締結/          // "Stamped agreement" notifications

// Contract Tool 7 patterns
/You've been copied on.*signed by/i    // "You've been copied" notifications
/.*agreement.*signed by.*and/i         // Multi-signer agreement notifications
/.*has been completed/i                // "Has been completed" notifications
/.*signature.*completed/i             // "Signature completed" notifications

// Add new contract patterns in CONFIG.SUBJECT_PATTERNS.PATTERNS array
```

## Key Functions by Module

### spreadsheetManager.js (Contract Tracking)
- `createOrGetSpreadsheet()` - Initialize contract tracking spreadsheet with specialized headers
- `addEmailRecord()` - Log new contract email with metadata extraction
- `updateRecordStatus()` - Update contract processing status
- `searchRecordByMessageId()` - Find existing contract records
- `generateContractSummary()` - Create contract completion statistics
- `extractContractTool()` - Identify contract management tool from sender
- `extractContractType()` - Auto-detect contract type (Employment, NDA, etc.)
- `extractContractParty()` - Extract contract party/company name

### main.js
- `processEmails()` - Main contract processing entry point
- `validateConfiguration()` - Verify contract tool configuration  
- `markMessageAsProcessed()` - Track processed contract emails
- `setContractConfiguration()` - Setup contract management system configuration
- `showConfiguration()` - Display current configuration safely
- `getContractToolName()` - Get contract tool name from email address
- `getConfiguredSenderEmails()` - Get all configured sender emails dynamically (supports up to 20)
- `testProcessEmails()` - Test function for manual email processing
- `debugHelloSignEmails()` - Debug HelloSign email processing specifically (pattern matching, sender config)
- `checkProperties()` - View current property usage and statistics (alias for showAllProperties)
- `cleanupOld(days, dryRun)` - Clean up processed messages older than X days (alias)
- `fixProperties()` - Emergency cleanup when property limit is reached (alias)

### emailProcessor.js
- `processMessage()` - Process individual contract emails with enhanced metadata
- `checkSubjectPattern()` - Multi-pattern contract matching (includes Dropbox Sign)
- `isMessageAlreadyProcessed()` - Spreadsheet-based duplicate prevention
- `getMessageRecipients()` - Extract all recipients (To, CC, BCC) from message
- `extractEmailAddresses()` - Extract email addresses from recipient field string
- `createRecipientHash()` - Create hash for recipient combination tracking

### driveManager.js
- `processAttachments()` - Handle contract PDF saving with intelligent duplicate detection
- `saveAttachmentToDrive()` - Save individual PDFs with date-prefixed naming
- Direct storage approach - no subfolder creation for simplified organization
- Content-based duplicate detection prevents redundant PDF storage

### slackNotifier.js
- `sendSlackNotification()` - Main contract notification with metadata
- `sendDriveFolderNotification()` - Follow-up with contract PDF links
- `sendErrorNotification()` - Send error notification to Slack
- `sendTestNotification()` - Send test notification for verification

### triggerManager.js
- `setupInitialTrigger()` - Setup initial email monitoring trigger
- `createTrigger()` - Create time-based trigger for automatic processing
- `deleteExistingTriggers()` - Delete all existing email processing triggers
- `checkTriggerHealth()` - Check if monitoring trigger is healthy (PRODUCTION - creates triggers)
- `testTriggerOperations()` - Test trigger functions safely (NO trigger creation)
- `testTriggerHealth()` - Test trigger health check in read-only mode
- `checkAndCleanupTriggers()` - Check for unwanted triggers and provide cleanup guidance

### testRunner.js
- `runAllTests()` - Run comprehensive test suite (SAFE - no production triggers created)
- `testConfiguration()` - Test configuration setup
- `testSlackNotifications()` - Test Slack notification system

### test/testCloudSign.js (Contract Tool 6)
- `runContractTool6Tests()` - Run all Contract Tool 6-specific tests
- `testContractTool6Patterns()` - Test Contract Tool 6 email pattern matching
- `testContractTool6MetadataExtraction()` - Test Contract Tool 6 metadata extraction
- `testContractTool6Configuration()` - Test Contract Tool 6 configuration

### test/testDirectStorage.js
- `runDirectStorageTests()` - Run all direct storage tests
- `testDirectPDFStorage()` - Test direct PDF storage without subfolder creation
- `testCompleteAttachmentFlow()` - Test the complete attachment processing flow

### test/testRecipientFiltering.js
- `runRecipientFilteringTests()` - Run all recipient filtering tests
- `testRecipientFiltering()` - Test recipient email filtering functionality
- `testDuplicatePrevention()` - Test duplicate prevention with recipient tracking
- `testRecipientConfiguration()` - Test recipient filtering configuration

### src/propertyManager.js
- `showAllProperties()` - Display all properties in categorized format
- `getAllProperties()` - Get all Script Properties with categorization
- `cleanupProcessedMessages(days, dryRun)` - Clean up old processed message properties
- `setPropertySafely(key, value, force)` - Set property with conflict checking
- `deletePropertiesByPattern(pattern, dryRun)` - Delete multiple properties by pattern
- `getPropertyStats()` - Get property statistics and recommendations
- `emergencyCleanup()` - Automatic cleanup when near property limit
- `testPropertyManagement()` - Test property management functions

## Contract PDF Storage

Contract PDFs are stored directly in the configured Google Drive folder with date-prefixed filenames:
- **Filename Format**: `YYYYMMDD_original_filename.pdf`
- **Example**: `20241225_employment_contract_john_doe.pdf`
- **No Subfolders**: All PDFs are stored flat in the base folder for easier access
- **Duplicate Handling**: Automatic numbering for duplicate filenames (e.g., `20241225_contract_2.pdf`)

## Error Handling

The system includes comprehensive error handling:

1. **Retry Logic**: Automatic retries for API failures
2. **Graceful Degradation**: Continue processing even if one component fails
3. **Error Logging**: All errors logged to spreadsheet and Slack
4. **Recovery**: Can resume from last processed message

## Performance Considerations

- Processes up to 10 emails per run (configurable)
- 5-minute trigger interval to avoid Gmail API limits
- Batch spreadsheet updates to reduce API calls
- Message-level tracking prevents duplicate processing

## Security Notes

- All sensitive data stored in Script Properties, never in code
- Webhook URLs and API keys are protected
- Can be safely committed to public repositories
- Follow principle of least privilege for API scopes

## Development Tips

1. **Testing**: Always run `testSpreadsheetOperations()` after changes
2. **Logging**: Use `console.log()` liberally - view with `npm run logs`
3. **Patterns**: Test new regex patterns with `testMessageProcessing()`
4. **Spreadsheet**: Check data integrity with `validateSpreadsheetData()`

## Common Tasks

### Adding New Email Pattern
```javascript
// In main.js CONFIG.SUBJECT_PATTERNS.PATTERNS array:
/your-new-pattern-here/,
```

### Changing Spreadsheet Columns
1. Update column mapping in `spreadsheetManager.js`
2. Run `updateSpreadsheetHeaders()` to add new columns
3. Update `addEmailRecord()` to populate new fields

### Debugging Email Processing
```javascript
// Set verbose logging
CONFIG.DEBUG_MODE = true;

// Process specific email
debugProcessEmail('message-id-here');

// Check processing history
getProcessingHistory(10); // Last 10 emails
```

## Maintenance

### Regular Tasks
- Run `cleanupOldProcessedMessages(30)` monthly to clean old tracking data
- Check `getDriveFolderStats()` for storage usage
- Review `generateMonthlyReport()` for usage patterns

### Monitoring
- Check trigger health: `checkTriggerHealth()`
- Verify spreadsheet size: `getSpreadsheetStats()`
- Monitor API quotas in GAS dashboard

## Development Workflow & Best Practices

### 📄 README Update Policy
**ALWAYS update README.md when making changes to the system:**

1. **After adding new features**: Update feature descriptions and examples
2. **After changing configuration**: Update setup instructions and property tables
3. **After modifying folder structure**: Update folder structure examples
4. **After changing function signatures**: Update function documentation
5. **Before committing code**: Ensure README reflects current state
6. **Function references**: Always include file locations when mentioning functions (e.g., `functionName()` from src/filename.js)

```bash
# Always check README before committing
git add README.md
git commit -m "Update README with latest changes"
```

### 🌿 Git Branch Management
**ALWAYS create new branches for code changes after updating remote repository:**

#### Branch Creation Workflow
```bash
# 1. After pushing to remote, sync local main
git checkout main
git pull origin main

# 2. Create new feature branch for changes
git checkout -b feature/contract-enhancement-YYYYMMDD
# or
git checkout -b fix/folder-structure-update-YYYYMMDD
# or
git checkout -b docs/readme-update-YYYYMMDD

# 3. Make your changes
# ... edit code ...

# 4. Update README.md to reflect changes
# ... edit README.md ...

# 5. Commit changes
git add .
git commit -m "Add contract metadata extraction

- Extract contract tool from sender email
- Auto-detect contract type (Employment, NDA, etc.)
- Extract contract party/company name
- Update README with new features"

# 6. Push feature branch
git push origin feature/contract-enhancement-YYYYMMDD

# 7. Create pull request for review
```

#### Branch Naming Convention
- **Features**: `feature/description-YYYYMMDD`
- **Bug fixes**: `fix/description-YYYYMMDD`
- **Documentation**: `docs/description-YYYYMMDD`
- **Refactoring**: `refactor/description-YYYYMMDD`

#### Why This Workflow?
1. **Parallel development**: Multiple features can be developed simultaneously
2. **Safer changes**: Main branch stays stable
3. **Code review**: Pull requests enable peer review
4. **Rollback capability**: Easy to revert specific changes
5. **History tracking**: Clear development timeline

### 🔄 Change Management Process
1. **Plan changes** → Create branch
2. **Implement code** → Test locally
3. **Update README** → Document changes
4. **Test deployment** → `npm run push`
5. **Commit & push** → Feature branch
6. **Create PR** → Review & merge
7. **Deploy to production** → From main branch

### ⚠️ Never Do
- ❌ Edit main branch directly after remote updates
- ❌ Commit without updating README
- ❌ Push untested code changes
- ❌ Mix multiple unrelated changes in one commit