# Contract Management Email Processor

A specialized Google Apps Script-based system designed for contract management workflows. Automatically processes contract completion emails from multiple contract management tools (Contract Tool 1, Contract Tool 2, Contract Tool 3, Contract Tool 4, Contract Tool 6, Contract Tool 7, etc.), saves PDF contracts directly to Google Drive with intelligent duplicate detection, sends notifications to Slack, and maintains a comprehensive contract tracking spreadsheet.

## Features

### 📋 Contract Management Integration
- 🏢 **Multi-Tool Support**: Works with Contract Tool 1, Contract Tool 2, Contract Tool 3, Contract Tool 4, Contract Tool 5, Contract Tool 6, Contract Tool 7, and more
- 📧 **Smart Pattern Recognition**: Recognizes contract completion emails in multiple languages (English/Japanese)
- 🔍 **Intelligent Classification**: Automatically categorizes contract types (Employment, NDA, Service Agreement, etc.)
- 🏭 **Company Extraction**: Identifies contract parties and company names from email content

### 📄 PDF Contract Processing
- 📱 **Automatic PDF Detection**: Only processes emails containing PDF contract documents
- 📁 **Direct Storage**: Saves PDFs directly to base folder with `YYYYMMDD_filename.pdf` naming
- 🔄 **Intelligent Duplicate Detection**: Prevents duplicate PDF storage using content-based detection
- 🔗 **Direct Access Links**: Generates clickable links to PDF files
- 📊 **File Metadata Tracking**: Records PDF filenames and storage locations

### 📊 Contract Tracking & Analytics
- 📈 **Comprehensive Spreadsheet**: Logs all contract data with 14 specialized columns including recipient tracking
- 🔍 **Advanced Search**: Filter by contract type, management tool, contract party, or recipient
- 📊 **Daily Summaries**: Generate contract completion statistics and reports
- 📈 **Performance Metrics**: Track success rates and processing times
- 📧 **Recipient Tracking**: Records which email address received each contract notification

### 🚨 Real-time Notifications
- 📱 **Slack Integration**: Instant notifications with contract details and PDF links
- 🎯 **Contract-Specific Alerts**: Tailored messages showing contract type and parties
- 📎 **Follow-up Messages**: Separate notifications with direct Drive folder access
- ⚠️ **Error Handling**: Immediate alerts for processing failures

### 🔧 System Reliability
- ⏰ **Automated Scheduling**: Continuous monitoring every 5 minutes
- 🔄 **Smart Duplicate Prevention**: Content-based detection prevents duplicate PDF storage while ensuring all recipients receive notifications
- 📊 **Unlimited Tracking**: Spreadsheet-based duplicate tracking with no Script Properties limitations
- 🏷️ **Intelligent Skip Labels**: Automatically skips non-contract emails to prevent repeated checking
- 🛡️ **Error Recovery**: Robust error handling with detailed logging
- 📊 **Performance Monitoring**: Built-in analytics and health checks

## Quick Start

### 1. Prerequisites
- Google Account (for Gmail, Drive, and Apps Script access)
- Slack Workspace (with Incoming Webhooks permissions)
- Node.js v14+ (for development environment)
- Clasp CLI

### 2. Local Development Setup
```bash
# Clone and setup
git clone https://github.com/yusuke-j0728/contract-management-processor.git
cd contract-management-processor

# Install dependencies
npm install

# Clasp authentication (first time only)
npm run login

# Create Google Apps Script project
clasp create --type standalone --title "Contract Management Email Processor"

# If you encounter "Project file already exists" error:
# rm .clasp.json
# clasp create --type standalone --title "Contract Management Email Processor"

# Verify project creation
cat .clasp.json  # Should show scriptId

# Deploy code to Google Apps Script
npm run push
```

### 3. Configuration
⚠️ **SECURITY NOTICE**: All sensitive information is stored in Script Properties, NOT in code!

#### 3.1 Slack Webhook Setup
1. Add Incoming Webhooks app to your Slack workspace
2. Select notification channel and generate Webhook URL
3. **Keep this URL secure** - it will be stored in Script Properties

#### 3.2 Drive Folder Setup (Optional - Set Before Configuration)

The system supports both automatic and manual Drive folder setup:

**🆕 New: Direct PDF Storage with Smart Duplicate Detection**
- 📦 **PDF files only**: Contract PDFs are saved directly to the base folder
- 📁 **Simplified structure**: PDFs saved with `YYYYMMDD_filename.pdf` naming (no subfolders)
- 🔄 **Duplicate Prevention**: Same contract PDFs are detected and storage is skipped
- 🔗 **Direct links**: Slack notifications include clickable PDF file links (existing or new)
- ⚠️ **Non-PDF files**: Other attachment types are noted but not saved to Drive

**Option 1: Automatic (Default)**
- System automatically creates a base folder named "契約書管理_Contract_Documents"
- Contract PDFs are saved directly to this folder with date-prefixed filenames
- No subfolders are created - all PDFs stored flat for easier access
- Intelligent duplicate detection prevents saving the same contract multiple times
- Folder ID is automatically saved to Script Properties

**Option 2: Manual Setup**
If you want to use an existing folder or specify a custom location:

1. **Create or locate your folder in Google Drive**
2. **Get the folder ID from the URL:**
   ```
   https://drive.google.com/drive/folders/1ABC123XYZ789def456
                                    ^^^^^^^^^^^^^^^^^^
                                    This is your folder ID
   ```
3. **Save the folder ID** - you'll set this in the next step along with other properties

**📁 Contract Storage Example:**
```
契約書管理_Contract_Documents/
├── 20250618_Employment_Agreement.pdf
├── 20250618_NDA_Document.pdf
├── 20250619_Service_Agreement.pdf
├── 20250620_Partnership_Contract.pdf
├── 20250621_ExampleCompany_Agreement.pdf
└── 20250622_NDA_Document_ContractTool.pdf
```

**Folder Permissions:**
- The folder will be accessible to the Google account running the script
- Shared folders are supported if you have edit permissions
- All PDFs are stored directly in the base folder with date-prefixed filenames

#### 3.4 Required Script Properties Configuration

Configure your contract management tool email addresses and Slack integration:

| Property Key | Description | Example Value | Required |
|--------------|-------------|---------------|----------|
| `SENDER_EMAIL_1` to `SENDER_EMAIL_20` | Contract tool notification emails | Various emails | ✅ At least one |
| Example: `SENDER_EMAIL_1` | Contract Tool 1 email | `noreply@contracttool1.example.com` | ⚪ Optional |
| Example: `SENDER_EMAIL_2` | Contract Tool 2 email | `noreply@contracttool2.example.com` | ⚪ Optional |
| Example: `SENDER_EMAIL_3` | Contract Tool 3 email | `noreply@contracttool3.example.com` | ⚪ Optional |
| Example: `SENDER_EMAIL_4` | Contract Tool 4 email | `noreply@contracttool4.example.com` | ⚪ Optional |
| Example: `SENDER_EMAIL_5` | Contract Tool 5 email | `noreply@contracttool5.example.com` | ⚪ Optional |
| Example: `SENDER_EMAIL_6` | Contract Tool 6 email | `noreply@contracttool6.example.com` | ⚪ Optional |
| Example: `SENDER_EMAIL_7` | Contract Tool 7 email | `noreply@contracttool7.example.com` | ⚪ Optional |
| Example: `SENDER_EMAIL_8` | Contract Tool 8 email | `noreply@contracttool8.example.com` | ⚪ Optional |
| `SLACK_WEBHOOK_URL` | Slack Incoming Webhook URL | `https://hooks.slack.com/services/...` | ✅ Yes |
| `SLACK_CHANNEL` | Slack channel for contract notifications | `#contracts` | ✅ Yes |
| `DRIVE_FOLDER_ID` | Google Drive folder ID for contracts | `1ABC123XYZ789def456` | ⚪ Auto-created |
| `SPREADSHEET_ID` | Google Spreadsheet ID for contract tracking | `1DEF456GHI789jkl012` | ⚪ Auto-created |

**📌 Flexible Email Configuration:**
- **Sender Emails**: Up to 20 contract tool emails (`SENDER_EMAIL_1` through `SENDER_EMAIL_20`)
- Only set the ones you need - system automatically detects configured emails
- No need to set emails in sequential order (e.g., can set 1, 3, and 7)
- Legacy `SENDER_EMAIL` property also supported for backward compatibility

**🔄 Smart Duplicate Prevention:**
- **Content-Based Detection**: Uses sender + date + subject + PDF filenames to detect duplicates
- **Per-Recipient Processing**: Each recipient receives Slack notifications and spreadsheet logging
- **PDF Storage Optimization**: Duplicate contract PDFs skip storage, use existing Drive URLs
- **Benefit**: Process same contract sent to multiple team members with intelligent PDF deduplication

#### 3.4 How to Set Script Properties

Choose one of these methods to configure the required properties:

**Method 1: Using Functions (Easy for Beginners)**
Execute this function in Google Apps Script editor:

```javascript
// Replace with your actual contract tool emails before executing!
function setContractConfiguration() {  // From: src/main.js
  // 1. Set up contract management tool emails (replace with your actual tool emails)
  setProperty('SENDER_EMAIL_1', 'noreply@contracttool1.example.com');
  setProperty('SENDER_EMAIL_2', 'noreply@contracttool2.example.com'); 
  setProperty('SENDER_EMAIL_3', 'noreply@contracttool3.example.com');
  // setProperty('SENDER_EMAIL_4', 'noreply@contracttool4.example.com');
  // setProperty('SENDER_EMAIL_5', 'noreply@contracttool5.example.com');
  // setProperty('SENDER_EMAIL_6', 'noreply@contracttool6.example.com'); 
  // setProperty('SENDER_EMAIL_7', 'noreply@contracttool7.example.com'); 
  // setProperty('SENDER_EMAIL_8', 'noreply@contracttool8.example.com');
  
  // 2. Send contract notifications to this Slack webhook
  setProperty('SLACK_WEBHOOK_URL', 'https://hooks.slack.com/services/YOUR/ACTUAL/WEBHOOK');
  
  // 3. Post contract notifications in this channel
  setProperty('SLACK_CHANNEL', '#contracts');
  
  // 4. Optional: Set custom Drive folder for contracts
  // setProperty('DRIVE_FOLDER_ID', '1ABC123XYZ789def456');
}
```

**Method 2: Manual Setup (Direct Property Setting)**
Go to Google Apps Script → Project Settings → Script Properties and manually add:
- Key: `SENDER_EMAIL_1`, Value: `noreply@contracttool1.example.com`
- Key: `SENDER_EMAIL_2`, Value: `noreply@contracttool2.example.com` (optional)
- Key: `SENDER_EMAIL_3`, Value: `noreply@contracttool3.example.com` (optional)
- Key: `SENDER_EMAIL_6`, Value: `noreply@contracttool6.example.com` (optional for Contract Tool 6)
- Key: `SLACK_WEBHOOK_URL`, Value: `https://hooks.slack.com/services/YOUR/WEBHOOK`
- Key: `SLACK_CHANNEL`, Value: `#contracts`
- Key: `DRIVE_FOLDER_ID`, Value: `1ABC123XYZ789def456` (auto-created if not set)

**Both methods achieve the same result** - choose whichever you find easier!

**Verification**
```javascript
// Check your configuration (safely displays without exposing secrets)
showConfiguration();  // From: src/main.js
```

Expected output:
```
=== Current Configuration ===
Sender Email: your-sender@example.com
Slack Channel: your-channel
Webhook URL: [SET]
Drive Folder ID: 1ABC123XYZ789def456
Spreadsheet ID: 1DEF456GHI789jkl012
Spreadsheet Logging: ENABLED

Pattern Settings:
Multiple patterns enabled: true
Total patterns: 4
```

**Note**: 
- Sender Email: Shows your configured email address
- Slack Channel: Shows your actual channel name (with or without # prefix)
- Webhook URL: Always shows "[SET]" to protect the secret URL
- Drive Folder ID: Shows the actual folder ID when auto-created
- Pattern Settings: Displays the current subject pattern configuration

#### 3.5 Environment File Reference (Optional - For Local Documentation Only)

The `.env.example` file is provided **only for documentation and local reference**. Google Apps Script doesn't read `.env` files - all configuration must be set in Script Properties.

**When you might use this:**
1. **Local documentation**: Keep track of your configuration values locally
2. **Team sharing**: Share configuration structure (without actual secrets) with team members
3. **Backup reference**: Remember what values you've configured

**Usage (Optional):**
```bash
# Copy the example file for your local reference
cp .env.example .env

# Edit .env with your actual values for documentation
# ⚠️ IMPORTANT: Never commit .env to git! (.gitignore already excludes it)
# ⚠️ The .env file is NOT read by Google Apps Script
# ⚠️ You still must set values in Script Properties using Methods 1 or 2 above
```

**Why we include this:**
- **Standard practice**: Most projects include `.env.example` for configuration documentation
- **Developer convenience**: Easier to track what configuration you've set
- **Future flexibility**: If you later build local development tools that integrate with this system

**Key point**: The `.env` file is purely for your convenience - Google Apps Script only reads from Script Properties!

### 4. Deploy
```bash
# Push code to Google Apps Script (if not done in step 2)
npm run push

# Verify deployment
npm run open
```

**🔧 Latest Updates:**
- **Message-level processing**: Same-subject emails are now handled correctly
- **PDF-only Drive folders**: Folders are created only when PDF attachments are found
- **Enhanced logging**: Detailed folder URLs and file paths in processing logs
- **Follow-up notifications**: Separate Slack messages with Drive folder links
- **🆕 Spreadsheet logging**: All emails are automatically logged to Google Spreadsheet
- **🆕 Email analytics**: Generate daily summaries and search historical email data

### 5. Initialize
Execute the following functions in Google Apps Script editor in order:

```javascript
// 1. Set up your configuration (REQUIRED FIRST!)
// Only run this if you chose Method 1 in step 3.3:
setContractConfiguration();  // From: src/main.js - Skip if you used Manual Setup (Method 2)

// 2. Verify configuration was set correctly
showConfiguration();  // From: src/main.js
// Expected: Shows your actual email and channel, [SET] for webhook

// 3. Test that all properties are accessible
testConfiguration();  // From: src/testRunner.js
// Expected: All validation checks pass

// 4. Test the complete system (SAFE - does not create production triggers)
runAllTests();  // From: src/testRunner.js
// Expected: 100% success rate on all tests

// 5. Enable automatic email monitoring (PRODUCTION SETUP)
setupInitialTrigger();  // From: src/triggerManager.js
// Expected: Trigger created, Slack notification of activation
```

**⚠️ Important Notes on Testing vs Production Setup:**

- **Testing Phase** (`runAllTests()`): Safe to run multiple times - does NOT create production triggers
- **Production Setup** (`setupInitialTrigger()`): Creates actual email monitoring triggers - only run when ready for production
- **Trigger Management**: 
  - Check triggers: `checkAndCleanupTriggers()` - shows current triggers and cleanup options
  - View triggers: `getTriggerInfo()` - returns trigger information
  - Clean up: `deleteExistingTriggers()` - removes all email monitoring triggers

**🔍 Troubleshooting Configuration:**
If any step fails, check:
1. All 3 Script Properties are set: `SENDER_EMAIL`, `SLACK_WEBHOOK_URL`, `SLACK_CHANNEL`
2. Webhook URL is valid and starts with `https://hooks.slack.com/`
3. Channel name is correct (can be with or without `#` prefix, e.g., `your-channel` or `#your-channel`)
4. Email address is the exact sender you want to monitor

**🆕 Testing PDF Processing:**
To test PDF attachment processing:
```javascript
// Test with actual emails (recommended)
testProcessEmails();  // From: src/testRunner.js
// or
processEmails();      // From: src/main.js

// Check Drive operations
testDriveOperations(); // From: src/driveManager.js

// Test Spreadsheet operations
testSpreadsheetOperations(); // From: src/spreadsheetManager.js
```

## 🆕 Contract Tracking Spreadsheet

The system automatically creates and maintains a specialized Google Spreadsheet for comprehensive contract tracking and analytics.

### Contract Spreadsheet Features

#### Automatic Setup
- **Auto-creation**: Spreadsheet named "契約管理_Contract_Tracking" is created automatically
- **Contract-specific schema**: 14 specialized columns for enhanced contract tracking
- **Intelligent data extraction**: Automatically categorizes contracts and extracts metadata

#### Contract Data Captured
Each contract email is logged with comprehensive information:

| Column | Field | Description |
|--------|-------|-------------|
| 受信日時 | Receipt Date | Contract completion timestamp |
| 契約管理ツール | Contract Tool | Tool used (Contract Tool 1, Contract Tool 2, etc.) |
| 送信者メール | Sender Email | Contract management tool email |
| 🆕 受信メール | Recipient Email | Email address that received the contract |
| 件名 | Subject | Original email subject |
| 契約タイプ | Contract Type | Auto-detected (Employment, NDA, Service Agreement, etc.) |
| 契約相手 | Contract Party | Extracted company/party name |
| PDFファイル名 | PDF Filename | Name of contract PDF file(s) |
| 🆕 PDF直接リンク | PDF Direct Links | Clickable direct links to PDF files |
| 処理状態 | Processing Status | Success/Error status |
| Slack通知済み | Slack Notified | Notification delivery status |
| 本文要約 | Body Summary | Email content summary |
| メッセージID | Message ID | Unique Gmail message identifier |
| エラーログ | Error Log | Processing error details |

#### Contract Analytics Functions

```javascript
// Generate contract summary report
generateContractSummary();  // From: src/spreadsheetManager.js
// Returns: Contract counts, success rates, tool breakdown, contract types

// Search for specific contracts
searchRecordByMessageId('message-id-here');  // From: src/spreadsheetManager.js

// Extract contract information
extractContractTool('noreply@contracttool1.example.com');  // From: src/spreadsheetManager.js - Returns: "Contract Tool 1"
extractContractType('Employment Agreement Completed', bodyText);  // From: src/spreadsheetManager.js - Returns: "Employment Agreement"
extractContractParty(subject, bodyText);  // From: src/spreadsheetManager.js - Returns: Company name

// Validate contract data integrity
validateSpreadsheetData();  // From: src/spreadsheetManager.js

// Export contract data to CSV
exportToCSV(startDate, endDate);  // From: src/spreadsheetManager.js
```

### Accessing Your Contract Spreadsheet

1. After first run, find your spreadsheet ID in the logs or configuration
2. Access directly: `https://docs.google.com/spreadsheets/d/YOUR-SPREADSHEET-ID`
3. Or find it in your Google Drive: "契約管理_Contract_Tracking"

### Contract Spreadsheet Benefits

- **📊 Analytics Dashboard**: Track contract completion trends by tool and type
- **🔍 Advanced Filtering**: Filter by date range, contract type, or management tool
- **📈 Performance Metrics**: Monitor processing success rates and error patterns
- **🏢 Company Tracking**: Identify frequent contract parties and partners
- **📋 Compliance Records**: Maintain audit trail of all contract processing

### Disabling Spreadsheet Logging

If you don't need spreadsheet logging:

```javascript
// In src/main.js CONFIG object:
ENABLE_SPREADSHEET_LOGGING: false,  // Disable spreadsheet logging
```

### 🆕 Upgrading Existing Spreadsheets

If you have an existing spreadsheet from a previous version, upgrade it to support the new features:

#### Automatic Upgrade
The system automatically detects and upgrades old spreadsheet formats when accessed.

#### Manual Upgrade
To force an immediate upgrade:

```javascript
// Force upgrade to latest format with recipient email and PDF direct links
forceUpgradeSpreadsheet();  // From: src/spreadsheetManager.js

// Alternative upgrade function (legacy compatibility)
upgradeToLatestFormat();    // From: src/spreadsheetManager.js
```

#### What Gets Upgraded
- **🆕 Recipient Email Column**: Track who received each contract
- **🆕 PDF Direct Links**: Click directly on PDF files instead of folder links
- **Enhanced Schema**: Updated column layout for better contract management
- **Backwards Compatibility**: Existing data remains intact

#### Verification
After upgrade, verify the new format:

```javascript
// Test the upgraded spreadsheet
testSpreadsheetOperations();  // From: src/spreadsheetManager.js

// Check spreadsheet configuration
showConfiguration();          // From: src/main.js
```

## Advanced Configuration

### 🆕 New Features Configuration

#### Email Content Display
```javascript
// In src/main.js CONFIG object:
SHOW_FULL_EMAIL_BODY: true,              // Show full email content (up to 7500 chars)
BODY_PREVIEW_LENGTH: 7500,               // Maximum email content length
SEND_DRIVE_FOLDER_NOTIFICATION: true,    // Send follow-up Drive folder links
```

#### Message-Level Duplicate Prevention
The system now tracks individual messages instead of email threads:
- ✅ **Same subject, new emails**: Processed correctly
- ✅ **Email thread replies**: Each message handled separately  
- ✅ **Reliable tracking**: Uses message IDs stored in Script Properties

#### Maintenance Functions
```javascript
// Clean up old processed message records (optional)
cleanupOldProcessedMessages(30); // From: src/main.js - Remove records older than 30 days
```

### Subject Pattern Customization
The system supports multiple regex patterns for complex Japanese email subjects:

```javascript
SUBJECT_PATTERNS: {
  ENABLE_MULTIPLE_PATTERNS: true,
  MATCH_MODE: 'any',  // 'any' or 'all'
  PATTERNS: [
    // Organization newsletter
    /\w+メルマガ[／\/].*お知らせ.*PR/,
    
    // Event notifications with brackets
    /【.*】第\d+回.*部会/,
    
    // Study sessions
    /勉強会.*『.*』.*※.*開/,
    
    // Re-send notifications
    /【再送】.*開催.*案内/,
    
    // Add your own patterns here
  ]
}

// 🆕 Enhanced processing: Only creates Drive folders for PDF attachments

// Legacy single pattern (still supported)
SUBJECT_PATTERN: /第\d+回.*部会|メルマガ|勉強会/
```

**Example subjects that will match:**
- `組織メルマガ／会員企業からのお知らせ＆PR`
- `【本日開催】第14回部会開催のご案内※6/5（木）15:00開催`
- `【再送】第14回部会開催のご案内※6/5（木）15:00開催`
- `勉強会『最新技術の動向』※5月27日(火)17:00開催`

### Changing Slack Notification Settings

#### Changing Notification Channel
To change which Slack channel receives notifications:

**Method 1: Using Google Apps Script Function**
```javascript
// Execute this function in Google Apps Script editor
function updateSlackChannel() {
  // Replace with your new channel name (with or without # prefix)
  setProperty('SLACK_CHANNEL', '#new-channel-name');
  console.log('Slack channel updated successfully');
}
```

**Method 2: Manual Script Properties Update**
1. Open Google Apps Script editor (`npm run open`)
2. Go to **Project Settings** → **Script Properties**
3. Find the `SLACK_CHANNEL` property
4. Update the value to your new channel (e.g., `#alerts`, `notifications`)
5. Save changes

**Method 3: Update and Re-run Configuration**
```javascript
// Update all settings including channel
function updateConfiguration() {
  setProperty('SENDER_EMAIL', 'your-sender@example.com');
  setProperty('SLACK_WEBHOOK_URL', 'your-webhook-url');
  setProperty('SLACK_CHANNEL', '#your-new-channel');  // ← Update this
  
  console.log('Configuration updated successfully');
}
```

**Verification:**
After updating the channel, verify the change:
```javascript
// Check current configuration
showConfiguration();  // From: src/main.js

// Send test notification to new channel
sendTestNotification();  // From: src/slackNotifier.js
```

#### Changing Webhook URL (Different Workspace/App)
If you need to change the entire webhook (different workspace or app):

1. **Create new Slack webhook:**
   - Go to your Slack workspace settings
   - Add "Incoming Webhooks" app (or use existing)
   - Select the desired channel
   - Copy the new webhook URL

2. **Update webhook in Script Properties:**
```javascript
function updateSlackWebhook() {
  const newWebhookUrl = 'https://hooks.slack.com/services/TXXXXXXXX/BXXXXXXXX/your-new-webhook-key';
  setProperty('SLACK_WEBHOOK_URL', newWebhookUrl);
  console.log('Slack webhook updated successfully');
}
```

3. **Test the new webhook:**
```javascript
// Verify webhook works with new channel
testSlackNotifications();  // From: src/testRunner.js
```

**Important Notes:**
- Channel names can be with or without `#` prefix (e.g., `alerts` or `#alerts`)
- Make sure the webhook has permission to post to the target channel
- Test changes with `sendTestNotification()` before processing real emails
- Changes take effect immediately - no code redeployment needed

### 🆕 Enhanced Slack Notifications

#### Dual Notification System
1. **Main Email Notification**: Complete email content with attachment summary
2. **Follow-up Drive Notification**: Dedicated message with clickable folder links (when PDFs are saved)

#### Notification Content Features
- 📝 **Full email content**: Up to 7,500 characters (configurable)
- 📎 **Smart attachment handling**: 
  - ✅ **PDF files**: Saved to organized Drive folders
  - ⚠️ **Other files**: Listed but marked as "skipped"
- 🔗 **Direct links**: Clickable file and folder URLs
- 📁 **Folder organization**: Date + subject naming

#### Attachment Display Format
```
📎 添付ファイル (2/3件 (1件スキップ))
• ✅ document.pdf (1.2MB) in 2025-06-18_組織メルマガ - 📄 File | 📁 Folder
• ⚠️ image.jpg (500KB) - スキップ: Not a PDF file
• ✅ report.pdf (800KB) in 2025-06-18_組織メルマガ - 📄 File | 📁 Folder
```

### Notification Customization
Slack message format can be adjusted in `src/slackNotifier.js`:
- Modify email content length (`formatEmailBody` function)
- Change attachment display format (`buildAttachmentText` function)
- Customize follow-up notification content (`sendDriveFolderNotification` function)
- Adjust color rules and message structure

## Email Processing Logic & Architecture

### 🆕 Enhanced Monitoring Strategy
The system now uses **message-level processing** to handle same-subject emails correctly:

#### 1. **Smart Gmail Search**
- **Method**: `from:{SENDER_EMAIL}` (checks ALL threads, not just unprocessed)
- **Enhancement**: Individual message tracking prevents skipping new emails in existing threads
- **Data Retrieved**: All email threads from specified sender

#### 2. **Message-Level Duplicate Prevention**
- **Purpose**: Track each individual message, not just threads
- **Method**: Store processed message IDs in Script Properties
- **Advantage**: Same-subject emails in existing threads are processed correctly

#### 3. **Advanced Subject Pattern Matching**
- **Purpose**: Precise identification of target emails using multiple regex patterns
- **Support**: Complex Japanese email formats and organizational patterns
- **Flexibility**: Easy addition of new patterns

#### 4. **Intelligent PDF Processing**
- **Detection**: Automatically identifies PDF attachments
- **Selective Processing**: Creates Drive folders ONLY for emails with PDFs
- **Organization**: Date + subject-based folder structure

### 🆕 Enhanced Processing Logic
New message-level processing flow:

```
Improved Email Processing:
├── Search ALL threads from sender
├── For each thread:
│   ├── Check each message individually
│   ├── Message ID already processed? ──┐
│   │                                    ├── Yes → Skip this message
│   │                                    └── No → Continue processing
│   ├── Subject pattern match? ──┐
│   │                             ├── Yes → Process message
│   │                             └── No → Skip message
│   ├── PDF attachments found? ──┐
│   │                              ├── Yes → Create Drive folder & save
│   │                              └── No → Mark as "skipped"
│   ├── Send main Slack notification
│   ├── Send follow-up Drive folder notification (if PDFs saved)
│   └── Mark message ID as processed
└── Apply thread label (if any messages were processed)
```

### Performance Optimizations
- **Batch Processing**: Process up to 10 emails at once
- **Error Isolation**: One email failure doesn't affect the entire system
- **Execution Time Monitoring**: Reliable completion within 6-minute limit
- **Duplicate Avoidance**: Ensure unique filenames in Drive

## Slack Notification System

### Notification Types & Triggers

#### 📧 **Email Notifications**
- **Condition**: When pattern-matched email is received
- **Content**: Sender, subject, body excerpt, attachment information
- **Color Coding**: Success=green, Attachment errors=yellow

#### 🚨 **Error Notifications**  
- **Condition**: When system errors occur
- **Content**: Error details, occurrence time
- **Color Coding**: Red (danger)

#### 📊 **Processing Summary**
- **Condition**: When processing completes with errors
- **Content**: Success/failure counts, execution time, success rate

#### ⚡ **Trigger Status Notifications**
- **Condition**: When triggers are created/deleted
- **Content**: Monitoring status, check interval

## Development & Maintenance

### Development Commands
```bash
# Development cycle
npm run push          # Deploy local changes to GAS
npm run pull          # Sync GAS changes to local
npm run logs          # View execution logs (for debugging)
npm run open          # Open GAS editor

# Authentication
npm run login         # Clasp CLI authentication
```

### File Structure
```
contract-management-processor/
├── package.json           # npm configuration and Clasp commands
├── appsscript.json       # GAS configuration (timezone, APIs)
├── CLAUDE.md            # Claude Code documentation
└── src/
    ├── main.js           # Contract tool configuration and entry point
    ├── emailProcessor.js # Contract email processing logic
    ├── driveManager.js   # Contract PDF storage management
    ├── slackNotifier.js  # Contract-specific Slack notifications
    ├── spreadsheetManager.js # Contract tracking spreadsheet
    ├── triggerManager.js # Automated contract monitoring
    └── testRunner.js     # Contract management test suite
```

### 🆕 Enhanced Key Functions

#### `main.js`
- `processEmails()`: Main processing entry point with message-level tracking
- `validateConfiguration()`: Configuration validation
- `getOrCreateDriveFolder()`: Drive folder management
- 🆕 `markMessageAsProcessed()`: Store processed message IDs
- 🆕 `cleanupOldProcessedMessages()`: Maintenance function for old records
- 🆕 `showSkippedEmailStats()` or `checkSkipped()`: Display statistics for skipped emails
- 🆕 `cleanupOldSkipLabels()`: Remove skip labels from old emails

#### `emailProcessor.js`
- `processMessage()`: Individual email processing with duplicate prevention and recipient tracking
- 🆕 `isMessageAlreadyProcessed()`: Message-level duplicate checking
- 🆕 `formatEmailBody()`: Smart email content formatting (up to 7500 chars)
- 🆕 `getMessageRecipient()`: Extract recipient email from message To field
- 🆕 `extractEmailAddresses()`: Parse multiple email addresses from string
- 🆕 **Auto-skip feature**: Adds `Contract_Skipped` label to non-matching emails

#### `driveManager.js`
- 🆕 `processAttachments()`: PDF-only contract processing
- `saveAttachmentToDrive()`: Execute Drive saving with flat folder structure
- 🆕 `createContractFolder()`: YYYYMMDD_Subject folder creation
- 🆕 `cleanSubjectForFolder()`: Safe folder name generation

#### `slackNotifier.js`
- `sendSlackNotification()`: Enhanced email notifications with full content
- 🆕 `sendDriveFolderNotification()`: Follow-up Drive folder notifications
- `sendErrorNotification()`: Error notifications
- 🆕 `buildAttachmentText()`: Smart attachment display (saved/skipped/failed)

#### `spreadsheetManager.js`
- 🆕 `createOrGetSpreadsheet()`: Initialize contract tracking spreadsheet with 14 columns
- 🆕 `addEmailRecord()`: Log new contract email with recipient and PDF direct links
- 🆕 `updateRecordStatus()`: Update contract processing status
- 🆕 `searchRecordByMessageId()`: Find existing contract records
- 🆕 `generateContractSummary()`: Create contract completion statistics
- 🆕 `extractContractTool()`: Identify contract management tool
- 🆕 `extractContractType()`: Auto-detect contract type
- 🆕 `extractContractParty()`: Extract contract party/company
- 🆕 `forceUpgradeSpreadsheet()`: Force upgrade existing spreadsheets to latest format
- 🆕 `upgradeToLatestFormat()`: Upgrade spreadsheet schema with recipient and PDF links

#### `triggerManager.js`
- `createTrigger()`: Create periodic execution trigger
- `checkTriggerHealth()`: Trigger health check

#### `testRunner.js`
- `runAllTests()`: Comprehensive contract management test suite
- `testProcessEmails()`: Test actual contract email processing
- `testDriveOperations()`: Test contract PDF storage operations
- `testSlackNotifications()`: Test contract notifications
- 🆕 `testSpreadsheetOperations()`: Test contract tracking functions

## Troubleshooting

### Common Issues & Solutions

#### 🔍 **Contract Tool Emails Not Being Processed**

**Symptom**: Emails from specific contract management tools (like HelloSign/Dropbox Sign) are not being processed

**Debug Steps**:
```javascript
// 1. Check if sender email is configured
showConfiguration();  // From: src/main.js

// 2. Debug specific contract tool (e.g., HelloSign)
debugHelloSignEmails();  // From: src/main.js
// This checks:
// - Sender email configuration
// - Gmail search results  
// - Subject pattern matching

// 3. Manually test email processing
processEmails();  // From: src/main.js
```

**Common Solutions**:
- **Missing sender email**: Add the exact sender email to Script Properties (e.g., `SENDER_EMAIL_7 = noreply@mail.hellosign.com`)
- **Pattern mismatch**: Check if email subjects match configured patterns
- **Already processed**: Emails with `Contract_Processed` label are skipped

#### 🔧 **Script Properties Limit (50 Properties)**
**Symptom**: "Your script has more than 50 properties" or property setting fails

**Immediate Solutions:**
```javascript
// Option 1: Quick view and cleanup
checkProperties();         // View current usage
cleanupOld(7);             // Clean messages older than 7 days
fixProperties();           // Emergency cleanup

// Option 2: Manual cleanup
cleanupProcessedMessages(3, false);  // Clean messages older than 3 days
deletePropertiesByPattern('PROCESSED_MSG_', false);  // Delete all processed messages
```

**Automatic Prevention:**
- System now auto-cleans when approaching limit
- `setProperty()` includes automatic cleanup
- Regular cleanup recommended: `cleanupOld(7)` weekly

**Property Categories:**
- Configuration: 4-6 properties (permanent)
- Sender Emails: Up to 20 properties (as needed)
- Recipient Emails: Up to 20 properties (optional)
- Processed Messages: Grows over time (clean regularly)

#### 🔧 **Project Setup Issues**
**Symptom**: `clasp push` fails with "Project settings not found"
```bash
# Solution: Create or recreate the GAS project
rm .clasp.json  # Remove existing file
clasp create --type standalone --title "Contract Management Email Processor"
cat .clasp.json  # Verify scriptId is set
npm run push
```

#### 🔧 **Contract Emails Not Being Processed**
**Symptom**: New contract emails exist but no Slack notifications
```javascript
// Debug steps
1. Run testEmailSearch() → Check contract email search query
2. Run testMessageProcessing() → Check contract subject patterns
3. Run quickHealthCheck() → Check overall configuration
```

#### 🔧 **Slack Notifications Not Received**
**Symptom**: System operates but no Slack notifications
```javascript
// Resolution steps
1. Run testSlackNotifications()
2. Check Webhook URL: PropertiesService.getScriptProperties().getProperty('SLACK_WEBHOOK_URL')
3. Check channel permissions
```

#### 🔧 **Attachments Not Saved**
**Symptom**: Email notifications received but no files in Drive
```javascript
// Check items
1. Run testDriveOperations()
2. Check Drive API enablement
3. Check folder permissions
```

#### 🔧 **Same-Subject Emails Not Processing**
**Symptom**: New emails with same subject are skipped
🆕 **SOLVED**: This issue has been fixed with message-level tracking!
```javascript
// Verification: Check that message-level processing is working
1. Run testProcessEmails() with same-subject emails
2. Check logs for "Message already processed" vs "Processing new message"
3. Verify Script Properties contain PROCESSED_MSG_ entries
```

#### 🔧 **No Drive Folders Created**
**Symptom**: Emails processed but no Drive folders
🆕 **Expected Behavior**: Folders only created for PDF attachments
```javascript
// Check if this is expected:
1. Verify emails actually contain PDF attachments
2. Check logs for "Found X PDF files" vs "No PDF files found"
3. Non-PDF attachments will show "skipped: Not a PDF file" in Slack
```

#### 🔧 **Spreadsheet Not Created**
**Symptom**: No spreadsheet appears in Google Drive
🆕 **Resolution**:
```javascript
// Check if spreadsheet logging is enabled
console.log(CONFIG.ENABLE_SPREADSHEET_LOGGING); // Should be true

// Manually create spreadsheet
createOrGetSpreadsheet(); // From: src/spreadsheetManager.js

// Check spreadsheet ID
const id = getProperty('SPREADSHEET_ID');
console.log(`Spreadsheet ID: ${id}`);
```

#### 🔧 **Spreadsheet Data Missing**
**Symptom**: Emails processed but not appearing in spreadsheet
```javascript
// Debug steps
1. Run validateSpreadsheetData() → Check data integrity
2. Check logs for "Email record added to spreadsheet"
3. Verify spreadsheet permissions
4. Run testSpreadsheetOperations() → Test basic operations
```

#### 🔧 **Duplicate Processing**
**Symptom**: Same email processed multiple times
```javascript
// Resolution
1. Run checkTriggerHealth() → Remove duplicate triggers
2. Run cleanupOldProcessedMessages() → Clean message tracking
3. Recreate trigger: setupInitialTrigger()
```

### 🆕 Enhanced Debugging Steps
1. **Log Check**: Use `npm run logs` to check latest execution logs
   - Look for "Checked: X messages, Processed: Y" statistics
   - Check for PDF detection: "Found X PDF files" or "No PDF files found"
   - Verify message tracking: "Message already processed" vs "Processing new message"

2. **Individual Tests**: Run test functions for each module
   - `testProcessEmails()`: Test with actual emails (recommended for PDF testing)
   - `testDriveOperations()`: Test folder creation and organization
   - `testSlackNotifications()`: Test both main and follow-up notifications

3. **Configuration Validation**: Use `testConfiguration()` to check basic settings

4. **Message Tracking Check**:
   ```javascript
   // Check processed message tracking
   const properties = PropertiesService.getScriptProperties().getProperties();
   Object.keys(properties).filter(key => key.startsWith('PROCESSED_MSG_')).length;
   ```

5. **Drive Folder Verification**:
   ```javascript
   // Check Drive folder structure
   getDriveFolderStats(); // From: src/driveManager.js
   ```

### Performance Monitoring
**Expected Values**:
- Processing time: 5-15 seconds per email
- Success rate: 95%+
- Trigger execution: Stable operation at 5-minute intervals

**🆕 Enhanced Monitoring Points**:
```javascript
// Check execution statistics
getDriveFolderStats()                    // File saving status
getTriggerInfo()                         // Trigger operation status
cleanupOldProcessedMessages(30)          // Maintenance: cleanup old records

// Check message tracking
const props = PropertiesService.getScriptProperties().getProperties();
const processedCount = Object.keys(props).filter(k => k.startsWith('PROCESSED_MSG_')).length;
console.log(`Currently tracking ${processedCount} processed messages`);
```

**Expected Log Output**:
```
Checked: 25 messages, Processed: 3, Errors: 0, Time: 4250ms
Found 2 PDF files, processing attachments...
Created/found email folder: 2025-06-18_組織メルマガ
Successfully saved: document.pdf
File URL: https://drive.google.com/file/d/...
Sending follow-up message for 2 saved PDFs...
```

## API Integration Details

### Gmail API Usage Scope
- **Search**: `GmailApp.search()` - Email search
- **Messages**: `message.getSubject()`, `getAttachments()` - Content retrieval
- **Labels**: `GmailApp.createLabel()` - Processed management

### Drive API Usage Scope  
- **Folders**: `DriveApp.createFolder()`, `getFolderById()` - Folder management
- **Files**: `folder.createFile()` - Attachment saving
- **Metadata**: `file.getUrl()`, `getId()` - Link generation

### Slack Webhook Specifications
- **Endpoint**: `https://hooks.slack.com/services/...`
- **Method**: POST
- **Format**: JSON
- **Limits**: 1 request per second recommended

## Security Considerations

### 🔒 Critical Security Features

#### Configuration Security
- **Script Properties Only**: All sensitive data stored in Google Apps Script Properties, never in code
- **Zero Hardcoding**: No webhook URLs, emails, or secrets in source code
- **Public Repository Safe**: Code can be safely published to public GitHub repositories
- **Environment Separation**: Use `.env.example` for documentation, never commit actual `.env` files

#### Access Control
- **Minimum Permissions**: Only essential Google API scopes enabled
- **Drive Isolation**: Access restricted to designated attachment folder only
- **Channel Restrictions**: Slack notifications only to configured channels
- **Email Filtering**: Only processes emails from specified senders

#### Data Protection
- **Log Sanitization**: Sensitive information automatically redacted from logs
- **Secure Property Access**: Configuration helper functions prevent accidental exposure
- **Error Handling**: Error messages don't leak sensitive configuration details
- **Webhook Protection**: URLs never logged or displayed in plain text

#### Repository Security Guidelines
- **`.gitignore`**: Comprehensive exclusion of sensitive files
- **Example Files**: Use `.example` suffix for template configurations
- **Documentation**: Clear instructions for secure setup without exposing secrets
- **Code Reviews**: Security-first approach in all development

### File Access Control
- **Drive Permissions**: Access only specified folders
- **File Sharing**: Default private settings
- **Naming Rules**: Generate unpredictable filenames

### Execution Environment Security
- **OAuth Authentication**: Use Google standard authentication
- **Execution Logs**: Automatic deletion after 30 days
- **Configuration Validation**: Secure validation without exposing values

## License & Support

**License**: MIT License

## Recent Updates

### 🆕 Version 2.6 - Intelligent Skip Labels & Enhanced Docusign Support
- **Smart Skip Labels**: Automatically adds `Contract_Skipped` label to non-matching emails
  - Prevents repeated checking of irrelevant emails
  - Reduces processing overhead and improves performance
  - Includes emails without PDF attachments (for Docusign)
- **Skip Email Management**:
  - `showSkippedEmailStats()` or `checkSkipped()` - View statistics of skipped emails
  - `cleanupOldSkipLabels(days, dryRun)` - Clean up old skip labels
  - Review skipped emails in Gmail with search: `label:Contract_Skipped`
- **Enhanced Search Query**: Now excludes both `Contract_Processed` and `Contract_Skipped` labels
- **Improved Docusign Detection**: Better handling of emails without PDF attachments

### 🆕 Version 2.5 - HelloSign Support & Testing Improvements
- **HelloSign/Dropbox Sign Integration**: Full support for HelloSign email processing
  - Fixed Gmail search query to exclude processed emails (`-label:Contract_Processed`)
  - Added dedicated debugging function: `debugHelloSignEmails()`
  - Enhanced pattern matching for "You've been copied on" notifications
- **Safe Testing Framework**: Improved test safety and debugging
  - `runAllTests()` no longer creates production triggers accidentally
  - Added `testTriggerOperations()` - safe trigger testing without side effects
  - New cleanup functions: `checkAndCleanupTriggers()` for trigger management
- **Duplicate PDF Handling**: Enhanced duplicate display
  - Shows "Duplicate" text instead of folder links for duplicate PDFs
  - Cleaner spreadsheet display with direct PDF access for unique files

### 🆕 Version 2.4 - Enhanced Contract Tracking & PDF Direct Links
- **Recipient Email Tracking**: New column to track who received each contract notification
  - Automatically extracts recipient from email To field
  - Enhanced duplicate prevention with recipient-aware processing
  - Better visibility into contract distribution
- **PDF Direct Links**: Revolutionary improvement to PDF access
  - Replace folder URLs with direct clickable PDF file links
  - One-click access to specific contract documents
  - Multiple PDF support with individual file links
  - Format: Direct URL only for easy clicking: `https://drive.google.com/file/d/FILE_ID/view`
- **Enhanced Spreadsheet Schema**: Upgraded from 13 to 14 columns
  - New "受信メール" (Recipient Email) column
  - Updated "PDF直接リンク" (PDF Direct Links) column
  - Automatic upgrade of existing spreadsheets
- **Upgrade Functions**: 
  - `forceUpgradeSpreadsheet()` - Manual spreadsheet format upgrade
  - `upgradeToLatestFormat()` - Automatic format detection and upgrade
  - Backwards compatibility maintained for existing data
- **Enhanced Testing**: Updated test functions with new data format
  - `testSpreadsheetOperations()` now tests recipient and PDF link features
  - Comprehensive upgrade testing included

### 🆕 Version 2.3 - Property Management & Cleanup
- **Automatic Property Management**: Handles GAS 50-property limit automatically
  - Auto-cleanup when approaching limit during property setting
  - Comprehensive property viewer with categorization
  - Emergency cleanup functions for when limit is reached
- **Property Management Functions**: 
  - `checkProperties()` - View all properties with usage stats
  - `cleanupOld(days)` - Clean up processed messages older than X days
  - `fixProperties()` - Emergency cleanup when at 50-property limit
- **Smart Property Setting**: `setProperty()` now includes automatic cleanup
- **New Module**: `src/propertyManager.js` for advanced property management

### 🆕 Version 2.2 - Recipient Filtering & Duplicate Prevention
- **Recipient Email Filtering**: Prevent duplicate processing when same contract is sent to multiple recipients
  - Support for up to 20 recipient emails (RECIPIENT_EMAIL_1 to RECIPIENT_EMAIL_20)
  - Smart duplicate prevention using message ID + recipient combination
  - Optional configuration - works with or without recipient filtering
- **Enhanced Tracking**: Message processing tracks specific recipient combinations
- **Improved Search Query**: Automatically filters by target recipients when configured
- **New Test Suite**: `runRecipientFilteringTests()` for testing duplicate prevention

### 🆕 Version 2.1 - Flexible Email Configuration
- **Flexible Sender Emails**: Support for up to 20 contract tool emails (SENDER_EMAIL_1 to SENDER_EMAIL_20)
  - No hardcoded emails - all configuration via Script Properties
  - Automatic detection of configured emails - no sequential requirement
  - Dynamic email loading with `getConfiguredSenderEmails()` function
- **Improved Configuration Display**: `showConfiguration()` now shows all configured emails dynamically
- **Better Error Handling**: Optional email properties don't cause errors

### 🆕 Version 2.0 - Direct PDF Storage & Contract Tool 6 Support
- **Contract Tool 6 Integration**: Added support for Contract Tool 6 (Japanese contract management tool)
  - Email patterns for "合意締結が完了しました" notifications
  - Automatic extraction of contract parties from Contract Tool 6 email subjects
  - Support for Japanese contract types (覚書, 機密保持契約書, 個別契約)
- **Simplified Storage**: PDFs now saved directly to base folder
  - No more subfolder creation - all PDFs stored flat
  - Filenames use `YYYYMMDD_original_filename.pdf` format
  - Easier to browse and search contract PDFs
- **Enhanced Functions**: 
  - `setContractConfiguration()` (src/main.js) - now includes Contract Tool 6 setup
  - Improved `showConfiguration()` (src/main.js) - displays all contract tools
  - New test suite: `runContractTool6Tests()` (test/testCloudSign.js)
  - Direct storage tests: `runDirectStorageTests()` (test/testDirectStorage.js)

**Support**: 
- Setup assistance: Follow detailed instructions in README.md
- Bug reports: Report in GitHub repository Issues
- Feature requests: Contributions welcome via Pull Requests

**Contribution Guidelines**:
1. Development following Clasp CLI workflow
2. Implementation of test functions for each feature
3. Comprehensive logging and error handling
4. Configuration value management via Script Properties