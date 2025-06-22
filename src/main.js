/**
 * Contract Management Email Processor - Main Configuration and Entry Point
 * 
 * This file contains the main configuration and entry point for processing contract-related emails.
 * It monitors multiple contract management tools, processes PDF contracts, and logs to Slack and Spreadsheet.
 */

// === CONFIGURATION ===
const CONFIG = {
  // Contract management system monitoring settings
  // IMPORTANT: Set actual sender emails in Script Properties, not here!
  // Dynamically loaded when accessed via getSenderEmails() function
  
  // Legacy support for single sender email (for backward compatibility)
  get SENDER_EMAIL() {
    return getProperty('SENDER_EMAIL', false) || getProperty('SENDER_EMAIL_1', false);
  },
  
  // Get all configured sender emails dynamically
  get SENDER_EMAILS() {
    return getConfiguredSenderEmails();
  },
  
  // Remove recipient filtering - now using content-based duplicate detection
  // get RECIPIENT_EMAILS() {
  //   return getConfiguredRecipientEmails();
  // },
  
  // Contract email subject pattern matching configuration
  // 契約関連メール件名パターンマッチング設定
  SUBJECT_PATTERNS: {
    // Enable multiple pattern matching
    ENABLE_MULTIPLE_PATTERNS: true,
    
    // Pattern matching mode: 'any' (match any pattern) or 'all' (match all patterns)
    MATCH_MODE: 'any',
    
    // Define patterns for contract management tools
    PATTERNS: [
      // DocuSign patterns
      /.*completed.*document/i,
      /.*signed.*agreement/i,
      /.*contract.*executed/i,
      /.*署名.*完了/,
      /.*契約.*締結/,
      
      // Adobe Sign patterns
      /.*agreement.*signed/i,
      /.*document.*completed/i,
      /.*合意.*完了/,
      
      // HelloSign patterns
      /.*signature.*complete/i,
      /.*document.*finalized/i,
      
      // PandaDoc patterns
      /.*document.*signed/i,
      /.*contract.*completed/i,
      
      // ContractWorks patterns
      /.*contract.*approved/i,
      /.*agreement.*final/i,
      
      // CloudSign patterns
      /.*の合意締結が完了しました$/,
      /.*合意締結.*完了/,
      /.*捺印済.*合意締結/,
      
      // Contract Tool 7 patterns
      /You've been copied on.*signed by/i,
      /.*has been completed/i,
      /.*signature.*completed/i,
      /.*agreement.*signed by.*and/i,
      
      // Generic contract patterns
      /contract.*pdf/i,
      /agreement.*attached/i,
      /.*契約書.*添付/,
      /.*合意書.*完了/,
      /.*覚書.*締結/
    ]
  },
  
  // Legacy single pattern support (for backward compatibility)
  SUBJECT_PATTERN: /completed.*document|signed.*agreement|contract.*executed|署名.*完了|契約.*締結/i,
  
  GMAIL_LABEL: 'Contract_Processed',  // 処理済み契約メールのラベル名
  
  // Slack integration settings
  // IMPORTANT: Set actual Slack channel in Script Properties, not here!
  SLACK_CHANNEL: getProperty('SLACK_CHANNEL') || '#general',  // 通知先Slackチャンネル
  
  // Google Drive settings for contract storage
  DRIVE_FOLDER_NAME: '契約書管理_Contract_Documents',  // 契約書PDF保存フォルダ名
  
  // Contract tracking spreadsheet settings
  ENABLE_SPREADSHEET_LOGGING: true,  // true: 契約情報のスプレッドシート記録を有効化
  CONTRACT_TRACKING_MODE: true,  // true: 契約管理特化モード
  
  // Processing settings
  MAX_EMAILS_PER_RUN: 10,  // 一回の実行で処理する最大メール数
  BODY_PREVIEW_LENGTH: 7500,  // Slackに表示する本文の最大文字数（Slack制限: 8000文字）
  SHOW_FULL_EMAIL_BODY: true,  // true: 全文表示（制限内）, false: 短縮表示
  SEND_DRIVE_FOLDER_NOTIFICATION: true,  // true: PDF保存後にDriveフォルダリンクをフォローアップ送信
  
  // Property management settings (legacy - now using spreadsheet tracking)
  // EMAIL_SEARCH_DAYS: 7,  // Removed - no longer using time-based filtering
  // MAX_PROCESSED_MESSAGES: 50,  // Removed - now using spreadsheet for unlimited tracking
  
  // Trigger settings
  TRIGGER_INTERVAL_MINUTES: 5  // トリガーの実行間隔（分）
};

// === SCRIPT PROPERTIES KEYS ===
// All sensitive configuration should be stored in Script Properties, not in code!
const PROPERTY_KEYS = {
  // Required properties (must be set in Script Properties)
  SLACK_WEBHOOK_URL: 'SLACK_WEBHOOK_URL',
  SLACK_CHANNEL: 'SLACK_CHANNEL',
  
  // Contract management system sender emails (set as many as needed)
  SENDER_EMAIL: 'SENDER_EMAIL',  // Legacy support
  SENDER_EMAIL_1: 'SENDER_EMAIL_1',  // DocuSign
  SENDER_EMAIL_2: 'SENDER_EMAIL_2',  // Adobe Sign
  SENDER_EMAIL_3: 'SENDER_EMAIL_3',  // HelloSign
  SENDER_EMAIL_4: 'SENDER_EMAIL_4',  // ContractWorks
  SENDER_EMAIL_5: 'SENDER_EMAIL_5',  // PandaDoc
  
  // Optional properties (auto-generated if not set)
  DRIVE_FOLDER_ID: 'DRIVE_FOLDER_ID',
  SPREADSHEET_ID: 'SPREADSHEET_ID'
};

/**
 * Main processing function - called by time-based trigger
 * メイン処理関数 - 時間ベーストリガーから呼び出される
 */
function processEmails() {
  const startTime = new Date().getTime();
  
  try {
    console.log('=== Contract Management Email Processor Starting ===');
    console.log(`Configuration: Senders=${CONFIG.SENDER_EMAILS.length}, Pattern Mode=${CONFIG.SUBJECT_PATTERNS.MATCH_MODE}`);
    console.log(`Contract tools monitored: ${CONFIG.SENDER_EMAILS.join(', ')}`);
    
    // Check required properties
    validateConfiguration();
    
    // Build search query for multiple contract management tool senders
    // 複数の契約管理ツールからのメールを検索
    const senderQueries = CONFIG.SENDER_EMAILS.map(email => `from:${email}`);
    let query = `(${senderQueries.join(' OR ')}) -label:${CONFIG.GMAIL_LABEL}`;
    
    // Remove recipient filtering - now using content-based duplicate detection
    console.log('Using content-based duplicate detection (no recipient filtering)');
    
    // Remove time-based filtering to avoid missing emails on initial startup
    // Duplicate tracking will be handled via spreadsheet instead of Script Properties
    
    console.log(`Search query: ${query}`);
    const threads = GmailApp.search(query, 0, CONFIG.MAX_EMAILS_PER_RUN * 3); // Get more threads to check individual messages
    
    console.log(`Found ${threads.length} email threads from contract management tools`);
    
    if (threads.length === 0) {
      console.log('No contract emails to process');
      return;
    }
    
    let processedCount = 0;
    let errorCount = 0;
    let checkedCount = 0;
    
    // Process each thread and check individual messages
    threads.forEach((thread, index) => {
      try {
        console.log(`Checking thread ${index + 1}/${threads.length}: ${thread.getFirstMessageSubject()}`);
        
        // Check all messages in thread for unprocessed ones
        // スレッド内の全メッセージをチェックして未処理のものを検索
        const messages = thread.getMessages();
        let threadHasNewMessages = false;
        
        messages.forEach((message, msgIndex) => {
          try {
            checkedCount++;
            
            if (message.isInTrash()) {
              console.log(`  Message ${msgIndex + 1} is in trash, skipping`);
              return;
            }
            
            // Check if this specific message was already processed
            if (isMessageAlreadyProcessed(message)) {
              console.log(`  Message ${msgIndex + 1} already processed, skipping`);
              return;
            }
            
            console.log(`  Processing new message ${msgIndex + 1}: ${message.getSubject()}`);
            
            if (processMessage(message)) {
              processedCount++;
              threadHasNewMessages = true;
              
              // Mark this specific message as processed
              markMessageAsProcessed(message);
            }
            
          } catch (msgError) {
            console.error(`Error processing message ${msgIndex + 1}:`, msgError);
            errorCount++;
          }
        });
        
        // Add thread label only if we processed new messages
        if (threadHasNewMessages) {
          addProcessedLabel(thread);
        }
        
      } catch (error) {
        console.error(`Error processing thread ${index + 1}:`, error);
        errorCount++;
      }
    });
    
    const endTime = new Date().getTime();
    const executionTime = endTime - startTime;
    
    console.log(`=== Processing Complete ===`);
    console.log(`Checked: ${checkedCount} messages, Processed: ${processedCount}, Errors: ${errorCount}, Time: ${executionTime}ms`);
    
    // Send summary notification if there were errors
    if (errorCount > 0) {
      sendErrorSummary(processedCount, errorCount, executionTime);
    }
    
  } catch (error) {
    console.error('Critical error in processEmails:', error);
    sendErrorNotification(`Critical error in contract processor: ${error.message}`);
    throw error;
  }
}

/**
 * Validate that all required configuration is present
 * 必要な設定がすべて存在することを確認
 */
function validateConfiguration() {
  console.log('Validating configuration...');
  
  // Check at least one sender email is configured
  const senderEmails = CONFIG.SENDER_EMAILS;
  if (senderEmails.length === 0) {
    throw new Error('No sender emails configured. Please set at least one SENDER_EMAIL_* property.');
  }
  console.log(`Found ${senderEmails.length} configured sender email(s)`);
  
  // Check Slack webhook URL
  const webhookUrl = getProperty(PROPERTY_KEYS.SLACK_WEBHOOK_URL);
  if (!webhookUrl || !webhookUrl.startsWith('https://hooks.slack.com/')) {
    throw new Error('Invalid Slack webhook URL in script properties');
  }
  
  // Check or create Drive folder
  const folderId = getOrCreateDriveFolder();
  console.log(`Drive folder ID: ${folderId}`);
  
  // Check or create Spreadsheet if logging is enabled
  if (CONFIG.ENABLE_SPREADSHEET_LOGGING) {
    const spreadsheetId = createOrGetSpreadsheet();
    console.log(`Spreadsheet ID: ${spreadsheetId}`);
  }
  
  console.log('Configuration validation complete');
}

/**
 * Get property from Script Properties with error handling
 * スクリプトプロパティから値を取得（エラーハンドリング付き）
 */
function getProperty(key, required = true) {
  try {
    const value = PropertiesService.getScriptProperties().getProperty(key);
    if (!value && required) {
      throw new Error(`Required property ${key} not found in script properties. Please set it using setProperty() or the setup functions.`);
    }
    return value;
  } catch (error) {
    if (required) {
      console.error(`Error getting required property ${key}:`, error);
      throw error;
    }
    return null;
  }
}

/**
 * Set property in Script Properties with intelligent property management
 * スクリプトプロパティに値を設定（インテリジェントプロパティ管理付き）
 */
function setProperty(key, value) {
  try {
    const allProperties = PropertiesService.getScriptProperties().getProperties();
    const currentCount = Object.keys(allProperties).length;
    const exists = allProperties.hasOwnProperty(key);
    
    // If we're near the absolute limit and this is a new property, attempt cleanup
    if (currentCount >= 48 && !exists) {
      console.log(`⚠️  Near property limit (${currentCount}/50). Attempting automatic cleanup...`);
      
      try {
        // For processed messages, use rotation instead of age-based cleanup
        if (key.startsWith('PROCESSED_MSG_')) {
          const rotated = rotateProcessedMessages(5); // Remove 5 oldest entries
          console.log(`Rotated ${rotated} processed message entries`);
        } else {
          // For other properties, use age-based cleanup
          const cleaned = cleanupOldProcessedMessages(3); // Clean messages older than 3 days
          console.log(`Cleaned up ${cleaned} old entries`);
        }
      } catch (cleanupError) {
        console.warn('Auto-cleanup failed, but continuing with property set:', cleanupError);
      }
    }
    
    PropertiesService.getScriptProperties().setProperty(key, value);
    
    // Only log for non-processed message properties to reduce noise
    if (!key.startsWith('PROCESSED_MSG_')) {
      console.log(`Property ${key} set successfully`);
    }
    
  } catch (error) {
    console.error(`Error setting property ${key}:`, error);
    
    if (error.message && error.message.includes('limit')) {
      console.log('💡 Property limit reached. Try running:');
      console.log('  - clearAllProcessedMessages() to free up space');
      console.log('  - emergencyCleanup() for automatic cleanup');
      console.log('  - rotateProcessedMessages(20) to remove old entries');
    }
    
    throw error;
  }
}

/**
 * Get or create Google Drive folder for attachments
 * 添付ファイル用のGoogle Driveフォルダを取得または作成
 */
function getOrCreateDriveFolder() {
  try {
    // Try to get existing folder ID from properties
    let folderId;
    try {
      folderId = getProperty(PROPERTY_KEYS.DRIVE_FOLDER_ID);
      // Verify folder still exists
      DriveApp.getFolderById(folderId);
      return folderId;
    } catch (error) {
      console.log('Drive folder not found in properties or folder deleted, creating new one...');
    }
    
    // Search for existing folder by name
    const folders = DriveApp.getFoldersByName(CONFIG.DRIVE_FOLDER_NAME);
    if (folders.hasNext()) {
      const folder = folders.next();
      folderId = folder.getId();
      console.log(`Found existing Drive folder: ${CONFIG.DRIVE_FOLDER_NAME}`);
    } else {
      // Create new folder
      const folder = DriveApp.createFolder(CONFIG.DRIVE_FOLDER_NAME);
      folderId = folder.getId();
      console.log(`Created new Drive folder: ${CONFIG.DRIVE_FOLDER_NAME}`);
    }
    
    // Save folder ID to properties
    setProperty(PROPERTY_KEYS.DRIVE_FOLDER_ID, folderId);
    
    return folderId;
    
  } catch (error) {
    console.error('Error managing Drive folder:', error);
    throw error;
  }
}

/**
 * Mark individual message as processed using spreadsheet tracking
 * スプレッドシート追跡で個別メッセージを処理済みとしてマーク
 * 
 * @param {GmailMessage} message - Gmail message to mark as processed
 */
function markMessageAsProcessed(message) {
  try {
    const messageId = message.getId();
    const subject = message.getSubject();
    const sender = message.getFrom();
    
    // Always mark by message ID (each recipient gets logged separately)
    markMessageProcessedInSpreadsheet(messageId, subject, sender);
    console.log(`Marked message as processed: ${messageId}`);
    
  } catch (error) {
    console.error('Error marking message as processed:', error);
    // Don't throw - this is not critical for main functionality
  }
}

/**
 * Cleanup old processed message records (optional maintenance)
 * 古い処理済みメッセージレコードのクリーンアップ（オプションのメンテナンス）
 * 
 * @param {number} daysOld - Remove records older than this many days
 */
function cleanupOldProcessedMessages(daysOld = 30) {
  try {
    console.log(`Cleaning up processed message records older than ${daysOld} days...`);
    
    const cutoffTime = new Date().getTime() - (daysOld * 24 * 60 * 60 * 1000);
    const properties = PropertiesService.getScriptProperties().getProperties();
    
    let cleanedCount = 0;
    
    for (const [key, value] of Object.entries(properties)) {
      if (key.startsWith('PROCESSED_MSG_')) {
        const timestamp = parseInt(value);
        if (timestamp < cutoffTime) {
          PropertiesService.getScriptProperties().deleteProperty(key);
          cleanedCount++;
        }
      }
    }
    
    console.log(`Cleaned up ${cleanedCount} old processed message records`);
    return cleanedCount;
    
  } catch (error) {
    console.error('Error cleaning up old processed messages:', error);
    return 0;
  }
}

/**
 * Add processed label to email thread
 * メールスレッドに処理済みラベルを追加
 */
function addProcessedLabel(thread) {
  try {
    let label = GmailApp.getUserLabelByName(CONFIG.GMAIL_LABEL);
    if (!label) {
      label = GmailApp.createLabel(CONFIG.GMAIL_LABEL);
      console.log(`Created new Gmail label: ${CONFIG.GMAIL_LABEL}`);
    }
    thread.addLabel(label);
    console.log(`Added processed label to thread: ${thread.getFirstMessageSubject()}`);
  } catch (error) {
    console.error('Error adding processed label:', error);
    // Don't throw - this is not critical
  }
}

/**
 * Setup all required Script Properties
 * 必要なスクリプトプロパティをすべて設定
 */
function setupConfiguration() {
  console.log('=== Setting up Configuration ===');
  
  try {
    // Example setup - replace with actual values
    const config = {
      'SENDER_EMAIL': 'your-actual-sender@example.com',
      'SLACK_WEBHOOK_URL': 'https://hooks.slack.com/services/YOUR/ACTUAL/WEBHOOK',
      'SLACK_CHANNEL': '#your-actual-channel'
    };
    
    Object.entries(config).forEach(([key, value]) => {
      setProperty(key, value);
      console.log(`Set ${key}: ${key === 'SLACK_WEBHOOK_URL' ? '[REDACTED]' : value}`);
    });
    
    console.log('✅ Configuration setup completed');
    console.log('⚠️  IMPORTANT: Update the values above with your actual configuration!');
    
  } catch (error) {
    console.error('Configuration setup failed:', error);
    throw error;
  }
}

/**
 * Setup contract management system configuration
 * 契約管理システムの設定をセットアップ
 */
function setContractConfiguration() {
  console.log('=== Setting up Contract Management Configuration ===');
  
  try {
    // Contract management tool sender emails
    // ⚠️ IMPORTANT: Replace these with your actual contract tool email addresses!
    const contractConfig = {
      // Primary contract tools - Update with your actual contract platform emails
      'SENDER_EMAIL_1': 'noreply@contracttool1.example.com',  // Contract Tool 1 (e.g., DocuSign)
      'SENDER_EMAIL_2': 'noreply@contracttool2.example.com',  // Contract Tool 2 (e.g., Adobe Sign)
      'SENDER_EMAIL_3': 'noreply@contracttool3.example.com',  // Contract Tool 3 (e.g., HelloSign)
      'SENDER_EMAIL_4': 'noreply@contracttool4.example.com',  // Contract Tool 4 (e.g., ContractWorks)
      'SENDER_EMAIL_5': 'noreply@contracttool5.example.com',  // Contract Tool 5 (e.g., PandaDoc)
      'SENDER_EMAIL_6': 'noreply@contracttool6.example.com',  // Contract Tool 6 (e.g., CloudSign)
      'SENDER_EMAIL_7': 'noreply@contracttool7.example.com',  // Contract Tool 7 (e.g., Dropbox Sign)
      'SENDER_EMAIL_8': 'noreply@contracttool8.example.com',  // Contract Tool 8 (alternative)
      
      // Note: Recipient filtering removed - now using content-based duplicate detection
      // 注意: 受信者フィルタリングを削除 - コンテンツベースの重複検知を使用
      
      // Slack configuration
      'SLACK_WEBHOOK_URL': 'https://hooks.slack.com/services/YOUR/ACTUAL/WEBHOOK',
      'SLACK_CHANNEL': '#contract-notifications'
    };
    
    // Set each property
    Object.entries(contractConfig).forEach(([key, value]) => {
      setProperty(key, value);
      console.log(`Set ${key}: ${key === 'SLACK_WEBHOOK_URL' ? '[REDACTED]' : value}`);
    });
    
    // Create Drive folder for contracts
    const folderId = getOrCreateDriveFolder();
    console.log(`Drive folder created/verified: ${folderId}`);
    
    // Create spreadsheet for contract tracking
    if (CONFIG.ENABLE_SPREADSHEET_LOGGING) {
      const spreadsheetId = createOrGetSpreadsheet();
      console.log(`Contract tracking spreadsheet created/verified: ${spreadsheetId}`);
    }
    
    console.log('\n✅ Contract management configuration completed!');
    console.log('\n⚠️  IMPORTANT NEXT STEPS:');
    console.log('1. Update SENDER_EMAIL_* with your actual contract tool emails');
    console.log('2. Update SLACK_WEBHOOK_URL with your actual webhook');
    console.log('3. Update SLACK_CHANNEL with your desired channel');
    console.log('4. Run showConfiguration() to verify settings');
    console.log('5. Run runAllTests() to test the system');
    console.log('6. Run setupInitialTrigger() to enable automatic processing');
    
  } catch (error) {
    console.error('Contract configuration setup failed:', error);
    throw error;
  }
}

/**
 * Display current configuration (safely, without exposing secrets)
 * 現在の設定を表示（機密情報を除く）
 */
function showConfiguration() {
  console.log('=== Current Configuration ===');
  
  try {
    // Display contract management tool emails
    console.log('\nContract Management Tool Emails:');
    const configuredEmails = getConfiguredSenderEmails();
    
    if (configuredEmails.length === 0) {
      console.log('  ⚠️  No sender emails configured!');
      console.log('  Please run setContractConfiguration() or set SENDER_EMAIL_* properties');
    } else {
      configuredEmails.forEach((email, index) => {
        const toolName = getContractToolName(email);
        console.log(`  ${index + 1}. ${email} (${toolName})`);
      });
      console.log(`  Total configured: ${configuredEmails.length}`);
    }
    
    // Display duplicate prevention method
    console.log('\nDuplicate Prevention:');
    console.log('  📋 Method: Content-based detection');
    console.log('  🔍 Criteria: Sender + Date + Subject + PDF filenames');
    console.log('  📁 PDF Saving: Skipped for duplicate content');
    console.log('  📊 Slack/Spreadsheet: Always logged (per recipient)');
    
    // Slack configuration
    const slackChannel = getProperty('SLACK_CHANNEL', false);
    const webhookUrl = getProperty('SLACK_WEBHOOK_URL', false);
    console.log(`\nSlack Configuration:`);
    console.log(`  Channel: ${slackChannel || 'NOT SET'}`);
    console.log(`  Webhook URL: ${webhookUrl ? '[SET]' : 'NOT SET'}`);
    
    // Storage configuration
    const folderId = getProperty('DRIVE_FOLDER_ID', false);
    const spreadsheetId = getProperty('SPREADSHEET_ID', false);
    console.log(`\nStorage Configuration:`);
    console.log(`  Drive Folder ID: ${folderId || 'NOT SET'}`);
    console.log(`  Spreadsheet ID: ${spreadsheetId || 'NOT SET'}`);
    console.log(`  Spreadsheet Logging: ${CONFIG.ENABLE_SPREADSHEET_LOGGING ? 'ENABLED' : 'DISABLED'}`);
    
    // Pattern settings
    console.log(`\nPattern Settings:`);
    console.log(`  Multiple patterns enabled: ${CONFIG.SUBJECT_PATTERNS?.ENABLE_MULTIPLE_PATTERNS}`);
    console.log(`  Total patterns: ${CONFIG.SUBJECT_PATTERNS?.PATTERNS?.length || 0}`);
    console.log(`  Match mode: ${CONFIG.SUBJECT_PATTERNS?.MATCH_MODE}`);
    
    // Processing settings
    console.log(`\nProcessing Settings:`);
    console.log(`  Max emails per run: ${CONFIG.MAX_EMAILS_PER_RUN}`);
    console.log(`  Trigger interval: ${CONFIG.TRIGGER_INTERVAL_MINUTES} minutes`);
    console.log(`  Duplicate tracking: Spreadsheet-based (unlimited)`);
    
    // Show processed message statistics from spreadsheet
    try {
      const stats = getProcessedMessageStats();
      console.log(`  Processed messages (total): ${stats.total}`);
      console.log(`  Processed messages (last 7 days): ${stats.recent}`);
      if (stats.errors > 0) {
        console.log(`  Processing errors: ${stats.errors}`);
      }
    } catch (error) {
      console.log(`  Processed message stats: Unable to retrieve`);
    }
    
  } catch (error) {
    console.error('Error showing configuration:', error);
  }
}

/**
 * Get contract tool name from email address
 * メールアドレスから契約ツール名を取得
 */
function getContractToolName(email) {
  const toolMap = {
    'noreply@contracttool1.example.com': 'Contract Tool 1',
    'noreply@contracttool2.example.com': 'Contract Tool 2',
    'noreply@contracttool3.example.com': 'Contract Tool 3',
    'noreply@contracttool4.example.com': 'Contract Tool 4',
    'noreply@contracttool5.example.com': 'Contract Tool 5',
    'noreply@contracttool6.example.com': 'Contract Tool 6',
    'noreply@contracttool7.example.com': 'Contract Tool 7',
    'noreply@contracttool8.example.com': 'Contract Tool 8'
  };
  
  return toolMap[email] || 'Unknown';
}

/**
 * Get all configured sender emails dynamically
 * 動的に設定された全ての送信者メールを取得
 * 
 * @returns {Array} Array of configured sender emails
 */
function getConfiguredSenderEmails() {
  const emails = [];
  
  // Check legacy SENDER_EMAIL first
  const legacyEmail = getProperty('SENDER_EMAIL', false);
  if (legacyEmail) {
    emails.push(legacyEmail);
  }
  
  // Check SENDER_EMAIL_1 through SENDER_EMAIL_20 (flexible)
  for (let i = 1; i <= 20; i++) {
    const email = getProperty(`SENDER_EMAIL_${i}`, false);
    if (email) {
      emails.push(email);
    }
  }
  
  // Remove duplicates and return
  return [...new Set(emails)];
}

/**
 * Generate content-based duplicate key for email
 * メールのコンテンツベース重複キーを生成
 * 
 * @param {string} sender - Email sender
 * @param {Date} date - Email date
 * @param {string} subject - Email subject
 * @param {Array} pdfNames - Array of PDF attachment names
 * @returns {string} - Duplicate detection key
 */
function generateContentDuplicateKey(sender, date, subject, pdfNames = []) {
  try {
    // Create a standardized key based on email content
    const dateKey = Utilities.formatDate(date, 'JST', 'yyyyMMdd_HHmm'); // Minute precision
    const subjectKey = subject.replace(/[^a-zA-Z0-9]/g, '').toLowerCase(); // Remove special chars
    const pdfKey = pdfNames.sort().join(',').replace(/[^a-zA-Z0-9.,]/g, '').toLowerCase();
    const senderKey = sender.replace(/[^a-zA-Z0-9@.]/g, '').toLowerCase();
    
    const combinedKey = `${senderKey}_${dateKey}_${subjectKey}_${pdfKey}`;
    
    // Generate hash for shorter key
    const hash = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, combinedKey));
    
    return `CONTENT_${hash.substring(0, 16)}`;
  } catch (error) {
    console.error('Error generating content duplicate key:', error);
    return `CONTENT_${new Date().getTime()}`; // Fallback
  }
}

/**
 * Test function for manual execution
 * 手動実行用のテスト関数
 */
function testProcessEmails() {
  console.log('=== TESTING Email Processing ===');
  try {
    processEmails();
    console.log('Test completed successfully');
  } catch (error) {
    console.error('Test failed:', error);
    throw error;
  }
}

/**
 * Test the new spreadsheet-based duplicate tracking system
 * 新しいスプレッドシートベースの重複追跡システムのテスト
 */
function testSpreadsheetTracking() {
  console.log('=== TESTING Spreadsheet-Based Duplicate Tracking ===');
  
  try {
    // Test search query generation (no time filtering)
    const senderEmails = CONFIG.SENDER_EMAILS;
    const recipientEmails = CONFIG.RECIPIENT_EMAILS;
    
    if (senderEmails.length === 0) {
      console.log('⚠️  No sender emails configured for testing');
      return;
    }
    
    const senderQueries = senderEmails.map(email => `from:${email}`);
    let query = `(${senderQueries.join(' OR ')})`;
    
    if (recipientEmails.length > 0) {
      const recipientQueries = recipientEmails.map(email => `to:${email}`);
      query += ` AND (${recipientQueries.join(' OR ')})`;
    }
    
    console.log(`Generated search query: ${query}`);
    console.log('No time filtering - searches all emails for complete coverage');
    
    // Test spreadsheet tracking functions
    console.log('\nTesting spreadsheet tracking functions...');
    
    // Get current stats
    const stats = getProcessedMessageStats();
    console.log(`Current processed message stats:`);
    console.log(`  Total: ${stats.total}`);
    console.log(`  Recent (7 days): ${stats.recent}`);
    console.log(`  Errors: ${stats.errors}`);
    
    // Test creating a mock processed entry
    const testMessageId = 'test_msg_' + new Date().getTime();
    const testSubject = 'Test Contract Completion';
    const testSender = 'test@contracttool.example.com';
    const testHash = recipientEmails.length > 0 ? 'test_hash_123' : null;
    
    console.log(`\nTesting mock entry creation:`);
    console.log(`  Message ID: ${testMessageId}`);
    console.log(`  Subject: ${testSubject}`);
    console.log(`  Sender: ${testSender}`);
    console.log(`  Recipient Hash: ${testHash || 'None (no recipient filtering)'}`);
    
    // Mark as processed
    markMessageProcessedInSpreadsheet(testMessageId, testSubject, testSender, testHash);
    
    // Check if it's marked as processed
    const isProcessed = isMessageProcessedInSpreadsheet(testMessageId, testHash);
    console.log(`  Verification: ${isProcessed ? '✅ Successfully tracked' : '❌ Failed to track'}`);
    
    console.log('\n✅ Spreadsheet tracking test completed successfully');
    console.log('\nBenefits of this approach:');
    console.log('  • No time filtering - complete email history coverage');
    console.log('  • Unlimited tracking capacity (spreadsheet storage)');
    console.log('  • No Script Properties limit issues');
    console.log('  • Persistent tracking across all time periods');
    console.log('  • Easy to view and manage processed messages');
    
  } catch (error) {
    console.error('❌ Spreadsheet tracking test failed:', error);
    throw error;
  }
}

/**
 * Quick property management functions (aliases for easier access)
 * クイックプロパティ管理関数（簡単アクセス用エイリアス）
 */

/**
 * Show current property usage and statistics
 * 現在のプロパティ使用状況と統計を表示
 */
function checkProperties() {
  showAllProperties();
}

/**
 * Debug HelloSign email processing specifically
 * HelloSignメール処理の専用デバッグ
 */
function debugHelloSignEmails() {
  console.log('=== DEBUGGING HELLOSIGN EMAIL PROCESSING ===');
  
  try {
    // Check if HelloSign sender email is configured
    const senderEmails = getConfiguredSenderEmails();
    console.log('Configured sender emails:', senderEmails);
    
    const helloSignEmail = senderEmails.find(email => 
      email && email.toLowerCase().includes('hellosign')
    );
    
    if (!helloSignEmail) {
      console.log('❌ HelloSign email not found in configuration');
      console.log('Please set a property like: SENDER_EMAIL_X = noreply@mail.hellosign.com');
      return;
    }
    
    console.log(`✓ HelloSign email configured: ${helloSignEmail}`);
    
    // Search for HelloSign emails specifically
    const query = `from:${helloSignEmail} -label:${CONFIG.GMAIL_LABEL}`;
    console.log(`Gmail search query: ${query}`);
    
    const threads = GmailApp.search(query, 0, 10);
    console.log(`Found ${threads.length} HelloSign email threads`);
    
    if (threads.length === 0) {
      console.log('No HelloSign emails found. Possible reasons:');
      console.log('1. No emails from this sender in your Gmail');
      console.log('2. All emails already processed (have the label)');
      console.log('3. Email address mismatch in configuration');
      
      // Try broader search
      const broadQuery = 'from:hellosign.com';
      const broadThreads = GmailApp.search(broadQuery, 0, 5);
      console.log(`Broader search (from:hellosign.com): ${broadThreads.length} results`);
      
      if (broadThreads.length > 0) {
        console.log('Sample HelloSign email senders:');
        broadThreads.slice(0, 3).forEach((thread, index) => {
          const messages = thread.getMessages();
          if (messages.length > 0) {
            const message = messages[0];
            console.log(`${index + 1}. From: ${message.getFrom()}`);
            console.log(`   Subject: ${message.getSubject()}`);
          }
        });
      }
      
      return;
    }
    
    // Test subject pattern matching on HelloSign emails
    console.log('\nTesting subject pattern matching:');
    threads.slice(0, 5).forEach((thread, index) => {
      const messages = thread.getMessages();
      if (messages.length === 0) {
        console.log(`\n--- HelloSign Email ${index + 1} ---`);
        console.log('No messages in thread');
        return;
      }
      
      const message = messages[0]; // Get first message in thread
      const subject = message.getSubject();
      const sender = message.getFrom();
      
      console.log(`\n--- HelloSign Email ${index + 1} ---`);
      console.log(`From: ${sender}`);
      console.log(`Subject: "${subject}"`);
      
      // Test pattern matching
      const patternResult = checkSubjectPattern(subject);
      console.log(`Pattern match: ${patternResult.isMatch ? '✅ MATCHES' : '❌ NO MATCH'}`);
      
      if (!patternResult.isMatch && patternResult.checkedPatterns) {
        console.log(`Checked ${patternResult.checkedPatterns.length} patterns`);
      }
      
      if (patternResult.isMatch && patternResult.matchedPattern) {
        console.log(`Matched pattern: ${patternResult.matchedPattern}`);
      }
    });
    
    console.log('\n=== HelloSign Debug Complete ===');
    
  } catch (error) {
    console.error('Error debugging HelloSign emails:', error);
    throw error;
  }
}

/**
 * Clean up old processed messages (default: 7 days)
 * 古い処理済みメッセージをクリーンアップ（デフォルト: 7日）
 * 
 * @param {number} days - Days to keep (default: 7)
 * @param {boolean} dryRun - Preview mode (default: false)
 */
function cleanupOld(days = 7, dryRun = false) {
  return cleanupProcessedMessages(days, dryRun);
}

/**
 * Emergency cleanup when property limit is reached
 * プロパティ制限到達時の緊急クリーンアップ
 */
function fixProperties() {
  return emergencyCleanup();
}

/**
 * Clean up old spreadsheet entries - alias for easier access
 * 古いスプレッドシートエントリのクリーンアップ - 簡単アクセス用エイリアス
 */
function cleanupSpreadsheetOld(days = 30) {
  return cleanupOldProcessedEntriesFromSpreadsheet(days);
}

/**
 * Clear all processed messages to free up property space
 * 全ての処理済みメッセージを削除してプロパティ領域を解放
 * 
 * @returns {number} Number of properties deleted
 */
function clearAllProcessedMessages() {
  try {
    const properties = PropertiesService.getScriptProperties().getProperties();
    let count = 0;
    
    console.log('=== 全ての処理済みメッセージを削除中 ===');
    console.log(`削除前のプロパティ数: ${Object.keys(properties).length}/50`);
    
    for (const key of Object.keys(properties)) {
      if (key.startsWith('PROCESSED_MSG_')) {
        try {
          PropertiesService.getScriptProperties().deleteProperty(key);
          count++;
        } catch (error) {
          console.error(`削除エラー ${key}:`, error);
        }
      }
    }
    
    console.log(`${count} 件の処理済みメッセージを削除しました`);
    
    const newProperties = PropertiesService.getScriptProperties().getProperties();
    const newTotal = Object.keys(newProperties).length;
    console.log(`削除後のプロパティ数: ${newTotal}/50`);
    console.log(`解放されたプロパティ数: ${Object.keys(properties).length - newTotal}`);
    
    if (newTotal < 45) {
      console.log('✅ プロパティ制限から余裕が生まれました');
    } else if (newTotal < 50) {
      console.log('⚠️  まだ制限に近いです。さらなるクリーンアップを検討してください');
    } else {
      console.log('🚨 まだ制限に達しています。他のプロパティの削除を検討してください');
    }
    
    return count;
    
  } catch (error) {
    console.error('処理済みメッセージの削除中にエラー:', error);
    throw error;
  }
}

/**
 * Rotate processed messages by removing oldest entries
 * 最古のエントリを削除して処理済みメッセージを回転
 * 
 * @param {number} removeCount - Number of oldest entries to remove
 * @returns {number} Number of entries actually removed
 */
function rotateProcessedMessages(removeCount = 10) {
  try {
    const allProperties = PropertiesService.getScriptProperties().getProperties();
    const processedMessages = [];
    
    // Collect all processed message properties with timestamps
    for (const [key, value] of Object.entries(allProperties)) {
      if (key.startsWith('PROCESSED_MSG_')) {
        const timestamp = parseInt(value);
        if (!isNaN(timestamp)) {
          processedMessages.push({ key, timestamp });
        }
      }
    }
    
    // Sort by timestamp (oldest first)
    processedMessages.sort((a, b) => a.timestamp - b.timestamp);
    
    // Remove the oldest entries
    const toRemove = processedMessages.slice(0, Math.min(removeCount, processedMessages.length));
    let removedCount = 0;
    
    toRemove.forEach(item => {
      try {
        PropertiesService.getScriptProperties().deleteProperty(item.key);
        removedCount++;
      } catch (error) {
        console.error(`Error removing property ${item.key}:`, error);
      }
    });
    
    console.log(`Rotated processed messages: removed ${removedCount} oldest entries`);
    return removedCount;
    
  } catch (error) {
    console.error('Error rotating processed messages:', error);
    return 0;
  }
}

/**
 * Enhanced property check with detailed breakdown
 * 詳細な内訳付きプロパティチェック
 */
function checkPropertiesDetailed() {
  try {
    const properties = PropertiesService.getScriptProperties().getProperties();
    const total = Object.keys(properties).length;
    
    console.log(`=== 詳細プロパティ確認 (${total}/50) ===`);
    
    const categories = {
      processed: [],
      config: [],
      sender: [],
      recipient: [],
      other: []
    };
    
    for (const [key, value] of Object.entries(properties)) {
      if (key.startsWith('PROCESSED_MSG_')) {
        categories.processed.push({ key, value, date: new Date(parseInt(value)) });
      } else if (key.startsWith('SENDER_EMAIL')) {
        categories.sender.push({ key, value });
      } else if (key.startsWith('RECIPIENT_EMAIL')) {
        categories.recipient.push({ key, value });
      } else if (['SLACK_WEBHOOK_URL', 'SLACK_CHANNEL', 'DRIVE_FOLDER_ID', 'SPREADSHEET_ID'].includes(key)) {
        categories.config.push({ key, value: key === 'SLACK_WEBHOOK_URL' ? '[HIDDEN]' : value });
      } else {
        categories.other.push({ key, value });
      }
    }
    
    console.log(`📊 カテゴリ別内訳:`);
    console.log(`  設定: ${categories.config.length}`);
    console.log(`  送信者: ${categories.sender.length}`);
    console.log(`  受信者: ${categories.recipient.length}`);
    console.log(`  処理済み: ${categories.processed.length}`);
    console.log(`  その他: ${categories.other.length}`);
    
    if (categories.processed.length > 0) {
      console.log(`\n📝 処理済みメッセージ (最新5件):`);
      categories.processed
        .sort((a, b) => b.date - a.date)
        .slice(0, 5)
        .forEach(item => {
          const messageId = item.key.replace('PROCESSED_MSG_', '').split('_')[0];
          console.log(`  ${messageId}: ${item.date.toLocaleString('ja-JP')}`);
        });
      
      if (categories.processed.length > 5) {
        console.log(`  ...他 ${categories.processed.length - 5} 件`);
      }
    }
    
    console.log(`\n💡 推奨アクション:`);
    if (total >= 50) {
      console.log(`  🚨 Script Properties制限到達: clearAllProcessedMessages() を実行`);
      console.log(`  📊 スプレッドシートベースの追跡システムが有効です`);
    } else if (total >= 45) {
      console.log(`  ⚠️  Script Properties制限接近: 定期的な cleanupOld() を推奨`);
      console.log(`  📊 新しいシステムではスプレッドシートを使用`);
    } else {
      console.log(`  ✅ Script Properties余裕あり: 現在の運用を継続`);
      console.log(`  📊 スプレッドシートベースの重複追跡が有効です`);
    }
    
    return {
      total,
      categories,
      nearLimit: total >= 45,
      atLimit: total >= 50,
      processedMessageCount: categories.processed.length,
      spreadsheetTrackingActive: true
    };
    
  } catch (error) {
    console.error('プロパティチェック中にエラー:', error);
    throw error;
  }
}