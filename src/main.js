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
      /^Completed:.*Complete with Docusign:/i,  // Docusign specific pattern
      /^完了:.*Complete with Docusign:/i,        // Japanese "完了: Complete with Docusign:"
      /^Completed:.*\.pdf$/i,                    // "Completed: filename.pdf"
      /^Completed:.*\.docx$/i,                   // "Completed: filename.docx"
      /^完了:.*\.pdf$/i,                         // Japanese "完了: filename.pdf"
      /^完了:.*\.docx$/i,                        // Japanese "完了: filename.docx"
      /^Fwd:.*Completed:.*Complete with Docusign:/i,  // Forwarded Docusign emails
      /^Fwd:.*完了:.*Complete with Docusign:/i,        // Forwarded Japanese Docusign emails
      
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
      /.*You've been copied on.*signed by/i,
      /.*has been completed/i,
      /.*signature.*completed/i,
      /.*agreement.*signed by.*and/i,
      /.*You just signed.*/i,
      
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
  
  // Docusign integration settings for flexible sender detection
  // Docusign uses various sender patterns, detect via both sender and subject
  DOCUSIGN_INTEGRATION: {
    ENABLE: true,  // Enable Docusign email detection
    
    // Sender patterns for Docusign emails
    SENDER_PATTERNS: [
      /.*@.*\.docusign\.net$/i,        // Any email from docusign.net domains
      /.*via Docusign.*/i,             // "Name via Docusign" in sender field
      /.*docusign.*/i                  // Any sender containing "docusign"
    ],
    
    // Subject patterns for additional verification
    SUBJECT_PATTERNS: [
      /^Completed:.*Complete with Docusign:/i, // "Completed: Complete with Docusign: filename"
      /^完了:.*Complete with Docusign:/i,       // Japanese "完了: Complete with Docusign: filename"
      /^Completed:.*\.pdf$/i,                  // "Completed: filename.pdf"
      /^Completed:.*\.docx$/i,                 // "Completed: filename.docx"
      /^Completed:.*\.doc$/i,                  // "Completed: filename.doc"
      /^完了:.*\.pdf$/i,                       // Japanese "完了: filename.pdf"
      /^完了:.*\.docx$/i,                      // Japanese "完了: filename.docx"
      /^完了:.*\.doc$/i,                       // Japanese "完了: filename.doc"
      /.*via Docusign.*completed/i,            // "Name via Docusign completed"
      /.*has been completed/i,                 // Generic completion messages
      /.*document.*signed/i                    // Document signing completion
    ],
    
    // Additional verification requirements
    REQUIRE_PDF_ATTACHMENT: true,             // Only process emails with PDF attachments
    
    // Detection mode: 'sender_or_subject' (either match), 'sender_and_subject' (both required)
    DETECTION_MODE: 'sender_or_subject'
  },
  
  // Dropbox Sign/HelloSign integration settings for organizational email forwarding
  // Handles emails sent via organizational email addresses but originating from Dropbox Sign
  DROPBOX_SIGN_INTEGRATION: {
    ENABLE: true,  // Enable Dropbox Sign/HelloSign email detection
    
    // Sender patterns for Dropbox Sign emails (including organizational forwarding)
    SENDER_PATTERNS: [
      /.*via Dropbox Sign.*/i,         // "Name via Dropbox Sign" in sender field
      /.*via HelloSign.*/i,            // "Name via HelloSign" in sender field (legacy)
      /'Dropbox Sign' via .*/i,        // "'Dropbox Sign' via organization" format
      /'HelloSign' via .*/i            // "'HelloSign' via organization" format (legacy)
    ],
    
    // Reply-to patterns for additional verification
    REPLY_TO_PATTERNS: [
      /.*@hellosign\.com$/i,           // Reply-to hellosign.com domain
      /noreply@hellosign\.com$/i       // Specific noreply address
    ],
    
    // Subject patterns for Dropbox Sign/HelloSign completion emails
    SUBJECT_PATTERNS: [
      /You've been copied on.*signed by/i,     // "You've been copied on X signed by Y"
      /.*has been completed/i,                 // Generic completion messages
      /.*signature.*completed/i,               // "Signature completed" notifications
      /.*agreement.*signed by.*and/i,          // Multi-signer agreement notifications
      /^You just signed.*/i,                   // "You just signed X" notifications
      /.*document.*signed/i                    // Document signing completion
    ],
    
    // Additional verification requirements
    REQUIRE_PDF_ATTACHMENT: false,            // Dropbox Sign emails may not always have PDF attachments
    
    // Detection mode: 'sender_or_subject' (either match), 'sender_and_subject' (both required)
    DETECTION_MODE: 'sender_or_subject'
  },
  
  GMAIL_LABEL: 'Contract_Processed',  // 処理済み契約メールのラベル名
  GMAIL_SKIP_LABEL: 'Contract_Skipped',  // パターン不一致でスキップしたメールのラベル名
  
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
    
    // Build search query for multiple contract management tool senders and Docusign
    // 複数の契約管理ツールからのメールとDocusignを検索
    const senderQueries = CONFIG.SENDER_EMAILS.map(email => `from:${email}`);
    
    // Add Docusign subject-based search if enabled
    const searchQueries = [senderQueries.join(' OR ')];
    
    if (CONFIG.DOCUSIGN_INTEGRATION?.ENABLE) {
      // Build flexible Docusign search query with both sender and subject patterns
      const docusignSearchTerms = [];
      
      // Add sender-based searches
      if (CONFIG.DOCUSIGN_INTEGRATION.SENDER_PATTERNS?.length > 0) {
        docusignSearchTerms.push('from:docusign.net');  // Domain-based search
        docusignSearchTerms.push('"via Docusign"');     // Sender name pattern
      }
      
      // Add subject-based searches
      if (CONFIG.DOCUSIGN_INTEGRATION.SUBJECT_PATTERNS?.length > 0) {
        docusignSearchTerms.push('subject:"Complete with Docusign"');
        docusignSearchTerms.push('subject:"via Docusign"');
        docusignSearchTerms.push('subject:"Completed:"');
        docusignSearchTerms.push('subject:"has been completed"');
        docusignSearchTerms.push('subject:"document signed"');
      }
      
      if (docusignSearchTerms.length > 0) {
        const docusignQuery = `(${docusignSearchTerms.join(' OR ')})`;
        searchQueries.push(docusignQuery);
        console.log(`Docusign detection enabled with ${docusignSearchTerms.length} search terms`);
      }
    }
    
    if (CONFIG.DROPBOX_SIGN_INTEGRATION?.ENABLE) {
      // Build flexible Dropbox Sign search query with sender and subject patterns
      const dropboxSignSearchTerms = [];
      
      // Add sender-based searches
      if (CONFIG.DROPBOX_SIGN_INTEGRATION.SENDER_PATTERNS?.length > 0) {
        dropboxSignSearchTerms.push('"via Dropbox Sign"');     // "Name via Dropbox Sign" pattern
        dropboxSignSearchTerms.push('"via HelloSign"');        // Legacy HelloSign pattern
        dropboxSignSearchTerms.push("'Dropbox Sign' via");     // "'Dropbox Sign' via organization" pattern
        dropboxSignSearchTerms.push("'HelloSign' via");        // "'HelloSign' via organization" pattern
      }
      
      // Add reply-to based searches for better detection
      if (CONFIG.DROPBOX_SIGN_INTEGRATION.REPLY_TO_PATTERNS?.length > 0) {
        dropboxSignSearchTerms.push('from:hellosign.com');     // Reply-to domain search
      }
      
      // Add subject-based searches
      if (CONFIG.DROPBOX_SIGN_INTEGRATION.SUBJECT_PATTERNS?.length > 0) {
        dropboxSignSearchTerms.push('subject:"You\'ve been copied on"');
        dropboxSignSearchTerms.push('subject:"signed by"');
        dropboxSignSearchTerms.push('subject:"has been completed"');
        dropboxSignSearchTerms.push('subject:"You just signed"');
      }
      
      if (dropboxSignSearchTerms.length > 0) {
        const dropboxSignQuery = `(${dropboxSignSearchTerms.join(' OR ')})`;
        searchQueries.push(dropboxSignQuery);
        console.log(`Dropbox Sign detection enabled with ${dropboxSignSearchTerms.length} search terms`);
      }
    }
    
    let query = `(${searchQueries.join(' OR ')}) -label:${CONFIG.GMAIL_LABEL} -label:${CONFIG.GMAIL_SKIP_LABEL}`;
    
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
    
    // Docusign integration settings
    console.log(`\nDocusign Integration:`);
    console.log(`  Docusign detection: ${CONFIG.DOCUSIGN_INTEGRATION?.ENABLE ? 'ENABLED' : 'DISABLED'}`);
    if (CONFIG.DOCUSIGN_INTEGRATION?.ENABLE) {
      console.log(`  Sender patterns: ${CONFIG.DOCUSIGN_INTEGRATION.SENDER_PATTERNS?.length || 0}`);
      console.log(`  Subject patterns: ${CONFIG.DOCUSIGN_INTEGRATION.SUBJECT_PATTERNS?.length || 0}`);
      console.log(`  Detection mode: ${CONFIG.DOCUSIGN_INTEGRATION.DETECTION_MODE || 'sender_or_subject'}`);
      console.log(`  PDF attachment required: ${CONFIG.DOCUSIGN_INTEGRATION.REQUIRE_PDF_ATTACHMENT ? 'YES' : 'NO'}`);
      console.log(`  Processing: All Docusign emails (no recipient filtering)`);
    }
    
    // Dropbox Sign integration settings
    console.log(`\nDropbox Sign Integration:`);
    console.log(`  Dropbox Sign detection: ${CONFIG.DROPBOX_SIGN_INTEGRATION?.ENABLE ? 'ENABLED' : 'DISABLED'}`);
    if (CONFIG.DROPBOX_SIGN_INTEGRATION?.ENABLE) {
      console.log(`  Sender patterns: ${CONFIG.DROPBOX_SIGN_INTEGRATION.SENDER_PATTERNS?.length || 0}`);
      console.log(`  Subject patterns: ${CONFIG.DROPBOX_SIGN_INTEGRATION.SUBJECT_PATTERNS?.length || 0}`);
      console.log(`  Reply-to patterns: ${CONFIG.DROPBOX_SIGN_INTEGRATION.REPLY_TO_PATTERNS?.length || 0}`);
      console.log(`  Detection mode: ${CONFIG.DROPBOX_SIGN_INTEGRATION.DETECTION_MODE || 'sender_or_subject'}`);
      console.log(`  PDF attachment required: ${CONFIG.DROPBOX_SIGN_INTEGRATION.REQUIRE_PDF_ATTACHMENT ? 'YES' : 'NO'}`);
      console.log(`  Processing: All Dropbox Sign emails via organizational forwarding`);
    }
    
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
 * Detect message source type (sender-based, Docusign, or Dropbox Sign)
 * メッセージのソースタイプを検出（送信者ベース、Docusign、またはDropbox Sign）
 * 
 * @param {GmailMessage} message - Gmail message object
 * @returns {Object} - {type: 'SENDER_BASED'|'DOCUSIGN'|'DROPBOX_SIGN', details: {...}}
 */
function detectMessageSource(message) {
  try {
    const sender = message.getFrom();
    const subject = message.getSubject();
    const to = message.getTo();
    const attachments = message.getAttachments();
    
    // Get reply-to header for additional verification
    let replyTo = '';
    try {
      const rawMessage = message.getRawContent();
      const replyToMatch = rawMessage.match(/Reply-To: (.+)/i);
      if (replyToMatch) {
        replyTo = replyToMatch[1].trim();
      }
    } catch (error) {
      console.log('Could not extract reply-to header, continuing without it');
    }
    
    // Check if it's from a configured sender email (traditional contract tools)
    const configuredEmails = getConfiguredSenderEmails();
    const isSenderBased = configuredEmails.some(email => 
      sender.toLowerCase().includes(email.toLowerCase())
    );
    
    if (isSenderBased) {
      return {
        type: 'SENDER_BASED',
        details: {
          configuredSender: configuredEmails.find(email => 
            sender.toLowerCase().includes(email.toLowerCase())
          )
        }
      };
    }
    
    // Check if it's a Docusign email (flexible sender and subject detection)
    if (CONFIG.DOCUSIGN_INTEGRATION?.ENABLE) {
      const senderPatterns = CONFIG.DOCUSIGN_INTEGRATION.SENDER_PATTERNS || [];
      const subjectPatterns = CONFIG.DOCUSIGN_INTEGRATION.SUBJECT_PATTERNS || [];
      const detectionMode = CONFIG.DOCUSIGN_INTEGRATION.DETECTION_MODE || 'sender_or_subject';
      
      // Test sender patterns
      const senderMatch = senderPatterns.some(pattern => pattern.test(sender));
      
      // Test subject patterns
      const subjectMatch = subjectPatterns.some(pattern => pattern.test(subject));
      
      let isDocusign = false;
      let detectedBy = [];
      
      if (detectionMode === 'sender_or_subject') {
        // Either sender OR subject must match
        isDocusign = senderMatch || subjectMatch;
      } else if (detectionMode === 'sender_and_subject') {
        // Both sender AND subject must match
        isDocusign = senderMatch && subjectMatch;
      }
      
      if (senderMatch) detectedBy.push('sender_pattern');
      if (subjectMatch) detectedBy.push('subject_pattern');
      
      if (isDocusign) {
        // Additional verification: check for PDF attachment if required
        if (CONFIG.DOCUSIGN_INTEGRATION.REQUIRE_PDF_ATTACHMENT) {
          const hasPdfAttachment = attachments.some(attachment => 
            attachment.getName().toLowerCase().endsWith('.pdf')
          );
          
          if (!hasPdfAttachment) {
            return {
              type: 'FILTERED_OUT',
              details: {
                reason: 'Docusign email detected but no PDF attachment found',
                senderMatch: senderMatch,
                subjectMatch: subjectMatch,
                attachmentCount: attachments.length,
                detectedBy: detectedBy.join(', ')
              }
            };
          }
          
          detectedBy.push('pdf_attachment_verified');
        }
        
        return {
          type: 'DOCUSIGN',
          details: {
            detectedBy: detectedBy.join(', '),
            senderMatch: senderMatch,
            subjectMatch: subjectMatch,
            detectionMode: detectionMode,
            hasPdfAttachment: CONFIG.DOCUSIGN_INTEGRATION.REQUIRE_PDF_ATTACHMENT ? 
              attachments.some(att => att.getName().toLowerCase().endsWith('.pdf')) : 'not_checked',
            attachmentCount: attachments.length,
            recipient: to
          }
        };
      }
    }
    
    // Check if it's a Dropbox Sign/HelloSign email (flexible sender and subject detection)
    if (CONFIG.DROPBOX_SIGN_INTEGRATION?.ENABLE) {
      const senderPatterns = CONFIG.DROPBOX_SIGN_INTEGRATION.SENDER_PATTERNS || [];
      const subjectPatterns = CONFIG.DROPBOX_SIGN_INTEGRATION.SUBJECT_PATTERNS || [];
      const replyToPatterns = CONFIG.DROPBOX_SIGN_INTEGRATION.REPLY_TO_PATTERNS || [];
      const detectionMode = CONFIG.DROPBOX_SIGN_INTEGRATION.DETECTION_MODE || 'sender_or_subject';
      
      // Test sender patterns
      const senderMatch = senderPatterns.some(pattern => pattern.test(sender));
      
      // Test subject patterns
      const subjectMatch = subjectPatterns.some(pattern => pattern.test(subject));
      
      // Test reply-to patterns for additional verification
      const replyToMatch = replyTo && replyToPatterns.some(pattern => pattern.test(replyTo));
      
      let isDropboxSign = false;
      let detectedBy = [];
      
      if (detectionMode === 'sender_or_subject') {
        // Either sender, subject, or reply-to must match
        isDropboxSign = senderMatch || subjectMatch || replyToMatch;
      } else if (detectionMode === 'sender_and_subject') {
        // Both sender AND subject must match (reply-to is additional)
        isDropboxSign = senderMatch && subjectMatch;
      }
      
      if (senderMatch) detectedBy.push('sender_pattern');
      if (subjectMatch) detectedBy.push('subject_pattern');
      if (replyToMatch) detectedBy.push('reply_to_pattern');
      
      if (isDropboxSign) {
        // Additional verification: check for PDF attachment if required
        if (CONFIG.DROPBOX_SIGN_INTEGRATION.REQUIRE_PDF_ATTACHMENT) {
          const hasPdfAttachment = attachments.some(attachment => 
            attachment.getName().toLowerCase().endsWith('.pdf')
          );
          
          if (!hasPdfAttachment) {
            return {
              type: 'FILTERED_OUT',
              details: {
                reason: 'Dropbox Sign email detected but no PDF attachment found (PDF required)',
                senderMatch: senderMatch,
                subjectMatch: subjectMatch,
                replyToMatch: replyToMatch,
                attachmentCount: attachments.length,
                detectedBy: detectedBy.join(', ')
              }
            };
          }
          
          detectedBy.push('pdf_attachment_verified');
        }
        
        return {
          type: 'DROPBOX_SIGN',
          details: {
            detectedBy: detectedBy.join(', '),
            senderMatch: senderMatch,
            subjectMatch: subjectMatch,
            replyToMatch: replyToMatch,
            detectionMode: detectionMode,
            hasPdfAttachment: CONFIG.DROPBOX_SIGN_INTEGRATION.REQUIRE_PDF_ATTACHMENT ? 
              attachments.some(att => att.getName().toLowerCase().endsWith('.pdf')) : 'not_checked',
            attachmentCount: attachments.length,
            recipient: to,
            replyTo: replyTo
          }
        };
      }
    }
    
    return {
      type: 'UNKNOWN',
      details: {
        sender: sender,
        subject: subject,
        replyTo: replyTo
      }
    };
    
  } catch (error) {
    console.error('Error detecting message source:', error);
    return {
      type: 'ERROR',
      details: {
        error: error.message
      }
    };
  }
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
 * Debug Docusign email processing specifically
 * Docusignメール処理の専用デバッグ
 */
function debugDocusignEmails() {
  console.log('=== DEBUGGING DOCUSIGN EMAIL PROCESSING ===');
  
  try {
    // Check if Docusign integration is enabled
    if (!CONFIG.DOCUSIGN_INTEGRATION?.ENABLE) {
      console.log('❌ Docusign integration is disabled in configuration');
      console.log('Enable by setting CONFIG.DOCUSIGN_INTEGRATION.ENABLE = true');
      return;
    }
    
    console.log('✓ Docusign integration enabled');
    console.log(`Sender patterns: ${CONFIG.DOCUSIGN_INTEGRATION.SENDER_PATTERNS?.length || 0}`);
    console.log(`Subject patterns: ${CONFIG.DOCUSIGN_INTEGRATION.SUBJECT_PATTERNS?.length || 0}`);
    console.log(`Detection mode: ${CONFIG.DOCUSIGN_INTEGRATION.DETECTION_MODE || 'sender_or_subject'}`);
    console.log(`PDF attachment required: ${CONFIG.DOCUSIGN_INTEGRATION.REQUIRE_PDF_ATTACHMENT ? 'YES' : 'NO'}`);
    
    // Build Docusign search query (same as in processEmails)
    const docusignSearchTerms = [];
    
    // Add sender-based searches
    if (CONFIG.DOCUSIGN_INTEGRATION.SENDER_PATTERNS?.length > 0) {
      docusignSearchTerms.push('from:docusign.net');  // Domain-based search
      docusignSearchTerms.push('"via Docusign"');     // Sender name pattern
    }
    
    // Add subject-based searches
    if (CONFIG.DOCUSIGN_INTEGRATION.SUBJECT_PATTERNS?.length > 0) {
      docusignSearchTerms.push('subject:"Complete with Docusign"');
      docusignSearchTerms.push('subject:"via Docusign"');
      docusignSearchTerms.push('subject:"Completed:"');
      docusignSearchTerms.push('subject:"has been completed"');
      docusignSearchTerms.push('subject:"document signed"');
    }
    
    let query = `(${docusignSearchTerms.join(' OR ')}) -label:${CONFIG.GMAIL_LABEL}`;
    
    console.log(`Gmail search query: ${query}`);
    
    const threads = GmailApp.search(query, 0, 10);
    console.log(`Found ${threads.length} potential Docusign email threads`);
    
    if (threads.length === 0) {
      console.log('No Docusign emails found. Possible reasons:');
      console.log('1. No Docusign completion emails in your Gmail');
      console.log('2. All emails already processed (have the Contract_Processed label)');
      console.log('3. Sender/subject patterns do not match actual Docusign emails');
      console.log('4. Detection mode requires both sender AND subject match');
      
      // Try broader search
      const broadQuery = 'subject:"Docusign"';
      const broadThreads = GmailApp.search(broadQuery, 0, 5);
      console.log(`Broader search (subject:"Docusign"): ${broadThreads.length} results`);
      
      if (broadThreads.length > 0) {
        console.log('Sample Docusign email subjects:');
        broadThreads.slice(0, 3).forEach((thread, index) => {
          const messages = thread.getMessages();
          if (messages.length > 0) {
            const message = messages[0];
            console.log(`${index + 1}. From: ${message.getFrom()}`);
            console.log(`   To: ${message.getTo()}`);
            console.log(`   Subject: ${message.getSubject()}`);
          }
        });
      }
      
      return;
    }
    
    // Test message source detection and pattern matching on Docusign emails
    console.log('\nTesting message source detection and pattern matching:');
    threads.slice(0, 5).forEach((thread, index) => {
      const messages = thread.getMessages();
      if (messages.length === 0) {
        console.log(`\n--- Docusign Email ${index + 1} ---`);
        console.log('No messages in thread');
        return;
      }
      
      const message = messages[0]; // Get first message in thread
      const subject = message.getSubject();
      const sender = message.getFrom();
      const to = message.getTo();
      
      console.log(`\n--- Docusign Email ${index + 1} ---`);
      console.log(`From: ${sender}`);
      console.log(`To: ${to}`);
      console.log(`Subject: "${subject}"`);
      
      // Test message source detection
      const messageSource = detectMessageSource(message);
      console.log(`Source detection: ${messageSource.type}`);
      console.log(`Details: ${JSON.stringify(messageSource.details)}`);
      
      // Test pattern matching
      const patternResult = checkSubjectPattern(subject, messageSource.type);
      console.log(`Pattern match: ${patternResult.isMatch ? '✅ MATCHES' : '❌ NO MATCH'}`);
      
      if (patternResult.isMatch && patternResult.matchedPattern) {
        console.log(`Matched pattern: ${patternResult.matchedPattern}`);
      }
      
      if (!patternResult.isMatch && patternResult.checkedPatterns) {
        console.log(`Checked ${patternResult.checkedPatterns.length} patterns`);
      }
    });
    
    console.log('\n=== Docusign Debug Complete ===');
    
  } catch (error) {
    console.error('Error debugging Docusign emails:', error);
    throw error;
  }
}

/**
 * Debug specific message ID search in spreadsheet
 * スプレッドシートでの特定メッセージID検索のデバッグ
 * 
 * @param {string} messageId - The message ID to search for
 */
function debugMessageIdSearch(messageId) {
  console.log(`=== DEBUGGING MESSAGE ID SEARCH: ${messageId} ===`);
  
  try {
    const spreadsheetId = getProperty('SPREADSHEET_ID', false);
    if (!spreadsheetId) {
      console.log('❌ No spreadsheet ID configured');
      return;
    }
    
    console.log(`📊 Spreadsheet ID: ${spreadsheetId}`);
    
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    console.log(`📋 Spreadsheet Name: ${spreadsheet.getName()}`);
    console.log(`🔗 Spreadsheet URL: ${spreadsheet.getUrl()}`);
    
    // Check main contract tracking sheet
    console.log('\n🔍 Checking Main Contract Tracking Sheet...');
    const mainSheet = spreadsheet.getSheetByName(SPREADSHEET_CONFIG.TAB_NAME);
    if (mainSheet) {
      console.log(`✓ Main sheet found: ${SPREADSHEET_CONFIG.TAB_NAME}`);
      const lastRow = mainSheet.getLastRow();
      console.log(`📊 Total rows: ${lastRow} (${lastRow - 1} data rows)`);
      
      if (lastRow > 1) {
        const dataRange = mainSheet.getRange(2, 1, lastRow - 1, mainSheet.getLastColumn());
        const values = dataRange.getValues();
        
        console.log(`🔍 Searching ${values.length} rows for message ID: ${messageId}`);
        
        let found = false;
        for (let i = 0; i < values.length; i++) {
          const row = values[i];
          const rowMessageId = row[12]; // Column M (index 12) is Message ID
          
          if (rowMessageId === messageId) {
            found = true;
            console.log(`✅ FOUND in main sheet at row ${i + 2}:`);
            console.log(`   Date: ${row[0]}`);
            console.log(`   Tool: ${row[1]}`);
            console.log(`   Sender: ${row[2]}`);
            console.log(`   Recipient: ${row[3]}`);
            console.log(`   Subject: ${row[4]}`);
            console.log(`   Status: ${row[9]}`);
            console.log(`   Message ID: ${row[12]}`);
            break;
          }
        }
        
        if (!found) {
          console.log(`❌ Message ID NOT FOUND in main sheet`);
          
          // Show sample message IDs for comparison
          console.log('\n📋 Sample message IDs from main sheet (first 5 rows):');
          for (let i = 0; i < Math.min(5, values.length); i++) {
            const rowMessageId = values[i][12];
            console.log(`   Row ${i + 2}: "${rowMessageId}"`);
          }
        }
      } else {
        console.log('📋 Main sheet is empty (only header row)');
      }
    } else {
      console.log(`❌ Main sheet not found: ${SPREADSHEET_CONFIG.TAB_NAME}`);
    }
    
    // Check processed messages sheet
    console.log('\n🔍 Checking Processed Messages Sheet...');
    const processedSheet = spreadsheet.getSheetByName(SPREADSHEET_CONFIG.PROCESSED_TAB_NAME);
    if (processedSheet) {
      console.log(`✓ Processed sheet found: ${SPREADSHEET_CONFIG.PROCESSED_TAB_NAME}`);
      const lastRow = processedSheet.getLastRow();
      console.log(`📊 Total rows: ${lastRow} (${lastRow - 1} data rows)`);
      
      if (lastRow > 1) {
        const dataRange = processedSheet.getRange(2, 1, lastRow - 1, processedSheet.getLastColumn());
        const values = dataRange.getValues();
        
        console.log(`🔍 Searching ${values.length} rows for message ID: ${messageId}`);
        
        let found = false;
        for (let i = 0; i < values.length; i++) {
          const row = values[i];
          const rowMessageId = row[1]; // Column B (index 1) is Message ID in processed sheet
          
          if (rowMessageId === messageId) {
            found = true;
            console.log(`✅ FOUND in processed sheet at row ${i + 2}:`);
            console.log(`   Processing Date: ${row[0]}`);
            console.log(`   Message ID: ${row[1]}`);
            console.log(`   Content Key: ${row[2]}`);
            console.log(`   Subject: ${row[3]}`);
            console.log(`   Sender: ${row[4]}`);
            console.log(`   Status: ${row[5]}`);
            break;
          }
        }
        
        if (!found) {
          console.log(`❌ Message ID NOT FOUND in processed sheet`);
          
          // Show sample message IDs for comparison
          console.log('\n📋 Sample message IDs from processed sheet (first 5 rows):');
          for (let i = 0; i < Math.min(5, values.length); i++) {
            const rowMessageId = values[i][1];
            console.log(`   Row ${i + 2}: "${rowMessageId}"`);
          }
        }
      } else {
        console.log('📋 Processed sheet is empty (only header row)');
      }
    } else {
      console.log(`❌ Processed sheet not found: ${SPREADSHEET_CONFIG.PROCESSED_TAB_NAME}`);
    }
    
    // Test the search function itself
    console.log('\n🧪 Testing searchRecordByMessageId() function...');
    const searchResult = searchRecordByMessageId(messageId);
    if (searchResult) {
      console.log(`✅ searchRecordByMessageId() returned result:`);
      console.log(`   Row: ${searchResult.row}`);
      console.log(`   Subject: ${searchResult.subject}`);
      console.log(`   Status: ${searchResult.status}`);
      console.log(`   Message ID: ${searchResult.messageId}`);
    } else {
      console.log(`❌ searchRecordByMessageId() returned null`);
    }
    
    // Test the processed message check function
    console.log('\n🧪 Testing isMessageProcessedInSpreadsheet() function...');
    const isProcessed = isMessageProcessedInSpreadsheet(messageId);
    console.log(`   Result: ${isProcessed ? 'TRUE (already processed)' : 'FALSE (not processed)'}`);
    
    console.log('\n=== DEBUG COMPLETE ===');
    
  } catch (error) {
    console.error('Error debugging message ID search:', error);
    console.error('Error details:', error.stack);
  }
}

/**
 * Search for partial message ID matches in spreadsheet
 * スプレッドシートで部分的なメッセージIDマッチを検索
 * 
 * @param {string} partialMessageId - Partial message ID to search for
 */
function searchPartialMessageId(partialMessageId) {
  console.log(`=== SEARCHING FOR PARTIAL MESSAGE ID: ${partialMessageId} ===`);
  
  try {
    const spreadsheetId = getProperty('SPREADSHEET_ID', false);
    if (!spreadsheetId) {
      console.log('❌ No spreadsheet ID configured');
      return;
    }
    
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    
    // Search in main sheet
    console.log('\n🔍 Searching Main Contract Tracking Sheet...');
    const mainSheet = spreadsheet.getSheetByName(SPREADSHEET_CONFIG.TAB_NAME);
    if (mainSheet && mainSheet.getLastRow() > 1) {
      const dataRange = mainSheet.getRange(2, 1, mainSheet.getLastRow() - 1, mainSheet.getLastColumn());
      const values = dataRange.getValues();
      
      console.log(`🔍 Searching ${values.length} rows for partial match: ${partialMessageId}`);
      
      let matches = [];
      for (let i = 0; i < values.length; i++) {
        const row = values[i];
        const rowMessageId = row[12]; // Column M (index 12) is Message ID
        
        if (rowMessageId && rowMessageId.includes(partialMessageId)) {
          matches.push({
            row: i + 2,
            messageId: rowMessageId,
            subject: row[4],
            date: row[0],
            status: row[9]
          });
        }
      }
      
      if (matches.length > 0) {
        console.log(`✅ Found ${matches.length} partial matches in main sheet:`);
        matches.forEach(match => {
          console.log(`   Row ${match.row}: ${match.messageId}`);
          console.log(`     Subject: ${match.subject}`);
          console.log(`     Date: ${match.date}`);
          console.log(`     Status: ${match.status}`);
        });
      } else {
        console.log(`❌ No partial matches found in main sheet`);
      }
    }
    
    // Search in processed sheet
    console.log('\n🔍 Searching Processed Messages Sheet...');
    const processedSheet = spreadsheet.getSheetByName(SPREADSHEET_CONFIG.PROCESSED_TAB_NAME);
    if (processedSheet && processedSheet.getLastRow() > 1) {
      const dataRange = processedSheet.getRange(2, 1, processedSheet.getLastRow() - 1, processedSheet.getLastColumn());
      const values = dataRange.getValues();
      
      let matches = [];
      for (let i = 0; i < values.length; i++) {
        const row = values[i];
        const rowMessageId = row[1]; // Column B (index 1) is Message ID in processed sheet
        
        if (rowMessageId && rowMessageId.includes(partialMessageId)) {
          matches.push({
            row: i + 2,
            messageId: rowMessageId,
            subject: row[3],
            date: row[0],
            status: row[5]
          });
        }
      }
      
      if (matches.length > 0) {
        console.log(`✅ Found ${matches.length} partial matches in processed sheet:`);
        matches.forEach(match => {
          console.log(`   Row ${match.row}: ${match.messageId}`);
          console.log(`     Subject: ${match.subject}`);
          console.log(`     Date: ${match.date}`);
          console.log(`     Status: ${match.status}`);
        });
      } else {
        console.log(`❌ No partial matches found in processed sheet`);
      }
    }
    
    console.log('\n=== PARTIAL SEARCH COMPLETE ===');
    
  } catch (error) {
    console.error('Error searching for partial message ID:', error);
  }
}

/**
 * Debug Dropbox Sign email processing specifically
 * Dropbox Signメール処理の専用デバッグ
 */
function debugDropboxSignEmails() {
  console.log('=== DEBUGGING DROPBOX SIGN EMAIL PROCESSING ===');
  
  try {
    // Check if Dropbox Sign integration is enabled
    if (!CONFIG.DROPBOX_SIGN_INTEGRATION?.ENABLE) {
      console.log('❌ Dropbox Sign integration is disabled in configuration');
      console.log('Enable by setting CONFIG.DROPBOX_SIGN_INTEGRATION.ENABLE = true');
      return;
    }
    
    console.log('✓ Dropbox Sign integration enabled');
    console.log(`Sender patterns: ${CONFIG.DROPBOX_SIGN_INTEGRATION.SENDER_PATTERNS?.length || 0}`);
    console.log(`Subject patterns: ${CONFIG.DROPBOX_SIGN_INTEGRATION.SUBJECT_PATTERNS?.length || 0}`);
    console.log(`Reply-to patterns: ${CONFIG.DROPBOX_SIGN_INTEGRATION.REPLY_TO_PATTERNS?.length || 0}`);
    console.log(`Detection mode: ${CONFIG.DROPBOX_SIGN_INTEGRATION.DETECTION_MODE || 'sender_or_subject'}`);
    console.log(`PDF attachment required: ${CONFIG.DROPBOX_SIGN_INTEGRATION.REQUIRE_PDF_ATTACHMENT ? 'YES' : 'NO'}`);
    
    // Build Dropbox Sign search query (same as in processEmails)
    const dropboxSignSearchTerms = [];
    
    // Add sender-based searches
    if (CONFIG.DROPBOX_SIGN_INTEGRATION.SENDER_PATTERNS?.length > 0) {
      dropboxSignSearchTerms.push('"via Dropbox Sign"');
      dropboxSignSearchTerms.push('"via HelloSign"');
      dropboxSignSearchTerms.push("'Dropbox Sign' via");
      dropboxSignSearchTerms.push("'HelloSign' via");
    }
    
    // Add reply-to based searches
    if (CONFIG.DROPBOX_SIGN_INTEGRATION.REPLY_TO_PATTERNS?.length > 0) {
      dropboxSignSearchTerms.push('from:hellosign.com');
    }
    
    // Add subject-based searches
    if (CONFIG.DROPBOX_SIGN_INTEGRATION.SUBJECT_PATTERNS?.length > 0) {
      dropboxSignSearchTerms.push('subject:"You\'ve been copied on"');
      dropboxSignSearchTerms.push('subject:"signed by"');
      dropboxSignSearchTerms.push('subject:"has been completed"');
      dropboxSignSearchTerms.push('subject:"You just signed"');
    }
    
    let query = `(${dropboxSignSearchTerms.join(' OR ')}) -label:${CONFIG.GMAIL_LABEL}`;
    
    console.log(`Gmail search query: ${query}`);
    
    const threads = GmailApp.search(query, 0, 10);
    console.log(`Found ${threads.length} potential Dropbox Sign email threads`);
    
    if (threads.length === 0) {
      console.log('No Dropbox Sign emails found. Possible reasons:');
      console.log('1. No Dropbox Sign/HelloSign completion emails in your Gmail');
      console.log('2. All emails already processed (have the Contract_Processed label)');
      console.log('3. Sender/subject patterns do not match actual emails');
      console.log('4. Detection mode requires both sender AND subject match');
      
      // Try broader search
      const broadQuery = 'subject:"You\'ve been copied on" OR "via Dropbox Sign" OR "via HelloSign"';
      const broadThreads = GmailApp.search(broadQuery, 0, 5);
      console.log(`Broader search: ${broadThreads.length} results`);
      
      if (broadThreads.length > 0) {
        console.log('Sample Dropbox Sign/HelloSign email details:');
        broadThreads.slice(0, 3).forEach((thread, index) => {
          const messages = thread.getMessages();
          if (messages.length > 0) {
            const message = messages[0];
            console.log(`${index + 1}. From: ${message.getFrom()}`);
            console.log(`   To: ${message.getTo()}`);
            console.log(`   Subject: ${message.getSubject()}`);
            console.log(`   Reply-To: ${message.getReplyTo()}`);
          }
        });
      }
      
      return;
    }
    
    // Test message source detection and pattern matching on Dropbox Sign emails
    console.log('\nTesting message source detection and pattern matching:');
    threads.slice(0, 5).forEach((thread, index) => {
      const messages = thread.getMessages();
      if (messages.length === 0) {
        console.log(`\n--- Dropbox Sign Email ${index + 1} ---`);
        console.log('No messages in thread');
        return;
      }
      
      const message = messages[0]; // Get first message in thread
      const subject = message.getSubject();
      const sender = message.getFrom();
      const to = message.getTo();
      const replyTo = message.getReplyTo();
      
      console.log(`\n--- Dropbox Sign Email ${index + 1} ---`);
      console.log(`From: ${sender}`);
      console.log(`To: ${to}`);
      console.log(`Reply-To: ${replyTo}`);
      console.log(`Subject: "${subject}"`);
      
      // Test message source detection
      const messageSource = detectMessageSource(message);
      console.log(`Source detection: ${messageSource.type}`);
      console.log(`Details: ${JSON.stringify(messageSource.details)}`);
      
      // Test pattern matching
      const patternResult = checkSubjectPattern(subject, messageSource.type);
      console.log(`Pattern match: ${patternResult.isMatch ? '✅ MATCHES' : '❌ NO MATCH'}`);
      
      if (patternResult.isMatch && patternResult.matchedPattern) {
        console.log(`Matched pattern: ${patternResult.matchedPattern}`);
      }
      
      if (!patternResult.isMatch && patternResult.checkedPatterns) {
        console.log(`Checked ${patternResult.checkedPatterns.length} patterns`);
      }
    });
    
    console.log('\n=== Dropbox Sign Debug Complete ===');
    
  } catch (error) {
    console.error('Error debugging Dropbox Sign emails:', error);
    throw error;
  }
}

/**
 * Test Dropbox Sign email processing with spreadsheet logging
 * Dropbox Signメール処理とスプレッドシートロギングのテスト
 */
function testDropboxSignSpreadsheetLogging() {
  console.log('=== TESTING DROPBOX SIGN SPREADSHEET LOGGING ===');
  
  try {
    // Check if spreadsheet logging is enabled
    console.log(`Spreadsheet logging enabled: ${CONFIG.ENABLE_SPREADSHEET_LOGGING}`);
    
    if (!CONFIG.ENABLE_SPREADSHEET_LOGGING) {
      console.log('❌ Spreadsheet logging is disabled');
      console.log('Enable by setting CONFIG.ENABLE_SPREADSHEET_LOGGING = true in main.js');
      return;
    }
    
    // Check if spreadsheet exists
    const spreadsheetId = getProperty('SPREADSHEET_ID', false);
    console.log(`Spreadsheet ID: ${spreadsheetId || 'NOT SET'}`);
    
    if (!spreadsheetId) {
      console.log('Creating spreadsheet...');
      const newSpreadsheetId = createOrGetSpreadsheet();
      console.log(`New spreadsheet created: ${newSpreadsheetId}`);
    }
    
    // Test with mock Dropbox Sign email data
    const mockDropboxSignData = {
      date: new Date(),
      sender: "'Dropbox Sign' via investment-abkk <investment-abkk@animocabrands.com>",
      recipient: 'investment-abkk@animocabrands.com',
      subject: "You've been copied on Shareholder Resolution - Approval of Fundraising - Convertible Notes and Warrants Issuance (ABKK) - signed by Yusuke Jindo",
      body: 'This is a test email body for Dropbox Sign organizational forwarding. The document has been completed and signed.',
      messageId: `dropbox-sign-test-${new Date().getTime()}`,
      attachmentCount: 1,
      pdfCount: 1,
      pdfFilename: 'Shareholder_Resolution_ABKK.pdf',
      pdfDirectLinks: 'https://drive.google.com/file/d/1TEST_DROPBOX_SIGN/view',
      status: 'Success',
      slackNotified: true,
      error: null
    };
    
    console.log('\n📝 Testing spreadsheet record addition:');
    console.log(`Message ID: ${mockDropboxSignData.messageId}`);
    console.log(`Sender: ${mockDropboxSignData.sender}`);
    console.log(`Subject: ${mockDropboxSignData.subject}`);
    
    // Test contract tool extraction
    const contractTool = extractContractTool(mockDropboxSignData.sender);
    console.log(`Extracted contract tool: ${contractTool}`);
    
    // Test contract type extraction
    const contractType = extractContractType(mockDropboxSignData.subject, mockDropboxSignData.body);
    console.log(`Extracted contract type: ${contractType}`);
    
    // Test contract party extraction
    const contractParty = extractContractParty(mockDropboxSignData.subject, mockDropboxSignData.body);
    console.log(`Extracted contract party: ${contractParty}`);
    
    // Add record to spreadsheet
    console.log('\n📊 Adding record to spreadsheet...');
    const addResult = addEmailRecord(mockDropboxSignData);
    
    if (addResult) {
      console.log('✅ Record added successfully to spreadsheet');
      
      // Search for the record
      console.log('\n🔍 Searching for the added record...');
      const searchResult = searchRecordByMessageId(mockDropboxSignData.messageId);
      
      if (searchResult) {
        console.log('✅ Record found in spreadsheet:');
        console.log(`  - Row: ${searchResult.row}`);
        console.log(`  - Status: ${searchResult.status}`);
        console.log(`  - Contract Tool: ${extractContractTool(searchResult.sender)}`);
      } else {
        console.log('❌ Record not found in spreadsheet');
      }
      
    } else {
      console.log('❌ Failed to add record to spreadsheet');
    }
    
    // Get spreadsheet URL for manual verification
    if (spreadsheetId) {
      const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
      console.log(`\n🔗 Spreadsheet URL: ${spreadsheet.getUrl()}`);
      console.log('Check the spreadsheet manually to verify the Dropbox Sign entry');
    }
    
    console.log('\n=== Dropbox Sign Spreadsheet Test Complete ===');
    
  } catch (error) {
    console.error('❌ Error testing Dropbox Sign spreadsheet logging:', error);
    console.error('Error details:', error.stack);
    throw error;
  }
}

/**
 * Debug specific Dropbox Sign email processing workflow
 * 特定のDropbox Signメール処理ワークフローのデバッグ
 */
function debugDropboxSignWorkflow() {
  console.log('=== DEBUGGING DROPBOX SIGN PROCESSING WORKFLOW ===');
  
  try {
    // Step 1: Check configuration
    console.log('Step 1: Checking configuration...');
    console.log(`ENABLE_SPREADSHEET_LOGGING: ${CONFIG.ENABLE_SPREADSHEET_LOGGING}`);
    console.log(`DROPBOX_SIGN_INTEGRATION.ENABLE: ${CONFIG.DROPBOX_SIGN_INTEGRATION?.ENABLE}`);
    
    // Step 2: Search for Dropbox Sign emails
    console.log('\nStep 2: Searching for Dropbox Sign emails...');
    const query = 'subject:"You\'ve been copied on" OR "via Dropbox Sign" OR "via HelloSign"';
    const threads = GmailApp.search(query, 0, 5);
    console.log(`Found ${threads.length} potential Dropbox Sign emails`);
    
    if (threads.length === 0) {
      console.log('No Dropbox Sign emails found for testing');
      return;
    }
    
    // Step 3: Test processing workflow on first email
    console.log('\nStep 3: Testing processing workflow...');
    const thread = threads[0];
    const messages = thread.getMessages();
    
    if (messages.length > 0) {
      const message = messages[0];
      const sender = message.getFrom();
      const subject = message.getSubject();
      const messageId = message.getId();
      
      console.log(`\nTesting message:`);
      console.log(`From: ${sender}`);
      console.log(`Subject: ${subject}`);
      console.log(`Message ID: ${messageId}`);
      
      // Step 4: Test message source detection
      console.log('\nStep 4: Testing message source detection...');
      const messageSource = detectMessageSource(message);
      console.log(`Message source: ${messageSource.type}`);
      console.log(`Detection details: ${JSON.stringify(messageSource.details)}`);
      
      // Step 5: Test subject pattern matching
      console.log('\nStep 5: Testing subject pattern matching...');
      const patternResult = checkSubjectPattern(subject, messageSource.type);
      console.log(`Pattern match: ${patternResult.isMatch ? '✅ MATCHES' : '❌ NO MATCH'}`);
      
      if (patternResult.isMatch) {
        console.log(`Matched pattern: ${patternResult.matchedPattern}`);
        
        // Step 6: Check if already processed
        console.log('\nStep 6: Checking if already processed...');
        const alreadyProcessed = isMessageAlreadyProcessed(message);
        console.log(`Already processed: ${alreadyProcessed ? 'YES' : 'NO'}`);
        
        if (!alreadyProcessed) {
          console.log('\nStep 7: This message would be processed and logged to spreadsheet');
          console.log('Run processEmails() to actually process this message');
        } else {
          console.log('\nStep 7: This message was already processed');
          
          // Check if it exists in spreadsheet
          const recordInSpreadsheet = searchRecordByMessageId(messageId);
          if (recordInSpreadsheet) {
            console.log('✅ Record exists in spreadsheet');
          } else {
            console.log('⚠️ Record NOT found in spreadsheet (possible issue)');
          }
        }
      }
    }
    
    console.log('\n=== Dropbox Sign Workflow Debug Complete ===');
    
  } catch (error) {
    console.error('Error debugging Dropbox Sign workflow:', error);
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

/**
 * Show statistics for skipped emails
 * スキップされたメールの統計を表示
 */
function showSkippedEmailStats() {
  console.log('=== Skipped Email Statistics ===');
  
  try {
    // Get skipped label
    const skipLabel = GmailApp.getUserLabelByName(CONFIG.GMAIL_SKIP_LABEL);
    
    if (!skipLabel) {
      console.log('No skipped emails found (label does not exist)');
      return;
    }
    
    // Get threads with skip label
    const threads = skipLabel.getThreads(0, 100); // Get up to 100 threads
    console.log(`Total skipped threads: ${threads.length}`);
    
    if (threads.length === 0) {
      return;
    }
    
    // Analyze skipped emails
    const stats = {
      total: 0,
      bySender: {},
      byReason: {
        pattern_mismatch: 0,
        no_pdf_attachment: 0,
        unknown_sender: 0,
        filtered_out: 0
      },
      recentExamples: []
    };
    
    threads.forEach((thread, index) => {
      const messages = thread.getMessages();
      messages.forEach(message => {
        stats.total++;
        
        const sender = message.getFrom();
        const subject = message.getSubject();
        const date = message.getDate();
        
        // Count by sender
        const senderEmail = sender.match(/<(.+?)>/) ? sender.match(/<(.+?)>/)[1] : sender;
        stats.bySender[senderEmail] = (stats.bySender[senderEmail] || 0) + 1;
        
        // Add recent examples (first 5)
        if (stats.recentExamples.length < 5) {
          stats.recentExamples.push({
            date: Utilities.formatDate(date, 'JST', 'yyyy-MM-dd HH:mm'),
            sender: sender,
            subject: subject
          });
        }
      });
    });
    
    // Display statistics
    console.log(`\nTotal skipped messages: ${stats.total}`);
    
    console.log('\nTop senders with skipped emails:');
    const sortedSenders = Object.entries(stats.bySender)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 10);
    
    sortedSenders.forEach(([sender, count]) => {
      console.log(`  ${sender}: ${count} messages`);
    });
    
    console.log('\nRecent skipped emails:');
    stats.recentExamples.forEach(example => {
      console.log(`  ${example.date} - ${example.sender}`);
      console.log(`    Subject: ${example.subject}`);
    });
    
    console.log('\nTo manually review skipped emails:');
    console.log('  1. Go to Gmail');
    console.log(`  2. Search for: label:${CONFIG.GMAIL_SKIP_LABEL}`);
    console.log('  3. Review if any patterns need to be added');
    
  } catch (error) {
    console.error('Error getting skipped email statistics:', error);
  }
}

/**
 * Verify specific message ID in both sheets and check for discrepancies
 * 両方のシートで特定のメッセージIDを検証し、不一致をチェック
 * 
 * @param {string} messageId - The message ID to verify
 */
function verifyMessageIdConsistency(messageId) {
  console.log(`=== VERIFYING MESSAGE ID CONSISTENCY: ${messageId} ===`);
  
  try {
    const spreadsheetId = getProperty('SPREADSHEET_ID', false);
    if (!spreadsheetId) {
      console.log('❌ No spreadsheet ID configured');
      return;
    }
    
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    console.log(`📊 Spreadsheet URL: ${spreadsheet.getUrl()}`);
    
    // Check 1: searchRecordByMessageId() function
    console.log('\n🔍 Test 1: searchRecordByMessageId() function');
    const searchResult = searchRecordByMessageId(messageId);
    const foundInMainSheet = searchResult !== null;
    console.log(`Result: ${foundInMainSheet ? '✅ FOUND' : '❌ NOT FOUND'}`);
    
    if (foundInMainSheet) {
      console.log(`   Row: ${searchResult.row}`);
      console.log(`   Subject: ${searchResult.subject}`);
      console.log(`   Date: ${searchResult.date}`);
      console.log(`   Status: ${searchResult.status}`);
    }
    
    // Check 2: isMessageProcessedInSpreadsheet() function
    console.log('\n🔍 Test 2: isMessageProcessedInSpreadsheet() function');
    const isProcessed = isMessageProcessedInSpreadsheet(messageId);
    console.log(`Result: ${isProcessed ? '✅ FOUND (processed)' : '❌ NOT FOUND (not processed)'}`);
    
    // Check 3: Manual search in main sheet
    console.log('\n🔍 Test 3: Manual search in main contract sheet');
    const mainSheet = spreadsheet.getSheetByName(SPREADSHEET_CONFIG.TAB_NAME);
    let foundInManualMain = false;
    
    if (mainSheet && mainSheet.getLastRow() > 1) {
      const dataRange = mainSheet.getRange(2, 1, mainSheet.getLastRow() - 1, mainSheet.getLastColumn());
      const values = dataRange.getValues();
      
      for (let i = 0; i < values.length; i++) {
        const row = values[i];
        const rowMessageId = row[12]; // Column M (index 12) is Message ID
        
        if (rowMessageId === messageId) {
          foundInManualMain = true;
          console.log(`✅ FOUND in main sheet at row ${i + 2}`);
          console.log(`   Message ID: "${rowMessageId}"`);
          console.log(`   Subject: ${row[4]}`);
          console.log(`   Date: ${row[0]}`);
          console.log(`   Status: ${row[9]}`);
          break;
        }
      }
    }
    
    if (!foundInManualMain) {
      console.log('❌ NOT FOUND in manual main sheet search');
    }
    
    // Check 4: Manual search in processed sheet
    console.log('\n🔍 Test 4: Manual search in processed messages sheet');
    const processedSheet = spreadsheet.getSheetByName(SPREADSHEET_CONFIG.PROCESSED_TAB_NAME);
    let foundInManualProcessed = false;
    
    if (processedSheet && processedSheet.getLastRow() > 1) {
      const dataRange = processedSheet.getRange(2, 1, processedSheet.getLastRow() - 1, processedSheet.getLastColumn());
      const values = dataRange.getValues();
      
      for (let i = 0; i < values.length; i++) {
        const row = values[i];
        const rowMessageId = row[1]; // Column B (index 1) is Message ID in processed sheet
        
        if (rowMessageId === messageId) {
          foundInManualProcessed = true;
          console.log(`✅ FOUND in processed sheet at row ${i + 2}`);
          console.log(`   Message ID: "${rowMessageId}"`);
          console.log(`   Subject: ${row[3]}`);
          console.log(`   Processing Date: ${row[0]}`);
          console.log(`   Status: ${row[5]}`);
          break;
        }
      }
    }
    
    if (!foundInManualProcessed) {
      console.log('❌ NOT FOUND in manual processed sheet search');
    }
    
    // Summary and analysis
    console.log('\n📊 CONSISTENCY ANALYSIS:');
    console.log(`searchRecordByMessageId(): ${foundInMainSheet ? 'FOUND' : 'NOT FOUND'}`);
    console.log(`isMessageProcessedInSpreadsheet(): ${isProcessed ? 'FOUND' : 'NOT FOUND'}`);
    console.log(`Manual main sheet search: ${foundInManualMain ? 'FOUND' : 'NOT FOUND'}`);
    console.log(`Manual processed sheet search: ${foundInManualProcessed ? 'FOUND' : 'NOT FOUND'}`);
    
    // Check for inconsistencies
    const hasInconsistency = (foundInMainSheet !== foundInManualMain) || 
                            (isProcessed !== foundInManualProcessed);
    
    if (hasInconsistency) {
      console.log('\n⚠️  INCONSISTENCY DETECTED!');
      console.log('This suggests there may be a bug in one of the search functions.');
      
      if (foundInMainSheet !== foundInManualMain) {
        console.log('- Discrepancy between searchRecordByMessageId() and manual main sheet search');
      }
      
      if (isProcessed !== foundInManualProcessed) {
        console.log('- Discrepancy between isMessageProcessedInSpreadsheet() and manual processed sheet search');
      }
    } else {
      console.log('\n✅ NO INCONSISTENCIES DETECTED');
      console.log('All search methods returned consistent results.');
    }
    
    // User reported issue check
    if (isProcessed && !foundInMainSheet && !foundInManualMain) {
      console.log('\n🚨 USER REPORTED ISSUE CONFIRMED:');
      console.log('Message shows as processed but cannot be found in main contract sheet');
      console.log('This means the message is only in the processed messages tracking sheet');
      console.log('but not in the main contract data sheet.');
      
      console.log('\nPOSSIBLE CAUSES:');
      console.log('1. Message was marked as processed but failed to add to main sheet');
      console.log('2. Message was deleted from main sheet but not from processed sheet');
      console.log('3. There was an error during the addEmailRecord() process');
      console.log('4. The message is in a different tab or sheet');
    }
    
    return {
      foundInMainSheet,
      isProcessed,
      foundInManualMain,
      foundInManualProcessed,
      hasInconsistency,
      userIssueConfirmed: isProcessed && !foundInMainSheet && !foundInManualMain
    };
    
  } catch (error) {
    console.error('Error verifying message ID consistency:', error);
    console.error('Error details:', error.stack);
  }
}

/**
 * Test the specific user-reported message ID issue
 * ユーザー報告の特定メッセージID問題をテスト
 */
function testUserReportedMessageId() {
  console.log('=== TESTING USER REPORTED MESSAGE ID ISSUE ===');
  
  const problematicMessageId = '197bf725ef40ab70';
  
  console.log(`Testing message ID: ${problematicMessageId}`);
  console.log('User reports: "appears in logs as found but not visible in spreadsheet"');
  
  return verifyMessageIdConsistency(problematicMessageId);
}

/**
 * Clean up old skip labels (optional maintenance function)
 * 古いスキップラベルをクリーンアップ（オプションのメンテナンス機能）
 * 
 * @param {number} daysOld - Remove skip labels older than this many days
 * @param {boolean} dryRun - If true, only show what would be removed
 */
function cleanupOldSkipLabels(daysOld = 30, dryRun = true) {
  console.log(`=== Cleanup Skip Labels (${dryRun ? 'DRY RUN' : 'ACTUAL'}) ===`);
  
  try {
    const skipLabel = GmailApp.getUserLabelByName(CONFIG.GMAIL_SKIP_LABEL);
    
    if (!skipLabel) {
      console.log('No skip label found');
      return;
    }
    
    const cutoffDate = new Date();
    cutoffDate.setDate(cutoffDate.getDate() - daysOld);
    
    const threads = skipLabel.getThreads();
    let removedCount = 0;
    
    threads.forEach(thread => {
      const lastMessageDate = thread.getLastMessageDate();
      
      if (lastMessageDate < cutoffDate) {
        if (!dryRun) {
          thread.removeLabel(skipLabel);
        }
        removedCount++;
        
        if (removedCount <= 5) {
          console.log(`  ${dryRun ? 'Would remove' : 'Removed'} label from: ${thread.getFirstMessageSubject()}`);
          console.log(`    Last message: ${Utilities.formatDate(lastMessageDate, 'JST', 'yyyy-MM-dd')}`);
        }
      }
    });
    
    console.log(`\n${dryRun ? 'Would remove' : 'Removed'} skip label from ${removedCount} threads older than ${daysOld} days`);
    
    if (dryRun && removedCount > 0) {
      console.log('\nTo actually remove labels, run: cleanupOldSkipLabels(30, false)');
    }
    
  } catch (error) {
    console.error('Error cleaning up skip labels:', error);
  }
}
// Aliases for convenience
const checkSkipped = showSkippedEmailStats;
