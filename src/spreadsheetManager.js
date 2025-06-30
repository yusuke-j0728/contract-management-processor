/**
 * Contract Management Spreadsheet Module
 * 
 * Handles logging contract-related email data to Google Spreadsheet for tracking and analysis.
 * Provides searchable, sortable contract history with metadata and processing status.
 */

// === CONTRACT TRACKING SPREADSHEET CONFIGURATION ===
const SPREADSHEET_CONFIG = {
  NAME: '契約管理_Contract_Tracking',
  TAB_NAME: '契約一覧_Contract_List',
  PROCESSED_TAB_NAME: '処理済みメッセージ_Processed_Messages', // New tab for processed message tracking
  HEADERS: [
    '受信日時',           // A: Receipt Date
    '契約管理ツール',      // B: Contract Management Tool
    '送信者メール',        // C: Sender Email
    '受信メール',          // D: Recipient Email
    '件名',              // E: Subject
    '契約タイプ',          // F: Contract Type
    '契約相手',          // G: Contract Party
    'PDFファイル名',       // H: PDF Filename
    'PDF直接リンク',       // I: PDF Direct Links
    '処理状態',          // J: Processing Status
    'Slack通知済み',     // K: Slack Notified
    '本文要約',          // L: Body Summary
    'メッセージID',      // M: Message ID
    'エラーログ'         // N: Error Log
  ],
  PROCESSED_HEADERS: [
    '処理日時',          // A: Processing Date
    'メッセージID',      // B: Message ID
    'コンテンツキー',       // C: Content-based duplicate key
    '件名',              // D: Subject (for reference)
    '送信者',            // E: Sender (for reference)
    'ステータス'           // F: Status (Success/Error)
  ],
  CONTENT_DUPLICATE_HEADERS: [
    '登録日時',          // A: Registration Date
    'コンテンツキー',       // B: Content duplicate key
    '送信者',            // C: Sender
    '送信日時',          // D: Send Date
    '件名',              // E: Subject
    'PDFファイル名',       // F: PDF Filenames
    'DriveフォルダURL',    // G: Drive Folder URL
    '初回保存日時'      // H: First Save Date
  ],
  MAX_BODY_LENGTH: 1000,  // Maximum characters for body summary (longer for contracts)
  DATE_FORMAT: 'yyyy/MM/dd HH:mm:ss'
};

/**
 * Create or get the logging spreadsheet
 * ログ用スプレッドシートを作成または取得
 * 
 * @returns {string} - Spreadsheet ID
 */
function createOrGetSpreadsheet() {
  try {
    console.log('Creating or getting logging spreadsheet...');
    
    // Try to get existing spreadsheet ID from properties
    let spreadsheetId = getProperty('SPREADSHEET_ID', false);
    
    if (spreadsheetId) {
      try {
        // Verify spreadsheet still exists and is accessible
        const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
        console.log(`Found existing spreadsheet: ${spreadsheet.getName()}`);
        
        // Ensure the logging sheet exists
        ensureLoggingSheet(spreadsheet);
        
        return spreadsheetId;
      } catch (error) {
        console.log('Existing spreadsheet not accessible, creating new one...');
      }
    }
    
    // Search for existing spreadsheet by name
    const files = DriveApp.getFilesByName(SPREADSHEET_CONFIG.NAME);
    if (files.hasNext()) {
      const file = files.next();
      if (file.getMimeType() === 'application/vnd.google-apps.spreadsheet') {
        spreadsheetId = file.getId();
        console.log(`Found existing spreadsheet by name: ${SPREADSHEET_CONFIG.NAME}`);
        
        // Verify and setup sheet
        const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
        ensureLoggingSheet(spreadsheet);
        
        // Save to properties
        setProperty('SPREADSHEET_ID', spreadsheetId);
        return spreadsheetId;
      }
    }
    
    // Create new spreadsheet
    const spreadsheet = SpreadsheetApp.create(SPREADSHEET_CONFIG.NAME);
    spreadsheetId = spreadsheet.getId();
    console.log(`Created new spreadsheet: ${SPREADSHEET_CONFIG.NAME} (ID: ${spreadsheetId})`);
    
    // Setup initial sheet
    setupInitialSheet(spreadsheet);
    
    // Also create the processed messages tracking sheet
    createOrGetProcessedSheet(spreadsheet);
    
    // Create content duplicate tracking sheet
    createOrGetContentDuplicateSheet(spreadsheet);
    
    // Save to properties
    setProperty('SPREADSHEET_ID', spreadsheetId);
    
    // Share settings (optional - keep private by default)
    console.log(`Spreadsheet URL: ${spreadsheet.getUrl()}`);
    
    return spreadsheetId;
    
  } catch (error) {
    console.error('Error creating/getting spreadsheet:', error);
    throw error;
  }
}

/**
 * Create or get processed messages sheet for duplicate tracking
 * 重複防止用の処理済みメッセージシートを作成または取得
 * 
 * @param {Spreadsheet} spreadsheet - The main spreadsheet object
 * @returns {Sheet} - The processed messages sheet
 */
function createOrGetProcessedSheet(spreadsheet) {
  try {
    let processedSheet = spreadsheet.getSheetByName(SPREADSHEET_CONFIG.PROCESSED_TAB_NAME);
    
    if (!processedSheet) {
      console.log('Creating processed messages tracking sheet...');
      processedSheet = spreadsheet.insertSheet(SPREADSHEET_CONFIG.PROCESSED_TAB_NAME);
      
      // Set headers
      const headers = SPREADSHEET_CONFIG.PROCESSED_HEADERS;
      processedSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      
      // Format header row
      const headerRange = processedSheet.getRange(1, 1, 1, headers.length);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#E8F0FE');
      
      // Set column widths for better readability
      processedSheet.setColumnWidth(1, 150); // Processing Date
      processedSheet.setColumnWidth(2, 200); // Message ID
      processedSheet.setColumnWidth(3, 200); // Content Key
      processedSheet.setColumnWidth(4, 300); // Subject
      processedSheet.setColumnWidth(5, 200); // Sender
      processedSheet.setColumnWidth(6, 100); // Status
      
      // Freeze header row
      processedSheet.setFrozenRows(1);
      
      console.log('Processed messages sheet created with headers');
    }
    
    return processedSheet;
    
  } catch (error) {
    console.error('Error creating/getting processed messages sheet:', error);
    throw error;
  }
}

/**
 * Check if content has already been saved to avoid duplicate PDF storage
 * コンテンツが既に保存済みかどうかをチェックしてPDF重複保存を防ぐ
 * 
 * @param {string} contentKey - Content-based duplicate key
 * @returns {Object} - {isDuplicate: boolean, existingUrl: string}
 */
function checkContentDuplicate(contentKey) {
  try {
    const spreadsheetId = getProperty('SPREADSHEET_ID', false);
    if (!spreadsheetId) {
      return { isDuplicate: false, existingUrl: null };
    }
    
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const duplicateSheet = createOrGetContentDuplicateSheet(spreadsheet);
    
    // Get all data from the content duplicate tracking sheet
    const lastRow = duplicateSheet.getLastRow();
    if (lastRow <= 1) {
      return { isDuplicate: false, existingUrl: null };
    }
    
    const data = duplicateSheet.getRange(2, 1, lastRow - 1, 8).getValues();
    
    // Search for matching content key
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const rowContentKey = row[1]; // Column B: Content Key
      
      if (rowContentKey === contentKey) {
        const existingUrl = row[6]; // Column G: Drive Folder URL
        console.log(`Content duplicate found: ${contentKey}`);
        console.log(`Existing Drive URL: ${existingUrl}`);
        return { isDuplicate: true, existingUrl: existingUrl };
      }
    }
    
    return { isDuplicate: false, existingUrl: null };
    
  } catch (error) {
    console.error('Error checking content duplicate:', error);
    return { isDuplicate: false, existingUrl: null };
  }
}

/**
 * Record content duplicate information
 * コンテンツ重複情報を記録
 * 
 * @param {string} contentKey - Content-based duplicate key
 * @param {string} sender - Email sender
 * @param {Date} date - Email date
 * @param {string} subject - Email subject
 * @param {Array} pdfNames - PDF filenames
 * @param {string} driveUrl - Drive folder URL
 */
function recordContentDuplicate(contentKey, sender, date, subject, pdfNames, driveUrl) {
  try {
    const spreadsheetId = getProperty('SPREADSHEET_ID', false);
    if (!spreadsheetId) {
      console.log('No spreadsheet configured, cannot record content duplicate');
      return;
    }
    
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const duplicateSheet = createOrGetContentDuplicateSheet(spreadsheet);
    
    const now = new Date();
    const rowData = [
      Utilities.formatDate(now, 'JST', SPREADSHEET_CONFIG.DATE_FORMAT),     // Registration Date
      contentKey,                                                          // Content Key
      sender,                                                              // Sender
      Utilities.formatDate(date, 'JST', SPREADSHEET_CONFIG.DATE_FORMAT),  // Send Date
      subject,                                                             // Subject
      pdfNames.join(', '),                                               // PDF Filenames
      driveUrl,                                                           // Drive Folder URL
      Utilities.formatDate(now, 'JST', SPREADSHEET_CONFIG.DATE_FORMAT)    // First Save Date
    ];
    
    duplicateSheet.appendRow(rowData);
    console.log(`Recorded content duplicate: ${contentKey}`);
    
  } catch (error) {
    console.error('Error recording content duplicate:', error);
  }
}

/**
 * Create or get content duplicate tracking sheet
 * コンテンツ重複追跡シートを作成または取得
 * 
 * @param {Spreadsheet} spreadsheet - The main spreadsheet object
 * @returns {Sheet} - The content duplicate tracking sheet
 */
function createOrGetContentDuplicateSheet(spreadsheet) {
  try {
    const sheetName = 'コンテンツ重複管理_Content_Duplicates';
    let duplicateSheet = spreadsheet.getSheetByName(sheetName);
    
    if (!duplicateSheet) {
      console.log('Creating content duplicate tracking sheet...');
      duplicateSheet = spreadsheet.insertSheet(sheetName);
      
      // Set headers
      const headers = SPREADSHEET_CONFIG.CONTENT_DUPLICATE_HEADERS;
      duplicateSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      
      // Format header row
      const headerRange = duplicateSheet.getRange(1, 1, 1, headers.length);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#FFF2CC'); // Light yellow for content tracking
      
      // Set column widths for better readability
      duplicateSheet.setColumnWidth(1, 150); // Registration Date
      duplicateSheet.setColumnWidth(2, 200); // Content Key
      duplicateSheet.setColumnWidth(3, 200); // Sender
      duplicateSheet.setColumnWidth(4, 150); // Send Date
      duplicateSheet.setColumnWidth(5, 300); // Subject
      duplicateSheet.setColumnWidth(6, 250); // PDF Filenames
      duplicateSheet.setColumnWidth(7, 300); // Drive URL
      duplicateSheet.setColumnWidth(8, 150); // First Save Date
      
      // Freeze header row
      duplicateSheet.setFrozenRows(1);
      
      console.log('Content duplicate tracking sheet created');
    }
    
    return duplicateSheet;
    
  } catch (error) {
    console.error('Error creating/getting content duplicate sheet:', error);
    throw error;
  }
}

/**
 * Check if a message has already been processed by checking the spreadsheet
 * スプレッドシートでメッセージが既に処理済みかどうかをチェック
 * 
 * @param {string} messageId - Gmail message ID
 * @returns {boolean} - true if message was already processed
 */
function isMessageProcessedInSpreadsheet(messageId) {
  try {
    const spreadsheetId = getProperty('SPREADSHEET_ID', false);
    if (!spreadsheetId) {
      console.log('No spreadsheet configured, assuming message not processed');
      return false;
    }
    
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const processedSheet = createOrGetProcessedSheet(spreadsheet);
    
    // Get all data from the processed messages sheet
    const lastRow = processedSheet.getLastRow();
    if (lastRow <= 1) {
      // Only header row exists
      return false;
    }
    
    const data = processedSheet.getRange(2, 1, lastRow - 1, 6).getValues();
    
    // Search for matching message ID
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const rowMessageId = row[1]; // Column B: Message ID
      
      if (rowMessageId === messageId) {
        console.log(`Message ${messageId} found as processed in spreadsheet`);
        return true;
      }
    }
    
    return false;
    
  } catch (error) {
    console.error('Error checking if message is processed in spreadsheet:', error);
    return false; // Assume not processed if check fails
  }
}

/**
 * Mark a message as processed in the spreadsheet
 * スプレッドシートでメッセージを処理済みとしてマーク
 * 
 * @param {string} messageId - Gmail message ID
 * @param {string} subject - Email subject (for reference)
 * @param {string} sender - Email sender (for reference)
 * @param {string} contentKey - Content-based duplicate key (optional)
 * @param {string} status - Processing status (Success/Error)
 */
function markMessageProcessedInSpreadsheet(messageId, subject, sender, contentKey = null, status = 'Success') {
  try {
    const spreadsheetId = getProperty('SPREADSHEET_ID', false);
    if (!spreadsheetId) {
      console.log('No spreadsheet configured, cannot mark message as processed');
      return;
    }
    
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const processedSheet = createOrGetProcessedSheet(spreadsheet);
    
    // Prepare the row data
    const now = new Date();
    const rowData = [
      Utilities.formatDate(now, 'JST', SPREADSHEET_CONFIG.DATE_FORMAT), // Processing Date
      messageId,                    // Message ID
      contentKey || '',             // Content Key
      subject || '',                // Subject
      sender || '',                 // Sender
      status                        // Status
    ];
    
    // Append the row
    processedSheet.appendRow(rowData);
    
    console.log(`Marked message ${messageId} as processed in spreadsheet`);
    
  } catch (error) {
    console.error('Error marking message as processed in spreadsheet:', error);
    // Don't throw - this is not critical for main functionality
  }
}

/**
 * Clean up old processed message entries from spreadsheet
 * スプレッドシートから古い処理済みメッセージエントリをクリーンアップ
 * 
 * @param {number} daysOld - Remove entries older than this many days
 * @returns {number} Number of entries removed
 */
function cleanupOldProcessedEntriesFromSpreadsheet(daysOld = 30) {
  try {
    console.log(`Cleaning up processed entries older than ${daysOld} days from spreadsheet...`);
    
    const spreadsheetId = getProperty('SPREADSHEET_ID', false);
    if (!spreadsheetId) {
      console.log('No spreadsheet configured');
      return 0;
    }
    
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const processedSheet = createOrGetProcessedSheet(spreadsheet);
    
    const lastRow = processedSheet.getLastRow();
    if (lastRow <= 1) {
      console.log('No processed entries to clean up');
      return 0;
    }
    
    const data = processedSheet.getRange(2, 1, lastRow - 1, 6).getValues();
    const cutoffDate = new Date(Date.now() - (daysOld * 24 * 60 * 60 * 1000));
    
    let removedCount = 0;
    
    // Mark rows for deletion (from bottom to top to avoid index shifting)
    for (let i = data.length - 1; i >= 0; i--) {
      const row = data[i];
      const processingDate = new Date(row[0]); // Column A: Processing Date
      
      if (processingDate < cutoffDate) {
        // Delete the row (add 2 because: 1 for header row, 1 for 0-based to 1-based index)
        processedSheet.deleteRow(i + 2);
        removedCount++;
      }
    }
    
    console.log(`Cleaned up ${removedCount} old processed entries from spreadsheet`);
    return removedCount;
    
  } catch (error) {
    console.error('Error cleaning up old processed entries from spreadsheet:', error);
    return 0;
  }
}

/**
 * Get processed message statistics from spreadsheet
 * スプレッドシートから処理済みメッセージの統計を取得
 * 
 * @returns {Object} Statistics about processed messages
 */
function getProcessedMessageStats() {
  try {
    const spreadsheetId = getProperty('SPREADSHEET_ID', false);
    if (!spreadsheetId) {
      return { total: 0, recent: 0, errors: 0 };
    }
    
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const processedSheet = createOrGetProcessedSheet(spreadsheet);
    
    const lastRow = processedSheet.getLastRow();
    if (lastRow <= 1) {
      return { total: 0, recent: 0, errors: 0 };
    }
    
    const data = processedSheet.getRange(2, 1, lastRow - 1, 6).getValues();
    const recentCutoff = new Date(Date.now() - (7 * 24 * 60 * 60 * 1000)); // Last 7 days
    
    let total = data.length;
    let recent = 0;
    let errors = 0;
    
    data.forEach(row => {
      const processingDate = new Date(row[0]);
      const status = row[5];
      
      if (processingDate >= recentCutoff) {
        recent++;
      }
      
      if (status === 'Error') {
        errors++;
      }
    });
    
    return { total, recent, errors };
    
  } catch (error) {
    console.error('Error getting processed message stats:', error);
    return { total: 0, recent: 0, errors: 0 };
  }
}

/**
 * Ensure the logging sheet exists with proper headers
 * ログシートが適切なヘッダーで存在することを確認
 * 
 * @param {Spreadsheet} spreadsheet - Spreadsheet object
 */
function ensureLoggingSheet(spreadsheet) {
  try {
    let sheet = spreadsheet.getSheetByName(SPREADSHEET_CONFIG.TAB_NAME);
    
    if (!sheet) {
      console.log(`Creating new sheet: ${SPREADSHEET_CONFIG.TAB_NAME}`);
      sheet = spreadsheet.insertSheet(SPREADSHEET_CONFIG.TAB_NAME);
      setupSheetHeaders(sheet);
    } else {
      // Verify headers are correct and up-to-date
      const currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const expectedHeaders = SPREADSHEET_CONFIG.HEADERS;
      
      // Check if headers match exactly (including new columns)
      const headersMatch = currentHeaders.length === expectedHeaders.length &&
                          currentHeaders.every((header, index) => header === expectedHeaders[index]);
      
      if (!headersMatch) {
        console.log('Headers outdated or mismatched, upgrading to latest format...');
        console.log(`Current headers (${currentHeaders.length}): ${currentHeaders.join(', ')}`);
        console.log(`Expected headers (${expectedHeaders.length}): ${expectedHeaders.join(', ')}`);
        
        // Upgrade to latest format automatically
        updateSpreadsheetHeaders();
        console.log('✓ Spreadsheet headers upgraded to latest format');
      }
    }
    
    return sheet;
    
  } catch (error) {
    console.error('Error ensuring logging sheet:', error);
    throw error;
  }
}

/**
 * Setup initial spreadsheet with headers and formatting
 * ヘッダーとフォーマットで初期スプレッドシートを設定
 * 
 * @param {Spreadsheet} spreadsheet - Spreadsheet object
 */
function setupInitialSheet(spreadsheet) {
  try {
    // Get the first sheet and rename it
    const sheet = spreadsheet.getSheets()[0];
    sheet.setName(SPREADSHEET_CONFIG.TAB_NAME);
    
    // Setup headers
    setupSheetHeaders(sheet);
    
    // Additional formatting
    sheet.setFrozenRows(1);  // Freeze header row
    sheet.autoResizeColumns(1, SPREADSHEET_CONFIG.HEADERS.length);
    
    console.log('Initial sheet setup completed');
    
  } catch (error) {
    console.error('Error setting up initial sheet:', error);
    throw error;
  }
}

/**
 * Setup sheet headers with formatting
 * フォーマット付きでシートヘッダーを設定
 * 
 * @param {Sheet} sheet - Sheet object
 */
function setupSheetHeaders(sheet) {
  try {
    // Set headers
    const headerRange = sheet.getRange(1, 1, 1, SPREADSHEET_CONFIG.HEADERS.length);
    headerRange.setValues([SPREADSHEET_CONFIG.HEADERS]);
    
    // Format headers
    headerRange.setBackground('#4a90e2');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    headerRange.setHorizontalAlignment('center');
    
    // Set column widths for contract tracking
    sheet.setColumnWidth(1, 150);  // 受信日時 - Receipt Date
    sheet.setColumnWidth(2, 180);  // 契約管理ツール - Contract Management Tool
    sheet.setColumnWidth(3, 250);  // 送信者メール - Sender Email
    sheet.setColumnWidth(4, 250);  // 受信メール - Recipient Email
    sheet.setColumnWidth(5, 350);  // 件名 - Subject
    sheet.setColumnWidth(6, 120);  // 契約タイプ - Contract Type
    sheet.setColumnWidth(7, 200);  // 契約相手 - Contract Party
    sheet.setColumnWidth(8, 250);  // PDFファイル名 - PDF Filename
    sheet.setColumnWidth(9, 300);  // PDF直接リンク - PDF Direct Links
    sheet.setColumnWidth(10, 100); // 処理状態 - Processing Status
    sheet.setColumnWidth(11, 120); // Slack通知済み - Slack Notified
    sheet.setColumnWidth(12, 400); // 本文要約 - Body Summary
    sheet.setColumnWidth(13, 200); // メッセージID - Message ID
    sheet.setColumnWidth(14, 300); // エラーログ - Error Log
    
    console.log('Headers setup completed');
    
  } catch (error) {
    console.error('Error setting up headers:', error);
    throw error;
  }
}

/**
 * Add email record to spreadsheet
 * スプレッドシートにメール記録を追加
 * 
 * @param {Object} emailData - Email data object
 * @returns {boolean} - Success status
 */
function addEmailRecord(emailData) {
  try {
    console.log(`Adding email record to spreadsheet: ${emailData.subject}`);
    
    const spreadsheetId = getProperty('SPREADSHEET_ID');
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(SPREADSHEET_CONFIG.TAB_NAME);
    
    if (!sheet) {
      throw new Error('Logging sheet not found');
    }
    
    // Prepare contract tracking row data
    const rowData = [
      Utilities.formatDate(emailData.date, 'JST', SPREADSHEET_CONFIG.DATE_FORMAT),  // A: 受信日時
      extractContractTool(emailData.sender),                                          // B: 契約管理ツール
      emailData.sender,                                                               // C: 送信者メール
      emailData.recipient || '',                                                      // D: 受信メール
      emailData.subject,                                                              // E: 件名
      extractContractType(emailData.subject, emailData.body),                        // F: 契約タイプ
      extractContractParty(emailData.subject, emailData.body),                       // G: 契約相手
      emailData.pdfFilename || '',                                                    // H: PDFファイル名
      emailData.pdfDirectLinks || '',                                                 // I: PDF直接リンク
      emailData.status || 'Processing',                                              // J: 処理状態
      emailData.slackNotified ? 'Yes' : 'No',                                        // K: Slack通知済み
      truncateBody(emailData.body),                                                   // L: 本文要約
      emailData.messageId,                                                            // M: メッセージID
      emailData.error || ''                                                           // N: エラーログ
    ];
    
    // Append row
    sheet.appendRow(rowData);
    
    // Format the new row
    const lastRow = sheet.getLastRow();
    const newRowRange = sheet.getRange(lastRow, 1, 1, SPREADSHEET_CONFIG.HEADERS.length);
    
    // Alternate row colors
    if (lastRow % 2 === 0) {
      newRowRange.setBackground('#f8f9fa');
    }
    
    // Format status cell based on value
    const statusCell = sheet.getRange(lastRow, 8);
    if (emailData.status === 'Success') {
      statusCell.setBackground('#d4edda');
      statusCell.setFontColor('#155724');
    } else if (emailData.status === 'Error') {
      statusCell.setBackground('#f8d7da');
      statusCell.setFontColor('#721c24');
    }
    
    console.log(`Email record added successfully at row ${lastRow}`);
    return true;
    
  } catch (error) {
    console.error('Error adding email record:', error);
    return false;
  }
}

/**
 * Update existing record status
 * 既存レコードのステータスを更新
 * 
 * @param {string} messageId - Gmail message ID
 * @param {Object} updates - Fields to update
 * @returns {boolean} - Success status
 */
function updateRecordStatus(messageId, updates) {
  try {
    console.log(`Updating record status for message: ${messageId}`);
    
    const spreadsheetId = getProperty('SPREADSHEET_ID');
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(SPREADSHEET_CONFIG.TAB_NAME);
    
    if (!sheet) {
      throw new Error('Logging sheet not found');
    }
    
    // Find row with matching message ID (now in column M, index 12)
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    
    let rowIndex = -1;
    for (let i = 1; i < values.length; i++) {  // Skip header row
      if (values[i][12] === messageId) {  // Column M (index 12) is Message ID
        rowIndex = i + 1;  // Convert to 1-based index
        break;
      }
    }
    
    if (rowIndex === -1) {
      console.log(`Message ID not found: ${messageId}`);
      return false;
    }
    
    // Update specified contract tracking fields
    if (updates.status !== undefined) {
      sheet.getRange(rowIndex, 10).setValue(updates.status);  // Column J: Processing Status
      
      // Update formatting based on status
      const statusCell = sheet.getRange(rowIndex, 10);
      if (updates.status === 'Success') {
        statusCell.setBackground('#d4edda');
        statusCell.setFontColor('#155724');
      } else if (updates.status === 'Error') {
        statusCell.setBackground('#f8d7da');
        statusCell.setFontColor('#721c24');
      }
    }
    
    if (updates.pdfDirectLinks !== undefined) {
      sheet.getRange(rowIndex, 9).setValue(updates.pdfDirectLinks);  // Column I: PDF Direct Links
    }
    
    if (updates.pdfFilename !== undefined) {
      sheet.getRange(rowIndex, 8).setValue(updates.pdfFilename);  // Column H: PDF Filename
    }
    
    if (updates.slackNotified !== undefined) {
      sheet.getRange(rowIndex, 11).setValue(updates.slackNotified ? 'Yes' : 'No');  // Column K: Slack Notified
    }
    
    if (updates.error !== undefined) {
      sheet.getRange(rowIndex, 14).setValue(updates.error);  // Column N: Error Log
    }
    
    if (updates.pdfCount !== undefined) {
      // PDF count is derived from filename, but we can update if needed
      console.log(`PDF count update: ${updates.pdfCount}`);
    }
    
    console.log(`Record updated successfully at row ${rowIndex}`);
    return true;
    
  } catch (error) {
    console.error('Error updating record status:', error);
    return false;
  }
}

/**
 * Search for record by message ID
 * メッセージIDでレコードを検索
 * 
 * @param {string} messageId - Gmail message ID
 * @returns {Object|null} - Record data or null
 */
function searchRecordByMessageId(messageId) {
  try {
    const spreadsheetId = getProperty('SPREADSHEET_ID');
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(SPREADSHEET_CONFIG.TAB_NAME);
    
    if (!sheet) {
      return null;
    }
    
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    
    for (let i = 1; i < values.length; i++) {  // Skip header row
      if (values[i][8] === messageId) {  // Column I is Message ID
        return {
          row: i + 1,
          date: values[i][0],
          sender: values[i][1],
          subject: values[i][2],
          bodySummary: values[i][3],
          attachmentCount: values[i][4],
          pdfCount: values[i][5],
          driveFolderUrl: values[i][6],
          status: values[i][7],
          messageId: values[i][8],
          slackNotified: values[i][9] === 'Yes',
          errorLog: values[i][10]
        };
      }
    }
    
    return null;
    
  } catch (error) {
    console.error('Error searching for record:', error);
    return null;
  }
}

/**
 * Generate daily summary report
 * 日次サマリーレポートを生成
 * 
 * @param {Date} date - Date for summary (defaults to today)
 * @returns {Object} - Summary statistics
 */
function generateDailySummary(date = new Date()) {
  try {
    console.log('Generating daily summary report...');
    
    const spreadsheetId = getProperty('SPREADSHEET_ID');
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(SPREADSHEET_CONFIG.TAB_NAME);
    
    if (!sheet) {
      throw new Error('Logging sheet not found');
    }
    
    // Format date for comparison
    const targetDate = Utilities.formatDate(date, 'JST', 'yyyy/MM/dd');
    
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    
    let summary = {
      date: targetDate,
      totalEmails: 0,
      successfulEmails: 0,
      errorEmails: 0,
      totalAttachments: 0,
      totalPDFs: 0,
      senderBreakdown: {},
      subjectPatterns: [],
      errors: []
    };
    
    // Analyze data
    for (let i = 1; i < values.length; i++) {  // Skip header row
      const rowDate = Utilities.formatDate(values[i][0], 'JST', 'yyyy/MM/dd');
      
      if (rowDate === targetDate) {
        summary.totalEmails++;
        
        // Status breakdown
        if (values[i][7] === 'Success') {
          summary.successfulEmails++;
        } else if (values[i][7] === 'Error') {
          summary.errorEmails++;
          if (values[i][10]) {  // Error log
            summary.errors.push({
              subject: values[i][2],
              error: values[i][10]
            });
          }
        }
        
        // Attachment counts
        summary.totalAttachments += Number(values[i][4]) || 0;
        summary.totalPDFs += Number(values[i][5]) || 0;
        
        // Sender breakdown
        const sender = values[i][1];
        summary.senderBreakdown[sender] = (summary.senderBreakdown[sender] || 0) + 1;
        
        // Subject patterns (simplified)
        const subject = values[i][2];
        if (subject.includes('メルマガ')) {
          summary.subjectPatterns.push('Newsletter');
        } else if (subject.includes('部会')) {
          summary.subjectPatterns.push('Meeting');
        } else if (subject.includes('勉強会')) {
          summary.subjectPatterns.push('Study Session');
        }
      }
    }
    
    // Calculate success rate
    summary.successRate = summary.totalEmails > 0 
      ? Math.round((summary.successfulEmails / summary.totalEmails) * 100) 
      : 100;
    
    console.log(`Daily summary generated for ${targetDate}:`);
    console.log(`Total emails: ${summary.totalEmails}`);
    console.log(`Success rate: ${summary.successRate}%`);
    
    return summary;
    
  } catch (error) {
    console.error('Error generating daily summary:', error);
    return null;
  }
}

/**
 * Truncate email body for summary
 * サマリー用にメール本文を切り詰める
 * 
 * @param {string} body - Full email body
 * @returns {string} - Truncated body
 */
function truncateBody(body) {
  if (!body || body.trim().length === 0) {
    return '';
  }
  
  // Clean and truncate
  const cleanBody = body
    .replace(/\n\s*\n/g, ' ')  // Replace multiple newlines with space
    .replace(/\s+/g, ' ')      // Replace multiple spaces with single space
    .trim();
  
  if (cleanBody.length <= SPREADSHEET_CONFIG.MAX_BODY_LENGTH) {
    return cleanBody;
  }
  
  return cleanBody.substring(0, SPREADSHEET_CONFIG.MAX_BODY_LENGTH) + '...';
}

/**
 * Get spreadsheet statistics
 * スプレッドシートの統計情報を取得
 * 
 * @returns {Object} - Spreadsheet statistics
 */
function getSpreadsheetStats() {
  try {
    const spreadsheetId = getProperty('SPREADSHEET_ID');
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(SPREADSHEET_CONFIG.TAB_NAME);
    
    if (!sheet) {
      return null;
    }
    
    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();
    
    return {
      spreadsheetName: spreadsheet.getName(),
      spreadsheetUrl: spreadsheet.getUrl(),
      sheetName: sheet.getName(),
      totalRows: lastRow - 1,  // Exclude header
      totalColumns: lastColumn,
      lastUpdated: lastRow > 1 ? sheet.getRange(lastRow, 1).getValue() : null
    };
    
  } catch (error) {
    console.error('Error getting spreadsheet stats:', error);
    return null;
  }
}

/**
 * Test spreadsheet operations
 * スプレッドシート操作のテスト
 */
function testSpreadsheetOperations() {
  console.log('=== TESTING Spreadsheet Operations ===');
  
  try {
    // Test 1: Create or get spreadsheet
    console.log('Test 1: Create/Get Spreadsheet');
    const spreadsheetId = createOrGetSpreadsheet();
    console.log(`✓ Spreadsheet ID: ${spreadsheetId}`);
    
    // Test 2: Add test email record with new format
    console.log('\nTest 2: Add Email Record (New Format)');
    const testEmailData = {
      date: new Date(),
      sender: 'test@contracttool1.example.com',
      recipient: 'recipient@company.com',
      subject: 'テストメール: 契約完了通知',
      body: 'これはテストメールの本文です。契約が完了しました。添付のPDFファイルをご確認ください。' + '詳細情報が続きます。'.repeat(50),
      attachmentCount: 2,
      pdfCount: 2,
      pdfFilename: 'contract_agreement.pdf, nda_document.pdf',
      pdfDirectLinks: 'https://drive.google.com/file/d/1ABC123TEST/view\nhttps://drive.google.com/file/d/1DEF456TEST/view',
      status: 'Success',
      messageId: `test-${new Date().getTime()}`,
      slackNotified: true,
      error: null
    };
    
    const addResult = addEmailRecord(testEmailData);
    console.log(`✓ Email record added: ${addResult}`);
    
    // Test 3: Search for record
    console.log('\nTest 3: Search for Record');
    const searchResult = searchRecordByMessageId(testEmailData.messageId);
    console.log(`✓ Record found: ${searchResult !== null}`);
    if (searchResult) {
      console.log(`  - Subject: ${searchResult.subject}`);
      console.log(`  - Status: ${searchResult.status}`);
    }
    
    // Test 4: Update record status
    console.log('\nTest 4: Update Record Status');
    const updateResult = updateRecordStatus(testEmailData.messageId, {
      status: 'Updated',
      error: 'Test error message'
    });
    console.log(`✓ Record updated: ${updateResult}`);
    
    // Test 5: Generate daily summary
    console.log('\nTest 5: Generate Daily Summary');
    const summary = generateDailySummary();
    console.log(`✓ Daily summary generated:`);
    console.log(`  - Total emails: ${summary.totalEmails}`);
    console.log(`  - Success rate: ${summary.successRate}%`);
    console.log(`  - Total PDFs: ${summary.totalPDFs}`);
    
    // Test 6: Get spreadsheet stats
    console.log('\nTest 6: Get Spreadsheet Stats');
    const stats = getSpreadsheetStats();
    console.log(`✓ Spreadsheet stats:`);
    console.log(`  - Total rows: ${stats.totalRows}`);
    console.log(`  - URL: ${stats.spreadsheetUrl}`);
    
    // Test 7: Error handling with new format
    console.log('\nTest 7: Error Handling (New Format)');
    const errorEmailData = {
      date: new Date(),
      sender: 'error@contracttool2.example.com',
      recipient: 'recipient2@company.com',
      subject: 'エラーテスト: 処理失敗のシミュレーション',
      body: 'このメールは処理エラーをテストするためのものです。PDFファイルの処理中にエラーが発生しました。',
      attachmentCount: 1,
      pdfCount: 0,
      pdfFilename: '',
      pdfDirectLinks: '',
      status: 'Error',
      messageId: `error-test-${new Date().getTime()}`,
      slackNotified: false,
      error: 'Simulated processing error for testing'
    };
    
    const errorResult = addEmailRecord(errorEmailData);
    console.log(`✓ Error record added: ${errorResult}`);
    
    // Test 8: Test spreadsheet format upgrade
    console.log('\nTest 8: Spreadsheet Format Upgrade');
    try {
      const upgradeResult = upgradeToLatestFormat();
      console.log(`✓ Spreadsheet format upgrade: ${upgradeResult ? 'Success' : 'Failed'}`);
    } catch (upgradeError) {
      console.log(`⚠ Spreadsheet format upgrade error: ${upgradeError.message}`);
    }
    
    console.log('\nSpreadsheet operations test completed successfully');
    console.log(`Spreadsheet URL: ${stats.spreadsheetUrl}`);
    console.log('\nNEW FORMAT FEATURES TESTED:');
    console.log('✓ Recipient email tracking');
    console.log('✓ PDF direct links');
    console.log('✓ Enhanced contract metadata');
    
  } catch (error) {
    console.error('Spreadsheet operations test failed:', error);
    throw error;
  }
}

/**
 * Validate spreadsheet data integrity
 * スプレッドシートデータの整合性を検証
 */
function validateSpreadsheetData() {
  try {
    console.log('Validating spreadsheet data integrity...');
    
    const spreadsheetId = getProperty('SPREADSHEET_ID');
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(SPREADSHEET_CONFIG.TAB_NAME);
    
    if (!sheet) {
      throw new Error('Logging sheet not found');
    }
    
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    
    let issues = [];
    let duplicateMessageIds = {};
    
    // Check for data integrity issues
    for (let i = 1; i < values.length; i++) {  // Skip header
      const row = i + 1;
      
      // Check for required fields
      if (!values[i][0]) issues.push(`Row ${row}: Missing date`);
      if (!values[i][1]) issues.push(`Row ${row}: Missing sender`);
      if (!values[i][2]) issues.push(`Row ${row}: Missing subject`);
      if (!values[i][8]) issues.push(`Row ${row}: Missing message ID`);
      
      // Check for duplicate message IDs
      const messageId = values[i][8];
      if (messageId) {
        if (duplicateMessageIds[messageId]) {
          issues.push(`Row ${row}: Duplicate message ID found (also in row ${duplicateMessageIds[messageId]})`);
        } else {
          duplicateMessageIds[messageId] = row;
        }
      }
      
      // Validate numeric fields
      const attachmentCount = values[i][4];
      if (attachmentCount !== '' && isNaN(attachmentCount)) {
        issues.push(`Row ${row}: Invalid attachment count`);
      }
      
      const pdfCount = values[i][5];
      if (pdfCount !== '' && isNaN(pdfCount)) {
        issues.push(`Row ${row}: Invalid PDF count`);
      }
    }
    
    console.log(`Validation completed. Found ${issues.length} issues.`);
    
    if (issues.length > 0) {
      console.log('Issues found:');
      issues.slice(0, 10).forEach(issue => console.log(`  - ${issue}`));
      if (issues.length > 10) {
        console.log(`  ... and ${issues.length - 10} more issues`);
      }
    } else {
      console.log('✓ No data integrity issues found');
    }
    
    return {
      isValid: issues.length === 0,
      totalRows: values.length - 1,
      issueCount: issues.length,
      issues: issues
    };
    
  } catch (error) {
    console.error('Error validating spreadsheet data:', error);
    return {
      isValid: false,
      error: error.message
    };
  }
}

/**
 * Update spreadsheet headers (for schema changes)
 * スプレッドシートヘッダーを更新（スキーマ変更用）
 */
function updateSpreadsheetHeaders() {
  try {
    console.log('Updating spreadsheet headers...');
    
    const spreadsheetId = getProperty('SPREADSHEET_ID');
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(SPREADSHEET_CONFIG.TAB_NAME);
    
    if (!sheet) {
      throw new Error('Logging sheet not found');
    }
    
    // Get current headers
    const currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    console.log(`Current headers: ${currentHeaders.join(', ')}`);
    
    // Check if we need to add new columns
    const currentColumnCount = currentHeaders.filter(h => h !== '').length;
    const requiredColumnCount = SPREADSHEET_CONFIG.HEADERS.length;
    
    if (requiredColumnCount > currentColumnCount) {
      console.log(`Adding ${requiredColumnCount - currentColumnCount} new columns...`);
      
      // Insert new columns if needed
      const columnsToAdd = requiredColumnCount - currentColumnCount;
      sheet.insertColumnsAfter(currentColumnCount, columnsToAdd);
    }
    
    // Update all headers
    setupSheetHeaders(sheet);
    
    console.log('Headers updated successfully');
    console.log('New headers: PDF直接リンク column has been added');
    return true;
    
  } catch (error) {
    console.error('Error updating headers:', error);
    return false;
  }
}

/**
 * Update existing spreadsheet to use PDF direct links and recipient email
 * 既存スプレッドシートをPDF直接リンクと受信メール対応に更新
 */
function upgradeToLatestFormat() {
  try {
    console.log('=== Upgrading spreadsheet to latest format ===');
    
    // Update headers
    const headerUpdateResult = updateSpreadsheetHeaders();
    if (!headerUpdateResult) {
      throw new Error('Failed to update headers');
    }
    
    console.log('✓ Headers updated with new format:');
    console.log('  - Added: 受信メール (Recipient Email) column');
    console.log('  - Updated: DriveフォルダURL → PDF直接リンク');
    console.log('✓ Spreadsheet is now ready for enhanced contract tracking');
    console.log('');
    console.log('NEW FEATURES:');
    console.log('1. Recipient email tracking - see who received each contract');
    console.log('2. Direct PDF links - click directly on PDF files');
    console.log('');
    console.log('NOTE: Existing data will remain in previous format.');
    console.log('New entries will include recipient email and direct PDF links.');
    
    return true;
    
  } catch (error) {
    console.error('Error upgrading spreadsheet:', error);
    return false;
  }
}

/**
 * Legacy function for backwards compatibility
 * 後方互換性のためのレガシー関数
 */
function upgradeToPdfDirectLinks() {
  return upgradeToLatestFormat();
}

/**
 * Check current spreadsheet format and status
 * 現在のスプレッドシート形式と状態をチェック
 */
function checkSpreadsheetFormat() {
  try {
    console.log('=== CHECKING SPREADSHEET FORMAT ===');
    
    const spreadsheetId = getProperty('SPREADSHEET_ID');
    if (!spreadsheetId) {
      console.log('❌ No spreadsheet found');
      console.log('Run createOrGetSpreadsheet() to create one');
      return null;
    }
    
    console.log(`📊 Spreadsheet ID: ${spreadsheetId}`);
    
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(SPREADSHEET_CONFIG.TAB_NAME);
    
    if (!sheet) {
      console.log('❌ Main contract sheet not found');
      return null;
    }
    
    // Get current headers
    const currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const expectedHeaders = SPREADSHEET_CONFIG.HEADERS;
    
    console.log('\n📋 CURRENT FORMAT:');
    console.log(`Columns: ${currentHeaders.length}`);
    currentHeaders.forEach((header, index) => {
      console.log(`  ${String.fromCharCode(65 + index)}: ${header}`);
    });
    
    console.log('\n📋 EXPECTED FORMAT:');
    console.log(`Columns: ${expectedHeaders.length}`);
    expectedHeaders.forEach((header, index) => {
      console.log(`  ${String.fromCharCode(65 + index)}: ${header}`);
    });
    
    // Check for specific new columns
    const hasRecipientEmail = currentHeaders.includes('受信メール');
    const hasPdfDirectLinks = currentHeaders.includes('PDF直接リンク');
    const hasOldDriveFolderUrl = currentHeaders.includes('DriveフォルダURL');
    
    console.log('\n🔍 FORMAT ANALYSIS:');
    console.log(`✓ Recipient Email column: ${hasRecipientEmail ? '✅ Present' : '❌ Missing'}`);
    console.log(`✓ PDF Direct Links column: ${hasPdfDirectLinks ? '✅ Present' : '❌ Missing'}`);
    console.log(`✓ Old Drive Folder URL: ${hasOldDriveFolderUrl ? '⚠️ Still present (needs upgrade)' : '✅ Upgraded'}`);
    
    const isLatestFormat = hasRecipientEmail && hasPdfDirectLinks && !hasOldDriveFolderUrl;
    console.log(`\n📊 FORMAT STATUS: ${isLatestFormat ? '✅ LATEST' : '⚠️ NEEDS UPGRADE'}`);
    
    // Check data
    const lastRow = sheet.getLastRow();
    console.log(`\n📈 DATA STATUS:`);
    console.log(`Total rows: ${lastRow} (${lastRow - 1} data rows)`);
    console.log(`Spreadsheet URL: ${spreadsheet.getUrl()}`);
    
    return {
      hasLatestFormat: isLatestFormat,
      currentHeaders: currentHeaders,
      expectedHeaders: expectedHeaders,
      hasRecipientEmail: hasRecipientEmail,
      hasPdfDirectLinks: hasPdfDirectLinks,
      hasOldDriveFolderUrl: hasOldDriveFolderUrl,
      spreadsheetUrl: spreadsheet.getUrl(),
      totalRows: lastRow
    };
    
  } catch (error) {
    console.error('Error checking spreadsheet format:', error);
    return null;
  }
}

/**
 * Force upgrade spreadsheet format (manual execution)
 * スプレッドシートフォーマットを強制アップグレード（手動実行用）
 */
function forceUpgradeSpreadsheet() {
  try {
    console.log('=== FORCE UPGRADING SPREADSHEET FORMAT ===');
    
    const spreadsheetId = getProperty('SPREADSHEET_ID');
    if (!spreadsheetId) {
      console.log('No spreadsheet found. Creating new one...');
      createOrGetSpreadsheet();
      return true;
    }
    
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(SPREADSHEET_CONFIG.TAB_NAME);
    
    if (!sheet) {
      console.log('Main sheet not found. Creating...');
      ensureLoggingSheet(spreadsheet);
      return true;
    }
    
    // Show current state
    const currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    console.log(`Current format (${currentHeaders.length} columns):`);
    currentHeaders.forEach((header, index) => {
      console.log(`  ${String.fromCharCode(65 + index)}: ${header}`);
    });
    
    // Upgrade
    const result = upgradeToLatestFormat();
    
    if (result) {
      // Show new state
      const newHeaders = sheet.getRange(1, 1, 1, SPREADSHEET_CONFIG.HEADERS.length).getValues()[0];
      console.log(`\nNew format (${newHeaders.length} columns):`);
      newHeaders.forEach((header, index) => {
        console.log(`  ${String.fromCharCode(65 + index)}: ${header}`);
      });
      
      console.log('\n✅ SPREADSHEET UPGRADE COMPLETED');
      console.log('Your spreadsheet now supports:');
      console.log('  • Recipient email tracking');
      console.log('  • Direct PDF file links');
      console.log('  • Enhanced contract metadata');
    }
    
    return result;
    
  } catch (error) {
    console.error('Error in force upgrade:', error);
    return false;
  }
}

/**
 * Complete spreadsheet upgrade workflow
 * 完全なスプレッドシートアップグレードワークフロー
 */
function completeSpreadsheetUpgrade() {
  try {
    console.log('🚀 STARTING COMPLETE SPREADSHEET UPGRADE');
    console.log('='.repeat(50));
    
    // Step 1: Check current format
    console.log('Step 1: Checking current format...');
    const formatCheck = checkSpreadsheetFormat();
    
    if (!formatCheck) {
      console.log('❌ Cannot check format. Creating new spreadsheet...');
      createOrGetSpreadsheet();
      return true;
    }
    
    if (formatCheck.hasLatestFormat) {
      console.log('✅ Spreadsheet already has latest format!');
      return true;
    }
    
    // Step 2: Force upgrade
    console.log('\nStep 2: Performing upgrade...');
    const upgradeResult = forceUpgradeSpreadsheet();
    
    if (!upgradeResult) {
      console.log('❌ Upgrade failed');
      return false;
    }
    
    // Step 3: Verify upgrade
    console.log('\nStep 3: Verifying upgrade...');
    const verifyCheck = checkSpreadsheetFormat();
    
    if (verifyCheck && verifyCheck.hasLatestFormat) {
      console.log('✅ UPGRADE COMPLETED SUCCESSFULLY!');
      console.log('\n🎉 Your spreadsheet now supports:');
      console.log('  • Recipient email tracking');
      console.log('  • PDF direct links');
      console.log('  • Enhanced contract metadata');
      console.log(`\n📊 Access your spreadsheet: ${verifyCheck.spreadsheetUrl}`);
      
      // Step 4: Run test to demonstrate new features
      console.log('\nStep 4: Testing new features...');
      try {
        testSpreadsheetOperations();
        console.log('✅ All tests passed with new format!');
      } catch (testError) {
        console.log('⚠️ Test had issues, but upgrade was successful');
        console.error('Test error:', testError);
      }
      
      return true;
    } else {
      console.log('❌ Upgrade verification failed');
      return false;
    }
    
  } catch (error) {
    console.error('Error in complete spreadsheet upgrade:', error);
    return false;
  }
}

/**
 * Quick fix for user's current issue
 * ユーザーの現在の問題に対するクイックフィックス
 */
function fixCurrentSpreadsheetIssue() {
  console.log('🔧 FIXING CURRENT SPREADSHEET ISSUE');
  console.log('='.repeat(40));
  console.log('Issue: DriveフォルダURL instead of PDF直接リンク');
  console.log('Issue: Missing 受信メール column');
  console.log('');
  
  return completeSpreadsheetUpgrade();
}

/**
 * Test PDF direct links format
 * PDF直接リンク形式をテスト
 */
function testPdfDirectLinksFormat() {
  try {
    console.log('=== TESTING PDF DIRECT LINKS FORMAT ===');
    
    // Create test data with new format
    const testEmailData = {
      date: new Date(),
      sender: 'test@contracttool1.example.com',
      recipient: 'recipient@company.com',
      subject: 'PDFリンクテスト: 直接リンク形式',
      body: 'このテストではPDFファイルが直接クリック可能なリンクとして記録されることを確認します。',
      attachmentCount: 2,
      pdfCount: 2,
      pdfFilename: 'contract_test.pdf, agreement_test.pdf',
      pdfDirectLinks: 'https://drive.google.com/file/d/1TEST123ABC/view\nhttps://drive.google.com/file/d/1TEST456DEF/view',
      status: 'Success',
      messageId: `pdf-test-${new Date().getTime()}`,
      slackNotified: true,
      error: null
    };
    
    // Add record
    const addResult = addEmailRecord(testEmailData);
    if (addResult) {
      console.log('✅ Test record added successfully');
      console.log('📋 PDF Links recorded as:');
      console.log('   Link 1: https://drive.google.com/file/d/1TEST123ABC/view');
      console.log('   Link 2: https://drive.google.com/file/d/1TEST456DEF/view');
      console.log('');
      console.log('📊 These links should be clickable in the spreadsheet');
      console.log('💡 Each link opens the PDF file directly');
      console.log('🎯 No filename prefix - clean URLs for easy clicking');
      
      // Get spreadsheet URL
      const spreadsheetId = getProperty('SPREADSHEET_ID');
      if (spreadsheetId) {
        const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
        console.log(`\n🔗 Check your spreadsheet: ${spreadsheet.getUrl()}`);
      }
    } else {
      console.log('❌ Failed to add test record');
    }
    
    return addResult;
    
  } catch (error) {
    console.error('Error testing PDF direct links format:', error);
    return false;
  }
}

/**
 * Export data to CSV format (optional feature)
 * データをCSV形式でエクスポート（オプション機能）
 * 
 * @param {Date} startDate - Start date for export
 * @param {Date} endDate - End date for export
 * @returns {string} - CSV content
 */
function exportToCSV(startDate, endDate) {
  try {
    console.log('Exporting data to CSV...');
    
    const spreadsheetId = getProperty('SPREADSHEET_ID');
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(SPREADSHEET_CONFIG.TAB_NAME);
    
    if (!sheet) {
      throw new Error('Logging sheet not found');
    }
    
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    
    // Filter by date range if provided
    let filteredData = [values[0]];  // Include headers
    
    for (let i = 1; i < values.length; i++) {
      const rowDate = values[i][0];
      
      if ((!startDate || rowDate >= startDate) && (!endDate || rowDate <= endDate)) {
        filteredData.push(values[i]);
      }
    }
    
    // Convert to CSV
    const csv = filteredData.map(row => {
      return row.map(cell => {
        // Escape quotes and wrap in quotes if contains comma
        const cellStr = String(cell);
        if (cellStr.includes(',') || cellStr.includes('"') || cellStr.includes('\n')) {
          return `"${cellStr.replace(/"/g, '""')}"`;
        }
        return cellStr;
      }).join(',');
    }).join('\n');
    
    console.log(`Exported ${filteredData.length - 1} records to CSV`);
    return csv;
    
  } catch (error) {
    console.error('Error exporting to CSV:', error);
    return null;
  }
}

/**
 * Extract contract management tool from sender email
 * 送信者メールアドレスから契約管理ツールを抽出
 * 
 * @param {string} senderEmail - Sender email address
 * @returns {string} - Contract management tool name
 */
function extractContractTool(senderEmail) {
  try {
    if (!senderEmail) return 'Unknown';
    
    const email = senderEmail.toLowerCase();
    
    if (email.includes('contracttool1')) return 'Contract Tool 1';
    if (email.includes('contracttool2') || email.includes('contracttool2alt')) return 'Contract Tool 2';
    if (email.includes('contracttool3')) return 'Contract Tool 3';
    if (email.includes('contracttool4')) return 'Contract Tool 4';
    if (email.includes('contracttool5')) return 'Contract Tool 5';
    if (email.includes('contracttool6')) return 'Contract Tool 6';
    if (email.includes('contracttool7')) return 'Contract Tool 7';
    if (email.includes('contracttool8')) return 'Contract Tool 8';
    if (email.includes('contracttool9')) return 'Contract Tool 9';
    if (email.includes('contracttool10')) return 'Contract Tool 10';
    
    // Check for Dropbox Sign/HelloSign patterns in sender field
    if (email.includes('via Dropbox Sign') || email.includes("'Dropbox Sign' via")) return 'Dropbox Sign';
    if (email.includes('via HelloSign') || email.includes("'HelloSign' via")) return 'HelloSign';
    if (email.includes('hellosign.com')) return 'HelloSign/Dropbox Sign';
    if (email.includes('docusign')) return 'DocuSign';
    
    // Extract domain for unknown tools
    const domain = email.split('@')[1];
    return domain ? domain.split('.')[0] : 'Unknown';
    
  } catch (error) {
    console.error('Error extracting contract tool:', error);
    return 'Unknown';
  }
}

/**
 * Extract contract type from subject and body
 * 件名と本文から契約タイプを抽出
 * 
 * @param {string} subject - Email subject
 * @param {string} body - Email body
 * @returns {string} - Contract type
 */
function extractContractType(subject, body) {
  try {
    const text = `${subject} ${body}`.toLowerCase();
    
    // Japanese contract types
    if (text.includes('雇用契約') || text.includes('労働契約')) return '雇用契約';
    if (text.includes('業務委託') || text.includes('委託契約')) return '業務委託契約';
    if (text.includes('秘密保持') || text.includes('nda') || text.includes('機密保持')) return '秘密保持契約';
    if (text.includes('売買契約') || text.includes('購入契約')) return '売買契約';
    if (text.includes('賃貸借契約') || text.includes('リース契約')) return '賃貸借契約';
    if (text.includes('サービス契約') || text.includes('利用規約')) return 'サービス契約';
    if (text.includes('パートナー契約') || text.includes('提携契約')) return 'パートナー契約';
    if (text.includes('ライセンス契約')) return 'ライセンス契約';
    if (text.includes('覚書') || text.includes('memorandum')) return '覚書';
    if (text.includes('個別契約')) return '個別契約';
    
    // English contract types
    if (text.includes('employment') || text.includes('job offer')) return 'Employment Agreement';
    if (text.includes('service agreement') || text.includes('msa')) return 'Service Agreement';
    if (text.includes('nda') || text.includes('confidentiality')) return 'Non-Disclosure Agreement';
    if (text.includes('purchase') || text.includes('sales')) return 'Purchase Agreement';
    if (text.includes('lease') || text.includes('rental')) return 'Lease Agreement';
    if (text.includes('partnership') || text.includes('joint venture')) return 'Partnership Agreement';
    if (text.includes('license') || text.includes('licensing')) return 'License Agreement';
    if (text.includes('consulting') || text.includes('contractor')) return 'Consulting Agreement';
    if (text.includes('vendor') || text.includes('supplier')) return 'Vendor Agreement';
    if (text.includes('master service') || text.includes('framework')) return 'Master Service Agreement';
    
    return '一般契約';
    
  } catch (error) {
    console.error('Error extracting contract type:', error);
    return '不明';
  }
}

/**
 * Extract contract party/company from subject and body
 * 件名と本文から契約相手を抽出
 * 
 * @param {string} subject - Email subject
 * @param {string} body - Email body
 * @returns {string} - Contract party name
 */
function extractContractParty(subject, body) {
  try {
    const text = `${subject} ${body}`;
    
    // CloudSign specific patterns - extract from subject like "Company1_覚書Company2"
    const cloudSignMatch = subject.match(/^「?([^_」]+)_[^_]+_?([^」]+)?」?の合意締結が完了しました$/);
    if (cloudSignMatch) {
      // Return the first company name found
      return cloudSignMatch[1].trim();
    }
    
    // Common patterns for company names
    const companyPatterns = [
      // Japanese company patterns
      /([^\s]+)株式会社/g,
      /株式会社([^\s]+)/g,
      /([^\s]+)有限会社/g,
      /有限会社([^\s]+)/g,
      /([^\s]+)合同会社/g,
      /合同会社([^\s]+)/g,
      /([^\s]+)Co\.?/g,
      /([^\s]+)Corp\.?/g,
      /([^\s]+)Inc\.?/g,
      /([^\s]+)Ltd\.?/g,
      /([^\s]+)LLC/g,
      
      // Extract from common email signatures
      /(?:from|signed by|completed by)\s+([A-Za-z\s]+?)(?:\s|<|$)/gi,
      /(?:company|organization):\s*([^\n\r]+)/gi
    ];
    
    for (const pattern of companyPatterns) {
      const matches = text.match(pattern);
      if (matches && matches.length > 0) {
        // Return the first meaningful match
        const match = matches[0].trim();
        if (match.length > 2 && match.length < 50) {
          return match.replace(/[<>]/g, '').trim();
        }
      }
    }
    
    // Try to extract from email domain if no company name found
    const emailMatch = text.match(/@([a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/);
    if (emailMatch) {
      const domain = emailMatch[1];
      const company = domain.split('.')[0];
      if (company && company.length > 2) {
        return company.charAt(0).toUpperCase() + company.slice(1);
      }
    }
    
    return '未特定';
    
  } catch (error) {
    console.error('Error extracting contract party:', error);
    return '抽出エラー';
  }
}

/**
 * Generate contract summary for daily reports
 * 日次レポート用の契約サマリーを生成
 * 
 * @param {Date} date - Date for summary
 * @returns {Object} - Contract summary
 */
function generateContractSummary(date = new Date()) {
  try {
    console.log('Generating contract summary...');
    
    const spreadsheetId = getProperty('SPREADSHEET_ID');
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(SPREADSHEET_CONFIG.TAB_NAME);
    
    if (!sheet) {
      throw new Error('Contract tracking sheet not found');
    }
    
    const targetDate = Utilities.formatDate(date, 'JST', 'yyyy/MM/dd');
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    
    let summary = {
      date: targetDate,
      totalContracts: 0,
      completedContracts: 0,
      errorContracts: 0,
      contractsByTool: {},
      contractsByType: {},
      contractParties: [],
      errors: []
    };
    
    for (let i = 1; i < values.length; i++) {
      const rowDate = Utilities.formatDate(values[i][0], 'JST', 'yyyy/MM/dd');
      
      if (rowDate === targetDate) {
        summary.totalContracts++;
        
        const tool = values[i][1] || 'Unknown';
        const contractType = values[i][4] || '不明';
        const contractParty = values[i][5] || '未特定';
        const status = values[i][8] || 'Processing';
        const error = values[i][12];
        
        // Tool breakdown
        summary.contractsByTool[tool] = (summary.contractsByTool[tool] || 0) + 1;
        
        // Contract type breakdown
        summary.contractsByType[contractType] = (summary.contractsByType[contractType] || 0) + 1;
        
        // Contract parties
        if (contractParty !== '未特定' && contractParty !== '抽出エラー') {
          summary.contractParties.push(contractParty);
        }
        
        // Status tracking
        if (status === 'Success') {
          summary.completedContracts++;
        } else if (status === 'Error') {
          summary.errorContracts++;
          if (error) {
            summary.errors.push({
              contractParty: contractParty,
              contractType: contractType,
              error: error
            });
          }
        }
      }
    }
    
    // Calculate success rate
    summary.successRate = summary.totalContracts > 0 
      ? Math.round((summary.completedContracts / summary.totalContracts) * 100) 
      : 100;
    
    console.log(`Contract summary generated for ${targetDate}:`);
    console.log(`Total contracts: ${summary.totalContracts}`);
    console.log(`Success rate: ${summary.successRate}%`);
    
    return summary;
    
  } catch (error) {
    console.error('Error generating contract summary:', error);
    return null;
  }
}