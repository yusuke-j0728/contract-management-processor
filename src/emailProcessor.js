/**
 * Email Processing Module
 * 
 * Handles individual email message processing, including subject pattern matching
 * and coordination of attachment and notification handling.
 */

/**
 * Process a single email message
 * 個別のメールメッセージを処理
 * 
 * @param {GmailMessage} message - Gmail message object
 * @returns {boolean} - true if message was processed successfully
 */
function processMessage(message) {
  try {
    console.log('--- Processing Message ---');
    
    const subject = message.getSubject();
    const sender = message.getFrom();
    const recipient = getMessageRecipient(message);
    const date = message.getDate();
    const body = message.getPlainBody();
    const attachments = message.getAttachments();
    const messageId = message.getId();
    
    console.log(`Subject: ${subject}`);
    console.log(`From: ${sender}`);
    console.log(`To: ${recipient}`);
    console.log(`Date: ${date}`);
    console.log(`Message ID: ${messageId}`);
    console.log(`Attachments: ${attachments.length}`);
    
    // Check subject pattern using advanced pattern matching
    const patternMatch = checkSubjectPattern(subject);
    if (!patternMatch.isMatch) {
      console.log(`Subject pattern mismatch: ${subject}`);
      console.log(`Checked patterns: ${patternMatch.checkedPatterns}`);
      return false;
    }
    
    console.log(`Subject pattern matched: ${patternMatch.matchedPattern}`);
    console.log('Processing email...');
    
    // Initialize email record for spreadsheet logging
    let emailRecord = null;
    if (CONFIG.ENABLE_SPREADSHEET_LOGGING) {
      emailRecord = {
        date: date,
        sender: sender,
        recipient: recipient,
        subject: subject,
        body: body,
        messageId: messageId,
        attachmentCount: attachments.length,
        pdfCount: 0,
        pdfDirectLinks: '',
        status: 'Processing',
        slackNotified: false,
        error: null
      };
      
      // Add initial record to spreadsheet
      try {
        addEmailRecord(emailRecord);
        console.log('Email record added to spreadsheet');
      } catch (error) {
        console.error('Error adding email record to spreadsheet:', error);
        // Continue processing even if spreadsheet logging fails
      }
    }
    
    // Process attachments if any
    const attachmentInfo = [];
    if (attachments.length > 0) {
      console.log(`Processing ${attachments.length} attachments...`);
      
      // Check if there are any PDF files before processing
      const pdfAttachments = attachments.filter(attachment => {
        const fileName = attachment.getName().toLowerCase();
        return fileName.endsWith('.pdf');
      });
      
      if (pdfAttachments.length > 0) {
        console.log(`Found ${pdfAttachments.length} PDF files, processing attachments...`);
        try {
          // Pass email date and sender for content-based duplicate detection
          attachmentInfo.push(...processAttachments(attachments, subject, date, sender));
          
          // Update spreadsheet record with PDF info
          if (CONFIG.ENABLE_SPREADSHEET_LOGGING && emailRecord) {
            emailRecord.pdfCount = pdfAttachments.length;
            
            // Collect all PDF direct links (URLs only for clickable links in spreadsheet)
            // Show "Duplicate" for duplicate PDFs instead of folder links
            const pdfDirectLinks = attachmentInfo
              .filter(att => !att.error && att.originalName && att.originalName.toLowerCase().endsWith('.pdf'))
              .map(att => {
                if (att.isDuplicate) {
                  return 'Duplicate';
                } else if (att.pdfDirectUrl) {
                  return att.pdfDirectUrl;
                }
                return null;
              })
              .filter(link => link !== null)
              .join('\n');
            
            if (pdfDirectLinks) {
              emailRecord.pdfDirectLinks = pdfDirectLinks;
            }
            
            // If multiple PDFs, list all filenames
            const allPdfNames = attachmentInfo
              .filter(att => !att.error && att.originalName && att.originalName.toLowerCase().endsWith('.pdf'))
              .map(att => {
                const name = att.originalName;
                return att.isDuplicate ? `${name} (duplicate)` : name;
              })
              .join(', ');
            
            if (allPdfNames) {
              emailRecord.pdfFilename = allPdfNames;
            }
          }
        } catch (error) {
          console.error('Error processing attachments:', error);
          // Continue with notification even if attachment processing fails
        }
      } else {
        console.log('No PDF files found in attachments, skipping Drive folder creation');
        // Still add attachment info for non-PDF files (for Slack notification)
        attachments.forEach(attachment => {
          attachmentInfo.push({
            originalName: attachment.getName(),
            savedName: null,
            size: attachment.getSize(),
            driveUrl: null,
            fileId: null,
            folderPath: null,
            skipped: 'Not a PDF file'
          });
        });
      }
    }
    
    // Send Slack notification
    let slackNotified = false;
    try {
      sendSlackNotification({
        subject: subject,
        sender: sender,
        date: date,
        body: formatEmailBody(body),
        attachments: attachmentInfo
      });
      slackNotified = true;
    } catch (error) {
      console.error('Error sending Slack notification:', error);
      // Don't throw - we still want to mark as processed
    }
    
    // Update spreadsheet record with final status
    if (CONFIG.ENABLE_SPREADSHEET_LOGGING && emailRecord) {
      try {
        updateRecordStatus(messageId, {
          status: 'Success',
          slackNotified: slackNotified,
          pdfCount: emailRecord.pdfCount,
          pdfDirectLinks: emailRecord.pdfDirectLinks,
          pdfFilename: emailRecord.pdfFilename
        });
        console.log('Spreadsheet record updated with final status');
      } catch (error) {
        console.error('Error updating spreadsheet record:', error);
      }
    }
    
    console.log(`Message processed successfully: ${messageId}`);
    return true;
    
  } catch (error) {
    console.error('Error processing message:', error);
    
    // Update spreadsheet record with error status
    if (CONFIG.ENABLE_SPREADSHEET_LOGGING && messageId) {
      try {
        updateRecordStatus(messageId, {
          status: 'Error',
          error: error.message
        });
      } catch (updateError) {
        console.error('Error updating spreadsheet record with error:', updateError);
      }
    }
    
    return false;
  }
}

/**
 * Format email body for Slack display with proper length handling
 * Slack表示用にメール本文を適切な長さでフォーマット
 * 
 * @param {string} body - Raw email body
 * @returns {string} - Formatted body for Slack
 */
function formatEmailBody(body) {
  try {
    if (!body || body.trim().length === 0) {
      return '_本文なし_';
    }
    
    // Clean up the body text
    let cleanBody = body
      // Remove excessive whitespace
      .replace(/\n\s*\n\s*\n/g, '\n\n')
      // Remove leading/trailing whitespace
      .trim();
    
    // If full body display is disabled, use shorter preview
    if (!CONFIG.SHOW_FULL_EMAIL_BODY && cleanBody.length > 500) {
      const shortPreview = cleanBody.substring(0, 500);
      const lastBreak = Math.max(
        shortPreview.lastIndexOf('\u3002'),
        shortPreview.lastIndexOf('.'),
        shortPreview.lastIndexOf('\n')
      );
      const cutPoint = lastBreak > 400 ? lastBreak + 1 : 500;
      return cleanBody.substring(0, cutPoint) + '...\n\n_[簡略表示モード]_';
    }
    
    // Check if body exceeds Slack field limit (approximately 8000 characters)
    if (cleanBody.length <= CONFIG.BODY_PREVIEW_LENGTH) {
      return cleanBody;
    }
    
    // If too long, truncate smartly
    const truncated = cleanBody.substring(0, CONFIG.BODY_PREVIEW_LENGTH);
    
    // Try to end at a sentence boundary
    const lastSentenceEnd = Math.max(
      truncated.lastIndexOf('\u3002'),  // Japanese period
      truncated.lastIndexOf('.'),   // English period
      truncated.lastIndexOf('\n\n')  // Paragraph break
    );
    
    if (lastSentenceEnd > CONFIG.BODY_PREVIEW_LENGTH * 0.8) {
      // If we found a good break point near the end, use it
      return truncated.substring(0, lastSentenceEnd + 1) + '\n\n_[以下略]_';
    } else {
      // Otherwise, just truncate and add indicator
      return truncated + '...\n\n_[以下略]_';
    }
    
  } catch (error) {
    console.error('Error formatting email body:', error);
    return '_本文表示エラー_';
  }
}

/**
 * Check if message has already been processed using spreadsheet tracking
 * スプレッドシート追跡でメッセージが既に処理済みかどうかをチェック
 * 
 * @param {GmailMessage} message - Gmail message object
 * @returns {boolean} - true if message was already processed
 */
function isMessageAlreadyProcessed(message) {
  try {
    const messageId = message.getId();
    
    // Check by message ID only (each recipient gets processed separately)
    return isMessageProcessedInSpreadsheet(messageId);
    
  } catch (error) {
    console.error('Error checking if message is processed:', error);
    return false; // Assume not processed if check fails
  }
}

/**
 * Get all recipients (To, CC, BCC) for a message
 * メッセージの全受信者（To、CC、BCC）を取得
 * 
 * @param {GmailMessage} message - Gmail message object
 * @returns {Array} - Array of recipient email addresses
 */
function getMessageRecipients(message) {
  try {
    const recipients = [];
    
    // Get To recipients
    const toField = message.getTo();
    if (toField) {
      recipients.push(...extractEmailAddresses(toField));
    }
    
    // Get CC recipients
    const ccField = message.getCc();
    if (ccField) {
      recipients.push(...extractEmailAddresses(ccField));
    }
    
    // Get BCC recipients (note: BCC may not be available in forwarded emails)
    const bccField = message.getBcc();
    if (bccField) {
      recipients.push(...extractEmailAddresses(bccField));
    }
    
    // Remove duplicates and return
    return [...new Set(recipients)];
    
  } catch (error) {
    console.error('Error getting message recipients:', error);
    return [];
  }
}

/**
 * Extract email addresses from a recipient field string
 * 受信者フィールド文字列からメールアドレスを抽出
 * 
 * @param {string} field - Recipient field string
 * @returns {Array} - Array of email addresses
 */
function extractEmailAddresses(field) {
  try {
    // Handle format like: "Name <email@domain.com>, Another Name <email2@domain.com>"
    const emailRegex = /([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/g;
    const matches = field.match(emailRegex) || [];
    return matches.map(email => email.toLowerCase().trim());
  } catch (error) {
    console.error('Error extracting email addresses:', error);
    return [];
  }
}

/**
 * Create a hash for recipient combination for tracking
 * 追跡用の受信者組み合わせハッシュを作成
 * 
 * @param {Array} recipients - Array of recipient email addresses
 * @returns {string} - Hash string for the recipient combination
 */
function createRecipientHash(recipients) {
  try {
    // Sort recipients to ensure consistent hash regardless of order
    const sortedRecipients = recipients.sort();
    const recipientString = sortedRecipients.join(',');
    
    // Create a simple hash (for Apps Script compatibility)
    const hash = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, recipientString));
    
    // Truncate to reasonable length for property key
    return hash.substring(0, 16);
  } catch (error) {
    console.error('Error creating recipient hash:', error);
    return 'default';
  }
}

/**
 * Advanced subject pattern checking with multiple pattern support
 * 複数パターン対応の高度な件名パターンチェック
 * 
 * @param {string} subject - Email subject
 * @returns {Object} - Pattern match result with details
 */
function checkSubjectPattern(subject) {
  try {
    console.log(`Checking subject pattern for: "${subject}"`);
    
    // Use advanced pattern matching if enabled
    if (CONFIG.SUBJECT_PATTERNS && CONFIG.SUBJECT_PATTERNS.ENABLE_MULTIPLE_PATTERNS) {
      return checkMultiplePatterns(subject);
    }
    
    // Fallback to legacy single pattern
    const isMatch = CONFIG.SUBJECT_PATTERN.test(subject);
    return {
      isMatch: isMatch,
      matchedPattern: isMatch ? CONFIG.SUBJECT_PATTERN.toString() : null,
      checkedPatterns: [CONFIG.SUBJECT_PATTERN.toString()],
      matchDetails: isMatch ? extractLegacyPatternInfo(subject) : null
    };
    
  } catch (error) {
    console.error('Error in subject pattern checking:', error);
    return {
      isMatch: false,
      error: error.message,
      checkedPatterns: []
    };
  }
}

/**
 * Check subject against multiple patterns
 * 複数パターンに対する件名チェック
 * 
 * @param {string} subject - Email subject
 * @returns {Object} - Detailed match result
 */
function checkMultiplePatterns(subject) {
  const patterns = CONFIG.SUBJECT_PATTERNS.PATTERNS || [];
  const matchMode = CONFIG.SUBJECT_PATTERNS.MATCH_MODE || 'any';
  
  const results = [];
  const checkedPatterns = [];
  
  for (let i = 0; i < patterns.length; i++) {
    const pattern = patterns[i];
    const patternString = pattern.toString();
    checkedPatterns.push(patternString);
    
    try {
      const isMatch = pattern.test(subject);
      
      results.push({
        pattern: patternString,
        isMatch: isMatch,
        matchDetails: isMatch ? subject.match(pattern) : null
      });
      
      console.log(`Pattern ${i + 1}: ${patternString} -> ${isMatch ? 'MATCH' : 'NO MATCH'}`);
      
      // For 'any' mode, return immediately on first match
      if (matchMode === 'any' && isMatch) {
        return {
          isMatch: true,
          matchedPattern: patternString,
          checkedPatterns: checkedPatterns,
          allResults: results,
          matchMode: matchMode
        };
      }
      
    } catch (error) {
      console.error(`Error testing pattern ${i + 1}:`, error);
      results.push({
        pattern: patternString,
        isMatch: false,
        error: error.message
      });
    }
  }
  
  // Determine final result based on match mode
  const hasMatches = results.some(r => r.isMatch);
  const allMatch = results.every(r => r.isMatch);
  
  const finalMatch = (matchMode === 'any') ? hasMatches : allMatch;
  
  return {
    isMatch: finalMatch,
    matchedPattern: finalMatch ? results.find(r => r.isMatch)?.pattern : null,
    checkedPatterns: checkedPatterns,
    allResults: results,
    matchMode: matchMode,
    summary: {
      totalPatterns: patterns.length,
      matchedCount: results.filter(r => r.isMatch).length,
      finalResult: finalMatch
    }
  };
}

/**
 * Extract subject pattern match information (legacy support)
 * 件名パターンマッチ情報を抽出（レガシーサポート）
 * 
 * @param {string} subject - Email subject
 * @returns {Object} - Pattern match information
 */
function extractLegacyPatternInfo(subject) {
  try {
    const match = subject.match(CONFIG.SUBJECT_PATTERN);
    if (!match) {
      return null;
    }
    
    return {
      fullMatch: match[0],
      originalSubject: subject,
      matchIndex: match.index
    };
  } catch (error) {
    console.error('Error extracting legacy pattern info:', error);
    return null;
  }
}

/**
 * Generate safe filename from subject and timestamp
 * 件名とタイムスタンプから安全なファイル名を生成
 * 
 * @param {string} subject - Email subject
 * @param {number} index - Attachment index
 * @param {string} originalName - Original attachment filename
 * @returns {string} - Safe filename
 */
function generateSafeFilename(subject, index, originalName) {
  try {
    const timestamp = Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd_HHmmss');
    
    // Clean subject for filename (remove special characters)
    const cleanSubject = subject
      .replace(/[<>:"/\\|?*]/g, '_')  // Replace invalid filename characters
      .substring(0, 50);  // Limit length
    
    // Extract extension from original name
    const extension = originalName.includes('.') 
      ? originalName.split('.').pop() 
      : '';
    
    const extensionPart = extension ? `.${extension}` : '';
    
    return `${cleanSubject}_${timestamp}_${index + 1}${extensionPart}`;
    
  } catch (error) {
    console.error('Error generating safe filename:', error);
    // Fallback to simple timestamp-based name
    const timestamp = Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd_HHmmss');
    return `attachment_${timestamp}_${index + 1}`;
  }
}

/**
 * Test function for message processing with organizational patterns
 * 組織パターンを含むメッセージ処理のテスト関数
 */
function testMessageProcessing() {
  console.log('=== TESTING Message Processing with Organizational Patterns ===');
  
  try {
    // Test organizational subject pattern matching
    const organizationalTestSubjects = [
      '組織メルマガ／会員企業からのお知らせ＆PR',
      '【本日開催】第14回部会開催のご案内※6/5（木）15:00開催',
      '【再送】第14回部会開催のご案内※6/5（木）15:00開催',
      '勉強会『最新技術の先行事例 ? 海外の成功例を中心に』、『関連取引所の取組みについて』他　※5月27日(火)17:00開',
      '第18回税制検討部会開催のご案内※5/28（水）13:00開催',
      '第2回ブロックチェーン技術部会開催のご案内※6/6（金）11:00開催',
      '普通のメール件名（マッチしないはず）',
      '第1回 月次報告（レガシーパターンテスト）'
    ];
    
    console.log('Testing organizational subject pattern matching:');
    console.log(`Multiple patterns enabled: ${CONFIG.SUBJECT_PATTERNS?.ENABLE_MULTIPLE_PATTERNS}`);
    console.log(`Match mode: ${CONFIG.SUBJECT_PATTERNS?.MATCH_MODE}`);
    console.log(`Total patterns: ${CONFIG.SUBJECT_PATTERNS?.PATTERNS?.length || 0}`);
    console.log('');
    
    organizationalTestSubjects.forEach((subject, index) => {
      console.log(`--- Test ${index + 1} ---`);
      console.log(`Subject: "${subject}"`);
      
      const result = checkSubjectPattern(subject);
      console.log(`Result: ${result.isMatch ? '✅ MATCH' : '❌ NO MATCH'}`);
      
      if (result.isMatch) {
        console.log(`Matched pattern: ${result.matchedPattern}`);
        if (result.summary) {
          console.log(`Patterns checked: ${result.summary.totalPatterns}, Matched: ${result.summary.matchedCount}`);
        }
      } else if (result.checkedPatterns) {
        console.log(`Checked ${result.checkedPatterns.length} patterns`);
      }
      
      if (result.error) {
        console.log(`Error: ${result.error}`);
      }
      
      console.log('');
    });
    
    // Test filename generation with organizational subjects
    console.log('Testing filename generation with organizational subjects:');
    const filenameTests = [
      { subject: '【本日開催】第14回部会開催のご案内※6/5（木）15:00開催', file: 'agenda.pdf' },
      { subject: '組織メルマガ／会員企業からのお知らせ＆PR', file: 'newsletter.docx' },
      { subject: '勉強会『最新技術の先行事例』※5月27日', file: 'presentation.pptx' }
    ];
    
    filenameTests.forEach((test, index) => {
      const filename = generateSafeFilename(test.subject, index, test.file);
      console.log(`${index + 1}. "${test.subject}" -> "${filename}"`);
    });
    
    console.log('\nMessage processing test completed successfully');
    
  } catch (error) {
    console.error('Message processing test failed:', error);
    throw error;
  }
}

/**
 * Test individual pattern matching for debugging
 * デバッグ用の個別パターンマッチングテスト
 */
function testIndividualPatterns() {
  console.log('=== TESTING Individual Patterns ===');
  
  if (!CONFIG.SUBJECT_PATTERNS?.PATTERNS) {
    console.log('No multiple patterns configured');
    return;
  }
  
  const testSubject = '【本日開催】第14回部会開催のご案内※6/5（木）15:00開催';
  console.log(`Test subject: "${testSubject}"`);
  console.log('');
  
  CONFIG.SUBJECT_PATTERNS.PATTERNS.forEach((pattern, index) => {
    try {
      const isMatch = pattern.test(testSubject);
      console.log(`Pattern ${index + 1}: ${pattern.toString()}`);
      console.log(`Result: ${isMatch ? '✅ MATCH' : '❌ NO MATCH'}`);
      
      if (isMatch) {
        const match = testSubject.match(pattern);
        console.log(`Matched text: "${match[0]}"`);
      }
      
      console.log('');
    } catch (error) {
      console.error(`Error in pattern ${index + 1}:`, error);
    }
  });
}

/**
 * Get message recipient (To field)
 * メッセージの受信者（To フィールド）を取得
 * 
 * @param {GmailMessage} message - Gmail message object
 * @returns {string} - Recipient email address
 */
function getMessageRecipient(message) {
  try {
    // Get the To field from the message
    const to = message.getTo();
    
    if (!to || to.trim().length === 0) {
      console.log('No To field found in message');
      return '';
    }
    
    // Extract the first email address if multiple recipients
    const emails = extractEmailAddresses(to);
    const primaryRecipient = emails.length > 0 ? emails[0] : to;
    
    console.log(`Primary recipient: ${primaryRecipient}`);
    return primaryRecipient;
    
  } catch (error) {
    console.error('Error getting message recipient:', error);
    return '';
  }
}

/**
 * Extract email addresses from a string field
 * 文字列フィールドからメールアドレスを抽出
 * 
 * @param {string} field - Email field string (can contain multiple addresses)
 * @returns {Array} - Array of email addresses
 */
function extractEmailAddresses(field) {
  try {
    if (!field || field.trim().length === 0) {
      return [];
    }
    
    // Regular expression to match email addresses
    const emailPattern = /[\w\.-]+@[\w\.-]+\.\w+/g;
    const matches = field.match(emailPattern);
    
    return matches || [];
    
  } catch (error) {
    console.error('Error extracting email addresses:', error);
    return [];
  }
}