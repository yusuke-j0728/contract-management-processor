/**
 * Google Drive Attachment Management Module
 * 
 * Handles saving email attachments to Google Drive with proper file organization
 * and error handling.
 */

/**
 * Process and save email attachments directly to Google Drive
 * メール添付ファイルを直接Google Driveに保存（サブフォルダ作成なし）
 * 
 * @param {Array} attachments - Array of Gmail attachment objects
 * @param {string} subject - Email subject for filename generation
 * @param {Date} emailDate - Email date for filename generation
 * @returns {Array} - Array of attachment info objects
 */
function processAttachments(attachments, subject, emailDate = new Date(), sender = '') {
  console.log(`Processing ${attachments.length} attachments for subject: ${subject}`);
  
  if (attachments.length === 0) {
    return [];
  }
  
  const attachmentInfo = [];
  const contractBaseFolderId = getProperty(PROPERTY_KEYS.DRIVE_FOLDER_ID);
  
  // Extract PDF filenames for duplicate detection
  const pdfNames = attachments
    .filter(attachment => attachment.getName().toLowerCase().endsWith('.pdf'))
    .map(attachment => attachment.getName());
  
  console.log(`Found ${pdfNames.length} PDF files: [${pdfNames.join(', ')}]`);
  
  // Check for content duplicate
  let contentKey = null;
  let duplicateInfo = { isDuplicate: false, existingUrl: null };
  
  if (pdfNames.length > 0) {
    contentKey = generateContentDuplicateKey(sender, emailDate, subject, pdfNames);
    duplicateInfo = checkContentDuplicate(contentKey);
    
    console.log(`Content duplicate check: ${duplicateInfo.isDuplicate ? 'DUPLICATE FOUND' : 'NEW CONTENT'}`);
    if (duplicateInfo.isDuplicate) {
      console.log(`Existing Drive URL: ${duplicateInfo.existingUrl}`);
    }
  }
  
  try {
    const contractBaseFolder = DriveApp.getFolderById(contractBaseFolderId);
    console.log(`Using contract base folder: ${contractBaseFolder.getName()} (ID: ${contractBaseFolderId})`);
    console.log(`Contract base folder URL: ${contractBaseFolder.getUrl()}`);
    
    attachments.forEach((attachment, index) => {
      try {
        const fileName = attachment.getName();
        console.log(`Processing attachment ${index + 1}/${attachments.length}: ${fileName}`);
        
        // Check if it's a PDF file
        const isPdf = fileName.toLowerCase().endsWith('.pdf');
        
        if (!isPdf) {
          console.log(`Skipping non-PDF file: ${fileName}`);
          attachmentInfo.push({
            originalName: fileName,
            savedName: null,
            size: attachment.getSize(),
            driveUrl: null,
            fileId: null,
            folderPath: null,
            skipped: 'Not a PDF file',
            contentKey: null,
            isDuplicate: false
          });
          return;
        }
        
        // If this is a duplicate content, skip PDF saving but provide existing URL
        if (duplicateInfo.isDuplicate) {
          console.log(`Skipping PDF save for duplicate content: ${fileName}`);
          attachmentInfo.push({
            originalName: fileName,
            savedName: null,
            size: attachment.getSize(),
            driveUrl: duplicateInfo.existingUrl, // This is the folder URL for existing content
            fileId: null,
            folderPath: null,
            folderUrl: duplicateInfo.existingUrl,
            pdfDirectUrl: duplicateInfo.existingUrl, // For existing PDFs, use the stored URL
            skipped: 'Duplicate content - PDF already saved',
            contentKey: contentKey,
            isDuplicate: true,
            existingUrl: duplicateInfo.existingUrl
          });
          return;
        }
        
        // Save PDF to Drive (new content)
        const info = saveAttachmentToDrive(attachment, subject, index, contractBaseFolder, emailDate);
        info.contentKey = contentKey;
        info.isDuplicate = false;
        attachmentInfo.push(info);
        
        console.log(`Successfully saved: ${info.savedName}`);
        console.log(`File URL: ${info.driveUrl}`);
        console.log(`Saved in folder: ${info.folderPath}`);
        console.log(`Folder URL: ${info.folderUrl}`);
        
        // Record this content as saved to prevent future duplicates
        if (index === 0 && contentKey) { // Only record once per email
          recordContentDuplicate(contentKey, sender, emailDate, subject, pdfNames, info.folderUrl || contractBaseFolder.getUrl());
        }
        
      } catch (error) {
        console.error(`Failed to save attachment ${index + 1} (${attachment.getName()}):`, error);
        
        // Add error info instead of skipping completely
        attachmentInfo.push({
          originalName: attachment.getName(),
          savedName: null,
          size: attachment.getSize(),
          driveUrl: null,
          fileId: null,
          folderPath: null,
          error: error.message,
          contentKey: contentKey,
          isDuplicate: false
        });
      }
    });
    
    const savedCount = attachmentInfo.filter(info => !info.error && !info.skipped).length;
    const duplicateCount = attachmentInfo.filter(info => info.isDuplicate).length;
    const skippedCount = attachmentInfo.filter(info => info.skipped && !info.isDuplicate).length;
    
    console.log(`Contract processing complete:`);
    console.log(`  Saved: ${savedCount}, Duplicates: ${duplicateCount}, Skipped: ${skippedCount}`);
    
  } catch (error) {
    console.error('Error accessing contract base folder:', error);
    throw new Error(`Contract base folder access failed: ${error.message}`);
  }
  
  return attachmentInfo;
}

/**
 * Create contract folder with date and subject (flat structure)
 * 日付と件名で契約フォルダを作成（フラット構造）
 * 
 * @param {DriveFolder} contractBaseFolder - Contract base folder
 * @param {string} subject - Email subject
 * @param {Date} emailDate - Email date
 * @returns {DriveFolder} - Created or existing contract folder
 */
function createContractFolder(contractBaseFolder, subject, emailDate) {
  try {
    // Format date as YYYYMMDD (no hyphens for contract folders)
    const dateStr = Utilities.formatDate(emailDate, 'JST', 'yyyyMMdd');
    
    // Clean subject for folder name
    const cleanSubject = cleanSubjectForFolder(subject);
    
    // Create folder name: "YYYYMMDD_Subject" (directly in contract base folder)
    const folderName = `${dateStr}_${cleanSubject}`;
    
    console.log(`Creating contract folder: ${folderName}`);
    console.log(`Date string: ${dateStr}`);
    console.log(`Clean subject: ${cleanSubject}`);
    console.log(`Full folder name: ${folderName}`);
    
    // Check if folder already exists in the contract base folder
    const existingFolders = contractBaseFolder.getFoldersByName(folderName);
    if (existingFolders.hasNext()) {
      const existingFolder = existingFolders.next();
      console.log(`Using existing contract folder: ${folderName}`);
      return existingFolder;
    }
    
    // Create new contract folder directly in base folder
    const contractFolder = contractBaseFolder.createFolder(folderName);
    console.log(`Created new contract folder: ${folderName}`);
    console.log(`New contract folder URL: ${contractFolder.getUrl()}`);
    console.log(`New contract folder ID: ${contractFolder.getId()}`);
    
    return contractFolder;
    
  } catch (error) {
    console.error('Error creating contract folder:', error);
    // Fallback to base folder if folder creation fails
    return contractBaseFolder;
  }
}

/**
 * Clean email subject for use as folder name
 * メール件名をフォルダ名として使用できるようにクリーニング
 * 
 * @param {string} subject - Email subject
 * @returns {string} - Cleaned subject suitable for folder name
 */
function cleanSubjectForFolder(subject) {
  try {
    return subject
      // Remove or replace invalid characters for folder names
      .replace(/[<>:"/\\|?*]/g, '_')
      // Replace multiple spaces with single space
      .replace(/\s+/g, ' ')
      // Remove leading/trailing spaces
      .trim()
      // Limit length to avoid very long folder names
      .substring(0, 100)
      // Replace Japanese brackets that might cause issues
      .replace(/[【】]/g, '_')
      .replace(/[「」]/g, '_')
      // Replace forward slashes (common in subjects)
      .replace(/[／\/]/g, '_');
  } catch (error) {
    console.error('Error cleaning subject for folder:', error);
    // Fallback to timestamp if subject cleaning fails
    return `Email_${Utilities.formatDate(new Date(), 'JST', 'HHmmss')}`;
  }
}

/**
 * Save individual attachment directly to Google Drive
 * 個別の添付ファイルを直接Google Driveに保存
 * 
 * @param {GmailAttachment} attachment - Gmail attachment object
 * @param {string} subject - Email subject
 * @param {number} index - Attachment index
 * @param {DriveFolder} folder - Drive folder object
 * @param {Date} emailDate - Email date for filename generation
 * @returns {Object} - Attachment info object
 */
function saveAttachmentToDrive(attachment, subject, index, folder, emailDate) {
  const startTime = new Date().getTime();
  
  try {
    const originalName = attachment.getName();
    const size = attachment.getSize();
    
    console.log(`Saving attachment: ${originalName} (${formatFileSize(size)})`);
    
    // Generate filename with date prefix (YYYYMMDD_filename)
    const dateStr = Utilities.formatDate(emailDate, 'JST', 'yyyyMMdd');
    const safeFileName = generateDatePrefixedFilename(dateStr, index, originalName);
    
    // Check for duplicate filenames
    const finalFileName = ensureUniqueFilename(folder, safeFileName);
    
    // Create file in Drive
    const blob = attachment.copyBlob().setName(finalFileName);
    const file = folder.createFile(blob);
    
    const endTime = new Date().getTime();
    const uploadTime = endTime - startTime;
    
    console.log(`Upload completed in ${uploadTime}ms`);
    
    return {
      originalName: originalName,
      savedName: finalFileName,
      size: size,
      driveUrl: file.getUrl(),
      fileId: file.getId(),
      folderPath: folder.getName(),
      folderUrl: folder.getUrl(),
      pdfDirectUrl: file.getUrl(), // Direct link to the PDF file
      uploadTime: uploadTime
    };
    
  } catch (error) {
    console.error('Error saving attachment to Drive:', error);
    throw error;
  }
}

/**
 * Generate filename with date prefix
 * 日付プレフィックス付きのファイル名を生成
 * 
 * @param {string} dateStr - Date string in YYYYMMDD format
 * @param {number} index - Attachment index
 * @param {string} originalName - Original filename
 * @returns {string} - Filename with date prefix
 */
function generateDatePrefixedFilename(dateStr, index, originalName) {
  try {
    const timestamp = Utilities.formatDate(new Date(), 'JST', 'HHmmss');
    
    if (originalName && originalName.trim()) {
      // Clean the original name
      const cleanName = originalName
        .replace(/[<>:"/\\|?*]/g, '_')
        .trim();
      
      // Add date prefix to filename
      return `${dateStr}_${cleanName}`;
    }
    
    // Fallback for unnamed attachments
    return `${dateStr}_attachment_${index + 1}_${timestamp}`;
    
  } catch (error) {
    console.error('Error generating date-prefixed filename:', error);
    return `${dateStr}_attachment_${index + 1}_${new Date().getTime()}`;
  }
}

/**
 * Generate simple filename for organized folder structure
 * 整理されたフォルダ構造用のシンプルなファイル名を生成
 * 
 * @param {number} index - Attachment index
 * @param {string} originalName - Original filename
 * @returns {string} - Simple filename
 */
function generateSimpleFilename(index, originalName) {
  try {
    const timestamp = Utilities.formatDate(new Date(), 'JST', 'HHmmss');
    
    // For organized folders, use simpler naming
    if (originalName && originalName.trim()) {
      // Clean the original name
      const cleanName = originalName
        .replace(/[<>:"/\\|?*]/g, '_')
        .trim();
      
      // If multiple attachments, add index
      if (index > 0) {
        const nameParts = cleanName.split('.');
        const extension = nameParts.length > 1 ? `.${nameParts.pop()}` : '';
        const baseName = nameParts.join('.');
        return `${baseName}_${index + 1}${extension}`;
      }
      
      return cleanName;
    }
    
    // Fallback for unnamed attachments
    return `attachment_${index + 1}_${timestamp}`;
    
  } catch (error) {
    console.error('Error generating simple filename:', error);
    return `attachment_${index + 1}_${new Date().getTime()}`;
  }
}

/**
 * Ensure filename is unique in the Drive folder
 * Driveフォルダ内でファイル名が一意になることを保証
 * 
 * @param {DriveFolder} folder - Drive folder object
 * @param {string} fileName - Desired filename
 * @returns {string} - Unique filename
 */
function ensureUniqueFilename(folder, fileName) {
  try {
    const files = folder.getFilesByName(fileName);
    
    if (!files.hasNext()) {
      // No conflict, use original name
      return fileName;
    }
    
    // File exists, generate unique name
    const nameParts = fileName.split('.');
    const extension = nameParts.length > 1 ? `.${nameParts.pop()}` : '';
    const baseName = nameParts.join('.');
    
    let counter = 1;
    let uniqueName;
    
    do {
      uniqueName = `${baseName}_${counter}${extension}`;
      const conflictFiles = folder.getFilesByName(uniqueName);
      
      if (!conflictFiles.hasNext()) {
        break;
      }
      
      counter++;
    } while (counter < 100); // Safety limit
    
    console.log(`Generated unique filename: ${uniqueName}`);
    return uniqueName;
    
  } catch (error) {
    console.error('Error ensuring unique filename:', error);
    // Fallback to timestamp-based uniqueness
    const timestamp = new Date().getTime();
    return `${fileName}_${timestamp}`;
  }
}

/**
 * Format file size for human reading
 * ファイルサイズを人間が読める形式にフォーマット
 * 
 * @param {number} bytes - File size in bytes
 * @returns {string} - Formatted file size
 */
function formatFileSize(bytes) {
  if (bytes === 0) return '0 B';
  
  const k = 1024;
  const sizes = ['B', 'KB', 'MB', 'GB', 'TB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  
  return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
}

/**
 * Get Drive folder usage statistics
 * Driveフォルダの使用状況統計を取得
 * 
 * @returns {Object} - Folder statistics
 */
function getDriveFolderStats() {
  try {
    const folderId = getProperty(PROPERTY_KEYS.DRIVE_FOLDER_ID);
    const folder = DriveApp.getFolderById(folderId);
    
    const files = folder.getFiles();
    let fileCount = 0;
    let totalSize = 0;
    
    while (files.hasNext()) {
      const file = files.next();
      fileCount++;
      totalSize += file.getSize();
    }
    
    return {
      folderName: folder.getName(),
      fileCount: fileCount,
      totalSize: totalSize,
      formattedSize: formatFileSize(totalSize),
      folderUrl: folder.getUrl()
    };
    
  } catch (error) {
    console.error('Error getting Drive folder stats:', error);
    return null;
  }
}

/**
 * Clean up old files in Drive folder (optional maintenance function)
 * Driveフォルダ内の古いファイルをクリーンアップ（オプションのメンテナンス関数）
 * 
 * @param {number} daysOld - Files older than this many days will be deleted
 * @returns {Object} - Cleanup results
 */
function cleanupOldFiles(daysOld = 90) {
  try {
    console.log(`Starting cleanup of files older than ${daysOld} days...`);
    
    const folderId = getProperty(PROPERTY_KEYS.DRIVE_FOLDER_ID);
    const folder = DriveApp.getFolderById(folderId);
    
    const cutoffDate = new Date();
    cutoffDate.setDate(cutoffDate.getDate() - daysOld);
    
    const files = folder.getFiles();
    let deletedCount = 0;
    let freedSpace = 0;
    
    while (files.hasNext()) {
      const file = files.next();
      const fileDate = file.getDateCreated();
      
      if (fileDate < cutoffDate) {
        console.log(`Deleting old file: ${file.getName()} (${fileDate})`);
        freedSpace += file.getSize();
        file.setTrashed(true);
        deletedCount++;
      }
    }
    
    console.log(`Cleanup completed. Deleted ${deletedCount} files, freed ${formatFileSize(freedSpace)}`);
    
    return {
      deletedCount: deletedCount,
      freedSpace: freedSpace,
      formattedFreedSpace: formatFileSize(freedSpace)
    };
    
  } catch (error) {
    console.error('Error during cleanup:', error);
    throw error;
  }
}

/**
 * Test function for Drive operations
 * Drive操作のテスト関数
 */
function testDriveOperations() {
  console.log('=== TESTING Drive Operations ===');
  
  try {
    // Test base folder access
    const folderId = getProperty(PROPERTY_KEYS.DRIVE_FOLDER_ID);
    const baseFolder = DriveApp.getFolderById(folderId);
    console.log(`✓ Base Drive folder accessible: ${baseFolder.getName()}`);
    
    // Test date-prefixed filename generation
    const testDate = new Date();
    const dateStr = Utilities.formatDate(testDate, 'JST', 'yyyyMMdd');
    const datePrefixedFileName = generateDatePrefixedFilename(dateStr, 0, 'test-contract.pdf');
    console.log(`✓ Date-prefixed filename generated: ${datePrefixedFileName}`);
    
    // Test simple filename generation
    const simpleFileName = generateSimpleFilename(0, 'test-document.pdf');
    console.log(`✓ Simple filename generated: ${simpleFileName}`);
    
    // Test unique filename generation
    const uniqueName = ensureUniqueFilename(baseFolder, datePrefixedFileName);
    console.log(`✓ Unique filename: ${uniqueName}`);
    
    // Test file size formatting
    const sizes = [0, 1024, 1048576, 1073741824];
    sizes.forEach(size => {
      console.log(`✓ ${size} bytes = ${formatFileSize(size)}`);
    });
    
    // Test folder stats
    const stats = getDriveFolderStats();
    if (stats) {
      console.log(`✓ Folder stats: ${stats.fileCount} files, ${stats.formattedSize}`);
    }
    
    // Test filename generation with various patterns
    console.log('\n--- Testing Filename Generation ---');
    const testFilenames = [
      'contract_agreement.pdf',
      'NDA_company_20240622.pdf',
      'employment_contract.docx'
    ];
    
    testFilenames.forEach((filename, index) => {
      const prefixedName = generateDatePrefixedFilename(dateStr, index, filename);
      console.log(`✓ Generated filename ${index + 1}: ${prefixedName}`);
    });
    
    console.log('Drive operations test completed successfully');
    
  } catch (error) {
    console.error('Drive operations test failed:', error);
    throw error;
  }
}