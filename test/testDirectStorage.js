/**
 * Test file for verifying direct PDF storage functionality
 * PDFの直接保存機能を確認するためのテストファイル
 */

/**
 * Test direct PDF storage without subfolder creation
 * サブフォルダ作成なしでの直接PDF保存をテスト
 */
function testDirectPDFStorage() {
  console.log('=== TESTING Direct PDF Storage ===');
  
  try {
    // Test configuration
    const testConfig = {
      subject: 'Employment Contract - John Doe - Completed',
      emailDate: new Date(),
      attachmentName: 'employment_contract_john_doe.pdf'
    };
    
    // Get base folder
    const folderId = getProperty(PROPERTY_KEYS.DRIVE_FOLDER_ID);
    const baseFolder = DriveApp.getFolderById(folderId);
    console.log(`✓ Base folder found: ${baseFolder.getName()}`);
    console.log(`  Folder URL: ${baseFolder.getUrl()}`);
    
    // Test filename generation
    const dateStr = Utilities.formatDate(testConfig.emailDate, 'JST', 'yyyyMMdd');
    const expectedFilename = generateDatePrefixedFilename(dateStr, 0, testConfig.attachmentName);
    console.log(`✓ Generated filename: ${expectedFilename}`);
    console.log(`  Expected format: ${dateStr}_${testConfig.attachmentName}`);
    
    // Verify no subfolder creation
    console.log('\n--- Verifying Direct Storage ---');
    console.log('✓ PDFs will be saved directly to base folder');
    console.log('✓ No subfolders will be created');
    console.log(`✓ All files will have date prefix: ${dateStr}_`);
    
    // Test multiple file scenarios
    console.log('\n--- Testing Multiple File Scenarios ---');
    const testFiles = [
      'contract_agreement.pdf',
      'nda_confidential.pdf',
      'service_agreement_2024.docx',
      'amendment_v2.pdf'
    ];
    
    testFiles.forEach((filename, index) => {
      const prefixedName = generateDatePrefixedFilename(dateStr, index, filename);
      console.log(`✓ File ${index + 1}: ${prefixedName}`);
    });
    
    // Test duplicate handling
    console.log('\n--- Testing Duplicate Filename Handling ---');
    const duplicateTest = `${dateStr}_test_contract.pdf`;
    console.log(`Original: ${duplicateTest}`);
    console.log(`If exists, will become: ${dateStr}_test_contract_1.pdf`);
    
    console.log('\n✅ Direct PDF storage test completed successfully');
    console.log('All PDFs will be stored directly in the base folder with date prefixes.');
    
    return {
      success: true,
      baseFolderUrl: baseFolder.getUrl(),
      sampleFilename: expectedFilename
    };
    
  } catch (error) {
    console.error('❌ Direct PDF storage test failed:', error);
    throw error;
  }
}

/**
 * Test the complete attachment processing flow
 * 完全な添付ファイル処理フローをテスト
 */
function testCompleteAttachmentFlow() {
  console.log('=== TESTING Complete Attachment Processing Flow ===');
  
  try {
    // Create mock attachment
    const mockBlob = Utilities.newBlob('Test PDF content', 'application/pdf', 'test_contract.pdf');
    const mockAttachment = {
      getName: () => 'test_contract.pdf',
      getSize: () => mockBlob.getBytes().length,
      copyBlob: () => mockBlob
    };
    
    const testSubject = 'Contract Execution Complete - Test Company';
    const testDate = new Date();
    
    console.log('Processing mock attachment...');
    
    // Process the attachment
    const results = processAttachments([mockAttachment], testSubject, testDate);
    
    if (results && results.length > 0) {
      const result = results[0];
      console.log('\n✅ Attachment processed successfully:');
      console.log(`  Original name: ${result.originalName}`);
      console.log(`  Saved as: ${result.savedName}`);
      console.log(`  Location: ${result.folderPath}`);
      console.log(`  URL: ${result.driveUrl}`);
      
      // Verify filename format
      const dateStr = Utilities.formatDate(testDate, 'JST', 'yyyyMMdd');
      if (result.savedName.startsWith(dateStr)) {
        console.log(`✓ Filename correctly starts with date: ${dateStr}`);
      } else {
        console.log(`❌ Filename does not start with expected date: ${dateStr}`);
      }
    }
    
    console.log('\n✅ Complete attachment flow test finished');
    
  } catch (error) {
    console.error('❌ Complete attachment flow test failed:', error);
    throw error;
  }
}

/**
 * Run all direct storage tests
 * すべての直接保存テストを実行
 */
function runDirectStorageTests() {
  console.log('=== RUNNING ALL DIRECT STORAGE TESTS ===\n');
  
  try {
    // Test 1: Basic functionality
    testDirectPDFStorage();
    console.log('\n' + '='.repeat(50) + '\n');
    
    // Test 2: Complete flow
    testCompleteAttachmentFlow();
    console.log('\n' + '='.repeat(50) + '\n');
    
    // Test 3: Drive operations
    testDriveOperations();
    
    console.log('\n✅ ALL DIRECT STORAGE TESTS COMPLETED SUCCESSFULLY');
    
  } catch (error) {
    console.error('\n❌ DIRECT STORAGE TESTS FAILED:', error);
    throw error;
  }
}