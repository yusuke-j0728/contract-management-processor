/**
 * Test file for content-based duplicate detection
 * コンテンツベース重複検知のテストファイル
 */

/**
 * Test content-based duplicate key generation
 * コンテンツベース重複キー生成をテスト
 */
function testContentDuplicateKeyGeneration() {
  console.log('=== TESTING Content Duplicate Key Generation ===');
  
  try {
    const testCases = [
      {
        sender: 'noreply@contracttool7.example.com',
        date: new Date('2024-06-22T10:30:00'),
        subject: "You've been copied on Loan Agreement - signed by John Doe",
        pdfNames: ['loan_agreement.pdf'],
        description: 'Single PDF, Contract Tool 7'
      },
      {
        sender: 'noreply@contracttool7.example.com',
        date: new Date('2024-06-22T10:30:00'),
        subject: "You've been copied on Loan Agreement - signed by John Doe",
        pdfNames: ['loan_agreement.pdf'],
        description: 'Exact duplicate of case 1'
      },
      {
        sender: 'noreply@contracttool1.example.com',
        date: new Date('2024-06-22T10:30:00'),
        subject: "Document completed: Employment Contract",
        pdfNames: ['employment_contract.pdf', 'nda.pdf'],
        description: 'Multiple PDFs, Contract Tool 1'
      },
      {
        sender: 'noreply@contracttool1.example.com',
        date: new Date('2024-06-22T10:30:00'),
        subject: "Document completed: Employment Contract",
        pdfNames: ['nda.pdf', 'employment_contract.pdf'], // Different order
        description: 'Same PDFs in different order (should generate same key)'
      },
      {
        sender: 'noreply@contracttool7.example.com',
        date: new Date('2024-06-22T10:31:00'), // Different time
        subject: "You've been copied on Loan Agreement - signed by John Doe",
        pdfNames: ['loan_agreement.pdf'],
        description: 'Same content, different time (1 minute later)'
      }
    ];
    
    console.log('Testing duplicate key generation:');
    const generatedKeys = [];
    
    testCases.forEach((testCase, index) => {
      console.log(`\nTest ${index + 1}: ${testCase.description}`);
      console.log(`  Sender: ${testCase.sender}`);
      console.log(`  Date: ${testCase.date.toISOString()}`);
      console.log(`  Subject: ${testCase.subject}`);
      console.log(`  PDFs: [${testCase.pdfNames.join(', ')}]`);
      
      const key = generateContentDuplicateKey(testCase.sender, testCase.date, testCase.subject, testCase.pdfNames);
      generatedKeys.push({ key, index, description: testCase.description });
      console.log(`  Generated Key: ${key}`);
    });
    
    // Check for expected duplicates
    console.log('\n--- Duplicate Analysis ---');
    console.log('Expected duplicates:');
    console.log('  Test 1 & 2: Should be identical (exact same content)');
    console.log('  Test 3 & 4: Should be identical (same PDFs, different order)');
    console.log('  Test 1 & 5: Should be different (different timestamps)');
    
    console.log('\nActual results:');
    const key1 = generatedKeys[0].key;
    const key2 = generatedKeys[1].key;
    const key3 = generatedKeys[2].key;
    const key4 = generatedKeys[3].key;
    const key5 = generatedKeys[4].key;
    
    console.log(`  Test 1 & 2 match: ${key1 === key2 ? '✅ YES' : '❌ NO'}`);
    console.log(`  Test 3 & 4 match: ${key3 === key4 ? '✅ YES' : '❌ NO'}`);
    console.log(`  Test 1 & 5 different: ${key1 !== key5 ? '✅ YES' : '❌ NO'}`);
    
    console.log('\n✅ Content duplicate key generation test completed');
    
  } catch (error) {
    console.error('❌ Content duplicate key generation test failed:', error);
    throw error;
  }
}

/**
 * Test content duplicate detection with spreadsheet
 * スプレッドシートでのコンテンツ重複検知をテスト
 */
function testContentDuplicateDetection() {
  console.log('\n=== TESTING Content Duplicate Detection ===');
  
  try {
    // Test data
    const testEmail = {
      sender: 'test@contracttool.example.com',
      date: new Date(),
      subject: 'Test Contract Agreement - signed by Test User',
      pdfNames: ['test_contract.pdf']
    };
    
    const contentKey = generateContentDuplicateKey(testEmail.sender, testEmail.date, testEmail.subject, testEmail.pdfNames);
    console.log(`Test content key: ${contentKey}`);
    
    // First check - should not be duplicate
    console.log('\n--- First Check (New Content) ---');
    const firstCheck = checkContentDuplicate(contentKey);
    console.log(`Is duplicate: ${firstCheck.isDuplicate ? 'YES' : 'NO'}`);
    console.log(`Existing URL: ${firstCheck.existingUrl || 'None'}`);
    
    if (firstCheck.isDuplicate) {
      console.log('⚠️  Content already exists - test may be using existing data');
    } else {
      console.log('✅ New content detected correctly');
    }
    
    // Record the content as saved
    const testDriveUrl = 'https://drive.google.com/drive/folders/test_folder_id';
    console.log('\n--- Recording Content ---');
    recordContentDuplicate(contentKey, testEmail.sender, testEmail.date, testEmail.subject, testEmail.pdfNames, testDriveUrl);
    console.log('Content recorded in spreadsheet');
    
    // Second check - should now be duplicate
    console.log('\n--- Second Check (Duplicate Content) ---');
    const secondCheck = checkContentDuplicate(contentKey);
    console.log(`Is duplicate: ${secondCheck.isDuplicate ? 'YES' : 'NO'}`);
    console.log(`Existing URL: ${secondCheck.existingUrl || 'None'}`);
    
    if (secondCheck.isDuplicate && secondCheck.existingUrl === testDriveUrl) {
      console.log('✅ Duplicate detection working correctly');
    } else {
      console.log('❌ Duplicate detection not working as expected');
    }
    
    console.log('\n✅ Content duplicate detection test completed');
    
  } catch (error) {
    console.error('❌ Content duplicate detection test failed:', error);
    throw error;
  }
}

/**
 * Test attachment processing with duplicate detection
 * 重複検知付き添付ファイル処理をテスト
 */
function testAttachmentProcessingWithDuplicates() {
  console.log('\n=== TESTING Attachment Processing with Duplicates ===');
  
  try {
    // Create mock attachment objects
    const mockPdfAttachment = {
      getName: () => 'test_contract.pdf',
      getSize: () => 1024000,
      getContentType: () => 'application/pdf',
      getBlob: () => ({ 
        getName: () => 'test_contract.pdf',
        getContentType: () => 'application/pdf'
      })
    };
    
    const mockDocAttachment = {
      getName: () => 'supporting_doc.docx',
      getSize: () => 512000,
      getContentType: () => 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      getBlob: () => ({ 
        getName: () => 'supporting_doc.docx',
        getContentType: () => 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
      })
    };
    
    const testAttachments = [mockPdfAttachment, mockDocAttachment];
    const testSubject = 'Test Contract Processing';
    const testDate = new Date();
    const testSender = 'test@contracttool.com';
    
    console.log('Testing attachment processing:');
    console.log(`  Subject: ${testSubject}`);
    console.log(`  Sender: ${testSender}`);
    console.log(`  Attachments: ${testAttachments.length}`);
    console.log(`    - ${mockPdfAttachment.getName()} (PDF)`);
    console.log(`    - ${mockDocAttachment.getName()} (DOCX)`);
    
    // Note: This test is limited because we can't actually save to Drive in test mode
    // We can test the logic flow but not the actual file operations
    console.log('\n--- Testing Processing Logic ---');
    
    // Test PDF name extraction
    const pdfNames = testAttachments
      .filter(attachment => attachment.getName().toLowerCase().endsWith('.pdf'))
      .map(attachment => attachment.getName());
    
    console.log(`Extracted PDF names: [${pdfNames.join(', ')}]`);
    
    // Test content key generation
    if (pdfNames.length > 0) {
      const contentKey = generateContentDuplicateKey(testSender, testDate, testSubject, pdfNames);
      console.log(`Generated content key: ${contentKey}`);
      
      // Test duplicate check
      const duplicateCheck = checkContentDuplicate(contentKey);
      console.log(`Duplicate check result: ${duplicateCheck.isDuplicate ? 'DUPLICATE' : 'NEW'}`);
      
      if (duplicateCheck.isDuplicate) {
        console.log(`  Existing URL: ${duplicateCheck.existingUrl}`);
        console.log('  ➜ Would skip PDF saving, use existing URL');
      } else {
        console.log('  ➜ Would save PDF to Drive, record new content');
      }
    }
    
    console.log('\n📝 Note: Full attachment processing test requires Drive access');
    console.log('   To test complete flow, run processEmails() with actual emails');
    
    console.log('\n✅ Attachment processing logic test completed');
    
  } catch (error) {
    console.error('❌ Attachment processing test failed:', error);
    throw error;
  }
}

/**
 * Test spreadsheet integration for content duplicates
 * コンテンツ重複のスプレッドシート統合をテスト
 */
function testSpreadsheetIntegration() {
  console.log('\n=== TESTING Spreadsheet Integration ===');
  
  try {
    console.log('Testing spreadsheet sheet creation...');
    
    // Test spreadsheet access
    const spreadsheetId = getProperty('SPREADSHEET_ID', false);
    if (!spreadsheetId) {
      console.log('⚠️  No spreadsheet configured - creating one for test');
      createOrGetSpreadsheet();
    }
    
    const spreadsheet = SpreadsheetApp.openById(getProperty('SPREADSHEET_ID'));
    console.log(`Using spreadsheet: ${spreadsheet.getName()}`);
    
    // Test content duplicate sheet creation
    const duplicateSheet = createOrGetContentDuplicateSheet(spreadsheet);
    console.log(`Content duplicate sheet: ${duplicateSheet.getName()}`);
    console.log(`Sheet has ${duplicateSheet.getLastRow()} rows (including header)`);
    
    // Test headers
    const headers = duplicateSheet.getRange(1, 1, 1, 8).getValues()[0];
    console.log('Sheet headers:');
    headers.forEach((header, index) => {
      console.log(`  ${String.fromCharCode(65 + index)}: ${header}`);
    });
    
    console.log('✅ Spreadsheet integration test completed');
    
  } catch (error) {
    console.error('❌ Spreadsheet integration test failed:', error);
    throw error;
  }
}

/**
 * Run comprehensive content duplicate detection tests
 * 包括的なコンテンツ重複検知テストを実行
 */
function runContentDuplicateTests() {
  console.log('=== RUNNING ALL CONTENT DUPLICATE DETECTION TESTS ===\n');
  
  try {
    // Test 1: Key generation
    testContentDuplicateKeyGeneration();
    console.log('\n' + '='.repeat(50) + '\n');
    
    // Test 2: Duplicate detection
    testContentDuplicateDetection();
    console.log('\n' + '='.repeat(50) + '\n');
    
    // Test 3: Attachment processing
    testAttachmentProcessingWithDuplicates();
    console.log('\n' + '='.repeat(50) + '\n');
    
    // Test 4: Spreadsheet integration
    testSpreadsheetIntegration();
    
    console.log('\n✅ ALL CONTENT DUPLICATE DETECTION TESTS COMPLETED SUCCESSFULLY');
    console.log('\nSystem Summary:');
    console.log('📧 Each recipient receives email processing and Slack notification');
    console.log('📄 PDF duplicates are detected by: Sender + Time + Subject + PDF names');
    console.log('💾 Duplicate PDFs skip Drive saving, use existing Drive URL');
    console.log('📊 All processing is logged to spreadsheet regardless of duplicates');
    console.log('🔄 Content duplicate tracking sheet maintains save history');
    
    console.log('\nNext steps:');
    console.log('1. Run processEmails() to test with actual contract emails');
    console.log('2. Verify duplicate detection in spreadsheet tabs');
    console.log('3. Check that Slack notifications still work for all recipients');
    
  } catch (error) {
    console.error('\n❌ CONTENT DUPLICATE DETECTION TESTS FAILED:', error);
    throw error;
  }
}

/**
 * Quick test for duplicate detection system
 * 重複検知システムのクイックテスト
 */
function quickDuplicateTest() {
  console.log('=== QUICK DUPLICATE DETECTION TEST ===');
  
  try {
    const testData = {
      sender: 'quicktest@example.com',
      date: new Date(),
      subject: 'Quick Test Contract',
      pdfNames: ['quick_test.pdf']
    };
    
    console.log('Testing with:', testData);
    
    const key = generateContentDuplicateKey(testData.sender, testData.date, testData.subject, testData.pdfNames);
    console.log(`Generated key: ${key}`);
    
    const check = checkContentDuplicate(key);
    console.log(`Is duplicate: ${check.isDuplicate}`);
    
    if (!check.isDuplicate) {
      console.log('Recording as new content...');
      recordContentDuplicate(key, testData.sender, testData.date, testData.subject, testData.pdfNames, 'https://test.url');
      
      const recheck = checkContentDuplicate(key);
      console.log(`After recording - Is duplicate: ${recheck.isDuplicate}`);
      console.log(`Existing URL: ${recheck.existingUrl}`);
    }
    
    console.log('✅ Quick test completed');
    
  } catch (error) {
    console.error('❌ Quick test failed:', error);
    throw error;
  }
}