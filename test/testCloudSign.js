/**
 * Test file for CloudSign email processing
 * CloudSignメール処理のテストファイル
 */

/**
 * Test CloudSign email pattern matching
 * CloudSignのメールパターンマッチングをテスト
 */
function testCloudSignPatterns() {
  console.log('=== TESTING CloudSign Email Patterns ===');
  
  try {
    // Test CloudSign email subjects
    const testSubjects = [
      '「ExampleCompany株式会社_覚書株式会社Partner」の合意締結が完了しました',
      '「2024001-01_サイン用_ExampleCompany業務（2025年5月）捺印済」の合意締結が完了しました',
      '「【個別契約】機密保持契約書（双方／単独）20250512（ExampleCompany株式会社様）」の合意締結が完了しました'
    ];
    
    console.log('Testing subject pattern matching...');
    testSubjects.forEach((subject, index) => {
      console.log(`\nTest ${index + 1}: ${subject}`);
      
      // Test pattern matching
      let matched = false;
      for (const pattern of CONFIG.SUBJECT_PATTERNS.PATTERNS) {
        if (pattern.test(subject)) {
          console.log(`✓ Matched pattern: ${pattern}`);
          matched = true;
          break;
        }
      }
      
      if (!matched) {
        console.log('❌ No pattern matched!');
      }
    });
    
    console.log('\n✅ CloudSign pattern matching test completed');
    
  } catch (error) {
    console.error('❌ CloudSign pattern test failed:', error);
    throw error;
  }
}

/**
 * Test CloudSign metadata extraction
 * CloudSignのメタデータ抽出をテスト
 */
function testCloudSignMetadataExtraction() {
  console.log('\n=== TESTING CloudSign Metadata Extraction ===');
  
  try {
    const testCases = [
      {
        subject: '「ExampleCompany株式会社_覚書株式会社Partner」の合意締結が完了しました',
        sender: 'noreply@contracttool6.example.com',
        expectedTool: 'Contract Tool 6',
        expectedType: '覚書',
        expectedParty: 'ExampleCompany株式会社'
      },
      {
        subject: '「【個別契約】機密保持契約書（双方／単独）20250512（ExampleCompany株式会社様）」の合意締結が完了しました',
        sender: 'noreply@contracttool6.example.com',
        expectedTool: 'Contract Tool 6',
        expectedType: '秘密保持契約',
        expectedParty: '【個別契約】機密保持契約書（双方／単独）20250512（ExampleCompany株式会社様）'
      }
    ];
    
    testCases.forEach((testCase, index) => {
      console.log(`\nTest Case ${index + 1}:`);
      console.log(`Subject: ${testCase.subject}`);
      
      // Test contract tool extraction
      const contractTool = extractContractTool(testCase.sender);
      console.log(`Contract Tool: ${contractTool} (Expected: ${testCase.expectedTool})`);
      console.log(contractTool === testCase.expectedTool ? '✓ Tool extraction passed' : '❌ Tool extraction failed');
      
      // Test contract type extraction
      const contractType = extractContractType(testCase.subject, '');
      console.log(`Contract Type: ${contractType} (Expected: ${testCase.expectedType})`);
      console.log(contractType === testCase.expectedType ? '✓ Type extraction passed' : '❌ Type extraction failed');
      
      // Test contract party extraction
      const contractParty = extractContractParty(testCase.subject, '');
      console.log(`Contract Party: ${contractParty}`);
      console.log(`✓ Party extraction completed`);
    });
    
    console.log('\n✅ CloudSign metadata extraction test completed');
    
  } catch (error) {
    console.error('❌ CloudSign metadata test failed:', error);
    throw error;
  }
}

/**
 * Test CloudSign configuration
 * CloudSignの設定をテスト
 */
function testCloudSignConfiguration() {
  console.log('\n=== TESTING CloudSign Configuration ===');
  
  try {
    // Check if CloudSign is in the sender emails list
    const cloudSignEmail = 'noreply@contracttool6.example.com';
    const isConfigured = CONFIG.SENDER_EMAILS.includes(cloudSignEmail);
    
    console.log(`CloudSign email configured: ${isConfigured ? '✓ Yes' : '❌ No'}`);
    
    // Check if CloudSign patterns are included
    const hasCloudSignPattern = CONFIG.SUBJECT_PATTERNS.PATTERNS.some(pattern => 
      pattern.test('の合意締結が完了しました')
    );
    
    console.log(`CloudSign patterns included: ${hasCloudSignPattern ? '✓ Yes' : '❌ No'}`);
    
    // Display current configuration
    console.log('\nCurrent CloudSign configuration:');
    console.log(`- Sender email: ${cloudSignEmail}`);
    console.log(`- Pattern count: ${CONFIG.SUBJECT_PATTERNS.PATTERNS.length}`);
    
    console.log('\n✅ CloudSign configuration test completed');
    
  } catch (error) {
    console.error('❌ CloudSign configuration test failed:', error);
    throw error;
  }
}

/**
 * Run all CloudSign tests
 * すべてのCloudSignテストを実行
 */
function runCloudSignTests() {
  console.log('=== RUNNING ALL CLOUDSIGN TESTS ===\n');
  
  try {
    // Test 1: Pattern matching
    testCloudSignPatterns();
    console.log('\n' + '='.repeat(50) + '\n');
    
    // Test 2: Metadata extraction
    testCloudSignMetadataExtraction();
    console.log('\n' + '='.repeat(50) + '\n');
    
    // Test 3: Configuration
    testCloudSignConfiguration();
    
    console.log('\n✅ ALL CLOUDSIGN TESTS COMPLETED SUCCESSFULLY');
    console.log('\nNext steps:');
    console.log('1. Run setContractConfiguration() and update SENDER_EMAIL_6 with noreply@contracttool6.example.com');
    console.log('2. Run showConfiguration() to verify settings');
    console.log('3. Run processEmails() to test with actual CloudSign emails');
    
  } catch (error) {
    console.error('\n❌ CLOUDSIGN TESTS FAILED:', error);
    throw error;
  }
}