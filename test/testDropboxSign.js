/**
 * Test file for Dropbox Sign integration
 * Dropbox Sign統合のテストファイル
 */

/**
 * Test Dropbox Sign email pattern matching
 * Dropbox Signメールパターンマッチングをテスト
 */
function testDropboxSignPatterns() {
  console.log('=== TESTING Dropbox Sign Pattern Matching ===');
  
  try {
    // Test subjects from contract tool emails
    const dropboxSignTestSubjects = [
      "You've been copied on Loan Agreement and Letter of Guarantee - CompanyA <> CompanyB - signed by John Smith and Ken Johnson",
      "You've been copied on ExampleCompany X Partner agreement - signed by Manager",
      "Document has been completed on Dropbox Sign",
      "Signature request completed on Dropbox",
      "Service Agreement signed by John Doe and Jane Smith",
      "NDA Agreement - signed by Alice Johnson",
      "Employment Contract has been completed - signed by Bob Wilson and HR Team",
      "Consulting Agreement - signed by Sarah Lee and Company Ltd",
      "普通のメール件名（マッチしないはず）"
    ];
    
    console.log('Testing Dropbox Sign subject pattern matching:');
    console.log(`Multiple patterns enabled: ${CONFIG.SUBJECT_PATTERNS?.ENABLE_MULTIPLE_PATTERNS}`);
    console.log(`Match mode: ${CONFIG.SUBJECT_PATTERNS?.MATCH_MODE}`);
    console.log(`Total patterns: ${CONFIG.SUBJECT_PATTERNS?.PATTERNS?.length || 0}`);
    console.log('');
    
    dropboxSignTestSubjects.forEach((subject, index) => {
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
    
    console.log('✅ Dropbox Sign pattern matching test completed successfully');
    
  } catch (error) {
    console.error('❌ Dropbox Sign pattern matching test failed:', error);
    throw error;
  }
}

/**
 * Test Dropbox Sign sender email recognition
 * Dropbox Sign送信者メール認識をテスト
 */
function testDropboxSignSenders() {
  console.log('\n=== TESTING Dropbox Sign Sender Recognition ===');
  
  try {
    const contractToolSenders = [
      'noreply@contracttool7.example.com',
      'noreply@contracttool8.example.com',
      'notifications@contracttool7.example.com', // May not be configured but test anyway
      'other@example.com' // Should not match
    ];
    
    console.log('Testing sender email recognition:');
    
    contractToolSenders.forEach((sender, index) => {
      console.log(`\nTest ${index + 1}: ${sender}`);
      
      // Check if this sender is in configured emails
      const configuredSenders = getConfiguredSenderEmails();
      const isConfigured = configuredSenders.includes(sender);
      console.log(`  Configured: ${isConfigured ? '✅ YES' : '❌ NO'}`);
      
      // Check tool name extraction
      const toolName = getContractToolName(sender);
      console.log(`  Tool name: ${toolName}`);
      
      // Test if this would be included in search query
      if (isConfigured) {
        console.log(`  Would be included in Gmail search: ✅ YES`);
      } else {
        console.log(`  Would be included in Gmail search: ❌ NO`);
      }
    });
    
    console.log('\n✅ Dropbox Sign sender recognition test completed successfully');
    
  } catch (error) {
    console.error('❌ Dropbox Sign sender recognition test failed:', error);
    throw error;
  }
}

/**
 * Test Dropbox Sign contract metadata extraction
 * Dropbox Sign契約メタデータ抽出をテスト
 */
function testDropboxSignMetadata() {
  console.log('\n=== TESTING Dropbox Sign Metadata Extraction ===');
  
  try {
    const testCases = [
      {
        subject: "You've been copied on Loan Agreement and Letter of Guarantee - CompanyA <> CompanyB - signed by John Smith and Ken Johnson",
        sender: "noreply@contracttool7.example.com",
        expectedTool: "Contract Tool 7",
        expectedType: "Unknown", // May need to add Loan Agreement detection
        expectedParty: "CompanyA, CompanyB" // May need custom extraction logic
      },
      {
        subject: "You've been copied on ExampleCompany X Partner agreement - signed by Manager",
        sender: "noreply@contracttool8.example.com", 
        expectedTool: "Contract Tool 8",
        expectedType: "Unknown",
        expectedParty: "ExampleCompany, Partner"
      },
      {
        subject: "Employment Contract has been completed - signed by Bob Wilson and HR Team",
        sender: "noreply@contracttool7.example.com",
        expectedTool: "Contract Tool 7", 
        expectedType: "Employment",
        expectedParty: "Bob Wilson, HR Team"
      }
    ];
    
    console.log('Testing metadata extraction:');
    
    testCases.forEach((testCase, index) => {
      console.log(`\nTest ${index + 1}: "${testCase.subject}"`);
      
      // Test contract tool extraction
      const extractedTool = extractContractTool(testCase.sender);
      console.log(`  Contract Tool: ${extractedTool} (expected: ${testCase.expectedTool})`);
      console.log(`  Tool Match: ${extractedTool === testCase.expectedTool ? '✅' : '❌'}`);
      
      // Test contract type extraction
      const extractedType = extractContractType(testCase.subject);
      console.log(`  Contract Type: ${extractedType} (expected: ${testCase.expectedType})`);
      console.log(`  Type Match: ${extractedType === testCase.expectedType ? '✅' : '❌'}`);
      
      // Test contract party extraction (if function exists)
      try {
        const extractedParty = extractContractParty(testCase.subject);
        console.log(`  Contract Party: ${extractedParty} (expected: ${testCase.expectedParty})`);
      } catch (error) {
        console.log(`  Contract Party: Function not available or error`);
      }
    });
    
    console.log('\n✅ Dropbox Sign metadata extraction test completed successfully');
    
  } catch (error) {
    console.error('❌ Dropbox Sign metadata extraction test failed:', error);
    throw error;
  }
}

/**
 * Run all Dropbox Sign tests
 * 全てのDropbox Signテストを実行
 */
function runDropboxSignTests() {
  console.log('=== RUNNING ALL DROPBOX SIGN TESTS ===\n');
  
  try {
    // Test 1: Pattern matching
    testDropboxSignPatterns();
    console.log('\n' + '='.repeat(50) + '\n');
    
    // Test 2: Sender recognition
    testDropboxSignSenders();
    console.log('\n' + '='.repeat(50) + '\n');
    
    // Test 3: Metadata extraction
    testDropboxSignMetadata();
    
    console.log('\n✅ ALL DROPBOX SIGN TESTS COMPLETED SUCCESSFULLY');
    console.log('\nNext steps:');
    console.log('1. Configure Contract Tool 7 sender emails: setProperty("SENDER_EMAIL_7", "noreply@contracttool7.example.com")');
    console.log('2. Run showConfiguration() to verify settings');
    console.log('3. Run processEmails() to test with actual emails');
    console.log('4. Consider adding custom party extraction logic for Contract Tool 7 format');
    
  } catch (error) {
    console.error('\n❌ DROPBOX SIGN TESTS FAILED:', error);
    throw error;
  }
}

/**
 * Quick setup function for Contract Tool 7 & 8
 * Contract Tool 7 & 8用のクイックセットアップ関数
 */
function setupContractTools() {
  console.log('=== SETTING UP CONTRACT TOOL CONFIGURATION ===');
  
  try {
    // Add Contract Tool sender emails
    setProperty('SENDER_EMAIL_7', 'noreply@contracttool7.example.com');
    setProperty('SENDER_EMAIL_8', 'noreply@contracttool8.example.com');
    
    console.log('✅ Contract Tool sender emails configured');
    console.log('Configured emails:');
    console.log('  - noreply@contracttool7.example.com');
    console.log('  - noreply@contracttool8.example.com');
    
    console.log('\n💡 To verify configuration:');
    console.log('  - Run showConfiguration()');
    console.log('  - Run runContractToolTests()');
    
  } catch (error) {
    console.error('❌ Contract Tool setup failed:', error);
    throw error;
  }
}