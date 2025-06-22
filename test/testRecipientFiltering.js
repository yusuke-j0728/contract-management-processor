/**
 * Test file for recipient filtering and duplicate prevention
 * 受信者フィルタリングと重複防止のテストファイル
 */

/**
 * Test recipient email filtering functionality
 * 受信者メールフィルタリング機能をテスト
 */
function testRecipientFiltering() {
  console.log('=== TESTING Recipient Email Filtering ===');
  
  try {
    // Test configuration
    console.log('Testing recipient configuration...');
    
    // Check current recipient configuration
    const recipientEmails = getConfiguredRecipientEmails();
    console.log(`Current recipient emails: ${recipientEmails.length}`);
    
    if (recipientEmails.length === 0) {
      console.log('⚠️  No recipient filtering configured');
      console.log('📝 System will process emails to ALL recipients');
      console.log('💡 To test recipient filtering:');
      console.log('   setProperty("RECIPIENT_EMAIL_1", "your-email@company.com");');
      console.log('   setProperty("RECIPIENT_EMAIL_2", "contracts@company.com");');
    } else {
      console.log('✅ Recipient filtering is configured:');
      recipientEmails.forEach((email, index) => {
        console.log(`  ${index + 1}. ${email}`);
      });
    }
    
    console.log('\n--- Testing Email Address Extraction ---');
    
    // Test email extraction from various formats
    const testRecipientFields = [
      'user@example.com',
      'User Name <user@example.com>',
      'user1@example.com, user2@example.com',
      'User One <user1@example.com>, User Two <user2@example.com>',
      'User Name <user@example.com>, another@test.com, "Third User" <third@example.com>'
    ];
    
    testRecipientFields.forEach((field, index) => {
      console.log(`\nTest ${index + 1}: "${field}"`);
      const extracted = extractEmailAddresses(field);
      console.log(`  Extracted: [${extracted.join(', ')}]`);
    });
    
    console.log('\n--- Testing Recipient Hash Generation ---');
    
    // Test hash generation for different recipient combinations
    const testRecipientSets = [
      ['user1@example.com'],
      ['user1@example.com', 'user2@example.com'],
      ['user2@example.com', 'user1@example.com'], // Different order, should be same hash
      ['user1@example.com', 'user2@example.com', 'user3@example.com']
    ];
    
    testRecipientSets.forEach((recipients, index) => {
      const hash = createRecipientHash(recipients);
      console.log(`Recipients: [${recipients.join(', ')}] → Hash: ${hash}`);
    });
    
    // Verify that order doesn't matter for hash
    const hash1 = createRecipientHash(['user1@example.com', 'user2@example.com']);
    const hash2 = createRecipientHash(['user2@example.com', 'user1@example.com']);
    console.log(`\nHash consistency check: ${hash1 === hash2 ? '✅ PASS' : '❌ FAIL'}`);
    
    console.log('\n✅ Recipient filtering test completed successfully');
    
  } catch (error) {
    console.error('❌ Recipient filtering test failed:', error);
    throw error;
  }
}

/**
 * Test duplicate prevention with recipient tracking
 * 受信者追跡による重複防止をテスト
 */
function testDuplicatePrevention() {
  console.log('\n=== TESTING Duplicate Prevention with Recipients ===');
  
  try {
    // Create mock message objects for testing
    const mockMessages = [
      {
        id: 'msg_001',
        to: 'user1@company.com',
        cc: '',
        bcc: '',
        subject: 'Test Contract 1'
      },
      {
        id: 'msg_002', 
        to: 'user1@company.com, user2@company.com',
        cc: '',
        bcc: '',
        subject: 'Test Contract 2'
      },
      {
        id: 'msg_001', // Same message ID but different recipients
        to: 'user3@company.com',
        cc: '',
        bcc: '',
        subject: 'Test Contract 1'
      }
    ];
    
    mockMessages.forEach((mockMsg, index) => {
      console.log(`\nTesting message ${index + 1}:`);
      console.log(`  ID: ${mockMsg.id}`);
      console.log(`  To: ${mockMsg.to}`);
      
      // Create mock Gmail message object
      const mockMessage = {
        getId: () => mockMsg.id,
        getTo: () => mockMsg.to,
        getCc: () => mockMsg.cc,
        getBcc: () => mockMsg.bcc,
        getSubject: () => mockMsg.subject
      };
      
      // Test recipient extraction
      const recipients = getMessageRecipients(mockMessage);
      console.log(`  Extracted recipients: [${recipients.join(', ')}]`);
      
      // Test hash generation
      const hash = createRecipientHash(recipients);
      console.log(`  Recipient hash: ${hash}`);
      
      // Test processing key generation
      const processedKey = CONFIG.RECIPIENT_EMAILS.length > 0 
        ? `PROCESSED_MSG_${mockMsg.id}_${hash}`
        : `PROCESSED_MSG_${mockMsg.id}`;
      console.log(`  Processing key: ${processedKey}`);
    });
    
    console.log('\n--- Duplicate Detection Logic ---');
    console.log('With recipient filtering enabled:');
    console.log('  • Same message ID + different recipients = NOT duplicate (will process)');
    console.log('  • Same message ID + same recipients = duplicate (will skip)');
    console.log('Without recipient filtering:');
    console.log('  • Same message ID = duplicate (will skip regardless of recipients)');
    
    console.log('\n✅ Duplicate prevention test completed successfully');
    
  } catch (error) {
    console.error('❌ Duplicate prevention test failed:', error);
    throw error;
  }
}

/**
 * Test recipient filtering configuration
 * 受信者フィルタリング設定をテスト
 */
function testRecipientConfiguration() {
  console.log('\n=== TESTING Recipient Configuration ===');
  
  try {
    // Test the configuration function
    console.log('Current recipient configuration:');
    const recipients = getConfiguredRecipientEmails();
    
    if (recipients.length === 0) {
      console.log('  No recipients configured - processing ALL emails');
    } else {
      console.log(`  ${recipients.length} recipients configured:`);
      recipients.forEach((email, index) => {
        console.log(`    ${index + 1}. ${email}`);
      });
    }
    
    // Show how to configure recipient filtering
    console.log('\n📝 To configure recipient filtering:');
    console.log('// In Google Apps Script, execute:');
    console.log('setProperty("RECIPIENT_EMAIL_1", "your-primary@company.com");');
    console.log('setProperty("RECIPIENT_EMAIL_2", "contracts@company.com");');
    console.log('setProperty("RECIPIENT_EMAIL_3", "team@company.com");');
    console.log('\n// Then run showConfiguration() to verify');
    
    // Show search query impact
    console.log('\n📋 Search Query Impact:');
    const senderEmails = CONFIG.SENDER_EMAILS;
    const senderQueries = senderEmails.map(email => `from:${email}`);
    let query = `(${senderQueries.join(' OR ')})`;
    
    if (recipients.length > 0) {
      const recipientQueries = recipients.map(email => `to:${email}`);
      query += ` AND (${recipientQueries.join(' OR ')})`;
    }
    
    console.log(`Current search query: ${query}`);
    
    console.log('\n✅ Recipient configuration test completed successfully');
    
  } catch (error) {
    console.error('❌ Recipient configuration test failed:', error);
    throw error;
  }
}

/**
 * Run all recipient filtering tests
 * 全ての受信者フィルタリングテストを実行
 */
function runRecipientFilteringTests() {
  console.log('=== RUNNING ALL RECIPIENT FILTERING TESTS ===\n');
  
  try {
    // Test 1: Basic functionality
    testRecipientFiltering();
    console.log('\n' + '='.repeat(50) + '\n');
    
    // Test 2: Duplicate prevention
    testDuplicatePrevention();
    console.log('\n' + '='.repeat(50) + '\n');
    
    // Test 3: Configuration
    testRecipientConfiguration();
    
    console.log('\n✅ ALL RECIPIENT FILTERING TESTS COMPLETED SUCCESSFULLY');
    console.log('\nNext steps:');
    console.log('1. Configure recipient emails if needed: setProperty("RECIPIENT_EMAIL_1", "your-email@company.com")');
    console.log('2. Run showConfiguration() to verify settings');
    console.log('3. Run processEmails() to test with actual emails');
    
  } catch (error) {
    console.error('\n❌ RECIPIENT FILTERING TESTS FAILED:', error);
    throw error;
  }
}