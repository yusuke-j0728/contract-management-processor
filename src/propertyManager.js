/**
 * Property Management Module
 * 
 * Handles Script Properties management when the 50 property limit is reached.
 * Provides functions to view, manage, and clean up properties programmatically.
 */

/**
 * Get all Script Properties with categorization
 * 全てのScript Propertiesをカテゴリ分けして取得
 * 
 * @returns {Object} Categorized properties object
 */
function getAllProperties() {
  try {
    const allProperties = PropertiesService.getScriptProperties().getProperties();
    const categorized = {
      configuration: {},
      processedMessages: {},
      senderEmails: {},
      recipientEmails: {},
      other: {},
      total: Object.keys(allProperties).length
    };
    
    for (const [key, value] of Object.entries(allProperties)) {
      if (key.startsWith('PROCESSED_MSG_')) {
        categorized.processedMessages[key] = {
          value: value,
          timestamp: parseInt(value),
          date: new Date(parseInt(value)).toLocaleString('ja-JP')
        };
      } else if (key.startsWith('SENDER_EMAIL')) {
        categorized.senderEmails[key] = value;
      } else if (key.startsWith('RECIPIENT_EMAIL')) {
        categorized.recipientEmails[key] = value;
      } else if (['SLACK_WEBHOOK_URL', 'SLACK_CHANNEL', 'DRIVE_FOLDER_ID', 'SPREADSHEET_ID'].includes(key)) {
        categorized.configuration[key] = key === 'SLACK_WEBHOOK_URL' ? '[HIDDEN]' : value;
      } else {
        categorized.other[key] = value;
      }
    }
    
    return categorized;
    
  } catch (error) {
    console.error('Error getting all properties:', error);
    throw error;
  }
}

/**
 * Display all properties in a readable format
 * 全てのプロパティを読みやすい形式で表示
 */
function showAllProperties() {
  try {
    console.log('=== SCRIPT PROPERTIES SUMMARY ===\n');
    
    const props = getAllProperties();
    
    console.log(`📊 Total Properties: ${props.total}/50 (${props.total >= 50 ? '⚠️ LIMIT REACHED' : '✅ OK'})\n`);
    
    // Configuration Properties
    console.log('🔧 Configuration Properties:');
    if (Object.keys(props.configuration).length === 0) {
      console.log('  (None configured)');
    } else {
      for (const [key, value] of Object.entries(props.configuration)) {
        console.log(`  ${key}: ${value}`);
      }
    }
    console.log('');
    
    // Sender Emails
    console.log('📧 Sender Emails:');
    if (Object.keys(props.senderEmails).length === 0) {
      console.log('  (None configured)');
    } else {
      for (const [key, value] of Object.entries(props.senderEmails)) {
        console.log(`  ${key}: ${value}`);
      }
    }
    console.log('');
    
    // Recipient Emails
    console.log('📬 Recipient Emails:');
    if (Object.keys(props.recipientEmails).length === 0) {
      console.log('  (None configured)');
    } else {
      for (const [key, value] of Object.entries(props.recipientEmails)) {
        console.log(`  ${key}: ${value}`);
      }
    }
    console.log('');
    
    // Processed Messages
    console.log(`📝 Processed Messages: ${Object.keys(props.processedMessages).length}`);
    if (Object.keys(props.processedMessages).length > 0) {
      console.log('  (Most recent 5 shown):');
      const recent = Object.entries(props.processedMessages)
        .sort(([,a], [,b]) => b.timestamp - a.timestamp)
        .slice(0, 5);
      
      recent.forEach(([key, data]) => {
        const messageId = key.replace('PROCESSED_MSG_', '').split('_')[0];
        console.log(`  ${messageId}: ${data.date}`);
      });
      
      if (Object.keys(props.processedMessages).length > 5) {
        console.log(`  ... and ${Object.keys(props.processedMessages).length - 5} more`);
      }
    }
    console.log('');
    
    // Other Properties
    if (Object.keys(props.other).length > 0) {
      console.log('🔍 Other Properties:');
      for (const [key, value] of Object.entries(props.other)) {
        console.log(`  ${key}: ${value}`);
      }
      console.log('');
    }
    
    // Recommendations
    if (props.total >= 45) {
      console.log('⚠️  RECOMMENDATIONS:');
      console.log('  • Run cleanupOldProcessedMessages() to remove old entries');
      console.log('  • Consider setting up automatic cleanup with triggers');
      if (props.total >= 50) {
        console.log('  • Property limit reached - use programmatic management only');
      }
    }
    
  } catch (error) {
    console.error('Error showing all properties:', error);
    throw error;
  }
}

/**
 * Clean up old processed message properties
 * 古い処理済みメッセージプロパティをクリーンアップ
 * 
 * @param {number} daysOld - Remove processed messages older than this many days
 * @param {boolean} dryRun - If true, only show what would be deleted without actually deleting
 * @returns {Object} Cleanup results
 */
function cleanupProcessedMessages(daysOld = 7, dryRun = false) {
  try {
    console.log(`=== ${dryRun ? 'DRY RUN: ' : ''}CLEANING UP PROCESSED MESSAGES ===`);
    console.log(`Target: Messages older than ${daysOld} days\n`);
    
    const allProperties = PropertiesService.getScriptProperties().getProperties();
    const cutoffTime = new Date().getTime() - (daysOld * 24 * 60 * 60 * 1000);
    
    const toDelete = [];
    const toKeep = [];
    
    for (const [key, value] of Object.entries(allProperties)) {
      if (key.startsWith('PROCESSED_MSG_')) {
        const timestamp = parseInt(value);
        const date = new Date(timestamp);
        
        if (timestamp < cutoffTime) {
          toDelete.push({
            key: key,
            timestamp: timestamp,
            date: date.toLocaleString('ja-JP'),
            messageId: key.replace('PROCESSED_MSG_', '').split('_')[0]
          });
        } else {
          toKeep.push({
            key: key,
            timestamp: timestamp,
            date: date.toLocaleString('ja-JP')
          });
        }
      }
    }
    
    console.log(`📊 Analysis Results:`);
    console.log(`  Total processed messages: ${toDelete.length + toKeep.length}`);
    console.log(`  To delete (older than ${daysOld} days): ${toDelete.length}`);
    console.log(`  To keep (newer than ${daysOld} days): ${toKeep.length}`);
    console.log('');
    
    if (toDelete.length > 0) {
      console.log('🗑️  Messages to delete:');
      toDelete.forEach((item, index) => {
        if (index < 10) { // Show first 10
          console.log(`  ${item.messageId}: ${item.date}`);
        }
      });
      if (toDelete.length > 10) {
        console.log(`  ... and ${toDelete.length - 10} more`);
      }
      console.log('');
      
      if (!dryRun) {
        console.log('🔄 Deleting old entries...');
        let deletedCount = 0;
        
        toDelete.forEach(item => {
          try {
            PropertiesService.getScriptProperties().deleteProperty(item.key);
            deletedCount++;
          } catch (error) {
            console.error(`Error deleting ${item.key}:`, error);
          }
        });
        
        console.log(`✅ Cleanup completed: ${deletedCount} entries deleted`);
      } else {
        console.log('📋 DRY RUN: No changes made. Run with dryRun=false to actually delete.');
      }
    } else {
      console.log('✅ No old entries found to clean up.');
    }
    
    // Show final count
    const finalProperties = PropertiesService.getScriptProperties().getProperties();
    const finalCount = Object.keys(finalProperties).length;
    console.log(`\n📊 Final property count: ${finalCount}/50`);
    
    return {
      totalBefore: Object.keys(allProperties).length,
      totalAfter: finalCount,
      deleted: dryRun ? 0 : toDelete.length,
      kept: toKeep.length,
      recommendations: finalCount >= 45 ? ['Consider more frequent cleanup', 'Set up automatic triggers'] : []
    };
    
  } catch (error) {
    console.error('Error cleaning up processed messages:', error);
    throw error;
  }
}

/**
 * Set property with conflict checking
 * 競合チェック付きでプロパティを設定
 * 
 * @param {string} key - Property key
 * @param {string} value - Property value
 * @param {boolean} force - Force set even if properties are near limit
 */
function setPropertySafely(key, value, force = false) {
  try {
    const allProperties = PropertiesService.getScriptProperties().getProperties();
    const currentCount = Object.keys(allProperties).length;
    const exists = allProperties.hasOwnProperty(key);
    
    console.log(`Setting property: ${key}`);
    console.log(`Current property count: ${currentCount}/50`);
    console.log(`Property exists: ${exists ? 'Yes' : 'No'}`);
    
    // Check if we're at the limit and this is a new property
    if (!exists && currentCount >= 50 && !force) {
      console.log('⚠️  Property limit reached!');
      console.log('Options:');
      console.log('1. Run cleanupProcessedMessages() to free up space');
      console.log('2. Use setPropertySafely(key, value, true) to force set');
      console.log('3. Delete old properties manually');
      throw new Error('Cannot add new property: 50 property limit reached');
    }
    
    // Set the property
    PropertiesService.getScriptProperties().setProperty(key, value);
    console.log(`✅ Property set successfully: ${key}`);
    
    // Show new count
    const newCount = Object.keys(PropertiesService.getScriptProperties().getProperties()).length;
    console.log(`New property count: ${newCount}/50`);
    
    if (newCount >= 48) {
      console.log('⚠️  Warning: Approaching property limit. Consider cleanup soon.');
    }
    
  } catch (error) {
    console.error('Error setting property safely:', error);
    throw error;
  }
}

/**
 * Delete multiple properties by pattern
 * パターンで複数のプロパティを削除
 * 
 * @param {string} pattern - Pattern to match (e.g., 'PROCESSED_MSG_')
 * @param {boolean} dryRun - If true, only show what would be deleted
 * @returns {number} Number of properties deleted
 */
function deletePropertiesByPattern(pattern, dryRun = false) {
  try {
    console.log(`=== ${dryRun ? 'DRY RUN: ' : ''}DELETING PROPERTIES BY PATTERN ===`);
    console.log(`Pattern: "${pattern}"\n`);
    
    const allProperties = PropertiesService.getScriptProperties().getProperties();
    const toDelete = [];
    
    for (const [key, value] of Object.entries(allProperties)) {
      if (key.includes(pattern)) {
        toDelete.push({ key, value });
      }
    }
    
    console.log(`Found ${toDelete.length} properties matching pattern:`);
    toDelete.forEach((item, index) => {
      if (index < 20) { // Show first 20
        const displayValue = item.value.length > 50 ? item.value.substring(0, 50) + '...' : item.value;
        console.log(`  ${item.key}: ${displayValue}`);
      }
    });
    
    if (toDelete.length > 20) {
      console.log(`  ... and ${toDelete.length - 20} more`);
    }
    
    if (toDelete.length === 0) {
      console.log('No properties found matching the pattern.');
      return 0;
    }
    
    if (!dryRun) {
      console.log('\n🔄 Deleting properties...');
      let deletedCount = 0;
      
      toDelete.forEach(item => {
        try {
          PropertiesService.getScriptProperties().deleteProperty(item.key);
          deletedCount++;
        } catch (error) {
          console.error(`Error deleting ${item.key}:`, error);
        }
      });
      
      console.log(`✅ Deletion completed: ${deletedCount} properties deleted`);
      
      const finalCount = Object.keys(PropertiesService.getScriptProperties().getProperties()).length;
      console.log(`📊 Final property count: ${finalCount}/50`);
      
      return deletedCount;
    } else {
      console.log('\n📋 DRY RUN: No changes made. Run with dryRun=false to actually delete.');
      return 0;
    }
    
  } catch (error) {
    console.error('Error deleting properties by pattern:', error);
    throw error;
  }
}

/**
 * Get property statistics and recommendations
 * プロパティ統計と推奨事項を取得
 * 
 * @returns {Object} Statistics and recommendations
 */
function getPropertyStats() {
  try {
    const props = getAllProperties();
    const stats = {
      total: props.total,
      configuration: Object.keys(props.configuration).length,
      senderEmails: Object.keys(props.senderEmails).length,
      recipientEmails: Object.keys(props.recipientEmails).length,
      processedMessages: Object.keys(props.processedMessages).length,
      other: Object.keys(props.other).length,
      percentUsed: Math.round((props.total / 50) * 100),
      nearLimit: props.total >= 45,
      atLimit: props.total >= 50
    };
    
    const recommendations = [];
    
    if (stats.atLimit) {
      recommendations.push('🚨 CRITICAL: Property limit reached - use programmatic management only');
      recommendations.push('🧹 Run cleanupProcessedMessages() immediately');
    } else if (stats.nearLimit) {
      recommendations.push('⚠️  WARNING: Approaching property limit');
      recommendations.push('🧹 Consider running cleanupProcessedMessages()');
    }
    
    if (stats.processedMessages > 20) {
      recommendations.push('📝 Large number of processed messages - consider more frequent cleanup');
    }
    
    if (stats.senderEmails === 0) {
      recommendations.push('📧 No sender emails configured - run setContractConfiguration()');
    }
    
    stats.recommendations = recommendations;
    
    return stats;
    
  } catch (error) {
    console.error('Error getting property stats:', error);
    throw error;
  }
}

/**
 * Emergency property management - automatically clean up when near limit
 * 緊急プロパティ管理 - 制限近くで自動クリーンアップ
 */
function emergencyCleanup() {
  try {
    console.log('=== EMERGENCY PROPERTY CLEANUP ===\n');
    
    const stats = getPropertyStats();
    console.log(`Current usage: ${stats.total}/50 (${stats.percentUsed}%)`);
    
    if (!stats.nearLimit) {
      console.log('✅ No emergency cleanup needed.');
      return { cleaned: 0, message: 'No cleanup needed' };
    }
    
    console.log('🚨 Emergency cleanup required!\n');
    
    // Clean up processed messages older than 3 days
    const result = cleanupProcessedMessages(3, false);
    
    const newStats = getPropertyStats();
    console.log(`\nAfter cleanup: ${newStats.total}/50 (${newStats.percentUsed}%)`);
    
    if (newStats.atLimit) {
      console.log('🚨 Still at limit after cleanup!');
      console.log('Consider:');
      console.log('  • More aggressive cleanup: cleanupProcessedMessages(1)');
      console.log('  • Remove unused sender/recipient emails');
      console.log('  • Review other properties for cleanup');
    } else {
      console.log('✅ Emergency cleanup successful!');
    }
    
    return {
      cleaned: result.deleted,
      beforeCount: stats.total,
      afterCount: newStats.total,
      message: newStats.atLimit ? 'Still at limit' : 'Cleanup successful'
    };
    
  } catch (error) {
    console.error('Error in emergency cleanup:', error);
    throw error;
  }
}

/**
 * Test property management functions
 * プロパティ管理関数のテスト
 */
function testPropertyManagement() {
  console.log('=== TESTING Property Management ===\n');
  
  try {
    // Test 1: Get statistics
    console.log('Test 1: Property Statistics');
    const stats = getPropertyStats();
    console.log(`✓ Total properties: ${stats.total}`);
    console.log(`✓ Processed messages: ${stats.processedMessages}`);
    console.log(`✓ Near limit: ${stats.nearLimit}`);
    console.log('');
    
    // Test 2: Show properties (condensed)
    console.log('Test 2: Property Display');
    showAllProperties();
    console.log('');
    
    // Test 3: Dry run cleanup
    console.log('Test 3: Cleanup Dry Run');
    cleanupProcessedMessages(30, true);
    console.log('');
    
    console.log('✅ Property management test completed successfully');
    
  } catch (error) {
    console.error('❌ Property management test failed:', error);
    throw error;
  }
}