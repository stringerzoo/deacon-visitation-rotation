/**
 * MODULE 5: GOOGLE CHAT NOTIFICATIONS (v1.1)
 * Deacon Visitation Rotation System - Chat Integration
 * 
 * This module contains:
 * - Google Chat webhook integration
 * - Weekly visitation summaries
 * - Day-before reminders
 * - Configuration management for chat spaces
 * - Test mode awareness and separate test chat integration
 * 
 * Integration with existing modules:
 * - Uses Module 1's getConfiguration() for schedule data
 * - Uses Module 2's getScheduleFromSheet() for visit information  
 * - Uses Module 3's getCurrentTestMode() for environment detection
 * - Extends Module 4's menu system with notification options
 */

// ===== CORE NOTIFICATION FUNCTIONS =====

function sendWeeklyVisitationChat() {
  /**
   * Main weekly summary function - sends 2-week lookahead using calendar week logic
   */
  try {
    // Check if notifications are configured
    if (!isNotificationConfigured()) {
      SpreadsheetApp.getUi().alert(
        'Notifications Not Configured',
        'Please configure the Google Chat webhook first:\n\n' +
        '📢 Notifications → 🔧 Configure Chat Webhook',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    // Validate that schedule matches current data
    if (!validateScheduleDataSync()) {
      return;
    }
    
    // Get current test mode from Module 3
    const currentTestMode = getCurrentTestMode();
    const chatPrefix = currentTestMode ? '🧪 TEST: ' : '';
    
    // NEW: Get visits for remainder of current week + next complete week
    const upcomingVisits = getVisitsForCalendarWeeks(2, true);
    
    if (upcomingVisits.length === 0) {
      const message = `${chatPrefix}📅 **Weekly Visitation Update**\n\nNo visits scheduled for the next 2 weeks. All caught up! 🎉`;
      
      sendToChatSpace(message, currentTestMode);
      console.log('No upcoming visits found for weekly summary');
      return;
    }
    
    // Build and send the weekly message with calendar week formatting
    const chatMessage = buildWeeklyCalendarSummary(upcomingVisits, currentTestMode);
    sendToChatSpace(chatMessage, currentTestMode);
    
    // Show success notification
    const ui = SpreadsheetApp.getUi();
    const weekRange = getWeekRangeDescription(upcomingVisits);
    
    ui.alert(
      currentTestMode ? '🧪 Test Weekly Summary Sent' : '📅 Weekly Summary Sent',
      `${chatPrefix}Sent 2-week summary for ${upcomingVisits.length} upcoming visits.\n\n` +
      `Chat space: ${currentTestMode ? 'TEST space' : 'Production space'}\n` +
      `Coverage: ${weekRange}`,
      ui.ButtonSet.OK
    );
    
    console.log(`Weekly calendar summary sent: ${upcomingVisits.length} visits across ${weekRange}`);
    
  } catch (error) {
    console.error('Failed to send weekly calendar summary:', error);
    SpreadsheetApp.getUi().alert(
      'Chat Notification Failed',
      `❌ Error sending to Google Chat: ${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

// ===== MESSAGE BUILDING FUNCTIONS =====
  
function buildWeeklyCalendarSummary(visits, isTestMode = false) {
  /**
   * v2.0 CORRECTED: Use actual visit dates instead of recalculating from next Sunday
   */
  const chatPrefix = isTestMode ? '🧪 TEST: ' : '';
  const today = new Date();
  const todayFormatted = today.toLocaleDateString('en-US', { 
    weekday: 'long', 
    month: 'long', 
    day: 'numeric' 
  });
  // Don't recalculate - use the visits we already have
// Split visits into first half and second half of the 2-week period
if (visits.length === 0) {
  return `${chatPrefix}📅 *Weekly Visitation Update*\n\nNo visits scheduled for the next 2 weeks.`;
}

// Calculate 2-week coverage starting from tomorrow
const tomorrow = new Date(today);
tomorrow.setDate(today.getDate() + 1);
tomorrow.setHours(0, 0, 0, 0);

const coverageEnd = new Date(tomorrow);
coverageEnd.setDate(tomorrow.getDate() + 13);  // 14 days total

// Week 1: Tomorrow through 6 days later
const week1Start = new Date(tomorrow);
const week1End = new Date(tomorrow);
week1End.setDate(tomorrow.getDate() + 6);

// Week 2: Day 8 through Day 14
const week2Start = new Date(tomorrow);
week2Start.setDate(tomorrow.getDate() + 7);
const week2End = new Date(coverageEnd);

// Split visits by week
const week1Visits = visits.filter(visit => {
  const visitDate = new Date(visit.date);
  visitDate.setHours(0, 0, 0, 0);
  return visitDate >= week1Start && visitDate <= week1End;
});

const week2Visits = visits.filter(visit => {
  const visitDate = new Date(visit.date);
  visitDate.setHours(0, 0, 0, 0);
  return visitDate >= week2Start && visitDate <= week2End;
});

// Build the message with Google Chat formatting
let message = `${chatPrefix}📅 *Weekly Visitation Update*\n`;
message += `Generated: ${todayFormatted}\n`;
message += `Coverage: ${tomorrow.toLocaleDateString('en-US', { month: 'short', day: 'numeric' })} - ${coverageEnd.toLocaleDateString('en-US', { month: 'short', day: 'numeric' })}\n\n`;
    
  // Week 1 Section
  message += `*Week of ${week1Start.toLocaleDateString('en-US', { 
    month: 'short', 
    day: 'numeric' 
  })} - ${week1End.toLocaleDateString('en-US', { 
    month: 'short', 
    day: 'numeric' 
  })}*\n\n`;
  
  if (week1Visits.length === 0) {
    message += `No visits scheduled for this week.\n\n`;
  } else {
    week1Visits.forEach(visit => {
      const visitDate = visit.date.toLocaleDateString('en-US', { 
        weekday: 'short', 
        month: 'short', 
        day: 'numeric' 
      });
      
      message += `🗓️ *${visit.deacon}* visits *${visit.household}*\n`;
      
      // Add contact info
      const contactInfo = [];
      if (visit.phone) contactInfo.push(`📞 ${visit.phone}`);
      if (visit.address) contactInfo.push(`🏠 ${visit.address}`);
      
      if (contactInfo.length > 0) {
        message += `   ${contactInfo.join(' • ')}\n`;
      }
      
      // Add links on separate lines with Google Chat format
      if (visit.breezeLink) {
        message += `   👤 <${visit.breezeLink}|Breeze Profile>\n`;
      }
      if (visit.notesLink) {
        message += `   📝 <${visit.notesLink}|Visit Notes>\n`;
      }
      
      message += '\n';
    });
  }
  
  message += '––––––\n\n';
  
  // Week 2 Section
  message += `*Week of ${week2Start.toLocaleDateString('en-US', { 
    month: 'short', 
    day: 'numeric' 
  })} - ${week2End.toLocaleDateString('en-US', { 
    month: 'short', 
    day: 'numeric' 
  })}*\n\n`;
  
  if (week2Visits.length === 0) {
    message += `No visits scheduled for this week.\n\n`;
  } else {
    week2Visits.forEach(visit => {
      const visitDate = visit.date.toLocaleDateString('en-US', { 
        weekday: 'short', 
        month: 'short', 
        day: 'numeric' 
      });
      
      message += `🗓️ *${visit.deacon}* visits *${visit.household}*\n`;
      
      // Add contact info
      const contactInfo = [];
      if (visit.phone) contactInfo.push(`📞 ${visit.phone}`);
      if (visit.address) contactInfo.push(`🏠 ${visit.address}`);
      
      if (contactInfo.length > 0) {
        message += `   ${contactInfo.join(' • ')}\n`;
      }
      
      // Add links on separate lines with Google Chat format
      if (visit.breezeLink) {
        message += `   👤 <${visit.breezeLink}|Breeze Profile>\n`;
      }
      if (visit.notesLink) {
        message += `   📝 <${visit.notesLink}|Visit Notes>\n`;
      }
      
      message += '\n';
    });
  }
  
  message += '––––––\n\n';
  
  // Footer with instructions and configurable calendar link
  message += `💡 *Instructions*: Call ahead to set up visit times then update the event in the Visitation Calendar.\n\n`;
  
  const calendarLink = getCalendarLinkFromSpreadsheet();
  if (calendarLink) {
    message += `📅 <${calendarLink}|Visitation Calendar>\n\n`;
  }
  
  const summaryLink = getSummaryLinkFromSpreadsheet();
  if (summaryLink) {
    message += `👀 <${summaryLink}|Schedule Summary> (quick look at all upcoming visits on the schedule, sorted by date and grouped by deacon)\n\n`;
  }
  
  const guideLink = getGuideLinkFromSpreadsheet();
  if (guideLink) {
    message += `📖 <${guideLink}|Visitation Guide>\n\n`;
  }
  
  message += `🔄 This update is sent weekly. It will not reflect any changes made to the visitation calendar. It only shows visits scheduled for the "week of". Reply in this thread with questions or scheduling conflicts.`;
  
  return message;
}

// ===== CHAT WEBHOOK FUNCTIONS =====

function sendToChatSpace(message, isTestMode = false) {
  /**
   * Send message to appropriate Google Chat space via webhook
   */
  try {
    const webhookUrl = getWebhookUrl(isTestMode);
    
    if (!webhookUrl) {
      throw new Error(`No webhook configured for ${isTestMode ? 'test' : 'production'} mode`);
    }
    
    const payload = {
      text: message
    };
    
    const response = UrlFetchApp.fetch(webhookUrl, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() !== 200) {
      throw new Error(`Chat webhook failed: ${response.getResponseCode()} - ${response.getContentText()}`);
    }
    
    console.log(`Message sent to ${isTestMode ? 'test' : 'production'} chat space`);
    
  } catch (error) {
    console.error('Error sending to chat space:', error);
    throw error;
  }
}

function getWebhookUrl(isTestMode = false) {
  /**
   * Get appropriate webhook URL based on current mode
   */
  const properties = PropertiesService.getScriptProperties();
  const key = isTestMode ? 'CHAT_WEBHOOK_TEST' : 'CHAT_WEBHOOK_PROD';
  return properties.getProperty(key);
}

// ===== VISIT DATA FUNCTIONS =====

function groupVisitsByDeacon(visits) {
  /**
   * Group visits by deacon name for organized display
   */
  const grouped = {};
  
  visits.forEach(visit => {
    if (!grouped[visit.deacon]) {
      grouped[visit.deacon] = [];
    }
    grouped[visit.deacon].push(visit);
  });
  
  return grouped;
}

function getBreezeLink(config, householdIndex) {
  /**
   * Get Breeze link for household
   */
  if (householdIndex < 0) return '';
  
  // Use Breeze number to build URL
  const breezeNumber = config.breezeLinks[householdIndex];  // Changed from breezeNumbers
  if (breezeNumber && breezeNumber.toString().trim().length > 0) {
    return buildBreezeUrl(breezeNumber);
  }
  
  return '';
}

function getNotesLink(config, householdIndex) {
  /**
   * Get Notes link for household
   */
  if (householdIndex < 0) return '';
  
  // Return full notes link
  const fullLink = config.notesLinks[householdIndex];
  if (fullLink && fullLink.toString().trim().length > 0) {
    return fullLink.toString().trim();
  }
  
  return '';
}

// ===== CONFIGURATION FUNCTIONS =====

function getCalendarLinkFromSpreadsheet() {
  /**
   * v1.1: Auto-generate calendar URL instead of reading from K19
   * Eliminates K19 dependency completely
   */
  try {
    return generateCalendarUrlDirect();
  } catch (error) {
    console.error('Error generating calendar URL:', error);
    return '';
  }
}

function generateCalendarUrlDirect() {
  /**
   * v1.1: Generates calendar URL directly from calendar system
   * Auto-detects current mode and user timezone
   */
  try {
    const currentTestMode = getCurrentTestMode();
    const calendarName = currentTestMode ? 'TEST - Deacon Visitation Schedule' : 'Deacon Visitation Schedule';
    
    const calendars = CalendarApp.getCalendarsByName(calendarName);
    if (calendars.length > 0) {
      const calendarId = calendars[0].getId();
      
      // Auto-detect user's timezone
      const userTimezone = Session.getScriptTimeZone();
      const encodedId = encodeURIComponent(calendarId);
      const encodedTimezone = encodeURIComponent(userTimezone);
      
      const calendarUrl = `https://calendar.google.com/calendar/embed?src=${encodedId}&ctz=${encodedTimezone}`;
      
      console.log(`Auto-generated calendar URL for notifications (${userTimezone}): ${calendarUrl}`);
      return calendarUrl;
    }
    
    console.log('No calendar found for URL generation - calendar link will not appear in notifications');
    return '';
    
  } catch (error) {
    console.error('Failed to generate calendar URL:', error);
    return '';
  }
}

function getGuideLinkFromSpreadsheet() {
  /**
   * Get visitation guide URL from configuration
   */
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const config = getConfiguration(sheet);
    
    if (config.guideUrl && config.guideUrl.trim().length > 0) {
      const trimmedUrl = config.guideUrl.trim();
      
      if (trimmedUrl.startsWith('http://') || trimmedUrl.startsWith('https://')) {
        return trimmedUrl;
      } else {
        console.warn('Visitation Guide URL does not appear to be valid:', trimmedUrl);
        return '';
      }
    }
    
    return '';
    
  } catch (error) {
    console.error('Error reading visitation guide URL:', error);
    return '';
  }
}

function getSummaryLinkFromSpreadsheet() {
  /**
   * Get schedule summary sheet URL from configuration
   */
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const config = getConfiguration(sheet);
    
    if (config.summaryUrl && config.summaryUrl.trim().length > 0) {
      const trimmedUrl = config.summaryUrl.trim();
      
      if (trimmedUrl.startsWith('http://') || trimmedUrl.startsWith('https://')) {
        return trimmedUrl;
      } else {
        console.warn('Schedule Summary Sheet URL does not appear to be valid:', trimmedUrl);
        return '';
      }
    }
    
    return '';
    
  } catch (error) {
    console.error('Error reading schedule summary sheet URL:', error);
    return '';
  }
}

function configureNotifications() {
  /**
   * UI for configuring notification settings
   */
  try {
    const ui = SpreadsheetApp.getUi();
    const properties = PropertiesService.getScriptProperties();
    
    // Get current webhook URLs
    const currentProdWebhook = properties.getProperty('CHAT_WEBHOOK_PROD') || '';
    const currentTestWebhook = properties.getProperty('CHAT_WEBHOOK_TEST') || '';
    
    // Show current configuration
    const currentConfig = `Current Configuration:
    
Production Chat Webhook: ${currentProdWebhook ? '✅ Configured' : '❌ Not configured'}
Test Chat Webhook: ${currentTestWebhook ? '✅ Configured' : '❌ Not configured'}

To get webhook URLs:
1. Go to Google Chat
2. Open your deacon chat space  
3. Click space name → Manage webhooks
4. Add incoming webhook
5. Copy the webhook URL

Configure which webhook?`;
    
    const response = ui.alert(
      'Configure Google Chat Notifications',
      currentConfig,
      ui.ButtonSet.YES_NO_CANCEL
    );
    
    if (response === ui.Button.YES) {
      configureProdWebhook();
    } else if (response === ui.Button.NO) {
      configureTestWebhook();
    }
    
  } catch (error) {
    SpreadsheetApp.getUi().alert(
      'Configuration Error',
      `❌ Error configuring notifications: ${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

function configureProdWebhook() {
  /**
   * Configure production chat webhook
   */
  const ui = SpreadsheetApp.getUi();
  const properties = PropertiesService.getScriptProperties();
  
  const response = ui.prompt(
    'Configure Production Chat Webhook',
    'Enter the webhook URL for your main deacon chat space:\n\n' +
    '(Should start with https://chat.googleapis.com/v1/spaces/...)',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() === ui.Button.OK) {
    const webhookUrl = response.getResponseText().trim();
    
    if (!webhookUrl.startsWith('https://chat.googleapis.com/')) {
      ui.alert('Invalid Webhook URL', 'Please enter a valid Google Chat webhook URL.');
      return;
    }
    
    properties.setProperty('CHAT_WEBHOOK_PROD', webhookUrl);
    ui.alert(
      'Production Webhook Configured',
      '✅ Production chat webhook has been saved!\n\nYou can now send notifications to your main deacon chat space.',
      ui.ButtonSet.OK
    );
  }
}

function configureTestWebhook() {
  /**
   * Configure test chat webhook
   */
  const ui = SpreadsheetApp.getUi();
  const properties = PropertiesService.getScriptProperties();
  
  const response = ui.prompt(
    'Configure Test Chat Webhook',
    'Enter the webhook URL for your test chat space:\n\n' +
    '(Should start with https://chat.googleapis.com/v1/spaces/...)',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() === ui.Button.OK) {
    const webhookUrl = response.getResponseText().trim();
    
    if (!webhookUrl.startsWith('https://chat.googleapis.com/')) {
      ui.alert('Invalid Webhook URL', 'Please enter a valid Google Chat webhook URL.');
      return;
    }
    
    properties.setProperty('CHAT_WEBHOOK_TEST', webhookUrl);
    ui.alert(
      'Test Webhook Configured',
      '✅ Test chat webhook has been saved!\n\nYou can now test notifications in your test chat space.',
      ui.ButtonSet.OK
    );
  }
}

function isNotificationConfigured() {
  /**
   * Check if notifications are properly configured for current mode
   */
  const currentTestMode = getCurrentTestMode();
  const webhookUrl = getWebhookUrl(currentTestMode);
  return webhookUrl && webhookUrl.length > 0;
}

// ===== TESTING FUNCTIONS =====

function testNotificationSystem() {
  /**
   * Test the notification system with current configuration
   */
  try {
    const currentTestMode = getCurrentTestMode();
    const ui = SpreadsheetApp.getUi();
    
    if (!isNotificationConfigured()) {
      ui.alert(
        'Test Failed - Not Configured',
        `❌ No webhook configured for ${currentTestMode ? 'test' : 'production'} mode.\n\n` +
        'Please configure the webhook first:\n📢 Notifications → 🔧 Configure Chat Webhook',
        ui.ButtonSet.OK
      );
      return;
    }
    
    // Send test message
    const testMessage = `${currentTestMode ? '🧪 TEST: ' : ''}🔔 **Notification System Test**\n\n` +
      `✅ Chat integration is working!\n` +
      `📅 Mode: ${currentTestMode ? 'Test' : 'Production'}\n` +
      `🕐 Time: ${new Date().toLocaleString()}\n\n` +
      `Ready to send visitation notifications! 🎉`;
    
    sendToChatSpace(testMessage, currentTestMode);
    
    ui.alert(
      'Test Notification Sent',
      `✅ Test message sent successfully!\n\n` +
      `Mode: ${currentTestMode ? 'Test' : 'Production'}\n` +
      `Check your ${currentTestMode ? 'test' : 'deacon'} chat space to verify delivery.`,
      ui.ButtonSet.OK
    );
    
    console.log(`Test notification sent successfully in ${currentTestMode ? 'test' : 'production'} mode`);
    
  } catch (error) {
    console.error('Test notification failed:', error);
    SpreadsheetApp.getUi().alert(
      'Test Failed',
      `❌ Test notification failed: ${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

function testCalendarLinkConfiguration() {
  /**
   * v1.1: Test calendar auto-detection and URL generation
   * Updated from K19 testing to auto-detection validation
   */
  try {
    const currentTestMode = getCurrentTestMode();
    const calendarName = currentTestMode ? 'TEST - Deacon Visitation Schedule' : 'Deacon Visitation Schedule';
    
    console.log(`Testing calendar detection for mode: ${currentTestMode ? 'TEST' : 'PRODUCTION'}`);
    
    // Test calendar existence
    const calendars = CalendarApp.getCalendarsByName(calendarName);
    if (calendars.length === 0) {
      SpreadsheetApp.getUi().alert(
        'Calendar Test Failed',
        `❌ Calendar not found: "${calendarName}"\n\n` +
        `📅 Expected calendar: "${calendarName}"\n` +
        `🔧 Solution: Run "📆 Full Calendar Regeneration" first\n\n` +
        `💡 The system automatically creates and manages calendars.`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    const calendar = calendars[0];
    const calendarId = calendar.getId();
    
    // Test URL generation
    const generatedUrl = generateCalendarUrlDirect();
    if (!generatedUrl) {
      SpreadsheetApp.getUi().alert(
        'URL Generation Failed',
        `❌ Failed to generate calendar URL\n\n` +
        `📅 Calendar found: "${calendarName}"\n` +
        `🆔 Calendar ID: ${calendarId}\n` +
        `❗ URL generation failed - check logs for details`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    // Test the related functions to ensure integration works
    const guideLink = getGuideLinkFromSpreadsheet();
    const summaryLink = getSummaryLinkFromSpreadsheet();
    
    // Success report
    SpreadsheetApp.getUi().alert(
      currentTestMode ? 'TEST: Calendar Detection Successful' : 'Calendar Detection Successful',
      `${currentTestMode ? '🧪 TEST: ' : ''}✅ Calendar auto-detection working perfectly!\n\n` +
      `📅 Calendar: "${calendarName}"\n` +
      `🆔 Calendar ID: ${calendarId}\n\n` +
      `🔗 Generated URL (for reference):\n${generatedUrl}\n\n` +
      `📋 Guide URL: ${guideLink ? '✅ Configured' : '❌ Not configured (K22)'}\n` +
      `📊 Summary URL: ${summaryLink ? '✅ Configured' : '❌ Not configured (K25)'}\n\n` +
      `💡 Calendar links now auto-generate - no need to manage K19!`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    console.log('Calendar test completed successfully');
    console.log(`Calendar: ${calendarName}`);
    console.log(`Calendar ID: ${calendarId}`);
    console.log(`Generated URL: ${generatedUrl}`);
    
  } catch (error) {
    console.error('Calendar test failed:', error);
    SpreadsheetApp.getUi().alert(
      'Calendar Test Error',
      `❌ Calendar test failed: ${error.message}\n\n` +
      `🔧 Check the execution logs for more details.\n` +
      `💡 Ensure you have calendar permissions and the calendar exists.`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

// ===== DATA VALIDATION FUNCTIONS =====

function validateScheduleDataSync() {
  /**
   * Check if schedule data matches current household and deacon configurations
   * Prevents notifications from failing due to mismatched data
   * Returns true if valid, false if mismatches found
   */
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const config = getConfiguration(sheet);
    const scheduleData = getScheduleFromSheet(sheet);
    
    if (scheduleData.length === 0) {
      SpreadsheetApp.getUi().alert(
        'No Schedule Found',
        '❌ No schedule data found.\n\nPlease generate a schedule first:\n🔄 Deacon Rotation → 📅 Generate Schedule',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return false;
    }
    
    // Get unique households and deacons from schedule
    const scheduleHouseholds = [...new Set(scheduleData.map(visit => visit.household))];
    const scheduleDeacons = [...new Set(scheduleData.map(visit => visit.deacon))];
    
    // Check for household mismatches
    const missingHouseholds = scheduleHouseholds.filter(household => 
      !config.households.includes(household)
    );
    
    const extraHouseholds = config.households.filter(household => 
      !scheduleHouseholds.includes(household)
    );
    
    // Check for deacon mismatches
    const missingDeacons = scheduleDeacons.filter(deacon => 
      !config.deacons.includes(deacon)
    );
    
    const extraDeacons = config.deacons.filter(deacon => 
      !scheduleDeacons.includes(deacon)
    );
    
    // Build mismatch report
    const mismatches = [];
    
    if (missingHouseholds.length > 0) {
      mismatches.push(`❌ Households in schedule but not in household list (column M):\n   ${missingHouseholds.join(', ')}`);
    }
    
    if (extraHouseholds.length > 0) {
      mismatches.push(`⚠️ Households in list (column M) but not in schedule:\n   ${extraHouseholds.join(', ')}`);
    }
    
    if (missingDeacons.length > 0) {
      mismatches.push(`❌ Deacons in schedule but not in deacon list (column L):\n   ${missingDeacons.join(', ')}`);
    }
    
    if (extraDeacons.length > 0) {
      mismatches.push(`⚠️ Deacons in list (column L) but not in schedule:\n   ${extraDeacons.join(', ')}`);
    }
    
    if (mismatches.length > 0) {
      const ui = SpreadsheetApp.getUi();
      const mismatchReport = mismatches.join('\n\n');
      
      const response = ui.alert(
        'Schedule Data Mismatch Detected',
        `🔄 The schedule doesn't match your current household and deacon lists:\n\n${mismatchReport}\n\n` +
        `This may cause missing contact information in notifications.\n\n` +
        `Choose your action:\n` +
        `• YES: Auto-fix by generating new schedule, then continue\n` +
        `• NO: Continue anyway with missing contact info\n` +
        `• CANCEL: Stop and let me fix this manually`,
        ui.ButtonSet.YES_NO_CANCEL
      );
      
      if (response === ui.Button.YES) {
        // Generate new schedule and continue
        try {
          ui.alert(
            'Generating New Schedule',
            '🔄 Creating fresh schedule with current data...\n\nThis may take a moment.',
            ui.ButtonSet.OK
          );
          generateRotationSchedule();
          
          ui.alert(
            'Schedule Updated Successfully',
            '✅ New schedule generated!\n\nProceeding with notification using updated data.',
            ui.ButtonSet.OK
          );
          return true; // Schedule regenerated, can proceed
        } catch (scheduleError) {
          ui.alert(
            'Schedule Generation Failed',
            `❌ Could not generate new schedule: ${scheduleError.message}\n\nNotification cancelled.`,
            ui.ButtonSet.OK
          );
          return false;
        }
      } else if (response === ui.Button.NO) {
        // User wants to proceed anyway
        console.log('User chose to proceed with mismatched data');
        ui.alert(
          'Proceeding with Current Data',
          '⚠️ Continuing with existing schedule.\n\nSome contact information may show as "not available" in the notification.',
          ui.ButtonSet.OK
        );
        return true;
      } else {
        // User cancelled
        console.log('User cancelled notification due to data mismatches');
        ui.alert(
          'Notification Cancelled',
          'ℹ️ Notification stopped.\n\nYou can fix the data mismatches and try again.',
          ui.ButtonSet.OK
        );
        return false;
      }
    }
    
    // No mismatches found
    console.log('Schedule data validation passed - no mismatches found');
    return true;
    
  } catch (error) {
    console.error('Error validating schedule data sync:', error);
    SpreadsheetApp.getUi().alert(
      'Validation Error',
      `❌ Error checking schedule data: ${error.message}\n\nProceeding anyway...`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return true; // Allow to proceed despite validation error
  }
}

function buildBreezeUrl(breezeNumber) {
  /**
   * Build full Breeze URL from number (reused from Module 4)
   */
  if (!breezeNumber || breezeNumber.toString().trim().length === 0) {
    return '';
  }
  
  const cleanNumber = breezeNumber.toString().trim();
  return `https://immanuelky.breezechms.com/people/view/${cleanNumber}`;
}

// ===== TRIGGER MANAGEMENT =====

function createWeeklyNotificationTrigger() {
  /**
   * Create automated weekly trigger - ROBUST VERSION
   * Handles Google Apps Script API inconsistencies gracefully
   */
  try {
    const ui = SpreadsheetApp.getUi();
    const sheet = SpreadsheetApp.getActiveSheet();
    
    console.log('Starting robust trigger creation...');
    
    // Check if trigger already exists - but don't try to delete it here
    const existingTrigger = getExistingWeeklyTrigger();
    if (existingTrigger) {
      console.log('Found existing trigger');
      const response = ui.alert(
        'Trigger Already Exists',
        'A weekly notification trigger is already active.\n\n' +
        'Would you like to create a new one? (This may result in duplicate triggers until you disable auto-send first)',
        ui.ButtonSet.YES_NO
      );
      
      if (response !== ui.Button.YES) {
        console.log('User declined to create additional trigger');
        return;
      }
    }
    
    // Get configuration
    const triggerConfig = getWeeklyTriggerConfig(sheet);
    console.log('Configuration loaded:', triggerConfig);
    
    if (!triggerConfig.isValid) {
      console.log('Invalid configuration');
      ui.alert(
        'Invalid Trigger Configuration',
        `❌ Please check your trigger settings:\n\n${triggerConfig.errors.join('\n')}`,
        ui.ButtonSet.OK
      );
      return;
    }
    
    // Confirm with user
    const response = ui.alert(
      'Configure Weekly Auto-Send',
      `Ready to set up automatic notifications:\n\n` +
      `📅 Day: ${triggerConfig.dayName} (from K11)\n` +
      `🕐 Time: ${triggerConfig.timeFormatted} (from K13)\n` +
      `💬 Chat: ${getCurrentTestMode() ? 'Test Chat Space' : 'Main Deacon Chat Space'}\n\n` +
      `Continue?`,
      ui.ButtonSet.YES_NO
    );
    
    if (response !== ui.Button.YES) {
      console.log('User cancelled');
      return;
    }
    
    // Validate notifications are configured
    if (!isNotificationConfigured()) {
      console.log('Notifications not configured');
      ui.alert(
        'Chat Webhook Required',
        'Please configure the Google Chat webhook first:\n\n📢 Notifications → 🔧 Configure Chat Webhook',
        ui.ButtonSet.OK
      );
      return;
    }
    
    // Create the trigger - this is the core operation
    console.log('Creating trigger...');
    let trigger = null;
    
    try {
      trigger = ScriptApp.newTrigger('sendWeeklyVisitationChat')
        .timeBased()
        .everyWeeks(1)
        .onWeekDay(triggerConfig.weekDay)
        .atHour(triggerConfig.hour)
        .create();
      
      console.log('Trigger created successfully:', trigger.getUniqueId());
      
    } catch (triggerError) {
      console.error('Failed to create trigger:', triggerError);
      throw new Error(`Trigger creation failed: ${triggerError.message}`);
    }
    
    // Store configuration - with error handling
    console.log('Storing configuration...');
    try {
      const properties = PropertiesService.getScriptProperties();
      properties.setProperties({
        'WEEKLY_TRIGGER_ID': trigger.getUniqueId(),
        'WEEKLY_TRIGGER_DAY': triggerConfig.dayName,
        'WEEKLY_TRIGGER_HOUR': triggerConfig.hour.toString(),
        'WEEKLY_TRIGGER_TIME_FORMATTED': triggerConfig.timeFormatted,
        'WEEKLY_TRIGGER_TIMEZONE': Session.getScriptTimeZone(),
        'WEEKLY_TRIGGER_CREATED': new Date().toISOString()
      });
      console.log('Configuration stored successfully');
    } catch (propertiesError) {
      console.warn('Failed to store configuration (non-critical):', propertiesError);
      // Don't fail the whole operation
    }
    
    ui.alert(
      '✅ Weekly Auto-Send Enabled',
      `Weekly notifications are now scheduled!\n\n` +
      `📅 Every ${triggerConfig.dayName} at ${triggerConfig.timeFormatted}\n` +
      `💬 Will send to: ${getCurrentTestMode() ? 'Test Chat Space' : 'Main Deacon Chat Space'}\n\n` +
      `Trigger ID: ${trigger.getUniqueId().substring(0, 8)}...\n\n` +
      `Use "📅 Show Auto-Send Schedule" to verify status.`,
      ui.ButtonSet.OK
    );
    
    console.log('Trigger creation completed successfully');
    
  } catch (error) {
    console.error('Trigger creation failed:', error);
    SpreadsheetApp.getUi().alert(
      'Trigger Creation Failed',
      `❌ Could not create weekly trigger: ${error.message}\n\n` +
      `This may be due to Google Apps Script API limitations.\n` +
      `The trigger might still have been created - check "📅 Show Auto-Send Schedule".`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
* Calendar Week Calculations Functions
*/

function getVisitsForCalendarWeeks(weeksAhead = 2, includeCurrentWeek = true) {
  /**
   * v2.0 CORRECTED: Get visits for calendar weeks with proper date handling
   * Default: Include remainder of current week + next week (true 2-week lookahead)
   */
  try {
    const config = getConfiguration();  // FIXED: No sheet parameter
    const scheduleData = getScheduleFromSheet();  // FIXED: No sheet parameter
    
    if (scheduleData.length === 0) {
      console.log('No schedule data found');
      return [];
    }
    
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const currentDayOfWeek = today.getDay(); // 0=Sunday, 6=Saturday
    
    // Determine start date based on includeCurrentWeek parameter
    let startDate;
    if (includeCurrentWeek) {
      // Start from today (include remainder of current week)
      startDate = new Date(today);
    } else {
      // Start from next Sunday
      const nextSunday = new Date(today);
      if (currentDayOfWeek === 0) {
        // If today is Sunday, next Sunday is 7 days away
        nextSunday.setDate(today.getDate() + 7);
      } else {
        // Otherwise, next Sunday is (7 - currentDayOfWeek) days away
        nextSunday.setDate(today.getDate() + (7 - currentDayOfWeek));
      }
      startDate = nextSunday;
    }
    startDate.setHours(0, 0, 0, 0);
    
    // End date: 2 weeks from start (14 days)
    const endDate = new Date(startDate);
    endDate.setDate(startDate.getDate() + 13); // 14 days total (start + 13 more)
    endDate.setHours(23, 59, 59, 999);
    
    // Debug logging
    console.log(`📅 Calendar Week Lookahead: ${startDate.toLocaleDateString()} to ${endDate.toLocaleDateString()}`);
    console.log(`Today: ${today.toLocaleDateString('en-US', { weekday: 'long' })} (day ${currentDayOfWeek})`);
    console.log(`Include current week: ${includeCurrentWeek}`);
    console.log(`Looking ahead ${weeksAhead} weeks from ${startDate.toLocaleDateString()}`);
    
    // Filter visits within the date range
    const upcomingVisits = scheduleData.filter(visit => {
      if (!visit.date || !(visit.date instanceof Date)) {
        console.warn(`Invalid date for visit: ${visit.household} - ${visit.date}`);
        return false;
      }
      
      const visitDate = new Date(visit.date);
      visitDate.setHours(0, 0, 0, 0);
      
      const inRange = visitDate >= startDate && visitDate <= endDate;
      
      // Debug logging for troubleshooting
      console.log(`${inRange ? '✅ INCLUDED' : '❌ EXCLUDED'}: ${visitDate.toLocaleDateString()} - ${visit.deacon} → ${visit.household}`);
      
      return inRange;
    });
    
    // Sort by date
    upcomingVisits.sort((a, b) => new Date(a.date) - new Date(b.date));
    
   // Enhance with contact information
const enhancedVisits = upcomingVisits.map(visit => {
  const householdIndex = config.households.indexOf(visit.household);
  
  return {
    ...visit,
    phone: (householdIndex >= 0 && config.phones) ? config.phones[householdIndex] : '',
    address: (householdIndex >= 0 && config.addresses) ? config.addresses[householdIndex] : '',
    breezeNumber: (householdIndex >= 0 && config.breezeNumbers) ? config.breezeNumbers[householdIndex] : '',notesLink: (householdIndex >= 0 && config.notesLinks) ? config.notesLinks[householdIndex] : '',
    breezeLink: (householdIndex >= 0 && config.breezeLinks) ? config.breezeLinks[householdIndex] : '',
      // Remove the breezeShortLink and notesShortLink lines entirely
    calendarWeek: getCalendarWeekInfo(visit.date)
  };
});
    
    console.log(`Found ${enhancedVisits.length} visits in ${weeksAhead}-week range`);
    return enhancedVisits;
    
  } catch (error) {
    console.error('Error getting calendar week visits:', error);
    return [];
  }
}

function getCalendarWeekInfo(date) {
  /**
   * Get human-readable calendar week information
   */
  const visitDate = new Date(date);
  const today = new Date();
  
  // Get start of this week (Sunday)
  const thisWeekStart = new Date(today);
  thisWeekStart.setDate(today.getDate() - today.getDay());
  thisWeekStart.setHours(0, 0, 0, 0);
  
  // Get start of visit week (Sunday)
  const visitWeekStart = new Date(visitDate);
  visitWeekStart.setDate(visitDate.getDate() - visitDate.getDay());
  visitWeekStart.setHours(0, 0, 0, 0);
  
  // Calculate week difference
  const weekDiff = Math.round((visitWeekStart - thisWeekStart) / (7 * 24 * 60 * 60 * 1000));
  
  let weekLabel;
  if (weekDiff === 0) {
    weekLabel = 'This Week';
  } else if (weekDiff === 1) {
    weekLabel = 'Next Week';
  } else if (weekDiff === 2) {
    weekLabel = 'Week After Next';
  } else if (weekDiff > 2) {
    weekLabel = `${weekDiff} Weeks Away`;
  } else {
    weekLabel = 'Past Week';
  }
  
  return {
    weekLabel,
    weekStart: visitWeekStart,
    weekEnd: new Date(visitWeekStart.getTime() + 6 * 24 * 60 * 60 * 1000),
    weeksFromNow: weekDiff
  };
}

function getWeekRangeDescription(visits) {
  /**
   * Get a human-readable description of the week range covered
   */
  if (visits.length === 0) return 'No visits';
  
  const weeks = [...new Set(visits.map(v => v.calendarWeek.weekLabel))];
  
  if (weeks.length === 1) {
    return weeks[0];
  } else if (weeks.length === 2) {
    if (weeks.includes('This Week') && weeks.includes('Next Week')) {
      return 'This Week & Next Week';
    } else {
      return weeks.join(' & ');
    }
  } else {
    return `${weeks.length} weeks`;
  }
}

function removeWeeklyNotificationTrigger() {
  /**
   * Remove weekly triggers - with robust error handling
   */
  try {
    const ui = SpreadsheetApp.getUi();
    const triggers = ScriptApp.getProjectTriggers();
    let weeklyTriggers = [];
    
    // Find weekly triggers
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'sendWeeklyVisitationChat') {
        weeklyTriggers.push(trigger);
      }
    });
    
    if (weeklyTriggers.length === 0) {
      ui.alert(
        'No Triggers Found',
        'No weekly notification triggers were found to remove.',
        ui.ButtonSet.OK
      );
      return;
    }
    
    console.log(`Found ${weeklyTriggers.length} weekly triggers to remove`);
    
    let deletedCount = 0;
    let failedCount = 0;
    
    // Try to delete each trigger individually
    weeklyTriggers.forEach(trigger => {
      try {
        console.log(`Attempting to delete trigger: ${trigger.getUniqueId()}`);
        ScriptApp.deleteTrigger(trigger);
        deletedCount++;
        console.log('Successfully deleted trigger');
      } catch (deleteError) {
        console.error('Failed to delete trigger:', deleteError);
        failedCount++;
      }
    });
    
    // Clear stored properties
    try {
      const properties = PropertiesService.getScriptProperties();
      properties.deleteProperty('WEEKLY_TRIGGER_ID');
      properties.deleteProperty('WEEKLY_TRIGGER_DAY');
      properties.deleteProperty('WEEKLY_TRIGGER_HOUR');
      properties.deleteProperty('WEEKLY_TRIGGER_TIME_FORMATTED');
      properties.deleteProperty('WEEKLY_TRIGGER_TIMEZONE');
      properties.deleteProperty('WEEKLY_TRIGGER_CREATED');
      console.log('Cleared stored trigger properties');
    } catch (propertiesError) {
      console.warn('Failed to clear properties:', propertiesError);
    }
    
    if (deletedCount > 0) {
      ui.alert(
        'Triggers Removed',
        `✅ Successfully removed ${deletedCount} weekly trigger(s).\n` +
        `${failedCount > 0 ? `⚠️ Failed to remove ${failedCount} trigger(s).` : ''}\n\n` +
        'Automatic weekly summaries are now disabled.',
        ui.ButtonSet.OK
      );
    } else {
      ui.alert(
        'Removal Failed',
        `❌ Could not remove triggers due to Google Apps Script API issues.\n\n` +
        `Triggers may still be active. Check "🔍 Inspect All Triggers" to verify.`,
        ui.ButtonSet.OK
      );
    }
    
    console.log(`Trigger removal completed: ${deletedCount} deleted, ${failedCount} failed`);
    
  } catch (error) {
    console.error('Trigger removal failed:', error);
    SpreadsheetApp.getUi().alert(
      'Removal Error',
      `❌ Could not remove triggers: ${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}


// ===== TRIGGER CONFIGURATION FUNCTIONS =====

function getWeeklyTriggerConfig(sheet) {
  /**
   * Read weekly trigger configuration from spreadsheet cells
   * Returns validation status and parsed values
   */
  try {
    // Read configuration values from K11 and K13
    const dayValue = sheet.getRange('K11').getValue();
    const timeValue = sheet.getRange('K13').getValue();
    
    const errors = [];
    let weekDay, hour, dayName, timeFormatted;
    
    // Validate day
    if (!dayValue || typeof dayValue !== 'string') {
      errors.push('K14 (Day): Please enter a day name (e.g., Sunday, Monday, Friday)');
    } else {
      const dayStr = dayValue.toString().trim().toLowerCase();
      const dayMap = {
        'sunday': ScriptApp.WeekDay.SUNDAY,
        'monday': ScriptApp.WeekDay.MONDAY,
        'tuesday': ScriptApp.WeekDay.TUESDAY,
        'wednesday': ScriptApp.WeekDay.WEDNESDAY,
        'thursday': ScriptApp.WeekDay.THURSDAY,
        'friday': ScriptApp.WeekDay.FRIDAY,
        'saturday': ScriptApp.WeekDay.SATURDAY
      };
      
      weekDay = dayMap[dayStr];
      if (weekDay === undefined) {
        errors.push(`K14 (Day): "${dayValue}" is not a valid day. Use: Sunday, Monday, Tuesday, etc.`);
      } else {
        dayName = dayStr.charAt(0).toUpperCase() + dayStr.slice(1);
      }
    }
    
    // Validate time
    if (timeValue === '' || timeValue === null || timeValue === undefined) {
      errors.push('K16 (Time): Please enter an hour (0-23)');
    } else {
      hour = parseInt(timeValue);
      if (isNaN(hour) || hour < 0 || hour > 23) {
        errors.push(`K16 (Time): "${timeValue}" is not valid. Enter a number 0-23 (e.g., 6 for 6 AM, 18 for 6 PM)`);
      } else {
        timeFormatted = formatHour(hour);
      }
    }
    
    return {
      isValid: errors.length === 0,
      errors: errors,
      weekDay: weekDay,
      hour: hour,
      dayName: dayName,
      timeFormatted: timeFormatted
    };
    
  } catch (error) {
    return {
      isValid: false,
      errors: [`Error reading configuration: ${error.message}`],
      weekDay: null,
      hour: null,
      dayName: null,
      timeFormatted: null
    };
  }
}

function getExistingWeeklyTrigger() {
  /**
   * Find existing weekly notification trigger
   */
  const triggers = ScriptApp.getProjectTriggers();
  return triggers.find(trigger => 
    trigger.getHandlerFunction() === 'sendWeeklyVisitationChat'
  );
}

function getDayName(weekDay) {
  /**
   * Convert ScriptApp.WeekDay to readable name
   */
  const dayMap = {
    [ScriptApp.WeekDay.SUNDAY]: 'Sunday',
    [ScriptApp.WeekDay.MONDAY]: 'Monday',
    [ScriptApp.WeekDay.TUESDAY]: 'Tuesday', 
    [ScriptApp.WeekDay.WEDNESDAY]: 'Wednesday',
    [ScriptApp.WeekDay.THURSDAY]: 'Thursday',
    [ScriptApp.WeekDay.FRIDAY]: 'Friday',
    [ScriptApp.WeekDay.SATURDAY]: 'Saturday'
  };
  return dayMap[weekDay] || 'Unknown';
}

function getHourFromTrigger(trigger) {
  /**
   * Extract hour from trigger (Google Apps Script doesn't provide direct access)
   * This is a workaround - we'll store hour in properties when creating
   */
  const properties = PropertiesService.getScriptProperties();
  return parseInt(properties.getProperty('WEEKLY_TRIGGER_HOUR') || '18');
}

function formatHour(hour) {
  /**
   * Format hour as readable time
   */
  if (hour === 0) return '12:00 AM (midnight)';
  if (hour < 12) return `${hour}:00 AM`;
  if (hour === 12) return '12:00 PM (noon)';
  return `${hour - 12}:00 PM`;
}

function showCurrentTriggerSchedule() {
  /**
   * Display current weekly trigger schedule and spreadsheet configuration
   */
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const existingTrigger = getExistingWeeklyTrigger();
    const triggerConfig = getWeeklyTriggerConfig(sheet);
    const ui = SpreadsheetApp.getUi();
    
    let message = '📋 Weekly Auto-Send Configuration:\n\n';
    
    // Show current trigger status
    if (existingTrigger) {
      message += '✅ Status: ACTIVE\n';
      message += `📅 Running: Weekly notifications enabled\n`;
      message += `💬 Destination: ${getCurrentTestMode() ? 'Test Chat Space' : 'Main Deacon Chat Space'}\n\n`;
    } else {
      message += '⏸️ Status: INACTIVE\n';
      message += `📅 Running: No weekly notifications scheduled\n\n`;
    }
    
    // Show spreadsheet configuration
    message += '📊 Spreadsheet Configuration:\n';
    if (triggerConfig.isValid) {
      message += `📅 Day: ${triggerConfig.dayName} (cell K11)\n`;
      message += `🕐 Time: ${triggerConfig.timeFormatted} (cell K13)\n\n`;
      
      if (!existingTrigger) {
        message += 'ℹ️ Use "🔄 Enable Weekly Auto-Send" to activate these settings.';
      } else {
        message += 'ℹ️ To change schedule: edit K11/K13, then disable and re-enable auto-send.';
      }
    } else {
      message += '❌ Invalid Configuration:\n';
      triggerConfig.errors.forEach(error => {
        message += `   • ${error}\n`;
      });
      message += '\n📝 Please fix the configuration in column K, then try again.';
    }
    
    ui.alert(
      'Weekly Auto-Send Schedule',
      message,
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    console.error('Failed to show trigger schedule:', error);
    SpreadsheetApp.getUi().alert(
      'Schedule Check Failed',
      `❌ Could not check schedule: ${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

function testNotificationNow() {
  /**
   * Test function to send notification immediately (bypasses trigger timing)
   */
  try {
    const ui = SpreadsheetApp.getUi();
    
    const response = ui.alert(
      'Test Notification',
      '🧪 This will send a test notification immediately, bypassing the scheduled trigger.\n\n' +
      'This helps verify that the notification system works, separate from timing issues.\n\n' +
      'Continue?',
      ui.ButtonSet.YES_NO
    );
    
    if (response === ui.Button.YES) {
      // Call the notification function directly
      sendWeeklyVisitationChat();
      
      ui.alert(
        'Test Complete',
        '✅ Test notification sent!\n\n' +
        'Check your Google Chat space to see if it arrived.\n\n' +
        'If this worked but the scheduled trigger didn\'t, it\'s likely a timezone issue.',
        ui.ButtonSet.OK
      );
    }
    
  } catch (error) {
    console.error('Test notification failed:', error);
    SpreadsheetApp.getUi().alert(
      'Test Failed',
      `❌ Test notification failed: ${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

function inspectAllTriggers() {
  /**
   * Detailed inspection of all project triggers
   */
  try {
    const ui = SpreadsheetApp.getUi();
    const triggers = ScriptApp.getProjectTriggers();
    
    if (triggers.length === 0) {
      ui.alert(
        'No Triggers Found',
        '❌ No triggers exist in this project.\n\n' +
        'This explains why notifications aren\'t being sent automatically.\n\n' +
        'Use "🔄 Enable Weekly Auto-Send" to create a trigger.',
        ui.ButtonSet.OK
      );
      return;
    }
    
    let message = `🔍 TRIGGER INSPECTION (${triggers.length} total):\n\n`;
    
    triggers.forEach((trigger, index) => {
      const handlerFunction = trigger.getHandlerFunction();
      const triggerSource = trigger.getTriggerSource();
      const eventType = trigger.getEventType();
      const triggerId = trigger.getUniqueId();
      
      message += `${index + 1}. Function: ${handlerFunction}\n`;
      message += `   Source: ${triggerSource}\n`;
      message += `   Type: ${eventType}\n`;
      message += `   ID: ${triggerId.substring(0, 8)}...\n`;
      
      // Special handling for time-based triggers
      if (eventType === ScriptApp.EventType.CLOCK) {
        try {
          // We can't directly access trigger schedule details, but we can check our stored properties
          const properties = PropertiesService.getScriptProperties();
          if (handlerFunction === 'sendWeeklyVisitationChat') {
            const storedDay = properties.getProperty('WEEKLY_TRIGGER_DAY');
            const storedHour = properties.getProperty('WEEKLY_TRIGGER_HOUR');
            const storedTimezone = properties.getProperty('WEEKLY_TRIGGER_TIMEZONE');
            const storedCreated = properties.getProperty('WEEKLY_TRIGGER_CREATED');
            
            if (storedDay) {
              message += `   ⏰ Schedule: ${storedDay} at ${storedHour}:00\n`;
              message += `   🕐 Timezone: ${storedTimezone || 'Unknown'}\n`;
              message += `   📅 Created: ${storedCreated ? new Date(storedCreated).toLocaleString() : 'Unknown'}\n`;
            }
          }
        } catch (detailError) {
          message += `   ⚠️ Could not read schedule details\n`;
        }
      }
      
      message += '\n';
    });
    
    // Check for weekly notification triggers specifically
    const weeklyTriggers = triggers.filter(t => t.getHandlerFunction() === 'sendWeeklyVisitationChat');
    
    if (weeklyTriggers.length === 0) {
      message += '❌ NO WEEKLY NOTIFICATION TRIGGERS FOUND!\n';
      message += 'This is why notifications aren\'t being sent.\n\n';
      message += 'Solution: Use "🔄 Enable Weekly Auto-Send"';
    } else if (weeklyTriggers.length > 1) {
      message += `⚠️ WARNING: ${weeklyTriggers.length} weekly triggers found!\n`;
      message += 'Multiple triggers may cause duplicate notifications.\n\n';
      message += 'Solution: Use "🛑 Disable Weekly Auto-Send" then re-enable.';
    } else {
      message += '✅ Found 1 weekly notification trigger (correct)';
    }
    
    ui.alert('Trigger Inspection Results', message, ui.ButtonSet.OK);
    
    console.log('Trigger inspection completed');
    console.log('Total triggers:', triggers.length);
    console.log('Weekly triggers:', weeklyTriggers.length);
    
  } catch (error) {
    console.error('Trigger inspection failed:', error);
    SpreadsheetApp.getUi().alert(
      'Inspection Error',
      `❌ Could not inspect triggers: ${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

function forceRecreateWeeklyTrigger() {
  /**
   * Force delete and recreate the weekly trigger
   */
  try {
    const ui = SpreadsheetApp.getUi();
    
    const response = ui.alert(
      'Force Recreate Trigger',
      '🔄 This will:\n' +
      '1. Delete ALL weekly notification triggers\n' +
      '2. Create a fresh trigger with current settings\n' +
      '3. Store new configuration data\n\n' +
      'Continue?',
      ui.ButtonSet.YES_NO
    );
    
    if (response !== ui.Button.YES) {
      return;
    }
    
    // Delete ALL weekly triggers (in case there are duplicates)
    const triggers = ScriptApp.getProjectTriggers();
    let deletedCount = 0;
    
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'sendWeeklyVisitationChat') {
        ScriptApp.deleteTrigger(trigger);
        deletedCount++;
      }
    });
    
    console.log(`Deleted ${deletedCount} existing weekly triggers`);
    
    // Get current configuration
    const sheet = SpreadsheetApp.getActiveSheet();
    const triggerConfig = getWeeklyTriggerConfig(sheet);
    
    if (!triggerConfig.isValid) {
      ui.alert(
        'Invalid Configuration',
        `❌ Cannot create trigger with invalid settings:\n\n${triggerConfig.errors.join('\n')}`,
        ui.ButtonSet.OK
      );
      return;
    }
    
    // Create fresh trigger
    const trigger = ScriptApp.newTrigger('sendWeeklyVisitationChat')
      .timeBased()
      .everyWeeks(1)
      .onWeekDay(triggerConfig.weekDay)
      .atHour(triggerConfig.hour)
      .create();
    
    // Store fresh configuration
    const scriptTimezone = Session.getScriptTimeZone();
    const properties = PropertiesService.getScriptProperties();
    properties.setProperties({
      'WEEKLY_TRIGGER_ID': trigger.getUniqueId(),
      'WEEKLY_TRIGGER_DAY': triggerConfig.dayName,
      'WEEKLY_TRIGGER_HOUR': triggerConfig.hour.toString(),
      'WEEKLY_TRIGGER_TIMEZONE': scriptTimezone,
      'WEEKLY_TRIGGER_CREATED': new Date().toISOString(),
      'WEEKLY_TRIGGER_RECREATED': 'true'
    });
    
    ui.alert(
      '✅ Trigger Recreated Successfully',
      `Fresh weekly trigger created!\n\n` +
      `📅 Schedule: ${triggerConfig.dayName} at ${triggerConfig.timeFormatted}\n` +
      `🕐 Timezone: ${scriptTimezone}\n` +
      `🗑️ Deleted: ${deletedCount} old trigger(s)\n` +
      `🆕 Created: 1 fresh trigger\n\n` +
      `Next firing: Next ${triggerConfig.dayName} at ${triggerConfig.timeFormatted}`,
      ui.ButtonSet.OK
    );
    
    console.log(`Force recreated trigger: ${triggerConfig.dayName} at ${triggerConfig.hour}:00`);
    
  } catch (error) {
    console.error('Force recreate failed:', error);
    ui.alert(
      'Recreate Failed',
      `❌ Could not recreate trigger: ${error.message}`,
      ui.ButtonSet.OK
    );
  }
}
