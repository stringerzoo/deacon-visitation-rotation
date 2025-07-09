/**
 * MODULE 5: GOOGLE CHAT NOTIFICATIONS (v25.0)
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
    const upcomingVisits = getVisitsForCalendarWeeks(2, false);
    
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
   * UPDATED: Build summary with configurable calendar link from spreadsheet
   */
  const chatPrefix = isTestMode ? '🧪 TEST: ' : '';
  const today = new Date();
  const todayFormatted = today.toLocaleDateString('en-US', { 
    weekday: 'long', 
    month: 'long', 
    day: 'numeric' 
  });
  
  // Calculate next Sunday
  const currentDayOfWeek = today.getDay();
  const nextSunday = new Date(today);
  if (currentDayOfWeek === 0) {
    nextSunday.setDate(today.getDate() + 7);
  } else {
    nextSunday.setDate(today.getDate() + (7 - currentDayOfWeek));
  }
  nextSunday.setHours(0, 0, 0, 0);
  
  // Calculate week boundaries
  const week1Start = new Date(nextSunday);
  const week1End = new Date(nextSunday);
  week1End.setDate(nextSunday.getDate() + 6);
  
  const week2Start = new Date(nextSunday);
  week2Start.setDate(nextSunday.getDate() + 7);
  const week2End = new Date(nextSunday);
  week2End.setDate(nextSunday.getDate() + 13);
  
  // Group visits by week
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
  message += `Coverage: 2 weeks starting ${nextSunday.toLocaleDateString('en-US', { month: 'short', day: 'numeric' })}\n\n`;
  
  // Week 1 Section
  message += `*Week 1* (${week1Start.toLocaleDateString('en-US', { 
    month: 'short', 
    day: 'numeric' 
  })} - ${week1End.toLocaleDateString('en-US', { 
    month: 'short', 
    day: 'numeric' 
  })})\n`;
  
  if (week1Visits.length === 0) {
    message += `No visits scheduled for Week 1.\n\n`;
  } else {
    week1Visits.forEach(visit => {
      const visitDate = visit.date.toLocaleDateString('en-US', { 
        weekday: 'short', 
        month: 'short', 
        day: 'numeric' 
      });
      
      message += `📅 ${visitDate} - *${visit.deacon}* visits *${visit.household}*\n`;
      
      // Add contact info
      const contactInfo = [];
      if (visit.phone) contactInfo.push(`📞 ${visit.phone}`);
      if (visit.address) contactInfo.push(`🏠 ${visit.address}`);
      
      if (contactInfo.length > 0) {
        message += `   ${contactInfo.join(' • ')}\n`;
      }
      
      // Add links on separate lines with Google Chat format
      if (visit.breezeShortLink) {
        message += `   🔗 <${visit.breezeShortLink}|Breeze Profile>\n`;
      }
      if (visit.notesShortLink) {
        message += `   📝 <${visit.notesShortLink}|Visit Notes>\n`;
      }
      
      message += '\n';
    });
  }
  
  message += '---\n\n';
  
  // Week 2 Section
  message += `*Week 2* (${week2Start.toLocaleDateString('en-US', { 
    month: 'short', 
    day: 'numeric' 
  })} - ${week2End.toLocaleDateString('en-US', { 
    month: 'short', 
    day: 'numeric' 
  })})\n`;
  
  if (week2Visits.length === 0) {
    message += `No visits scheduled for Week 2.\n\n`;
  } else {
    week2Visits.forEach(visit => {
      const visitDate = visit.date.toLocaleDateString('en-US', { 
        weekday: 'short', 
        month: 'short', 
        day: 'numeric' 
      });
      
      message += `📅 ${visitDate} - *${visit.deacon}* visits *${visit.household}*\n`;
      
      // Add contact info
      const contactInfo = [];
      if (visit.phone) contactInfo.push(`📞 ${visit.phone}`);
      if (visit.address) contactInfo.push(`🏠 ${visit.address}`);
      
      if (contactInfo.length > 0) {
        message += `   ${contactInfo.join(' • ')}\n`;
      }
      
      // Add links on separate lines with Google Chat format
      if (visit.breezeShortLink) {
        message += `   🔗 <${visit.breezeShortLink}|Breeze Profile>\n`;
      }
      if (visit.notesShortLink) {
        message += `   📝 <${visit.notesShortLink}|Visit Notes>\n`;
      }
      
      message += '\n';
    });
  }
  
  // Footer with instructions and configurable calendar link
  message += `💡 *Instructions*: Call ahead to confirm visit times. Contact families 1-2 days before your scheduled visit.\n\n`;
  
  // NEW: Add calendar link from spreadsheet configuration
  const calendarLink = getCalendarLinkFromSpreadsheet();
  if (calendarLink) {
    message += `📅 <${calendarLink}|View Visitation Calendar>\n\n`;
  }
  
  message += `🔄 This update is sent weekly. Reply here with questions or scheduling conflicts.`;
  
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
   * Get Breeze link for household (prefer shortened, fallback to full)
   */
  if (householdIndex < 0) return '';
  
  // Try shortened link first
  const shortLink = config.breezeShortLinks[householdIndex];
  if (shortLink && shortLink.trim().length > 0) {
    return shortLink;
  }
  
  // Fallback to building from Breeze number
  const breezeNumber = config.breezeNumbers[householdIndex];
  if (breezeNumber && breezeNumber.trim().length > 0) {
    return buildBreezeUrl(breezeNumber);
  }
  
  return '';
}

function getNotesLink(config, householdIndex) {
  /**
   * Get Notes link for household (prefer shortened, fallback to full)
   */
  if (householdIndex < 0) return '';
  
  // Try shortened link first
  const shortLink = config.notesShortLinks[householdIndex];
  if (shortLink && shortLink.trim().length > 0) {
    return shortLink;
  }
  
  // Fallback to full link
  const fullLink = config.notesLinks[householdIndex];
  if (fullLink && fullLink.trim().length > 0) {
    return fullLink;
  }
  
  return '';
}

// ===== CONFIGURATION FUNCTIONS =====

function getCalendarLinkFromSpreadsheet() {
  /**
   * Get calendar URL from configuration (now handled by Module 1)
   */
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const config = getConfiguration(sheet);
    
    if (config.calendarUrl && config.calendarUrl.trim().length > 0) {
      const trimmedUrl = config.calendarUrl.trim();
      
      if (trimmedUrl.startsWith('http://') || trimmedUrl.startsWith('https://')) {
        return trimmedUrl;
      } else {
        console.warn('Calendar URL does not appear to be valid:', trimmedUrl);
        return '';
      }
    }
    
    return '';
    
  } catch (error) {
    console.error('Error reading calendar URL:', error);
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
   * Test function to verify calendar link configuration
   */
  try {
    const calendarUrl = getCalendarLinkFromSpreadsheet();
    const ui = SpreadsheetApp.getUi();
    
    let message = '📅 CALENDAR LINK CONFIGURATION TEST\n\n';
    
    if (calendarUrl) {
      message += `✅ Calendar URL found in K19:\n${calendarUrl}\n\n`;
      message += `This link will appear in weekly chat messages as:\n📅 View Visitation Calendar`;
    } else {
      message += `❌ No calendar URL found in K19\n\n`;
      message += `To configure:\n`;
      message += `1. Paste your Google Calendar embed URL in cell K19\n`;
      message += `2. For testing: Use test calendar URL\n`;
      message += `3. For production: Use production calendar URL`;
    }
    
    ui.alert('Calendar Link Test Results', message, ui.ButtonSet.OK);
    
  } catch (error) {
    console.error('Calendar link test failed:', error);
    SpreadsheetApp.getUi().alert(
      'Test Error',
      `❌ Test failed: ${error.message}`,
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
  if (!breezeNumber || breezeNumber.trim().length === 0) {
    return '';
  }
  
  const cleanNumber = breezeNumber.trim();
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

function getVisitsForCalendarWeeks(weeksAhead = 2, includeCurrentWeek = false) {
  /**
   * CORRECTED: Get visits for proper calendar weeks starting from NEXT Sunday
   * For 2-week lookahead: Next Sunday through Saturday 2 weeks later
   */
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const config = getConfiguration(sheet);
    const scheduleData = getScheduleFromSheet(sheet);
    
    if (scheduleData.length === 0) {
      console.log('No schedule data found');
      return [];
    }
    
    const today = new Date();
    const currentDayOfWeek = today.getDay(); // 0=Sunday, 6=Saturday
    
    // ALWAYS start from next Sunday for the lookahead
    const nextSunday = new Date(today);
    if (currentDayOfWeek === 0) {
      // If today is Sunday, next Sunday is 7 days away
      nextSunday.setDate(today.getDate() + 7);
    } else {
      // Otherwise, next Sunday is (7 - currentDayOfWeek) days away
      nextSunday.setDate(today.getDate() + (7 - currentDayOfWeek));
    }
    nextSunday.setHours(0, 0, 0, 0);
    
    // End date: Saturday of the final week
    const endDate = new Date(nextSunday);
    endDate.setDate(nextSunday.getDate() + (7 * weeksAhead) - 1); // -1 to end on Saturday
    endDate.setHours(23, 59, 59, 999);
    
    // Debug logging
    console.log(`📅 Calendar Week Lookahead: ${nextSunday.toLocaleDateString()} to ${endDate.toLocaleDateString()}`);
    console.log(`Today: ${today.toLocaleDateString('en-US', { weekday: 'long' })} (day ${currentDayOfWeek})`);
    console.log(`Next Sunday: ${nextSunday.toLocaleDateString('en-US', { weekday: 'long' })}`);
    console.log(`Looking ahead ${weeksAhead} complete calendar weeks`);
    
    // Filter visits within the calendar week range
    const upcomingVisits = scheduleData.filter(visit => {
      if (!visit.date || !(visit.date instanceof Date)) {
        console.warn(`Invalid date for visit: ${visit.household} - ${visit.date}`);
        return false;
      }
      
      const visitDate = new Date(visit.date);
      visitDate.setHours(0, 0, 0, 0);
      
      const inRange = visitDate >= nextSunday && visitDate <= endDate;
      
      // Debug logging for troubleshooting
      console.log(`${inRange ? '✅ INCLUDED' : '❌ EXCLUDED'}: ${visitDate.toLocaleDateString()} - ${visit.deacon} → ${visit.household}`);
      
      return inRange;
    });
    
    // Sort by date
    upcomingVisits.sort((a, b) => new Date(a.date) - new Date(b.date));
    
    // Enhance with contact information and calendar week context
    const enhancedVisits = upcomingVisits.map(visit => {
      const householdIndex = config.households.indexOf(visit.household);
      
      return {
        ...visit,
        phone: householdIndex >= 0 ? config.phones[householdIndex] : '',
        address: householdIndex >= 0 ? config.addresses[householdIndex] : '',
        breezeNumber: householdIndex >= 0 ? config.breezeNumbers[householdIndex] : '',
        notesLink: householdIndex >= 0 ? config.notesLinks[householdIndex] : '',
        breezeShortLink: householdIndex >= 0 ? config.breezeShortLinks[householdIndex] : '',
        notesShortLink: householdIndex >= 0 ? config.notesShortLinks[householdIndex] : '',
        calendarWeek: getCalendarWeekInfoFromNextSunday(visit.date, nextSunday)
      };
    });
    
    console.log(`Found ${enhancedVisits.length} visits in 2-week calendar range`);
    return enhancedVisits;
    
  } catch (error) {
    console.error('Error getting calendar week visits:', error);
    return [];
  }
}

function getCalendarWeekInfoFromNextSunday(visitDate, nextSunday) {
  /**
   * Get calendar week info relative to the next Sunday starting point
   */
  const visit = new Date(visitDate);
  const startSunday = new Date(nextSunday);
  
  // Calculate which week this visit falls in relative to next Sunday
  const daysDiff = Math.floor((visit - startSunday) / (24 * 60 * 60 * 1000));
  const weekNumber = Math.floor(daysDiff / 7);
  
  let weekLabel;
  if (weekNumber === 0) {
    weekLabel = 'Week 1';
  } else if (weekNumber === 1) {
    weekLabel = 'Week 2';
  } else if (weekNumber >= 2) {
    weekLabel = `Week ${weekNumber + 1}`;
  } else {
    weekLabel = 'Before Range'; // Shouldn't happen
  }
  
  // Calculate week boundaries
  const weekStart = new Date(startSunday);
  weekStart.setDate(startSunday.getDate() + (weekNumber * 7));
  
  const weekEnd = new Date(weekStart);
  weekEnd.setDate(weekStart.getDate() + 6);
  
  return {
    weekLabel,
    weekStart: weekStart,
    weekEnd: weekEnd,
    weekNumber: weekNumber + 1
  };
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

function debugWeeklyCalendarSummary() {
  /**
   * Debug the 2-week lookahead logic specifically
   */
  try {
    const today = new Date();
    const ui = SpreadsheetApp.getUi();
    
    console.log('=== WEEKLY CALENDAR SUMMARY DEBUG ===');
    console.log(`Today: ${today.toLocaleDateString()} (${today.toLocaleDateString('en-US', { weekday: 'long' })})`);
    console.log(`Day of week: ${today.getDay()} (0=Sunday, 6=Saturday)`);
    
    // Test the exact logic used in weekly summary
    const weeklyVisits = getVisitsForCalendarWeeks(2, true);
    
    console.log(`\nFound ${weeklyVisits.length} visits for 2-week lookahead:`);
    weeklyVisits.forEach(visit => {
      console.log(`  ${visit.calendarWeek.weekLabel} - ${visit.date.toLocaleDateString()} - ${visit.deacon} → ${visit.household}`);
    });
    
    // Group by week to show the structure
    const visitsByWeek = {};
    weeklyVisits.forEach(visit => {
      const weekKey = visit.calendarWeek.weekLabel;
      if (!visitsByWeek[weekKey]) visitsByWeek[weekKey] = [];
      visitsByWeek[weekKey].push(visit);
    });
    
    let message = `📅 WEEKLY SUMMARY DEBUG\n\n`;
    message += `Today: ${today.toLocaleDateString('en-US', { weekday: 'long', month: 'long', day: 'numeric' })}\n\n`;
    message += `🔍 2-Week Lookahead Results:\n`;
    message += `Total visits found: ${weeklyVisits.length}\n\n`;
    
    Object.keys(visitsByWeek).forEach(weekKey => {
      const weekVisits = visitsByWeek[weekKey];
      message += `**${weekKey}**: ${weekVisits.length} visits\n`;
      
      weekVisits.slice(0, 3).forEach(visit => {
        message += `  • ${visit.date.toLocaleDateString()} - ${visit.deacon} → ${visit.household}\n`;
      });
      
      if (weekVisits.length > 3) {
        message += `  ... and ${weekVisits.length - 3} more\n`;
      }
      message += '\n';
    });
    
    if (weeklyVisits.length === 0) {
      message += `❌ NO VISITS FOUND!\n\n`;
      message += `This suggests:\n`;
      message += `• Schedule may need to be regenerated\n`;
      message += `• Current date range may be outside schedule\n`;
      message += `• Date filtering logic may have issues\n\n`;
      message += `Check the console for detailed logging.`;
    } else {
      message += `✅ System is working correctly!\n`;
      message += `This is what would be sent to the chat space.`;
    }
    
    ui.alert('Weekly Summary Debug Results', message, ui.ButtonSet.OK);
    
  } catch (error) {
    console.error('Weekly calendar debug failed:', error);
    SpreadsheetApp.getUi().alert(
      'Debug Error',
      `❌ Debug failed: ${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

function debugCalendarWeekLookahead() {
  /**
   * Debug function specifically for the calendar week lookahead logic
   */
  try {
    const today = new Date();
    const currentDayOfWeek = today.getDay();
    const ui = SpreadsheetApp.getUi();
    
    // Calculate next Sunday
    const nextSunday = new Date(today);
    if (currentDayOfWeek === 0) {
      nextSunday.setDate(today.getDate() + 7);
    } else {
      nextSunday.setDate(today.getDate() + (7 - currentDayOfWeek));
    }
    nextSunday.setHours(0, 0, 0, 0);
    
    // Calculate 2-week range
    const week1Start = new Date(nextSunday);
    const week1End = new Date(nextSunday);
    week1End.setDate(nextSunday.getDate() + 6);
    
    const week2Start = new Date(nextSunday);
    week2Start.setDate(nextSunday.getDate() + 7);
    const week2End = new Date(nextSunday);
    week2End.setDate(nextSunday.getDate() + 13);
    
    console.log('=== CALENDAR WEEK LOOKAHEAD DEBUG ===');
    console.log(`Today: ${today.toLocaleDateString()} (${today.toLocaleDateString('en-US', { weekday: 'long' })})`);
    console.log(`Day of week: ${currentDayOfWeek} (0=Sunday)`);
    console.log(`Next Sunday: ${nextSunday.toLocaleDateString()}`);
    console.log(`Week 1: ${week1Start.toLocaleDateString()} to ${week1End.toLocaleDateString()}`);
    console.log(`Week 2: ${week2Start.toLocaleDateString()} to ${week2End.toLocaleDateString()}`);
    
    // Test the function
    const weeklyVisits = getVisitsForCalendarWeeks(2, false);
    
    console.log(`\nFound ${weeklyVisits.length} visits for 2-week calendar lookahead:`);
    weeklyVisits.forEach(visit => {
      console.log(`  ${visit.calendarWeek.weekLabel} - ${visit.date.toLocaleDateString()} - ${visit.deacon} → ${visit.household}`);
    });
    
    // Show all schedule data for comparison
    const sheet = SpreadsheetApp.getActiveSheet();
    const scheduleData = getScheduleFromSheet(sheet);
    console.log(`\nAll schedule entries (first 10 of ${scheduleData.length}):`);
    scheduleData.slice(0, 10).forEach(visit => {
      console.log(`  ${visit.date.toLocaleDateString()} - ${visit.deacon} → ${visit.household}`);
    });
    
    let message = `📅 CALENDAR WEEK LOOKAHEAD DEBUG\n\n`;
    message += `Today: ${today.toLocaleDateString('en-US', { weekday: 'long', month: 'long', day: 'numeric' })}\n\n`;
    message += `📊 Calendar Week Boundaries:\n`;
    message += `• Next Sunday: ${nextSunday.toLocaleDateString()}\n`;
    message += `• Week 1: ${week1Start.toLocaleDateString()} - ${week1End.toLocaleDateString()}\n`;
    message += `• Week 2: ${week2Start.toLocaleDateString()} - ${week2End.toLocaleDateString()}\n\n`;
    message += `🔍 Results:\n`;
    message += `• Total visits found: ${weeklyVisits.length}\n`;
    message += `• Total schedule entries: ${scheduleData.length}\n\n`;
    
    if (weeklyVisits.length > 0) {
      // Group by week
      const week1Count = weeklyVisits.filter(v => v.calendarWeek.weekNumber === 1).length;
      const week2Count = weeklyVisits.filter(v => v.calendarWeek.weekNumber === 2).length;
      
      message += `📝 Visits by Week:\n`;
      message += `• Week 1: ${week1Count} visits\n`;
      message += `• Week 2: ${week2Count} visits\n\n`;
      
      message += `📋 Visit Details:\n`;
      weeklyVisits.slice(0, 5).forEach(visit => {
        message += `• ${visit.calendarWeek.weekLabel} - ${visit.date.toLocaleDateString()} - ${visit.deacon} → ${visit.household}\n`;
      });
      if (weeklyVisits.length > 5) {
        message += `... and ${weeklyVisits.length - 5} more\n`;
      }
    } else {
      message += `❌ NO VISITS FOUND!\n\n`;
      message += `Expected to find visits between:\n`;
      message += `${week1Start.toLocaleDateString()} and ${week2End.toLocaleDateString()}\n\n`;
      message += `Check console for detailed filtering logs.`;
    }
    
    ui.alert('Calendar Week Lookahead Debug', message, ui.ButtonSet.OK);
    
  } catch (error) {
    console.error('Calendar week lookahead debug failed:', error);
    SpreadsheetApp.getUi().alert(
      'Debug Error',
      `❌ Debug failed: ${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

// End of Calendar Week Calculation Functions

/**
* Trigger Debug function
*/

function debugTriggerCreation() {
  /**
   * Step-by-step debugging function to isolate the issue
   */
  try {
    const ui = SpreadsheetApp.getUi();
    
    console.log('=== DEBUGGING TRIGGER CREATION ===');
    
    // Step 1: Check configuration
    const sheet = SpreadsheetApp.getActiveSheet();
    const triggerConfig = getWeeklyTriggerConfig(sheet);
    console.log('Step 1 - Configuration:', triggerConfig);
    
    if (!triggerConfig.isValid) {
      ui.alert('Debug Result', `❌ Configuration invalid:\n${triggerConfig.errors.join('\n')}`, ui.ButtonSet.OK);
      return;
    }
    
    // Step 2: Test basic trigger creation
    console.log('Step 2 - Testing basic trigger creation...');
    let testTrigger = null;
    
    try {
      testTrigger = ScriptApp.newTrigger('sendWeeklyVisitationChat')
        .timeBased()
        .everyWeeks(1)
        .onWeekDay(triggerConfig.weekDay)
        .atHour(triggerConfig.hour)
        .create();
      
      console.log('Step 2 - SUCCESS: Trigger created with ID:', testTrigger.getUniqueId());
      
      // Step 3: Test trigger properties access
      console.log('Step 3 - Testing trigger property access...');
      const triggerId = testTrigger.getUniqueId();
      const handlerFunction = testTrigger.getHandlerFunction();
      console.log('Step 3 - SUCCESS: ID:', triggerId, 'Function:', handlerFunction);
      
      // Step 4: Test properties storage
      console.log('Step 4 - Testing properties storage...');
      const properties = PropertiesService.getScriptProperties();
      properties.setProperty('DEBUG_TRIGGER_TEST', triggerId);
      const stored = properties.getProperty('DEBUG_TRIGGER_TEST');
      console.log('Step 4 - SUCCESS: Stored and retrieved:', stored);
      
      // Clean up test trigger
      ScriptApp.deleteTrigger(testTrigger);
      properties.deleteProperty('DEBUG_TRIGGER_TEST');
      console.log('Step 5 - SUCCESS: Cleaned up test trigger');
      
      ui.alert(
        'Debug Complete',
        '✅ All trigger creation steps passed!\n\n' +
        'The error might be elsewhere. Check the console logs for details.',
        ui.ButtonSet.OK
      );
      
    } catch (triggerError) {
      console.error('Step 2 FAILED - Trigger creation error:', triggerError);
      ui.alert(
        'Debug Failed',
        `❌ Trigger creation failed at step 2:\n\n${triggerError.message}`,
        ui.ButtonSet.OK
      );
      
      // Clean up if needed
      if (testTrigger) {
        try {
          ScriptApp.deleteTrigger(testTrigger);
        } catch (cleanupError) {
          console.error('Cleanup failed:', cleanupError);
        }
      }
    }
    
    console.log('=== DEBUG COMPLETE ===');
    
  } catch (error) {
    console.error('Debug function failed:', error);
    SpreadsheetApp.getUi().alert(
      'Debug Error',
      `❌ Debug failed: ${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}
/**
* End Trigger Debug function
*/

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

// ===== TIMEZONE DIAGNOSTIC FUNCTIONS =====

function checkTimezoneSettings() {
  /**
   * Diagnostic function to check timezone settings and current time
   */
  try {
    const ui = SpreadsheetApp.getUi();
    
    // Get script timezone
    const scriptTimezone = Session.getScriptTimeZone();
    
    // Get current time in script timezone
    const now = new Date();
    const scriptTime = Utilities.formatDate(now, scriptTimezone, 'yyyy-MM-dd HH:mm:ss z');
    
    // Get spreadsheet timezone (if different)
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const spreadsheetTimezone = spreadsheet.getSpreadsheetTimeZone();
    const spreadsheetTime = Utilities.formatDate(now, spreadsheetTimezone, 'yyyy-MM-dd HH:mm:ss z');
    
    // Check existing triggers
    const triggers = ScriptApp.getProjectTriggers();
    const weeklyTriggers = triggers.filter(t => t.getHandlerFunction() === 'sendWeeklyVisitationChat');
    
    let triggerInfo = 'No weekly triggers found';
    if (weeklyTriggers.length > 0) {
      const trigger = weeklyTriggers[0];
      const triggerSource = trigger.getTriggerSource();
      const eventType = trigger.getEventType();
      triggerInfo = `Found ${weeklyTriggers.length} weekly trigger(s)`;
    }
    
    // Get configuration from spreadsheet
    const sheet = SpreadsheetApp.getActiveSheet();
    const dayValue = sheet.getRange('K11').getValue();
    const timeValue = sheet.getRange('K13').getValue();
    
    const message = `🕐 TIMEZONE DIAGNOSTIC REPORT:\n\n` +
      `📍 Script Timezone: ${scriptTimezone}\n` +
      `⏰ Current Script Time: ${scriptTime}\n\n` +
      `📊 Spreadsheet Timezone: ${spreadsheetTimezone}\n` +
      `🕒 Current Spreadsheet Time: ${spreadsheetTime}\n\n` +
      `⚙️ Trigger Status: ${triggerInfo}\n\n` +
      `📋 Configuration:\n` +
      `• Day (K11): ${dayValue}\n` +
      `• Hour (K13): ${timeValue}\n\n` +
      `💡 If timezones don't match your location, triggers may fire at unexpected times!`;
    
    ui.alert('Timezone Diagnostic', message, ui.ButtonSet.OK);
    
    console.log('Timezone diagnostic completed');
    console.log('Script timezone:', scriptTimezone);
    console.log('Script time:', scriptTime);
    console.log('Spreadsheet timezone:', spreadsheetTimezone);
    console.log('Spreadsheet time:', spreadsheetTime);
    
  } catch (error) {
    console.error('Timezone diagnostic failed:', error);
    SpreadsheetApp.getUi().alert(
      'Diagnostic Error',
      `❌ Could not check timezone settings: ${error.message}`,
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

function createWeeklyNotificationTriggerWithTimezone() {
  /**
   * Enhanced trigger creation that accounts for timezone settings
   */
  try {
    const ui = SpreadsheetApp.getUi();
    const sheet = SpreadsheetApp.getActiveSheet();
    
    // Show timezone info first
    const scriptTimezone = Session.getScriptTimeZone();
    const now = new Date();
    const currentTime = Utilities.formatDate(now, scriptTimezone, 'HH:mm z');
    
    // Get configuration
    const triggerConfig = getWeeklyTriggerConfig(sheet);
    
    if (!triggerConfig.isValid) {
      ui.alert(
        'Invalid Configuration',
        `❌ Please fix your trigger settings:\n\n${triggerConfig.errors.join('\n')}`,
        ui.ButtonSet.OK
      );
      return;
    }
    
    // Show timezone-aware confirmation
    const response = ui.alert(
      'Create Trigger with Timezone Info',
      `🕐 TIMEZONE AWARENESS:\n` +
      `Script timezone: ${scriptTimezone}\n` +
      `Current time: ${currentTime}\n\n` +
      `📅 Trigger will fire:\n` +
      `• Day: ${triggerConfig.dayName}\n` +
      `• Time: ${triggerConfig.timeFormatted} ${scriptTimezone}\n\n` +
      `⚠️ If this timezone doesn't match your location, the trigger will fire at the wrong time!\n\n` +
      `Continue anyway?`,
      ui.ButtonSet.YES_NO
    );
    
    if (response !== ui.Button.YES) {
      return;
    }
    
    // Remove existing triggers
    const existingTriggers = ScriptApp.getProjectTriggers()
      .filter(t => t.getHandlerFunction() === 'sendWeeklyVisitationChat');
    
    existingTriggers.forEach(trigger => {
      ScriptApp.deleteTrigger(trigger);
    });
    
    // Create new trigger
    const trigger = ScriptApp.newTrigger('sendWeeklyVisitationChat')
      .timeBased()
      .everyWeeks(1)
      .onWeekDay(triggerConfig.weekDay)
      .atHour(triggerConfig.hour)
      .create();
    
    // Store configuration
    const properties = PropertiesService.getScriptProperties();
    properties.setProperties({
      'WEEKLY_TRIGGER_ID': trigger.getUniqueId(),
      'WEEKLY_TRIGGER_DAY': triggerConfig.dayName,
      'WEEKLY_TRIGGER_HOUR': triggerConfig.hour.toString(),
      'WEEKLY_TRIGGER_TIMEZONE': scriptTimezone,
      'WEEKLY_TRIGGER_CREATED': new Date().toISOString()
    });
    
    ui.alert(
      '✅ Trigger Created with Timezone Info',
      `Weekly notifications scheduled!\n\n` +
      `📅 Every ${triggerConfig.dayName} at ${triggerConfig.timeFormatted}\n` +
      `🕐 Timezone: ${scriptTimezone}\n\n` +
      `🧪 Use "Test Notification Now" to verify the system works.\n` +
      `📋 Use "Check Timezone Settings" to diagnose timing issues.`,
      ui.ButtonSet.OK
    );
    
    console.log(`Created trigger: ${triggerConfig.dayName} at ${triggerConfig.hour}:00 ${scriptTimezone}`);
    
  } catch (error) {
    console.error('Failed to create timezone-aware trigger:', error);
    ui.alert(
      'Trigger Creation Failed',
      `❌ ${error.message}`,
      ui.ButtonSet.OK
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
// ===== END of TIMEZONE DIAGNOSTIC FUNCTIONS =====
