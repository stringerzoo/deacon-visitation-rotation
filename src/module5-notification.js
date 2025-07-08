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
   * Main function to send weekly visitation summary to Google Chat
   * Respects test mode and uses appropriate chat space
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
      return; // validateScheduleDataSync() shows its own error dialog
    }
    
    // Get current test mode from Module 3
    const currentTestMode = getCurrentTestMode();
    const chatPrefix = currentTestMode ? '🧪 TEST: ' : '';
    
    // Get upcoming visits (next 2-3 weeks for bi-weekly rhythm)
    const upcomingVisits = getUpcomingVisits(7, 21);
    
    if (upcomingVisits.length === 0) {
      const message = `${chatPrefix}📅 **No scheduled visits for next week**\n\nAll caught up on visitations! 🎉`;
      
      sendToChatSpace(message, currentTestMode);
      console.log('No upcoming visits found for weekly summary');
      return;
    }
    
    // Build and send the weekly message
    const chatMessage = buildWeeklyMessage(upcomingVisits, currentTestMode);
    sendToChatSpace(chatMessage, currentTestMode);
    
    // Show success notification
    const ui = SpreadsheetApp.getUi();
    ui.alert(
      currentTestMode ? '🧪 Test Chat Summary Sent' : '📅 Weekly Chat Summary Sent',
      `${chatPrefix}Sent summary for ${upcomingVisits.length} upcoming visits to Google Chat.\n\n` +
      `Chat space: ${currentTestMode ? 'TEST space' : 'Production space'}\n` +
      `Week of: ${upcomingVisits[0]?.date?.toLocaleDateString() || 'Unknown'}`,
      ui.ButtonSet.OK
    );
    
    console.log(`Weekly chat summary sent: ${upcomingVisits.length} visits`);
    
  } catch (error) {
    console.error('Failed to send weekly chat summary:', error);
    SpreadsheetApp.getUi().alert(
      'Chat Notification Failed',
      `❌ Error sending to Google Chat: ${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

function sendTomorrowReminders() {
  /**
   * Send reminders for visits happening tomorrow
   * Asks user confirmation if no visits are scheduled
   */
  try {
    if (!isNotificationConfigured()) {
      SpreadsheetApp.getUi().alert(
        'Notifications Not Configured',
        'Please configure the Google Chat webhook first.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    // Validate that schedule matches current data
    if (!validateScheduleDataSync()) {
      return; // validateScheduleDataSync() shows its own error dialog
    }
    
    const currentTestMode = getCurrentTestMode();
    const chatPrefix = currentTestMode ? '🧪 TEST: ' : '';
    const ui = SpreadsheetApp.getUi();
    
    // Get visits for tomorrow (1-2 days out)
    const tomorrowVisits = getUpcomingVisits(1, 2);
    
    if (tomorrowVisits.length === 0) {
      // Ask user if they want to send "no visits" message
      const response = ui.alert(
        'No Visits Tomorrow',
        'There are no visits scheduled for tomorrow.\n\n' +
        'Do you still want to send a notification to the group to communicate this?',
        ui.ButtonSet.YES_NO
      );
      
      if (response === ui.Button.YES) {
        const message = `${chatPrefix}⏰ No visits scheduled for tomorrow\n\nEnjoy your day! 😊`;
        sendToChatSpace(message, currentTestMode);
        
        ui.alert(
          'No-Visits Notification Sent',
          `${chatPrefix}Sent "no visits tomorrow" message to chat.`,
          ui.ButtonSet.OK
        );
      } else {
        console.log('User chose not to send no-visits notification');
      }
      return;
    }
    
    // Build tomorrow's reminder message
    const reminderMessage = buildTomorrowMessage(tomorrowVisits, currentTestMode);
    sendToChatSpace(reminderMessage, currentTestMode);
    
    ui.alert(
      currentTestMode ? '🧪 Test Tomorrow Reminders Sent' : '⏰ Tomorrow Reminders Sent',
      `${chatPrefix}Sent reminders for ${tomorrowVisits.length} visits happening tomorrow.`,
      ui.ButtonSet.OK
    );
    
    console.log(`Tomorrow reminders sent: ${tomorrowVisits.length} visits`);
    
  } catch (error) {
    console.error('Failed to send tomorrow reminders:', error);
    SpreadsheetApp.getUi().alert(
      'Reminder Failed',
      `❌ Error sending reminders: ${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

// ===== MESSAGE BUILDING FUNCTIONS =====

function buildWeeklyMessage(visits, isTestMode = false) {
  /**
   * Build rich weekly summary message for Google Chat
   * Uses 2-week lookahead to match bi-weekly visitation rhythm
   */
  const chatPrefix = isTestMode ? '🧪 TEST: ' : '';
  
  // Group visits by week for better organization
  const visitsByWeek = {};
  visits.forEach(visit => {
    const weekKey = visit.date.toLocaleDateString();
    if (!visitsByWeek[weekKey]) {
      visitsByWeek[weekKey] = [];
    }
    visitsByWeek[weekKey].push(visit);
  });
  
  const weekKeys = Object.keys(visitsByWeek).sort((a, b) => new Date(a) - new Date(b));
  
  let message = `${chatPrefix}📅 Deacon Visitation Schedule - Next 2 Weeks\n\n`;
  
  weekKeys.forEach((weekKey, index) => {
    const weekVisits = visitsByWeek[weekKey];
    const weekLabel = index === 0 ? 'This Week' : index === 1 ? 'Next Week' : `Week of ${weekKey}`;
    
    message += `📆 ${weekLabel} (${weekKey}):\n`;
    
    // Group by deacon for this week
    const deaconGroups = groupVisitsByDeacon(weekVisits);
    Object.keys(deaconGroups).sort().forEach(deaconName => {
      const deaconVisits = deaconGroups[deaconName];
      
      deaconVisits.forEach(visit => {
        message += `   👤 ${visit.deacon} → ${visit.household}\n`;
        message += `      📞 ${visit.phone || 'Phone not available'}\n`;
        
        if (visit.breezeLink) {
          message += `      🔗 <${visit.breezeLink}|Breeze Profile>\n`;
        }
        
        if (visit.notesLink) {
          message += `      📝 <${visit.notesLink}|Visit Notes>\n`;
        }
      });
    });
    
    message += `\n`;
  });
  
  message += `📋 Contact families 1-2 days before your week to schedule\n`;
  message += `📅 Update the shared calendar with your confirmed time\n`;
  message += `📝 Document visit in notes page after completing`;
  
  return message;
}

function buildTomorrowMessage(visits, isTestMode = false) {
  /**
   * Build reminder message for tomorrow's visits
   */
  const chatPrefix = isTestMode ? '🧪 TEST: ' : '';
  const tomorrowDate = new Date();
  tomorrowDate.setDate(tomorrowDate.getDate() + 1);
  
  let message = `${chatPrefix}⏰ **Reminder: Visits Tomorrow (${tomorrowDate.toLocaleDateString()})**\n\n`;
  
  visits.forEach(visit => {
    message += `👤 **${visit.deacon}** → ${visit.household}\n`;
    message += `📞 ${visit.phone || 'Phone not available'}\n`;
    
    if (visit.breezeLink) {
      message += `🔗 <${visit.breezeLink}|Breeze Profile>\n`;
    }
    
    message += `\n`;
  });
  
  message += `🕐 *Don't forget to confirm your visit time!*`;
  
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

function getUpcomingVisits(startDays, endDays) {
  /**
   * Get visits within specified day range
   * startDays: how many days from now to start looking
   * endDays: how many days from now to stop looking
   */
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const config = getConfiguration(sheet);
    const scheduleData = getScheduleFromSheet(sheet);
    
    if (scheduleData.length === 0) {
      console.log('No schedule data found');
      return [];
    }
    
    // Debug logging
    console.log(`Households in config: [${config.households.join(', ')}]`);
    console.log(`Phone numbers count: ${config.phones.length}`);
    console.log(`Phones array: [${config.phones.join(', ')}]`);
    
    // Calculate date range
    const today = new Date();
    const startDate = new Date();
    startDate.setDate(today.getDate() + startDays);
    startDate.setHours(0, 0, 0, 0);
    
    const endDate = new Date();
    endDate.setDate(today.getDate() + endDays);
    endDate.setHours(23, 59, 59, 999);
    
    // Filter visits within date range
    const upcomingVisits = scheduleData.filter(visit => {
      const visitDate = new Date(visit.date);
      return visitDate >= startDate && visitDate <= endDate;
    });
    
    // Enhance with contact information and debug each household
    const enhancedVisits = upcomingVisits.map(visit => {
      const householdIndex = config.households.indexOf(visit.household);
      
      // Debug logging for each visit
      console.log(`Processing visit: ${visit.household}, Index: ${householdIndex}`);
      if (householdIndex >= 0) {
        console.log(`  Phone: "${config.phones[householdIndex]}"`);
        console.log(`  Address: "${config.addresses[householdIndex]}"`);
      } else {
        console.log(`  ERROR: Household "${visit.household}" not found in config.households`);
      }
      
      return {
        ...visit,
        phone: householdIndex >= 0 ? config.phones[householdIndex] : '',
        address: householdIndex >= 0 ? config.addresses[householdIndex] : '',
        breezeLink: getBreezeLink(config, householdIndex),
        notesLink: getNotesLink(config, householdIndex)
      };
    });
    
    console.log(`Found ${enhancedVisits.length} visits between ${startDate.toLocaleDateString()} and ${endDate.toLocaleDateString()}`);
    return enhancedVisits.sort((a, b) => new Date(a.date) - new Date(b.date));
    
  } catch (error) {
    console.error('Error getting upcoming visits:', error);
    return [];
  }
}

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
   * Create automated weekly trigger using configuration from spreadsheet cells
   */
  try {
    const ui = SpreadsheetApp.getUi();
    const sheet = SpreadsheetApp.getActiveSheet();
    
    // Get configuration from spreadsheet cells
    const triggerConfig = getWeeklyTriggerConfig(sheet);
    
    if (!triggerConfig.isValid) {
      ui.alert(
        'Invalid Trigger Configuration',
        `❌ Please check your trigger settings in column K:\n\n` +
        `${triggerConfig.errors.join('\n')}\n\n` +
        `Valid examples:\n` +
        `• Day: Sunday, Monday, Friday, etc.\n` +
        `• Time: 6, 18, 20 (hour in 24-format)`,
        ui.ButtonSet.OK
      );
      return;
    }
    
    // Show what will be configured and confirm
    const response = ui.alert(
      'Configure Weekly Auto-Send',
      `Ready to set up automatic notifications:\n\n` +
      `📅 Day: ${triggerConfig.dayName}\n` +
      `🕐 Time: ${triggerConfig.timeFormatted}\n` +
      `💬 Chat: ${getCurrentTestMode() ? 'Test space' : 'Production space'}\n\n` +
      `Create this weekly trigger?`,
      ui.ButtonSet.YES_NO
    );
    
    if (response !== ui.Button.YES) {
      return;
    }
    
    // Delete existing weekly triggers to avoid duplicates
    removeWeeklyNotificationTrigger();
    
    // Create new trigger with spreadsheet configuration
    ScriptApp.newTrigger('sendWeeklyVisitationChat')
      .timeBased()
      .everyWeeks(1)
      .onWeekDay(triggerConfig.weekDay)
      .atHour(triggerConfig.hour)
      .create();
    
    ui.alert(
      'Weekly Auto-Send Enabled',
      `✅ Automatic weekly notifications are now active!\n\n` +
      `📅 Schedule: Every ${triggerConfig.dayName} at ${triggerConfig.timeFormatted}\n` +
      `📝 Function: 2-week visitation lookahead\n\n` +
      `To modify: Update cells K14 and K16, then run this again\n` +
      `To disable: Use "🛑 Disable Weekly Auto-Send"`,
      ui.ButtonSet.OK
    );
    
    console.log(`Weekly notification trigger created: ${triggerConfig.dayName} at ${triggerConfig.timeFormatted}`);
    
  } catch (error) {
    console.error('Failed to create weekly trigger:', error);
    SpreadsheetApp.getUi().alert(
      'Trigger Creation Failed',
      `❌ Could not create weekly trigger: ${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

function removeWeeklyNotificationTrigger() {
  /**
   * Remove automated weekly trigger
   */
  try {
    const triggers = ScriptApp.getProjectTriggers();
    let deletedCount = 0;
    
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'sendWeeklyVisitationChat') {
        ScriptApp.deleteTrigger(trigger);
        deletedCount++;
      }
    });
    
    if (deletedCount > 0) {
      SpreadsheetApp.getUi().alert(
        'Weekly Trigger Removed',
        `✅ Removed ${deletedCount} weekly notification trigger(s).\n\n` +
        'Automatic weekly summaries are now disabled.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      console.log(`Removed ${deletedCount} weekly notification triggers`);
    } else {
      SpreadsheetApp.getUi().alert(
        'No Triggers Found',
        'No weekly notification triggers were found to remove.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    
  } catch (error) {
    console.error('Failed to remove triggers:', error);
    SpreadsheetApp.getUi().alert(
      'Trigger Removal Failed',
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
    // Read configuration values from K14 and K16 (K11-K12 are used for test mode indicators)
    const dayValue = sheet.getRange('K14').getValue();
    const timeValue = sheet.getRange('K16').getValue();
    
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
      message += `💬 Destination: ${getCurrentTestMode() ? 'Test chat space' : 'Production chat space'}\n\n`;
    } else {
      message += '❌ Status: INACTIVE\n';
      message += '📅 Running: No automatic notifications\n\n';
    }
    
    // Show spreadsheet configuration
    message += '⚙️ Spreadsheet Configuration:\n';
    if (triggerConfig.isValid) {
      message += `📅 Day (K14): ${triggerConfig.dayName}\n`;
      message += `🕐 Time (K16): ${triggerConfig.timeFormatted}\n\n`;
      
      if (!existingTrigger) {
        message += '💡 Ready to enable: Use "🔄 Enable Weekly Auto-Send"';
      } else {
        message += '💡 To modify: Update K14/K16, then use "🔄 Enable Weekly Auto-Send"';
      }
    } else {
      message += '❌ Configuration Issues:\n';
      triggerConfig.errors.forEach(error => {
        message += `   • ${error}\n`;
      });
      message += '\n💡 Fix the issues above, then use "🔄 Enable Weekly Auto-Send"';
    }
    
    ui.alert(
      'Weekly Auto-Send Status',
      message,
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    console.error('Error checking trigger schedule:', error);
    SpreadsheetApp.getUi().alert(
      'Error',
      `❌ Could not check trigger schedule: ${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}
