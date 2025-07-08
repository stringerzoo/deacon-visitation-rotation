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
