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
    
    // Get current test mode from Module 3
    const currentTestMode = getCurrentTestMode();
    const chatPrefix = currentTestMode ? '🧪 TEST: ' : '';
    
    // Get upcoming visits (next 7-14 days)
    const upcomingVisits = getUpcomingVisits(7, 14);
    
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
   * Useful for day-before preparation
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
    
    const currentTestMode = getCurrentTestMode();
    const chatPrefix = currentTestMode ? '🧪 TEST: ' : '';
    
    // Get visits for tomorrow (1-2 days out)
    const tomorrowVisits = getUpcomingVisits(1, 2);
    
    if (tomorrowVisits.length === 0) {
      const message = `${chatPrefix}⏰ **No visits scheduled for tomorrow**\n\nEnjoy your day! 😊`;
      sendToChatSpace(message, currentTestMode);
      return;
    }
    
    // Build tomorrow's reminder message
    const reminderMessage = buildTomorrowMessage(tomorrowVisits, currentTestMode);
    sendToChatSpace(reminderMessage, currentTestMode);
    
    const ui = SpreadsheetApp.getUi();
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
   */
  const chatPrefix = isTestMode ? '🧪 TEST: ' : '';
  const weekOf = visits[0]?.date?.toLocaleDateString() || 'Unknown';
  
  let message = `${chatPrefix}📅 **Deacon Visits - Week of ${weekOf}**\n\n`;
  
  // Group visits by deacon for better organization
  const visitsByDeacon = groupVisitsByDeacon(visits);
  
  Object.keys(visitsByDeacon).sort().forEach(deaconName => {
    const deaconVisits = visitsByDeacon[deaconName];
    message += `👤 **${deaconName}**:\n`;
    
    deaconVisits.forEach(visit => {
      message += `   • ${visit.household}\n`;
      message += `     📞 ${visit.phone || 'Phone not available'}\n`;
      
      if (visit.breezeLink) {
        message += `     🔗 <${visit.breezeLink}|Breeze Profile>\n`;
      }
      
      if (visit.notesLink) {
        message += `     📝 <${visit.notesLink}|Visit Notes>\n`;
      }
      
      message += `\n`;
    });
  });
  
  message += `💡 *Remember to call 1-2 days before to confirm timing*\n`;
  message += `📱 *All links work on mobile for easy field access*`;
  
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
    
    // Enhance with contact information
    const enhancedVisits = upcomingVisits.map(visit => {
      const householdIndex = config.households.indexOf(visit.household);
      
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

// ===== HELPER FUNCTIONS =====

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
   * Create automated weekly trigger for notifications
   * Run this once to set up automatic weekly summaries
   */
  try {
    // Delete existing weekly triggers to avoid duplicates
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'sendWeeklyVisitationChat') {
        ScriptApp.deleteTrigger(trigger);
      }
    });
    
    // Create new weekly trigger (Sundays at 6 PM)
    ScriptApp.newTrigger('sendWeeklyVisitationChat')
      .timeBased()
      .everyWeeks(1)
      .onWeekDay(ScriptApp.WeekDay.SUNDAY)
      .atHour(18) // 6 PM
      .create();
    
    SpreadsheetApp.getUi().alert(
      'Weekly Trigger Created',
      '✅ Automatic weekly notifications are now enabled!\n\n' +
      '📅 Schedule: Every Sunday at 6:00 PM\n' +
      '📝 Function: Weekly visitation summary\n\n' +
      'You can disable this later by deleting the trigger.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    console.log('Weekly notification trigger created successfully');
    
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

// END OF MODULE 5
