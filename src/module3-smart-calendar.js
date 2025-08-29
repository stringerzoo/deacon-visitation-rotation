/**
 * MODULE 3: SMART CALENDAR & TEST MODE DETECTION (v2.0)
 * Deacon Visitation Rotation System - Advanced Calendar Integration
 * 
 * This module contains:
 * - Smart calendar update functions
 * - Test mode detection and management
 * - Mode indicators and visual feedback
 * - Intelligent calendar synchronization
 * 
 * Key Features:
 * - Contact info only updates (preserves scheduling)
 * - Future events only updates (protects current week)
 * - Automatic test/production mode detection
 * - Visual mode indicators in spreadsheet
 */

/**
 * Smart test mode detection with user communication
 */
function detectAndCommunicateTestMode() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    
    // Detection logic
    const testIndicators = [
      {
        name: 'Test household names',
        check: () => {
          const households = sheet.getRange('M2:M10').getValues().flat().filter(cell => cell !== '');
          const testPatterns = ['Alan & Alexa Adams', 'Barbara Baker', 'Chloe Cooper', 'test', 'sample'];
          return households.some(household => 
            testPatterns.some(pattern => 
              String(household).toLowerCase().includes(pattern.toLowerCase())
            )
          );
        }
      },
      {
        name: 'Test phone numbers (555)',
        check: () => {
          const phones = sheet.getRange('N2:N10').getValues().flat().filter(cell => cell !== '');
          return phones.some(phone => String(phone).includes('555'));
        }
      },
      {
        name: 'Test Breeze numbers (12345)',
        check: () => {
          const breezeNumbers = sheet.getRange('P2:P10').getValues().flat().filter(cell => cell !== '');
          return breezeNumbers.some(number => String(number).startsWith('12345'));
        }
      },
      {
        name: 'Spreadsheet name contains "test"',
        check: () => {
          const spreadsheetName = SpreadsheetApp.getActiveSpreadsheet().getName().toLowerCase();
          return spreadsheetName.includes('test') || spreadsheetName.includes('sample');
        }
      }
    ];
    
    // Find which indicators triggered
    const triggeredIndicators = testIndicators.filter(indicator => indicator.check());
    const isTestMode = triggeredIndicators.length > 0;
    
    // Console logging
    if (isTestMode) {
      console.log('🧪 TEST MODE detected automatically');
      console.log('Triggers found:', triggeredIndicators.map(i => i.name).join(', '));
      console.log('Calendar will be: "TEST - Deacon Visitation Schedule"');
    } else {
      console.log('✅ PRODUCTION MODE detected');
      console.log('Calendar will be: "Deacon Visitation Schedule"');
    }
    
    // Store mode info for later display
    const properties = PropertiesService.getScriptProperties();
    properties.setProperty('DETECTED_MODE', isTestMode ? 'TEST' : 'PRODUCTION');
    properties.setProperty('MODE_TRIGGERS', triggeredIndicators.map(i => i.name).join(', '));
    
    return isTestMode;
    
  } catch (error) {
    console.warn('Could not detect test mode, defaulting to production:', error);
    PropertiesService.getScriptProperties().setProperty('DETECTED_MODE', 'PRODUCTION (default)');
    return false;
  }
}

/**
 * Enhanced show mode notification that refreshes detection first
 */
function showModeNotification() {
  // Refresh detection before showing notification
  refreshModeDetection();
  
  const properties = PropertiesService.getScriptProperties();
  const mode = properties.getProperty('DETECTED_MODE') || 'UNKNOWN';
  const triggers = properties.getProperty('MODE_TRIGGERS') || 'None';
  
  const isTestMode = mode.includes('TEST');
  
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    isTestMode ? '🧪 Test Mode Detected' : '✅ Production Mode Active',
    `${isTestMode ? '🧪 TEST MODE' : '✅ PRODUCTION MODE'} has been automatically detected.\n\n` +
    `Calendar: "${isTestMode ? 'TEST - ' : ''}Deacon Visitation Schedule"\n\n` +
    (isTestMode ? `Detected because: ${triggers}\n\n` : '') +
    `${isTestMode ? 
      '• Test calendar will be used for safety\n• Sample data detected - perfect for testing!\n• Switch to real member data to enable production mode' : 
      '• Live calendar will be used\n• Real member data detected\n• All calendar functions ready for pastoral care'
    }`,
    ui.ButtonSet.OK
  );
}

/**
 * Add mode indicator to spreadsheet
 */
function addModeIndicatorToSheet() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const properties = PropertiesService.getScriptProperties();
  const mode = properties.getProperty('DETECTED_MODE') || 'UNKNOWN';
  const isTestMode = mode.includes('TEST');
  
  // Add indicator in a visible but non-intrusive location (K16)
  const indicatorCell = sheet.getRange('K16');
  indicatorCell.setValue(isTestMode ? '🧪 TEST MODE' : '✅ PRODUCTION');
  indicatorCell.setBackground(isTestMode ? '#ffeb3b' : '#4caf50'); // Yellow for test, green for production
  indicatorCell.setFontColor(isTestMode ? '#333' : 'white');
  indicatorCell.setFontWeight('bold');
  indicatorCell.setHorizontalAlignment('center');
  
  // Add label above it
  const labelCell = sheet.getRange('K15');
  labelCell.setValue('Current Mode:');
  labelCell.setFontWeight('bold');
  labelCell.setBackground('#fff2cc');
}

// TEST MODE CONFIGURATION
const TEST_MODE = detectAndCommunicateTestMode();
const TEST_CALENDAR_NAME = TEST_MODE ? 'TEST - Deacon Visitation Schedule' : 'Deacon Visitation Schedule';

// ===== SMART CALENDAR UPDATE FUNCTIONS =====

function updateContactInfoOnly() {
  /**
   * Updates ONLY contact information in existing calendar events
   * Preserves: Custom times, dates, guests, locations, scheduling details
   * Updates: Phone numbers, addresses, Breeze links, Notes links, instructions
   * Enhanced smart calendar functions that check mode dynamically
   */
  const currentTestMode = getCurrentTestMode();
  const currentCalendarName = currentTestMode ? 'TEST - Deacon Visitation Schedule' : 'Deacon Visitation Schedule';
  
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const config = getConfiguration(sheet);
    const scheduleData = getScheduleFromSheet(sheet);
    
    if (scheduleData.length === 0) {
      SpreadsheetApp.getUi().alert('No Schedule', 'Please generate a schedule first.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    // Validate that schedule matches current data
    if (!validateScheduleDataSync()) {
      return; // validateScheduleDataSync() shows its own error dialog
    }

    // Confirm with user (showing current mode)
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      currentTestMode ? 'TEST: Update Contact Info Only' : 'Update Contact Info Only',
      `${currentTestMode ? '🧪 TEST MODE: ' : ''}This will update contact information in existing calendar events.\n\n` +
      `Calendar: "${currentCalendarName}"\n\n` +
      `✅ PRESERVES: Custom times, dates, guest lists, locations\n` +
      `🔄 UPDATES: Contact info, Breeze links, Notes links, instructions\n\n` +
      `💡 TIP: If you've updated Breeze numbers (column P) or Notes links (column Q),\n` +
      `run "🔗 Generate Shortened URLs" first for optimal results.\n\n` +
      `Continue?`,
      ui.ButtonSet.YES_NO
    );
    
    if (response !== ui.Button.YES) return;
    
    // Get calendar using current mode
    const calendar = getOrCreateCalendar(currentCalendarName);
    if (!calendar) return;
    
    // Get existing events
    const startDate = new Date();
    startDate.setFullYear(startDate.getFullYear() - 1);
    const endDate = new Date();
    endDate.setFullYear(endDate.getFullYear() + 2);
    
    const existingEvents = calendar.getEvents(startDate, endDate);
    console.log(`Found ${existingEvents.length} existing events to update`);
    
    let updatedCount = 0;
    const updateStartTime = new Date().getTime();
    
    existingEvents.forEach((event, index) => {
      try {
        // Extract household name from event title: "[Deacon] visits [Household]"
        const title = event.getTitle();
        const visitMatch = title.match(/(.+) visits (.+)/);
        
        if (visitMatch) {
          const household = visitMatch[2].trim();
          
          // Find household in current data
          const householdIndex = config.households.indexOf(household);
          
          if (householdIndex >= 0) {
            // Get current contact info
            const phone = config.phones[householdIndex] || 'Phone not available';
            const address = config.addresses[householdIndex] || 'Address not available';
            
            // Get Breeze link (use shortened if available)
            let breezeLink = 'Not available';
            const shortBreezeLink = config.breezeShortLinks[householdIndex];
            if (shortBreezeLink && shortBreezeLink.trim().length > 0) {
              breezeLink = shortBreezeLink;
            } else {
              const breezeNumber = config.breezeNumbers[householdIndex];
              if (breezeNumber && breezeNumber.trim().length > 0) {
                breezeLink = buildBreezeUrl(breezeNumber);
              }
            }
            
            // Get Notes link (use shortened if available)
            let notesLink = 'Not available';
            const shortNotesLink = config.notesShortLinks[householdIndex];
            if (shortNotesLink && shortNotesLink.trim().length > 0) {
              notesLink = shortNotesLink;
            } else {
              const fullNotesLink = config.notesLinks[householdIndex];
              if (fullNotesLink && fullNotesLink.trim().length > 0) {
                notesLink = fullNotesLink;
              }
            }
            
            // Create updated description
            const updatedDescription = `Household: ${household}
Breeze Profile: ${breezeLink}

Contact Information:
Phone: ${phone}
Address: ${address}

Visit Notes: ${notesLink}

Instructions:
${config.calendarInstructions}`;
            
            // Update event description (preserves all other details)
            event.setDescription(updatedDescription);
            updatedCount++;
            
            // Rate limiting - pause every 25 updates
            if ((index + 1) % 25 === 0) {
              Utilities.sleep(1000);
              console.log(`Updated ${index + 1} events so far...`);
            }
          }
        }
      } catch (eventError) {
        console.error(`Failed to update event: ${event.getTitle()}`, eventError);
      }
    });
    
    console.log(`Updated ${updatedCount} events in ${new Date().getTime() - updateStartTime}ms`);
    
    ui.alert(
      'Contact Info Update Complete',
      `✅ Updated contact information in ${updatedCount} calendar events!\n\n` +
      `📞 Refreshed: Phone numbers, addresses, Breeze links, Notes links\n` +
      `🔒 Preserved: All custom scheduling details, times, dates, guests\n\n` +
      `All events now have current contact information.`,
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    SpreadsheetApp.getUi().alert(
      'Contact Info Update Failed',
      `❌ ${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

function updateFutureEventsOnly() {
  /**
   * Updates events starting from next week, preserving current week scheduling
   * Useful for mid-week contact updates without affecting ongoing visits
   */
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const config = getConfiguration(sheet);
    const scheduleData = getScheduleFromSheet(sheet);
    
    if (scheduleData.length === 0) {
      SpreadsheetApp.getUi().alert('No Schedule', 'Please generate a schedule first.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    // Validate that schedule matches current data
    if (!validateScheduleDataSync()) {
      return; // validateScheduleDataSync() shows its own error dialog
    }
    
    // Calculate cutoff date (next week)
    const cutoffDate = new Date();
    cutoffDate.setDate(cutoffDate.getDate() + 7);
    cutoffDate.setHours(0, 0, 0, 0);
    
    // Confirm with user
    const currentTestMode = getCurrentTestMode();
    const currentCalendarName = currentTestMode ? 'TEST - Deacon Visitation Schedule' : 'Deacon Visitation Schedule';
    
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      currentTestMode ? 'TEST: Update Future Events Only' : 'Update Future Events Only',
      `${currentTestMode ? '🧪 TEST MODE: ' : ''}This will update calendar events starting from ${cutoffDate.toLocaleDateString()}.\n\n` +
      `Calendar: "${currentCalendarName}"\n\n` +      `✅ PRESERVES: This week's scheduling details and customizations\n` +
      `🔄 UPDATES: Future events with current contact info and assignments\n\n` +
      `💡 TIP: For contact changes only, use "📞 Update Contact Info Only" instead.\n\n` +
      `Continue?`,
      ui.ButtonSet.YES_NO
    );
    
    if (response !== ui.Button.YES) return;
    
    // Get calendar
    const calendar = getOrCreateCalendar(currentCalendarName);
    if (!calendar) return;
    
    // Get existing future events
    const farFutureDate = new Date();
    farFutureDate.setFullYear(farFutureDate.getFullYear() + 2);
    
    const existingEvents = calendar.getEvents(cutoffDate, farFutureDate);
    console.log(`Found ${existingEvents.length} future events to update`);
    
    // Delete future events
    let deletedCount = 0;
    const deleteStartTime = new Date().getTime();
    
    existingEvents.forEach((event, index) => {
      event.deleteEvent();
      deletedCount++;
      
      // Rate limiting for deletions
      if ((index + 1) % 10 === 0) {
        Utilities.sleep(500);
        console.log(`Deleted ${index + 1} future events...`);
      }
    });
    
    console.log(`Deleted ${deletedCount} future events in ${new Date().getTime() - deleteStartTime}ms`);
    
    // Cooldown before creating new events
    if (deletedCount > 0) {
      console.log('Waiting for API cooldown...');
      Utilities.sleep(2000);
    }
    
    // Filter schedule data for future events
    const futureScheduleData = scheduleData.filter(visit => {
      const visitDate = new Date(visit.date);
      return visitDate >= cutoffDate;
    });
    
    console.log(`Creating ${futureScheduleData.length} new future events...`);
    
    // Create new future events with current data
    let eventsCreated = 0;
    const createStartTime = new Date().getTime();
    
    futureScheduleData.forEach((visit, index) => {
      try {
        const eventTitle = `${visit.deacon} visits ${visit.household}`;
        
        // Get current contact info
        const householdIndex = config.households.indexOf(visit.household);
        const phone = householdIndex >= 0 ? config.phones[householdIndex] || 'Phone not available' : 'Phone not available';
        const address = householdIndex >= 0 ? config.addresses[householdIndex] || 'Address not available' : 'Address not available';
        
        // Get current Breeze link
        let breezeLink = 'Not available';
        if (householdIndex >= 0) {
          const shortBreezeLink = config.breezeShortLinks[householdIndex];
          if (shortBreezeLink && shortBreezeLink.trim().length > 0) {
            breezeLink = shortBreezeLink;
          } else {
            const breezeNumber = config.breezeNumbers[householdIndex];
            if (breezeNumber && breezeNumber.trim().length > 0) {
              breezeLink = buildBreezeUrl(breezeNumber);
            }
          }
        }
        
        // Get current Notes link
        let notesLink = 'Not available';
        if (householdIndex >= 0) {
          const shortNotesLink = config.notesShortLinks[householdIndex];
          if (shortNotesLink && shortNotesLink.trim().length > 0) {
            notesLink = shortNotesLink;
          } else {
            const fullNotesLink = config.notesLinks[householdIndex];
            if (fullNotesLink && fullNotesLink.trim().length > 0) {
              notesLink = fullNotesLink;
            }
          }
        }
        
        // Create event description
        const eventDescription = `Household: ${visit.household}
Breeze Profile: ${breezeLink}

Contact Information:
Phone: ${phone}
Address: ${address}

Visit Notes: ${notesLink}

Instructions:
${config.calendarInstructions}`;
        
        // Set event timing (default 2-3 PM)
        const startTime = new Date(visit.date);
        startTime.setHours(14, 0, 0, 0);
        const endTime = new Date(startTime);
        endTime.setHours(15, 0, 0, 0);
        
        calendar.createEvent(eventTitle, startTime, endTime, {
          description: eventDescription,
          guests: '',
          sendInvites: false
        });
        
        eventsCreated++;
        
        // Rate limiting for creation
        if ((index + 1) % 25 === 0) {
          Utilities.sleep(1000);
          console.log(`Created ${index + 1} future events so far...`);
        }
        
      } catch (eventError) {
        console.error(`Failed to create future event for ${visit.deacon} -> ${visit.household}:`, eventError);
      }
    });
    
    console.log(`Created ${eventsCreated} future events in ${new Date().getTime() - createStartTime}ms`);
    
    ui.alert(
      'Future Events Update Complete',
      `✅ Updated future calendar events starting ${cutoffDate.toLocaleDateString()}!\n\n` +
      `🗑️ Deleted: ${deletedCount} old future events\n` +
      `📅 Created: ${eventsCreated} updated future events\n` +
      `🔒 Preserved: This week's scheduling details\n\n` +
      `Future events now have current contact info and assignments.`,
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    SpreadsheetApp.getUi().alert(
      'Future Events Update Failed',
      `❌ ${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

// ===== HELPER FUNCTIONS =====

/**
 * Helper function for calendar access with dynamic mode
 */
function getOrCreateCalendar(calendarName = null) {
  // Use provided name or detect current mode
  if (!calendarName) {
    const currentTestMode = getCurrentTestMode();
    calendarName = currentTestMode ? 'TEST - Deacon Visitation Schedule' : 'Deacon Visitation Schedule';
  }
  
  try {
    const calendars = CalendarApp.getCalendarsByName(calendarName);
    if (calendars.length > 0) {
      return calendars[0];
    } else {
      SpreadsheetApp.getUi().alert(
        'Calendar Not Found',
        `Calendar "${calendarName}" not found.\n\nPlease run "📆 Export to Google Calendar" first to create the calendar.`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return null;
    }
  } catch (error) {
    throw new Error(`Calendar access failed: ${error.message}`);
  }
}

function buildBreezeUrl(breezeNumber) {
  /**
   * Constructs full Breeze URL from 8-digit number
   */
  if (!breezeNumber || breezeNumber.trim().length === 0) {
    return '';
  }
  
  const cleanNumber = breezeNumber.trim();
  return `https://immanuelky.breezechms.com/people/view/${cleanNumber}`;
}

/**
 * Refresh mode detection and update all indicators
 * Call this whenever you want to check/update the current mode
 */
function refreshModeDetection() {
  // Re-run detection
  const isTestMode = detectAndCommunicateTestMode();
  
  // Update the spreadsheet indicator immediately
  addModeIndicatorToSheet();
  
  // Update the global variables (for any functions that use them)
  // Note: This won't change the const declarations, but stores current state
  const properties = PropertiesService.getScriptProperties();
  properties.setProperty('CURRENT_TEST_MODE', isTestMode.toString());
  properties.setProperty('CURRENT_CALENDAR_NAME', isTestMode ? 'TEST - Deacon Visitation Schedule' : 'Deacon Visitation Schedule');
  
  return isTestMode;
}

/**
 * Get current test mode (use this instead of the const when you need fresh detection)
 */
function getCurrentTestMode() {
  return refreshModeDetection();
}

/**
 * v2.0 COMPATIBLE CALENDAR EXPORT FUNCTION
 * Updated to work with v2.0 data structures and variable frequencies
 */

function exportToGoogleCalendar() {
  /**
   * v2.0 COMPATIBLE: Creates calendar events for each deacon's visits with enhanced contact information
   * including shortened Breeze and Notes links, with variable frequency support
   */
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const config = getConfiguration(); // v2.0 doesn't need sheet parameter
    const scheduleData = getScheduleFromSheet();
    
    if (config.deacons.length === 0) {
      SpreadsheetApp.getUi().alert('Error', 'No deacons found. Please generate a schedule first.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    if (scheduleData.length === 0) {
      SpreadsheetApp.getUi().alert('No Schedule', 'Please generate a schedule first.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    // NEW v2.0: Detect test mode and set appropriate calendar name
    const testMode = detectAndCommunicateTestMode();
    const calendarName = testMode ? 'TEST - Deacon Visitation Schedule' : 'Deacon Visitation Schedule';
    
    console.log(`Calendar: "${calendarName}"`);
    console.log(`v2.0 Variable frequency support: ${config.hasCustomFrequencies ? 'ENABLED' : 'DISABLED'}`);
    
    // Check if shortened URLs need to be generated
    const hasShortUrls = config.breezeShortLinks.some(link => link && link.trim().length > 0) || 
                        config.notesShortLinks.some(link => link && link.trim().length > 0);
    
    if (!hasShortUrls) {
      const response = SpreadsheetApp.getUi().alert(
        'Generate Shortened URLs?',
        'No shortened URLs found in columns R and S.\n\n' +
        'Would you like to generate them now before creating calendar events?\n\n' +
        '(This will improve the calendar event descriptions)',
        SpreadsheetApp.getUi().ButtonSet.YES_NO_CANCEL
      );
      
      if (response === SpreadsheetApp.getUi().Button.CANCEL) {
        return;
      }
      
      if (response === SpreadsheetApp.getUi().Button.YES) {
        generateShortUrls(); // v2.0 function name
        // Reload config to get the newly generated short URLs
        const updatedConfig = getConfiguration();
        config.breezeShortLinks = updatedConfig.breezeShortLinks;
        config.notesShortLinks = updatedConfig.notesShortLinks;
      }
    }
    
    // Create or get the deacon visitation calendar
    let calendar;
    
    try {
      const calendars = CalendarApp.getCalendarsByName(calendarName);
      if (calendars.length > 0) {
        calendar = calendars[0];
        
        const response = SpreadsheetApp.getUi().alert(
          'Calendar Exists',
          `Calendar "${calendarName}" already exists. Clear existing events and add new ones?`,
          SpreadsheetApp.getUi().ButtonSet.YES_NO_CANCEL
        );
        
        if (response === SpreadsheetApp.getUi().Button.CANCEL) return;
        
        if (response === SpreadsheetApp.getUi().Button.YES) {
          const startDate = new Date();
          startDate.setFullYear(startDate.getFullYear() - 1);
          const endDate = new Date();
          endDate.setFullYear(endDate.getFullYear() + 2);
          
          const existingEvents = calendar.getEvents(startDate, endDate);
          console.log(`Deleting ${existingEvents.length} existing events...`);
          
          // Delete events in smaller batches to avoid rate limiting
          const deleteStartTime = new Date().getTime();
          let deletedCount = 0;
          
          existingEvents.forEach((event, index) => {
            event.deleteEvent();
            deletedCount++;
            
            // Add small delay every 10 deletions
            if ((index + 1) % 10 === 0) {
              Utilities.sleep(500); // 0.5 second pause
            }
          });
          
          console.log(`Deleted ${deletedCount} events in ${new Date().getTime() - deleteStartTime}ms`);
          
          // Wait a bit longer before creating new events
          console.log('Waiting for API cooldown before creating new events...');
          Utilities.sleep(2000); // 2 second pause before creating new events
        }
      } else {
        calendar = CalendarApp.createCalendar(calendarName);
        const calendarDescription = testMode ? 
          'TEST CALENDAR - Automated schedule for deacon household visitations with contact information and management links' :
          'Automated schedule for deacon household visitations with contact information and management links';
        calendar.setDescription(calendarDescription);
        calendar.setColor(testMode ? CalendarApp.Color.RED : CalendarApp.Color.BLUE);
      }
    } catch (calError) {
      throw new Error(`Calendar access failed: ${calError.message}. Make sure you have calendar permissions.`);
    }
    
    // Create enhanced events with contact information and links
    let eventsCreated = 0;
    const maxEvents = 500;
    
    console.log(`Creating ${Math.min(scheduleData.length, maxEvents)} calendar events...`);
    const createStartTime = new Date().getTime();
    
    scheduleData.slice(0, maxEvents).forEach((visit, index) => {
      try {
        // Enhanced event title format with test mode indicator
        let eventTitle = `${visit.deacon} visits ${visit.household}`;
        if (testMode) {
          eventTitle = `TEST: ${eventTitle}`;
        }
        
        // Find household index to get all contact info and links
        const householdIndex = config.households.indexOf(visit.household);
        const phone = householdIndex >= 0 ? config.phones[householdIndex] || 'Phone not available' : 'Phone not available';
        const address = householdIndex >= 0 ? config.addresses[householdIndex] || 'Address not available' : 'Address not available';
        
        // Get Breeze link (use shortened if available, otherwise build from number)
        let breezeLink = 'Not available';
        if (householdIndex >= 0) {
          const shortBreezeLink = config.breezeShortLinks[householdIndex];
          if (shortBreezeLink && shortBreezeLink.trim().length > 0) {
            breezeLink = shortBreezeLink;
          } else {
            // Check for both possible property names for backward compatibility
            const breezeNumber = config.breezeLinks[householdIndex] || config.breezeNumbers[householdIndex];
            if (breezeNumber && String(breezeNumber).trim().length > 0) {
              breezeLink = `https://tinyurl.com/breeze${breezeNumber}`; // Fallback format if buildBreezeUrl missing
            }
          }
        }
        
        // Get Notes link (use shortened if available)
        let notesLink = 'Not available';
        if (householdIndex >= 0) {
          const shortNotesLink = config.notesShortLinks[householdIndex];
          if (shortNotesLink && shortNotesLink.trim().length > 0) {
            notesLink = shortNotesLink;
          } else {
            const originalNotesLink = config.notesLinks[householdIndex];
            if (originalNotesLink && originalNotesLink.trim().length > 0) {
              notesLink = originalNotesLink;
            }
          }
        }
        
        // Build enhanced event description with v2.0 frequency information
        let eventDescription = `Household: ${visit.household}\n`;
        
        // Add frequency information for v2.0
        if (config.hasCustomFrequencies) {
          const householdFreq = config.householdFrequencies.find(hf => hf.household === visit.household);
          if (householdFreq) {
            const freqLabel = householdFreq.isCustom ? 
              `${householdFreq.frequency} weeks (custom)` : 
              `${householdFreq.frequency} weeks (default)`;
            eventDescription += `Visit Frequency: ${freqLabel}\n`;
          }
        }
        
        eventDescription += `Breeze Profile: ${breezeLink}\n\n`;
        eventDescription += `Contact Information:\n`;
        eventDescription += `Phone: ${phone}\n`;
        eventDescription += `Address: ${address}\n\n`;
        eventDescription += `Visit Notes: ${notesLink}\n\n`;
        eventDescription += `Instructions:\n`;
        eventDescription += config.calendarInstructions || 'Please call to confirm visit time. Contact family 1-2 days before scheduled date to arrange convenient time.';
        
        // Add test mode warning if applicable
        if (testMode) {
          eventDescription = `TEST EVENT - This is for testing purposes only.\n\n` + eventDescription;
        }
        
        // Create the calendar event
        const eventStart = new Date(visit.date);
        eventStart.setHours(14, 0, 0, 0); // 2:00 PM default
        
        const eventEnd = new Date(visit.date);
        eventEnd.setHours(15, 0, 0, 0); // 3:00 PM default (1 hour duration)
        
        const event = calendar.createEvent(eventTitle, eventStart, eventEnd, {
          description: eventDescription
        });
        
        eventsCreated++;
        
        // Add small delay every 25 events to avoid rate limiting
        if ((index + 1) % 25 === 0) {
          Utilities.sleep(1000); // 1 second pause
          console.log(`Created ${index + 1} of ${scheduleData.length} events...`);
        }
        
      } catch (eventError) {
        console.error(`Failed to create event for ${visit.deacon} -> ${visit.household}:`, eventError);
      }
    });
    
    const totalTime = new Date().getTime() - createStartTime;
    console.log(`Created ${eventsCreated} calendar events in ${totalTime}ms`);
    
    // Show success message
    const successMessage = `Calendar Export Complete!\n\n` +
      `Calendar: "${calendarName}"\n` +
      `Events Created: ${eventsCreated}\n` +
      `Time: ${(totalTime/1000).toFixed(1)} seconds\n\n` +
      `${config.hasCustomFrequencies ? 'Variable frequency households marked with custom frequency info' : 'Uniform frequency schedule'}`;
    
    SpreadsheetApp.getUi().alert('Export Complete', successMessage, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    console.error('Calendar export failed:', error);
    SpreadsheetApp.getUi().alert(
      'Export Failed',
      `Could not export to calendar:\n\n${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

function getScheduleFromSheet() {
  const sheet = SpreadsheetApp.getActiveSheet();
  
  const data = sheet.getRange('A2:E1000').getValues()
    .filter(row => {
      return row[0] !== '' && row[0] != null &&  // Column A: Cycle
             row[1] !== '' && row[1] != null &&  // Column B: Week  
             row[3] !== '' && row[3] != null &&  // Column D: Household
             row[4] !== '' && row[4] != null &&  // Column E: Deacon
             typeof row[0] === 'number' &&       // Cycle should be a number
             typeof row[1] === 'number';         // Week should be a number
    });
  
  console.log(`Reading ${data.length} valid schedule entries from spreadsheet`);
  
  return data.map(row => ({
    cycle: row[0],
    week: row[1],
    date: row[2] ? new Date(row[2]) : new Date(), // Handle date parsing with fallback
    household: row[3],
    deacon: row[4]
  }));
}

function getOrCreateCalendar() {
  /**
   * Helper function for other modules to get the correct calendar
   */
  const testMode = detectAndCommunicateTestMode();
  const calendarName = testMode ? 'TEST - Deacon Visitation Schedule' : 'Deacon Visitation Schedule';
  
  const calendars = CalendarApp.getCalendarsByName(calendarName);
  if (calendars.length > 0) {
    return calendars[0];
  } else {
    const calendar = CalendarApp.createCalendar(calendarName);
    const calendarDescription = testMode ? 
      'TEST CALENDAR - Automated schedule for deacon household visitations' :
      'Automated schedule for deacon household visitations';
    calendar.setDescription(calendarDescription);
    calendar.setColor(testMode ? CalendarApp.Color.RED : CalendarApp.Color.BLUE);
    return calendar;
  }
}
