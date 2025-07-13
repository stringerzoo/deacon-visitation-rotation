/**
 * MODULE 4: CALENDAR EXPORT, MENU SYSTEM, AND UTILITIES (v1.1 - PROPERLY FIXED)
 * Deacon Visitation Rotation System - Calendar Integration & Menu Management
 * 
 * PROPERLY FIXED VERSION: Eliminated infinite loop WITHOUT caching
 * - ✅ Fixed: Infinite recursion by removing circular calls
 * - ✅ Maintained: Dynamic test mode detection based on current data
 * - ✅ Preserved: All v1.1 feature improvements
 * 
 * Key Fix: getCurrentTestMode() now does direct detection without calling other functions
 * that might call it back, while preserving the responsive nature of mode detection.
 */

// ===== GOOGLE CALENDAR EXPORT FUNCTIONS =====

function exportToGoogleCalendar() {
  /**
   * Full calendar regeneration with enhanced warnings
   */
  try {
    const currentTestMode = getCurrentTestMode();
    const sheet = SpreadsheetApp.getActiveSheet();
    const scheduleData = getScheduleFromSheet(sheet);
    
    if (scheduleData.length === 0) {
      SpreadsheetApp.getUi().alert(
        'No Schedule Found',
        'Please generate a schedule first using "📅 Generate Schedule".',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    let calendar;
    const calendarName = currentTestMode ? 
      'TEST - Deacon Visitation Schedule' : 'Deacon Visitation Schedule';
    
    try {
      const calendars = CalendarApp.getCalendarsByName(calendarName);
      if (calendars.length > 0) {
        calendar = calendars[0];
        
        const response = SpreadsheetApp.getUi().alert(
          currentTestMode ? 'TEST: Full Calendar Regeneration' : 'Full Calendar Regeneration',
          `${currentTestMode ? '🧪 TEST: ' : ''}⚠️ This will DELETE ALL existing events in "${calendarName}" and recreate them.\n\n` +
          '🛡️ Deacon scheduling customizations (times, guests, locations) will be LOST.\n\n' +
          '💡 Consider "📞 Update Contact Info Only" or "🔄 Update Future Events Only" instead.\n\n' +
          'Continue with full regeneration?',
          SpreadsheetApp.getUi().ButtonSet.YES_NO
        );
        
        if (response !== SpreadsheetApp.getUi().Button.YES) {
          return;
        }
        
        // Delete all existing events
        const events = calendar.getEvents(new Date('2020-01-01'), new Date('2030-12-31'));
        events.forEach(event => event.deleteEvent());
        console.log(`Deleted ${events.length} existing events`);
        
      } else {
        calendar = CalendarApp.createCalendar(calendarName);
        console.log(`Created new calendar: ${calendarName}`);
      }
    } catch (calendarError) {
      throw new Error(`Calendar setup failed: ${calendarError.message}`);
    }
    
    // Export all schedule data to calendar
    const config = getConfiguration(sheet);
    
    let eventsCreated = 0;
    scheduleData.forEach(([dateStr, deacon, household]) => {
      try {
        const visitDate = new Date(dateStr);
        const householdIndex = config.households.indexOf(household);
        
        // Build event description with contact info and links
        let description = `Household: ${household}\n`;
        
        if (householdIndex !== -1) {
          const phone = config.phones[householdIndex];
          const address = config.addresses[householdIndex];
          const breezeNumber = config.breezeNumbers[householdIndex];
          const notesLink = config.notesLinks[householdIndex];
          
          description += `\nContact Information:\n`;
          if (phone) description += `Phone: ${phone}\n`;
          if (address) description += `Address: ${address}\n`;
          
          if (breezeNumber) {
            const breezeUrl = buildBreezeUrl(breezeNumber);
            description += `\nBreeze Profile: ${breezeUrl}\n`;
          }
          
          if (notesLink) {
            description += `\nVisit Notes: ${notesLink}\n`;
          }
        }
        
        // Add custom instructions
        const customInstructions = config.customInstructions || 
          'Please call to set up a day and time for your visit in the coming week. ' +
          'You can update this event with the actual day and time and copy it to your personal calendar.';
        
        description += `\nInstructions:\n${customInstructions}`;
        
        // Create calendar event
        const event = calendar.createEvent(
          `${deacon} visits ${household}`,
          new Date(visitDate.getFullYear(), visitDate.getMonth(), visitDate.getDate(), 14, 0), // 2:00 PM
          new Date(visitDate.getFullYear(), visitDate.getMonth(), visitDate.getDate(), 15, 0), // 3:00 PM
          {
            description: description,
            location: config.addresses[householdIndex] || ''
          }
        );
        
        eventsCreated++;
        
      } catch (eventError) {
        console.error(`Failed to create event for ${deacon} -> ${household}:`, eventError);
      }
    });
    
    SpreadsheetApp.getUi().alert(
      currentTestMode ? 'TEST Calendar Export Complete' : 'Calendar Export Complete',
      `${currentTestMode ? '🧪 TEST: ' : ''}✅ Calendar export completed!\n\n` +
      `📅 Calendar: "${calendarName}"\n` +
      `📊 Events created: ${eventsCreated}\n` +
      `📱 View calendar: ${calendar.getDescription() ? calendar.getDescription() : 'Check Google Calendar'}\n\n` +
      'Events are scheduled for 2:00-3:00 PM by default. Deacons can adjust times as needed.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    console.error('Calendar export failed:', error);
    SpreadsheetApp.getUi().alert(
      'Export Failed',
      `❌ Calendar export failed: ${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

function updateFutureEventsOnly() {
  /**
   * Updates ONLY future events (next Monday onwards)
   * Preserves: This week's scheduling details completely
   * Updates: All future weeks with current assignments and contact info
   */
  try {
    const currentTestMode = getCurrentTestMode();
    const calendarName = currentTestMode ? 
      'TEST - Deacon Visitation Schedule' : 'Deacon Visitation Schedule';
    
    const sheet = SpreadsheetApp.getActiveSheet();
    const scheduleData = getScheduleFromSheet(sheet);
    
    if (scheduleData.length === 0) {
      SpreadsheetApp.getUi().alert('No Schedule', 'Please generate a schedule first.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    // Validate that schedule matches current data
    if (!validateScheduleDataSync()) {
      return; // validateScheduleDataSync() shows its own error dialog
    }

    const calendar = getOrCreateCalendar(calendarName);
    if (!calendar) return;

    // Confirm with user (showing current mode)
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      currentTestMode ? 'TEST: Update Future Events Only' : 'Update Future Events Only',
      `${currentTestMode ? '🧪 TEST: ' : ''}🔄 This will update future calendar events (starting next Monday) with current assignments.\n\n` +
      '🛡️ This week\'s events will be preserved completely\n' +
      '📞 Contact information will be updated\n' +
      '👥 Deacon assignments will reflect current schedule\n\n' +
      'Continue?',
      ui.ButtonSet.YES_NO
    );
    
    if (response !== ui.Button.YES) {
      return;
    }

    // Calculate cutoff date (next Monday)
    const today = new Date();
    const nextMonday = new Date(today);
    const daysUntilMonday = (8 - today.getDay()) % 7;
    nextMonday.setDate(today.getDate() + daysUntilMonday);
    nextMonday.setHours(0, 0, 0, 0);

    console.log(`Updating events from ${nextMonday.toDateString()} onwards`);

    // Get future events
    const endDate = new Date();
    endDate.setFullYear(endDate.getFullYear() + 2); // Look ahead 2 years
    const futureEvents = calendar.getEvents(nextMonday, endDate);

    console.log(`Found ${futureEvents.length} future events to update`);

    // Delete existing future events
    let deletedCount = 0;
    futureEvents.forEach(event => {
      try {
        event.deleteEvent();
        deletedCount++;
      } catch (deleteError) {
        console.error('Failed to delete event:', deleteError);
      }
    });

    console.log(`Deleted ${deletedCount} future events`);

    // Create updated future events
    const config = getConfiguration(sheet);
    const cutoffDate = nextMonday;
    
    let eventsCreated = 0;
    const createStartTime = new Date().getTime();
    
    scheduleData.forEach(([dateStr, deacon, household]) => {
      const visitDate = new Date(dateStr);
      
      // Only create events for dates from next Monday onwards
      if (visitDate >= cutoffDate) {
        try {
          const householdIndex = config.households.indexOf(household);
          
          // Build event description with contact info and links
          let description = `Household: ${household}\n`;
          
          if (householdIndex !== -1) {
            const phone = config.phones[householdIndex];
            const address = config.addresses[householdIndex];
            const breezeNumber = config.breezeNumbers[householdIndex];
            const notesLink = config.notesLinks[householdIndex];
            
            description += `\nContact Information:\n`;
            if (phone) description += `Phone: ${phone}\n`;
            if (address) description += `Address: ${address}\n`;
            
            if (breezeNumber) {
              const breezeUrl = buildBreezeUrl(breezeNumber);
              description += `\nBreeze Profile: ${breezeUrl}\n`;
            }
            
            if (notesLink) {
              description += `\nVisit Notes: ${notesLink}\n`;
            }
          }
          
          // Add custom instructions
          const customInstructions = config.customInstructions || 
            'Please call to set up a day and time for your visit in the coming week. ' +
            'You can update this event with the actual day and time and copy it to your personal calendar.';
          
          description += `\nInstructions:\n${customInstructions}`;
          
          // Create calendar event
          calendar.createEvent(
            `${deacon} visits ${household}`,
            new Date(visitDate.getFullYear(), visitDate.getMonth(), visitDate.getDate(), 14, 0), // 2:00 PM
            new Date(visitDate.getFullYear(), visitDate.getMonth(), visitDate.getDate(), 15, 0), // 3:00 PM
            {
              description: description,
              location: config.addresses[householdIndex] || ''
            }
          );
          
          eventsCreated++;
          
        } catch (eventError) {
          console.error(`Failed to create future event for ${deacon} -> ${household}:`, eventError);
        }
      }
    });
    
    console.log(`Created ${eventsCreated} future events in ${new Date().getTime() - createStartTime}ms`);
    
    ui.alert(
      currentTestMode ? 'TEST Future Events Updated' : 'Future Events Updated',
      `${currentTestMode ? '🧪 TEST: ' : ''}🔄 Future events update complete!\n\n` +
      `✅ Created: ${eventsCreated} future events\n` +
      `🛡️ Preserved: Current week events\n` +
      `📅 Starting: ${nextMonday.toLocaleDateString()}`,
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    console.error('Future events update failed:', error);
    SpreadsheetApp.getUi().alert(
      'Update Failed',
      `❌ Future events update failed: ${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

// ===== ARCHIVE AND EXPORT FUNCTIONS =====

function archiveCurrentSchedule() {
  /**
   * Archive current schedule to a new sheet before regeneration
   */
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const scheduleData = getScheduleFromSheet(sheet);
    
    if (scheduleData.length === 0) {
      SpreadsheetApp.getUi().alert(
        'No Schedule to Archive',
        'No schedule found to archive. Please generate a schedule first.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    const timestamp = new Date().toISOString().slice(0, 10); // YYYY-MM-DD format
    const archiveSheetName = `Archive_${timestamp}`;
    
    // Create archive sheet
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const archiveSheet = spreadsheet.insertSheet(archiveSheetName);
    
    // Copy current schedule to archive
    const sourceRange = sheet.getRange('A1:E' + (scheduleData.length + 1));
    const targetRange = archiveSheet.getRange('A1:E' + (scheduleData.length + 1));
    sourceRange.copyTo(targetRange);
    
    // Copy deacon reports too (from location G-I) - dynamically find the range
    const lastRow = sheet.getLastRow();
    const reportRange = `G1:I${Math.max(lastRow, 300)}`;
    const reportData = sheet.getRange(reportRange).getValues().filter(row => row.some(cell => cell !== ''));
    if (reportData.length > 0) {
      archiveSheet.getRange('G1:I' + reportData.length).setValues(reportData);
      console.log(`Archived ${reportData.length} rows of deacon reports from range ${reportRange}`);
    }
    
    SpreadsheetApp.getUi().alert(
      'Archive Created',
      `✅ Current schedule archived to sheet: "${archiveSheetName}"\n\n` +
      `You can now safely generate a new schedule.`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    SpreadsheetApp.getUi().alert(
      'Archive Failed',
      `❌ Could not archive schedule: ${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

function exportIndividualSchedules() {
  /**
   * Create individual schedule sheets for each deacon
   */
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const scheduleData = getScheduleFromSheet(sheet);
    
    if (scheduleData.length === 0) {
      SpreadsheetApp.getUi().alert(
        'No Schedule Found',
        'Please generate a schedule first using "📅 Generate Schedule".',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    const config = getConfiguration(sheet);
    const currentTestMode = getCurrentTestMode();
    
    // Create new spreadsheet for individual schedules
    const timestamp = new Date().toISOString().slice(0, 10);
    const newSpreadsheet = SpreadsheetApp.create(
      `${currentTestMode ? 'TEST - ' : ''}Individual Deacon Schedules - ${timestamp}`
    );
    
    // Group visits by deacon
    const deaconVisits = {};
    scheduleData.forEach(([date, deacon, household]) => {
      if (!deaconVisits[deacon]) {
        deaconVisits[deacon] = [];
      }
      deaconVisits[deacon].push([date, household]);
    });
    
    // Remove the default sheet and create individual sheets
    const defaultSheet = newSpreadsheet.getSheets()[0];
    
    Object.keys(deaconVisits).forEach(deacon => {
      const deaconSheet = newSpreadsheet.insertSheet(deacon);
      
      // Add headers
      deaconSheet.getRange('A1:C1').setValues([['Visit Date', 'Household', 'Contact Information']]);
      deaconSheet.getRange('A1:C1').setFontWeight('bold');
      
      // Add visit data with contact information
      const visits = deaconVisits[deacon];
      const enrichedVisits = visits.map(([date, household]) => {
        const householdIndex = config.households.indexOf(household);
        const contactInfo = householdIndex !== -1 ? 
          `${config.phones[householdIndex] || 'No phone'} | ${config.addresses[householdIndex] || 'No address'}` :
          'Contact info not found';
        
        return [new Date(date).toLocaleDateString(), household, contactInfo];
      });
      
      if (enrichedVisits.length > 0) {
        deaconSheet.getRange(2, 1, enrichedVisits.length, 3).setValues(enrichedVisits);
      }
      
      // Auto-resize columns
      deaconSheet.autoResizeColumns(1, 3);
    });
    
    // Delete default sheet
    newSpreadsheet.deleteSheet(defaultSheet);
    
    const spreadsheetUrl = newSpreadsheet.getUrl();
    
    SpreadsheetApp.getUi().alert(
      currentTestMode ? 'TEST Individual Schedules Created' : 'Individual Schedules Created',
      `${currentTestMode ? '🧪 TEST: ' : ''}✅ Individual deacon schedules created!\n\n` +
      `📊 Created ${Object.keys(deaconVisits).length} individual sheets\n` +
      `📱 Access at: ${spreadsheetUrl}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    SpreadsheetApp.getUi().alert(
      'Export Failed',
      `❌ Could not create individual schedules: ${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

// ===== UTILITY FUNCTIONS =====

function validateScheduleDataSync() {
  /**
   * Validates that current schedule matches configuration data
   * Returns true if synced, false if out of sync (shows error dialog)
   */
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const config = getConfiguration(sheet);
    const scheduleData = getScheduleFromSheet(sheet);
    
    if (scheduleData.length === 0) {
      return true; // No schedule to validate
    }
    
    // Check if deacons in schedule match current configuration
    const scheduleDeacons = new Set(scheduleData.map(row => row[1]));
    const configDeacons = new Set(config.deacons);
    
    const missingFromConfig = [...scheduleDeacons].filter(deacon => !configDeacons.has(deacon));
    const extraInConfig = [...configDeacons].filter(deacon => !scheduleDeacons.has(deacon));
    
    // Check if households in schedule match current configuration
    const scheduleHouseholds = new Set(scheduleData.map(row => row[2]));
    const configHouseholds = new Set(config.households);
    
    const missingHouseholdsFromConfig = [...scheduleHouseholds].filter(household => !configHouseholds.has(household));
    const extraHouseholdsInConfig = [...configHouseholds].filter(household => !scheduleHouseholds.has(household));
    
    if (missingFromConfig.length > 0 || extraInConfig.length > 0 || 
        missingHouseholdsFromConfig.length > 0 || extraHouseholdsInConfig.length > 0) {
      
      let message = '⚠️ Schedule data appears out of sync with current configuration:\n\n';
      
      if (missingFromConfig.length > 0) {
        message += `Deacons in schedule but not in config: ${missingFromConfig.join(', ')}\n`;
      }
      if (extraInConfig.length > 0) {
        message += `Deacons in config but not in schedule: ${extraInConfig.join(', ')}\n`;
      }
      if (missingHouseholdsFromConfig.length > 0) {
        message += `Households in schedule but not in config: ${missingHouseholdsFromConfig.join(', ')}\n`;
      }
      if (extraHouseholdsInConfig.length > 0) {
        message += `Households in config but not in schedule: ${extraHouseholdsInConfig.join(', ')}\n`;
      }
      
      message += '\n💡 Consider regenerating the schedule to sync with current configuration.';
      
      SpreadsheetApp.getUi().alert(
        'Data Sync Warning',
        message,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      
      return false;
    }
    
    return true;
    
  } catch (error) {
    console.error('Schedule validation failed:', error);
    return true; // Don't block operations if validation fails
  }
}

function generateShortUrlsFromMenu() {
  /**
   * Menu wrapper for generateShortUrls function
   */
  try {
    generateShortUrls();
  } catch (error) {
    SpreadsheetApp.getUi().alert(
      'URL Generation Failed',
      `❌ Could not generate shortened URLs: ${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

function showSetupInstructions() {
  /**
   * Display setup instructions for new users
   */
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'Setup Instructions',
    '📋 DEACON VISITATION ROTATION SETUP (v1.1)\n\n' +
    '1️⃣ Configure Settings (Column K):\n' +
    '   • K2: Start date (Monday)\n' +
    '   • K4: Visit frequency (weeks)\n' +
    '   • K6: Schedule length (weeks)\n' +
    '   • K8: Calendar event instructions\n\n' +
    '2️⃣ Configure Notifications (Column K):\n' +
    '   • K11: Notification day (dropdown)\n' +
    '   • K13: Notification time (0-23 hour)\n\n' +
    '3️⃣ Configure Resource Links (Column K):\n' +
    '   • K19: Google Calendar URL\n' +
    '   • K22: Visitation Guide URL\n' +
    '   • K25: Schedule Summary URL\n\n' +
    '4️⃣ Add Deacons (Column L):\n' +
    '   • L2, L3, L4... list all deacon names\n\n' +
    '5️⃣ Add Households (Column M):\n' +
    '   • M2, M3, M4... list all household names\n\n' +
    '6️⃣ Add Contact Info (Optional):\n' +
    '   • Column N: Phone numbers\n' +
    '   • Column O: Addresses\n' +
    '   • Column P: Breeze numbers\n' +
    '   • Column Q: Notes links\n\n' +
    '7️⃣ Generate Schedule:\n' +
    '   • Use "📅 Generate Schedule" menu\n\n' +
    '8️⃣ Export to Calendar:\n' +
    '   • Use "🚨 Full Calendar Regeneration"\n\n' +
    '9️⃣ Configure Notifications:\n' +
    '   • Use "📢 Notifications → 🔧 Configure Chat Webhook"\n\n' +
    '📖 For detailed instructions, see the project documentation.',
    ui.ButtonSet.OK
  );
}

function runSystemTests() {
  /**
   * Comprehensive system testing and diagnostics (v1.1)
   */
  const ui = SpreadsheetApp.getUi();
  const results = [];
  
  try {
    console.log('=== SYSTEM TESTS v1.1 ===');
    
    // Test 1: Configuration validation
    console.log('Test 1: Configuration validation...');
    try {
      const sheet = SpreadsheetApp.getActiveSheet();
      const config = getConfiguration(sheet);
      
      if (config.deacons.length > 0 && config.households.length > 0) {
        console.log(`Configuration passed: ${config.deacons.length} deacons, ${config.households.length} households`);
        results.push(`✅ Configuration: PASSED (${config.deacons.length} deacons, ${config.households.length} households)`);
      } else {
        console.log('Configuration failed: Missing deacons or households');
        results.push('❌ Configuration: FAILED - No deacons or households configured');
      }
    } catch (configError) {
      console.error('Configuration test failed:', configError);
      results.push(`❌ Configuration: FAILED - ${configError.message}`);
    }
    
    // Test 2: Mode detection
    console.log('Test 2: Mode detection...');
    try {
      const currentTestMode = getCurrentTestMode();
      const mode = currentTestMode ? 'TEST' : 'PRODUCTION';
      console.log(`Mode detection passed: ${mode}`);
      results.push(`✅ Mode detection: PASSED (${mode} mode detected)`);
    } catch (modeError) {
      console.error('Mode detection failed:', modeError);
      results.push(`❌ Mode detection: FAILED - ${modeError.message}`);
    }
    
    // Test 3: Schedule generation (basic validation)
    console.log('Test 3: Schedule generation...');
    try {
      const sheet = SpreadsheetApp.getActiveSheet();
      const scheduleData = getScheduleFromSheet(sheet);
      
      if (scheduleData.length > 0) {
        console.log(`Schedule test passed: ${scheduleData.length} schedule entries found`);
        results.push(`✅ Schedule: PASSED (${scheduleData.length} entries)`);
      } else {
        console.log('Schedule test: No schedule found');
        results.push('⚠️ Schedule: NO DATA - Generate a schedule to test');
      }
    } catch (scheduleError) {
      console.error('Schedule test failed:', scheduleError);
      results.push(`❌ Schedule: FAILED - ${scheduleError.message}`);
    }
    
    // Test 4: Calendar access
    console.log('Test 4: Calendar access...');
    try {
      const currentTestMode = getCurrentTestMode();
      const calendarName = currentTestMode ? 'TEST - Deacon Visitation Schedule' : 'Deacon Visitation Schedule';
      const calendars = CalendarApp.getCalendarsByName(calendarName);
      
      if (calendars.length > 0) {
        console.log(`Calendar access passed: Found calendar "${calendarName}"`);
        results.push(`✅ Calendar access: PASSED (${calendarName})`);
      } else {
        console.log(`Calendar test: Calendar "${calendarName}" not found`);
        results.push(`⚠️ Calendar access: NOT FOUND - Create calendar "${calendarName}" to test`);
      }
    } catch (calendarError) {
      console.error('Calendar test failed:', calendarError);
      results.push(`❌ Calendar access: FAILED - ${calendarError.message}`);
    }
    
    // Test 5: Script permissions
    console.log('Test 5: Script permissions...');
    try {
      PropertiesService.getScriptProperties().setProperty('test', 'value');
      const testValue = PropertiesService.getScriptProperties().getProperty('test');
      if (testValue === 'value') {
        PropertiesService.getScriptProperties().deleteProperty('test');
        console.log('Script permissions passed');
        results.push('✅ Script permissions: PASSED');
      } else {
        console.log(`Script permissions failed: expected 'value', got '${testValue}'`);
        results.push('❌ Script permissions: FAILED - Cannot read/write properties');
      }
    } catch (permError) {
      console.error('Script permissions failed:', permError);
      results.push(`❌ Script permissions: FAILED - ${permError.message}`);
    }
    
    console.log('=== SYSTEM TESTS COMPLETED ===');
    console.log('Results:', results);
    
    ui.alert(
      'System Test Results (v1.1)',
      results.join('\n\n'),
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    console.error('System tests crashed:', error);
    ui.alert(
      'System Tests Failed',
      `❌ Error running tests: ${error.message}\n\nCheck the Apps Script logs for details.`,
      ui.ButtonSet.OK
    );
  }
}

// ===== AUTO-EDIT TRIGGER (currently disabled) =====

function onEdit(e) {
  // AUTO-REGENERATION IS DISABLED
  // To re-enable automatic updates when deacon/household lists change,
  // uncomment the lines below and customize as needed
  
  /*
  try {
    const range = e.range;
    const sheet = range.getSheet();
    
    // Check if edit was in deacon or household configuration areas
    const editedColumn = range.getColumn();
    const editedRow = range.getRow();
    
    // Columns A-B (deacons) or D+ (households) in configuration rows
    if (editedRow >= 4 && editedRow <= 20) {
      if ((editedColumn >= 1 && editedColumn <= 2) || editedColumn >= 4) {
        console.log(`Configuration change detected at ${range.getA1Notation()}`);
        
        // Optional: Auto-regenerate schedule
        // generateRotationSchedule();
        
        // Optional: Auto-update calendar
        // updateContactInfoOnly();
      }
    }
  } catch (error) {
    console.error('onEdit error:', error);
  }
  */
}

// ===== MODE DETECTION AND INDICATORS (PROPERLY FIXED) =====

/**
 * PROPERLY FIXED: Get current test mode based on live data analysis
 * This function does direct detection without circular calls or caching
 * Test mode is determined fresh each time based on current spreadsheet state
 */
function getCurrentTestMode() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    
    // Direct detection logic - no external function calls to avoid recursion
    const testIndicators = [
      // Test household names
      () => {
        const households = sheet.getRange('M2:M10').getValues().flat().filter(cell => cell !== '');
        const testPatterns = ['Alan & Alexa Adams', 'Barbara Baker', 'Chloe Cooper', 'test', 'sample'];
        return households.some(household => 
          testPatterns.some(pattern => 
            String(household).toLowerCase().includes(pattern.toLowerCase())
          )
        );
      },
      // Test phone numbers (555)
      () => {
        const phones = sheet.getRange('N2:N10').getValues().flat().filter(cell => cell !== '');
        return phones.some(phone => String(phone).includes('555'));
      },
      // Test Breeze numbers (12345)
      () => {
        const breezeNumbers = sheet.getRange('P2:P10').getValues().flat().filter(cell => cell !== '');
        return breezeNumbers.some(number => String(number).startsWith('12345'));
      },
      // Spreadsheet name contains "test"
      () => {
        const spreadsheetName = SpreadsheetApp.getActiveSpreadsheet().getName().toLowerCase();
        return spreadsheetName.includes('test') || spreadsheetName.includes('sample');
      }
    ];
    
    // Check if any test indicator is true
    const isTestMode = testIndicators.some(check => check());
    
    // Simple console logging (no external calls)
    if (isTestMode) {
      console.log('🧪 TEST MODE detected');
    } else {
      console.log('✅ PRODUCTION MODE detected');
    }
    
    return isTestMode;
    
  } catch (error) {
    console.warn('Could not detect test mode, defaulting to production:', error);
    return false; // Default to production mode if detection fails
  }
}

function addModeIndicatorToSheet() {
  /**
   * Add visual mode indicator to spreadsheet
   */
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const currentTestMode = getCurrentTestMode();
    
    // Update mode indicator in K15-K16 (moved from K11-K12 to make room for notification config)
    const modeLabel = currentTestMode ? '🧪 TEST MODE' : '✅ PRODUCTION';
    const modeDetails = currentTestMode ? 
      'Using test data patterns' : 
      'Using production data';
    
    sheet.getRange('K15').setValue(modeLabel);
    sheet.getRange('K16').setValue(modeDetails);
    
    // Style the mode indicator
    const modeRange = sheet.getRange('K15:K16');
    if (currentTestMode) {
      modeRange.setBackground('#ffeb3b').setFontColor('#d84315'); // Yellow background, red text
    } else {
      modeRange.setBackground('#e8f5e8').setFontColor('#2e7d32'); // Light green background, dark green text
    }
    modeRange.setFontWeight('bold');
    
  } catch (error) {
    console.error('Failed to add mode indicator:', error);
  }
}

function showModeNotification() {
  /**
   * Show current mode in a user-friendly popup
   */
  try {
    const currentTestMode = getCurrentTestMode();
    const ui = SpreadsheetApp.getUi();
    const modeTitle = currentTestMode ? '🧪 Test Mode Active' : '✅ Production Mode Active';
    
    let message = `**Current Mode:** ${currentTestMode ? 'TEST' : 'PRODUCTION'}\n\n`;
    message += `Calendar: "${currentTestMode ? 'TEST - ' : ''}Deacon Visitation Schedule"\n\n`;
    
    if (currentTestMode) {
      message += '🧪 TEST MODE detected because:\n';
      message += '• Sample data patterns found\n';
      message += '• Test calendar will be used for safety\n\n';
      message += '💡 Switch to real member data to enable production mode';
    } else {
      message += '✅ PRODUCTION MODE - using live data\n';
      message += '• Real member data detected\n';
      message += '• Live calendar will be used\n';
      message += '• All functions ready for pastoral care';
    }
    
    ui.alert(modeTitle, message, ui.ButtonSet.OK);
    
  } catch (error) {
    console.error('Failed to show mode notification:', error);
    SpreadsheetApp.getUi().alert(
      'Mode Detection Error',
      `❌ Could not determine current mode: ${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

// ===== ENHANCED MENU SYSTEM (v1.1) WITH CLEANED FUNCTIONS =====

function createMenuItems() {
  const ui = SpreadsheetApp.getUi();
  
  // Get current mode for menu display
  let modeIcon = '❓';
  try {
    const currentTestMode = getCurrentTestMode();
    modeIcon = currentTestMode ? '🧪' : '✅';
  } catch (error) {
    console.warn('Could not determine mode for menu:', error);
  }
  
  ui.createMenu('🔄 Deacon Rotation')
    .addItem('📅 Generate Schedule', 'generateRotationSchedule')
    .addSeparator()
    .addItem('🔗 Generate Shortened URLs', 'generateShortUrlsFromMenu')
    .addSeparator()
    .addSubMenu(ui.createMenu('📆 Calendar Functions')
      .addItem('📞 Update Contact Info Only', 'updateContactInfoOnly')
      .addItem('🔄 Update Future Events Only', 'updateFutureEventsOnly')
      .addSeparator()
      .addItem('🚨 Full Calendar Regeneration', 'exportToGoogleCalendar'))
    .addSeparator()
    .addSubMenu(ui.createMenu('📢 Notifications')
      .addItem('💬 Send Weekly Chat Summary', 'sendWeeklyVisitationChat')
      .addSeparator()
      .addItem('🔄 Enable Weekly Auto-Send', 'createWeeklyNotificationTrigger')
      .addItem('📅 Show Auto-Send Schedule', 'showCurrentTriggerSchedule')
      .addItem('🛑 Disable Weekly Auto-Send', 'removeWeeklyNotificationTrigger')
      .addSeparator()
      .addItem('🔧 Configure Chat Webhook', 'configureNotifications')
      .addItem('📋 Test Notification System', 'testNotificationSystem')
      .addSeparator()
      .addItem('🔍 Inspect All Triggers', 'inspectAllTriggers')
      .addItem('🔄 Force Recreate Trigger', 'forceRecreateWeeklyTrigger')
      .addItem('🧪 Test Calendar Link Config', 'testCalendarLinkConfiguration'))
    .addSeparator()
    .addItem('📊 Export Individual Schedules', 'exportIndividualSchedules')
    .addSeparator()
    .addItem('📁 Archive Current Schedule', 'archiveCurrentSchedule')
    .addSeparator()
    .addItem('🔧 Validate Setup', 'validateSetupOnly')
    .addItem('🧪 Run Tests', 'runSystemTests')
    .addSeparator()
    .addItem(`${modeIcon} Show Current Mode`, 'showModeNotification')
    .addItem('❓ Setup Instructions', 'showSetupInstructions')
    .addToUi();
}

// ===== HELPER FUNCTIONS =====

function getOrCreateCalendar(calendarName = null) {
  /**
   * Helper function for calendar access with dynamic mode
   */
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
        `Calendar "${calendarName}" not found.\n\nPlease run "📆 Full Calendar Regeneration" first to create the calendar.`,
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

// END OF MODULE 4
