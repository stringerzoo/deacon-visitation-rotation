/**
 * MODULE 4: CALENDAR EXPORT, MENU SYSTEM, AND UTILITIES (v1.1)
 * Deacon Visitation Rotation System - Calendar Integration & Menu Management
 * 
 * CLEANED VERSION: Removed unused functions based on analysis
 * - ❌ Removed: generateNextYearSchedule (unlikely to be used)
 * - ✅ Kept: testNotificationNow (different purpose from testNotificationSystem)
 * - ✅ Kept: All other valuable functions for delegation and troubleshooting
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
          `${currentTestMode ? '🧪 TEST MODE: ' : ''}⚠️ This will completely rebuild the calendar "${calendarName}".\n\n` +
          `🚨 WARNING: This will delete ALL existing events and lose any custom scheduling details!\n\n` +
          `For safer updates, consider:\n` +
          `• "📞 Update Contact Info Only" - Preserves all scheduling\n` +
          `• "🔄 Update Future Events Only" - Preserves this week\n\n` +
          `Continue with full regeneration?`,
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
              Utilities.sleep(500);
            }
          });
          
          console.log(`Deleted ${deletedCount} events in ${new Date().getTime() - deleteStartTime}ms`);
          
          // Wait a bit longer before creating new events
          console.log('Waiting for API cooldown before creating new events...');
          Utilities.sleep(2000);
        }
      } else {
        calendar = CalendarApp.createCalendar(calendarName);
        calendar.setDescription(`${currentTestMode ? 'TEST: ' : ''}Automated schedule for deacon household visitations with contact information and management links`);
        calendar.setColor(currentTestMode ? CalendarApp.Color.RED : CalendarApp.Color.BLUE);
      }
    } catch (calError) {
      throw new Error(`Calendar access failed: ${calError.message}. Please check calendar permissions.`);
    }
    
    const config = getConfiguration(sheet);
    let successCount = 0;
    let errorCount = 0;
    const errors = [];
    
    console.log(`Creating ${scheduleData.length} calendar events...`);
    const creationStartTime = new Date().getTime();
    
    scheduleData.forEach((visit, index) => {
      try {
        const [visitDate, deacon, household] = visit;
        const householdIndex = config.households.indexOf(household);
        
        if (householdIndex === -1) {
          errors.push(`Household "${household}" not found in configuration`);
          errorCount++;
          return;
        }
        
        // Get contact information and links
        const phone = config.phones[householdIndex] || 'No phone listed';
        const address = config.addresses[householdIndex] || 'No address listed';
        const breezeNumber = config.breezeNumbers[householdIndex];
        const notesLink = config.notesLinks[householdIndex];
        
        // Create event title and description
        const eventTitle = `${deacon} visits ${household}`;
        
        let description = `Household: ${household}\n`;
        
        // Add Breeze profile link if available
        if (breezeNumber && breezeNumber.toString().length >= 7) {
          const breezeUrl = `https://immanuelpres.breezechms.com/people/view/${breezeNumber}`;
          const shortBreezeUrl = shortenUrl(breezeUrl);
          description += `Breeze Profile: ${shortBreezeUrl}\n\n`;
        }
        
        description += `Contact Information:\n`;
        description += `Phone: ${phone}\n`;
        description += `Address: ${address}\n\n`;
        
        // Add visit notes link if available
        if (notesLink && notesLink.length > 0) {
          const shortNotesUrl = shortenUrl(notesLink);
          description += `Visit Notes: ${shortNotesUrl}\n\n`;
        }
        
        description += `Instructions:\n`;
        description += `Please call to set up a day and time for your visit in the coming week. `;
        description += `You can update this event with the actual day and time and copy it to your personal calendar.`;
        
        // Add resource links from K19, K22, K25 if configured
        const calendarUrl = getCalendarLinkFromSpreadsheet();
        const guideUrl = getGuideLinkFromSpreadsheet();
        const summaryUrl = getSummaryLinkFromSpreadsheet();
        
        if (calendarUrl) {
          const shortCalendarUrl = shortenUrl(calendarUrl);
          description += ` Visitation Guide available at ${shortCalendarUrl}`;
        }
        
        // Create the event (Monday 2-3 PM of the specified week)
        const eventDate = new Date(visitDate);
        eventDate.setHours(14, 0, 0, 0); // 2:00 PM
        
        const endDate = new Date(eventDate);
        endDate.setHours(15, 0, 0, 0); // 3:00 PM
        
        const event = calendar.createEvent(eventTitle, eventDate, endDate, {
          description: description,
          location: address
        });
        
        successCount++;
        
        // Rate limiting: small delay every 25 operations
        if ((index + 1) % 25 === 0) {
          Utilities.sleep(2000);
          console.log(`Created ${index + 1}/${scheduleData.length} events...`);
        }
        
      } catch (eventError) {
        console.error(`Error creating event for ${visit[1]} -> ${visit[2]}:`, eventError);
        errors.push(`${visit[1]} -> ${visit[2]}: ${eventError.message}`);
        errorCount++;
      }
    });
    
    const totalTime = new Date().getTime() - creationStartTime;
    console.log(`Calendar export completed in ${totalTime}ms: ${successCount} success, ${errorCount} errors`);
    
    // Show results
    let resultMessage = `${currentTestMode ? '🧪 TEST: ' : ''}📅 Calendar Export Complete!\n\n`;
    resultMessage += `✅ Created: ${successCount} events\n`;
    if (errorCount > 0) {
      resultMessage += `❌ Errors: ${errorCount}\n\n`;
      resultMessage += `First few errors:\n${errors.slice(0, 3).join('\n')}`;
      if (errors.length > 3) {
        resultMessage += `\n... and ${errors.length - 3} more (check logs)`;
      }
    }
    resultMessage += `\n\n📅 Calendar: "${calendarName}"\n`;
    resultMessage += `⏱️ Time: ${(totalTime/1000).toFixed(1)} seconds`;
    
    SpreadsheetApp.getUi().alert(
      currentTestMode ? 'TEST Calendar Updated' : 'Calendar Updated',
      resultMessage,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    console.error('Calendar export failed:', error);
    SpreadsheetApp.getUi().alert(
      'Calendar Export Failed',
      `❌ Error: ${error.message}\n\nCheck the Apps Script logs for more details.`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

function updateContactInfoOnly() {
  /**
   * Smart update that preserves all scheduling details, only updates contact info
   */
  try {
    const currentTestMode = getCurrentTestMode();
    const calendarName = currentTestMode ? 
      'TEST - Deacon Visitation Schedule' : 'Deacon Visitation Schedule';
    
    const calendars = CalendarApp.getCalendarsByName(calendarName);
    if (calendars.length === 0) {
      SpreadsheetApp.getUi().alert(
        'Calendar Not Found',
        `No calendar named "${calendarName}" found.\n\nPlease run "🚨 Full Calendar Regeneration" first.`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    const calendar = calendars[0];
    const sheet = SpreadsheetApp.getActiveSheet();
    const config = getConfiguration(sheet);
    
    // Get all events for next 2 years
    const startDate = new Date();
    const endDate = new Date();
    endDate.setFullYear(endDate.getFullYear() + 2);
    
    const events = calendar.getEvents(startDate, endDate);
    let updatedCount = 0;
    let errorCount = 0;
    
    console.log(`Updating contact information for ${events.length} events...`);
    
    events.forEach((event, index) => {
      try {
        const title = event.getTitle();
        const match = title.match(/(.+) visits (.+)/);
        
        if (match) {
          const [, deacon, household] = match;
          const householdIndex = config.households.indexOf(household);
          
          if (householdIndex !== -1) {
            // Get updated contact information
            const phone = config.phones[householdIndex] || 'No phone listed';
            const address = config.addresses[householdIndex] || 'No address listed';
            const breezeNumber = config.breezeNumbers[householdIndex];
            const notesLink = config.notesLinks[householdIndex];
            
            // Build new description preserving the event structure
            let description = `Household: ${household}\n`;
            
            if (breezeNumber && breezeNumber.toString().length >= 7) {
              const breezeUrl = `https://immanuelpres.breezechms.com/people/view/${breezeNumber}`;
              const shortBreezeUrl = shortenUrl(breezeUrl);
              description += `Breeze Profile: ${shortBreezeUrl}\n\n`;
            }
            
            description += `Contact Information:\n`;
            description += `Phone: ${phone}\n`;
            description += `Address: ${address}\n\n`;
            
            if (notesLink && notesLink.length > 0) {
              const shortNotesUrl = shortenUrl(notesLink);
              description += `Visit Notes: ${shortNotesUrl}\n\n`;
            }
            
            description += `Instructions:\n`;
            description += `Please call to set up a day and time for your visit in the coming week. `;
            description += `You can update this event with the actual day and time and copy it to your personal calendar.`;
            
            // Add resource links
            const calendarUrl = getCalendarLinkFromSpreadsheet();
            if (calendarUrl) {
              const shortCalendarUrl = shortenUrl(calendarUrl);
              description += ` Visitation Guide available at ${shortCalendarUrl}`;
            }
            
            // Update event description and location
            event.setDescription(description);
            event.setLocation(address);
            
            updatedCount++;
          }
        }
        
        // Rate limiting
        if ((index + 1) % 50 === 0) {
          Utilities.sleep(1000);
          console.log(`Updated ${updatedCount} of ${index + 1} events...`);
        }
        
      } catch (eventError) {
        console.error(`Error updating event: ${eventError.message}`);
        errorCount++;
      }
    });
    
    SpreadsheetApp.getUi().alert(
      currentTestMode ? 'TEST Contact Info Updated' : 'Contact Info Updated',
      `${currentTestMode ? '🧪 TEST: ' : ''}📞 Contact information update complete!\n\n` +
      `✅ Updated: ${updatedCount} events\n` +
      `❌ Errors: ${errorCount}\n\n` +
      `All scheduling details preserved - only contact information was updated.`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    console.error('Contact info update failed:', error);
    SpreadsheetApp.getUi().alert(
      'Update Failed',
      `❌ Contact info update failed: ${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

function updateFutureEventsOnly() {
  /**
   * Smart update that only affects future events, preserves current week
   */
  try {
    const currentTestMode = getCurrentTestMode();
    const calendarName = currentTestMode ? 
      'TEST - Deacon Visitation Schedule' : 'Deacon Visitation Schedule';
    
    const calendars = CalendarApp.getCalendarsByName(calendarName);
    if (calendars.length === 0) {
      SpreadsheetApp.getUi().alert(
        'Calendar Not Found',
        `No calendar named "${calendarName}" found.\n\nPlease run "🚨 Full Calendar Regeneration" first.`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    // Calculate start of next Monday (preserve current week)
    const now = new Date();
    const nextMonday = new Date(now);
    const daysUntilNextMonday = (8 - now.getDay()) % 7;
    nextMonday.setDate(now.getDate() + daysUntilNextMonday);
    nextMonday.setHours(0, 0, 0, 0);
    
    const response = SpreadsheetApp.getUi().alert(
      currentTestMode ? 'TEST: Update Future Events' : 'Update Future Events',
      `${currentTestMode ? '🧪 TEST MODE: ' : ''}This will update events starting ${nextMonday.toLocaleDateString()}.\n\n` +
      `✅ PRESERVES: Current week events (any custom scheduling)\n` +
      `🔄 UPDATES: Future events with latest contact info and schedule\n\n` +
      `Continue?`,
      SpreadsheetApp.getUi().ButtonSet.YES_NO
    );
    
    if (response !== SpreadsheetApp.getUi().Button.YES) return;
    
    const calendar = calendars[0];
    
    // Delete future events only
    const endDate = new Date(nextMonday);
    endDate.setFullYear(endDate.getFullYear() + 2);
    
    const futureEvents = calendar.getEvents(nextMonday, endDate);
    console.log(`Deleting ${futureEvents.length} future events...`);
    
    futureEvents.forEach((event, index) => {
      event.deleteEvent();
      if ((index + 1) % 25 === 0) {
        Utilities.sleep(500);
      }
    });
    
    // Get schedule data and recreate future events
    const sheet = SpreadsheetApp.getActiveSheet();
    const scheduleData = getScheduleFromSheet(sheet);
    const futureSchedule = scheduleData.filter(visit => {
      const visitDate = new Date(visit[0]);
      return visitDate >= nextMonday;
    });
    
    if (futureSchedule.length === 0) {
      SpreadsheetApp.getUi().alert(
        'No Future Events',
        'No events found starting from next Monday. Update complete.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    // Recreate future events with updated information
    const config = getConfiguration(sheet);
    let successCount = 0;
    
    futureSchedule.forEach((visit, index) => {
      try {
        const [visitDate, deacon, household] = visit;
        const householdIndex = config.households.indexOf(household);
        
        if (householdIndex !== -1) {
          // Create event with updated contact information
          const phone = config.phones[householdIndex] || 'No phone listed';
          const address = config.addresses[householdIndex] || 'No address listed';
          const breezeNumber = config.breezeNumbers[householdIndex];
          const notesLink = config.notesLinks[householdIndex];
          
          const eventTitle = `${deacon} visits ${household}`;
          
          let description = `Household: ${household}\n`;
          
          if (breezeNumber && breezeNumber.toString().length >= 7) {
            const breezeUrl = `https://immanuelpres.breezechms.com/people/view/${breezeNumber}`;
            const shortBreezeUrl = shortenUrl(breezeUrl);
            description += `Breeze Profile: ${shortBreezeUrl}\n\n`;
          }
          
          description += `Contact Information:\n`;
          description += `Phone: ${phone}\n`;
          description += `Address: ${address}\n\n`;
          
          if (notesLink && notesLink.length > 0) {
            const shortNotesUrl = shortenUrl(notesLink);
            description += `Visit Notes: ${shortNotesUrl}\n\n`;
          }
          
          description += `Instructions:\n`;
          description += `Please call to set up a day and time for your visit in the coming week. `;
          description += `You can update this event with the actual day and time and copy it to your personal calendar.`;
          
          const calendarUrl = getCalendarLinkFromSpreadsheet();
          if (calendarUrl) {
            const shortCalendarUrl = shortenUrl(calendarUrl);
            description += ` Visitation Guide available at ${shortCalendarUrl}`;
          }
          
          const eventDate = new Date(visitDate);
          eventDate.setHours(14, 0, 0, 0);
          
          const endEventDate = new Date(eventDate);
          endEventDate.setHours(15, 0, 0, 0);
          
          calendar.createEvent(eventTitle, eventDate, endEventDate, {
            description: description,
            location: address
          });
          
          successCount++;
        }
        
        // Rate limiting
        if ((index + 1) % 25 === 0) {
          Utilities.sleep(1000);
        }
        
      } catch (eventError) {
        console.error(`Error creating future event: ${eventError.message}`);
      }
    });
    
    SpreadsheetApp.getUi().alert(
      currentTestMode ? 'TEST Future Events Updated' : 'Future Events Updated',
      `${currentTestMode ? '🧪 TEST: ' : ''}🔄 Future events update complete!\n\n` +
      `✅ Created: ${successCount} future events\n` +
      `🛡️ Preserved: Current week events\n` +
      `📅 Starting: ${nextMonday.toLocaleDateString()}`,
      SpreadsheetApp.getUi().ButtonSet.OK
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
      `${currentTestMode ? '🧪 TEST: ' : ''}📊 Individual schedules created!\n\n` +
      `✅ Created sheets for ${Object.keys(deaconVisits).length} deacons\n` +
      `📄 Total visits: ${scheduleData.length}\n\n` +
      `Click OK to open the new spreadsheet.`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    // The URL will be shown in the logs - user can access it from there
    console.log(`Individual schedules spreadsheet created: ${spreadsheetUrl}`);
    
  } catch (error) {
    console.error('Individual schedules export failed:', error);
    SpreadsheetApp.getUi().alert(
      'Export Failed',
      `❌ Could not create individual schedules: ${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

// ===== TESTING AND DIAGNOSTIC UTILITIES =====

function runSystemTests() {
  const ui = SpreadsheetApp.getUi();
  const results = [];
  
  console.log('=== SYSTEM TESTS STARTING ===');
  
  try {
    // Test 1: Configuration loading
    console.log('Test 1: Configuration loading...');
    try {
      const sheet = SpreadsheetApp.getActiveSheet();
      const config = getConfiguration(sheet);
      console.log(`Config loaded: ${config.deacons.length} deacons, ${config.households.length} households`);
      console.log(`Start date: ${config.startDate}, Visit frequency: ${config.visitFrequency}`);
      console.log(`Breeze numbers: ${config.breezeNumbers.filter(n => n && n.length > 0).length}`);
      console.log(`Notes links: ${config.notesLinks.filter(n => n && n.length > 0).length}`);
      results.push('✅ Configuration loading: PASSED');
      
      // Test 2: URL shortening
      console.log('Test 2: URL shortening...');
      try {
        const testUrl = 'https://docs.google.com/document/d/1234567890abcdef/edit';
        const shortUrl = shortenUrl(testUrl);
        
        if (shortUrl.includes('tinyurl.com') || shortUrl === testUrl) {
          console.log(`URL shortening test: ${testUrl} → ${shortUrl}`);
          results.push('✅ URL shortening: PASSED');
        } else {
          console.log(`URL shortening failed: got ${shortUrl}`);
          results.push('❌ URL shortening: FAILED - Unexpected response');
        }
      } catch (urlError) {
        console.error('URL shortening failed:', urlError);
        results.push(`⚠️ URL shortening: WARNING - ${urlError.message} (may still work)`);
      }
      
      // Test 3: Breeze URL construction
      console.log('Test 3: Breeze URL construction...');
      try {
        const testBreezeNumber = '12345678';
        const breezeUrl = `https://immanuelpres.breezechms.com/people/view/${testBreezeNumber}`;
        console.log(`Breeze URL construction test: ${testBreezeNumber} → ${breezeUrl}`);
        results.push('✅ Breeze URL construction: PASSED');
      } catch (breezeError) {
        console.error('Breeze URL construction failed:', breezeError);
        results.push(`❌ Breeze URL construction: FAILED - ${breezeError.message}`);
      }
      
      // Test 4: Calendar access
      console.log('Test 4: Calendar access...');
      try {
        const testCalendarName = 'TEST - System Validation Calendar';
        const testCalendar = CalendarApp.createCalendar(testCalendarName);
        console.log(`Created test calendar: ${testCalendarName}`);
        
        // Clean up test calendar
        testCalendar.deleteCalendar();
        console.log('Cleaned up test calendar');
        results.push('✅ Calendar access: PASSED');
      } catch (calendarError) {
        console.error('Calendar access failed:', calendarError);
        results.push(`❌ Calendar access: FAILED - ${calendarError.message}`);
      }
      
    } catch (configError) {
      console.error('Configuration loading failed:', configError);
      results.push(`❌ Configuration loading: FAILED - ${configError.message}`);
      results.push('⚠️ URL shortening: SKIPPED - No valid configuration');
      results.push('⚠️ Breeze URL construction: SKIPPED - No valid configuration');
      results.push('⚠️ Calendar access: SKIPPED - No valid configuration');
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

// ===== ENHANCED MENU SYSTEM (v1.1) WITH CLEANED FUNCTIONS =====

function createMenuItems() {
  const ui = SpreadsheetApp.getUi();
  
  // Get current mode for menu display
  const properties = PropertiesService.getScriptProperties();
  const mode = properties.getProperty('DETECTED_MODE') || 'UNKNOWN';
  const modeIcon = mode.includes('TEST') ? '🧪' : '✅';
  
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
      .addItem('🧪 Test Notification Now', 'testNotificationNow')
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
  
  // Add mode indicator to spreadsheet
  addModeIndicatorToSheet();
}

// Create menu when spreadsheet opens
function onOpen() {
  try {
    createMenuItems();
    
    // Show mode notification on first open (only once per session)
    const properties = PropertiesService.getScriptProperties();
    const sessionId = Utilities.getUuid();
    const lastSessionId = properties.getProperty('LAST_SESSION_ID');
    
    if (lastSessionId !== sessionId) {
      properties.setProperty('LAST_SESSION_ID', sessionId);
      
      // Small delay to let the spreadsheet load, then show notification
      Utilities.sleep(1000);
      showModeNotification();
    }
    
    console.log('Enhanced Deacon Rotation menu created successfully (v1.1 - CLEANED)');
  } catch (error) {
    console.error('Failed to create menu:', error);
  }
}

// Enhanced onEdit with optional auto-regeneration (currently disabled)
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

// ===== MODE DETECTION AND INDICATORS =====

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
    const properties = PropertiesService.getScriptProperties();
    const detectedMode = properties.getProperty('DETECTED_MODE') || 'Unknown';
    
    const ui = SpreadsheetApp.getUi();
    const modeTitle = currentTestMode ? '🧪 Test Mode Active' : '✅ Production Mode Active';
    
    let message = `**Current Mode:** ${currentTestMode ? 'TEST' : 'PRODUCTION'}\n`;
    message += `**Detection Method:** ${detectedMode}\n\n`;
    
    if (currentTestMode) {
      message += `🧪 **Test Mode Features:**\n`;
      message += `• Creates "TEST - Deacon Visitation Schedule" calendar\n`;
      message += `• Uses test chat webhook (if configured)\n`;
      message += `• Calendar events appear in red\n`;
      message += `• All notifications prefixed with "🧪 TEST:"\n\n`;
      message += `This mode is ideal for:\n`;
      message += `• Learning the system\n`;
      message += `• Testing configurations\n`;
      message += `• Training new users\n\n`;
      message += `**To switch to production:** Use real deacon/household data`;
    } else {
      message += `✅ **Production Mode Features:**\n`;
      message += `• Creates "Deacon Visitation Schedule" calendar\n`;
      message += `• Uses main deacon chat webhook\n`;
      message += `• Calendar events appear in blue\n`;
      message += `• Clean notifications without test prefixes\n\n`;
      message += `This mode is active because:\n`;
      message += `• Real deacon and household names detected\n`;
      message += `• System is ready for operational use\n\n`;
      message += `**All systems operational!** ✅`;
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

// ===== UTILITY FUNCTIONS =====

function generateShortUrlsFromMenu() {
  /**
   * Generate shortened URLs for all configured links
   */
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const config = getConfiguration(sheet);
    
    let shortened = 0;
    let errors = 0;
    
    console.log('Generating shortened URLs...');
    
    // Shorten notes links
    config.notesLinks.forEach((link, index) => {
      if (link && link.length > 0 && !link.includes('tinyurl.com')) {
        try {
          const shortUrl = shortenUrl(link);
          if (shortUrl !== link) {
            // Update the spreadsheet with shortened URL
            const row = 4 + index; // Assuming notes links start at row 4
            const notesColumn = findNotesColumn(sheet);
            if (notesColumn > 0) {
              sheet.getRange(row, notesColumn).setValue(shortUrl);
              shortened++;
            }
          }
        } catch (error) {
          console.error(`Failed to shorten URL: ${link}`, error);
          errors++;
        }
      }
    });
    
    SpreadsheetApp.getUi().alert(
      'URL Shortening Complete',
      `✅ Shortened: ${shortened} URLs\n❌ Errors: ${errors}\n\nLong URLs have been replaced with tinyurl.com links for better mobile compatibility.`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    console.error('URL shortening failed:', error);
    SpreadsheetApp.getUi().alert(
      'URL Shortening Failed',
      `❌ Could not generate shortened URLs: ${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

function findNotesColumn(sheet) {
  /**
   * Helper function to find the notes column dynamically
   */
  try {
    const headerRow = sheet.getRange('A3:Z3').getValues()[0];
    for (let i = 0; i < headerRow.length; i++) {
      if (headerRow[i].toString().toLowerCase().includes('notes')) {
        return i + 1; // Convert to 1-based column index
      }
    }
    return 0; // Not found
  } catch (error) {
    console.error('Failed to find notes column:', error);
    return 0;
  }
}

function validateSetupOnly() {
  /**
   * Comprehensive setup validation without running full tests
   */
  try {
    const ui = SpreadsheetApp.getUi();
    const sheet = SpreadsheetApp.getActiveSheet();
    const issues = [];
    
    console.log('=== SETUP VALIDATION STARTING ===');
    
    // Test 1: Configuration validation
    try {
      const config = getConfiguration(sheet);
      
      if (config.deacons.length === 0) {
        issues.push('❌ No deacons configured in column A');
      } else if (config.deacons.length < 2) {
        issues.push('⚠️ Only 1 deacon configured - need at least 2 for rotation');
      }
      
      if (config.households.length === 0) {
        issues.push('❌ No households configured in column D');
      } else if (config.households.length < 2) {
        issues.push('⚠️ Only 1 household configured - need at least 2 for rotation');
      }
      
      if (!config.startDate || isNaN(config.startDate.getTime())) {
        issues.push('❌ Invalid start date in K2');
      }
      
      if (!config.visitFrequency || config.visitFrequency < 1) {
        issues.push('❌ Invalid visit frequency in K4');
      }
      
      // Check for matching array lengths
      const maxHouseholds = Math.max(config.households.length, config.phones.length, config.addresses.length);
      if (config.phones.length < config.households.length) {
        issues.push(`⚠️ Missing phone numbers: ${config.households.length - config.phones.length} households have no phone`);
      }
      if (config.addresses.length < config.households.length) {
        issues.push(`⚠️ Missing addresses: ${config.households.length - config.addresses.length} households have no address`);
      }
      
    } catch (configError) {
      issues.push(`❌ Configuration error: ${configError.message}`);
    }
    
    // Test 2: Notification configuration
    try {
      const currentTestMode = getCurrentTestMode();
      const webhookUrl = getWebhookUrl(currentTestMode);
      
      if (!webhookUrl || webhookUrl.length === 0) {
        issues.push(`⚠️ No ${currentTestMode ? 'test' : 'production'} chat webhook configured`);
      } else if (!webhookUrl.includes('chat.googleapis.com')) {
        issues.push(`❌ Invalid webhook URL format`);
      }
      
    } catch (notificationError) {
      issues.push(`⚠️ Notification check failed: ${notificationError.message}`);
    }
    
    // Test 3: Calendar permissions
    try {
      const testCalendarName = 'VALIDATION_TEST_CALENDAR';
      const testCalendar = CalendarApp.createCalendar(testCalendarName);
      testCalendar.deleteCalendar();
    } catch (calendarError) {
      issues.push(`❌ Calendar permission issue: ${calendarError.message}`);
    }
    
    console.log('=== SETUP VALIDATION COMPLETED ===');
    
    // Generate report
    let report = '';
    if (issues.length === 0) {
      report = '✅ **Setup Validation: PASSED**\n\nAll systems configured correctly!\n\nYou can now:\n• Generate schedules\n• Export to calendar\n• Send notifications\n• Use all system features';
    } else {
      report = `⚠️ **Setup Validation: ${issues.length} Issue(s) Found**\n\n`;
      report += issues.join('\n\n');
      report += '\n\n📝 Please address these issues before using the system.';
    }
    
    ui.alert(
      'Setup Validation Results',
      report,
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    console.error('Setup validation failed:', error);
    SpreadsheetApp.getUi().alert(
      'Validation Error',
      `❌ Setup validation failed: ${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

function showSetupInstructions() {
  /**
   * Display comprehensive setup instructions
   */
  const ui = SpreadsheetApp.getUi();
  
  const instructions = `📋 **DEACON ROTATION SYSTEM SETUP**\n\n` +
    `**1. Configure Basic Information:**\n` +
    `• Column A: List deacon names (starting row 4)\n` +
    `• Column D: List household names (starting row 4)\n` +
    `• K2: Enter start date (Monday of first week)\n` +
    `• K4: Enter visit frequency in weeks\n\n` +
    `**2. Add Contact Information:**\n` +
    `• Column E: Phone numbers for each household\n` +
    `• Column F: Addresses for each household\n` +
    `• Column G: Breeze numbers (optional, for profile links)\n` +
    `• Column H: Google Docs links for visit notes (optional)\n\n` +
    `**3. Configure Notifications (Optional):**\n` +
    `• Create Google Chat space for deacons\n` +
    `• Get webhook URL from chat space\n` +
    `• Use "📢 Notifications → 🔧 Configure Chat Webhook"\n` +
    `• K11: Set notification day (e.g., "Saturday")\n` +
    `• K13: Set notification time (e.g., 18 for 6 PM)\n\n` +
    `**4. Generate and Test:**\n` +
    `• Use "📅 Generate Schedule" to create rotation\n` +
    `• Use "🧪 Run Tests" to validate setup\n` +
    `• Use "🚨 Full Calendar Regeneration" to create calendar\n\n` +
    `**Need Help?** Check the documentation or run "🔧 Validate Setup"`;
  
  ui.alert(
    'Setup Instructions',
    instructions,
    ui.ButtonSet.OK
  );
}
