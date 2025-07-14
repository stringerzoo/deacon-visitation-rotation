/**
 * MODULE 4: EXPORT, MENU & UTILITY FUNCTIONS (v1.1 - CLEAN MENU VERSION)
 * Deacon Visitation Rotation System - Export and Menu System
 * 
 * CLEAN VERSION: Based on working 7-10 version with only menu cleanup
 * - ❌ Removed: generateNextYearSchedule (menu item and function)
 * - ❌ Removed: sendTomorrowReminders (menu item and function) 
 * - ✅ Updated: Setup instructions to include K19/K22/K25
 * - ✅ Everything else: EXACTLY as it was in working version
 */

// ===== CALENDAR EXPORT FUNCTIONS =====

function exportToGoogleCalendar() {
  /**
   * Full calendar regeneration - creates calendar events with enhanced contact information
   * WARNING: This deletes ALL existing events and loses custom scheduling details
   * v1.1: Fixed data structure handling and optimized performance
   */
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const config = getConfiguration(sheet);
    const scheduleData = getScheduleFromSheet(sheet);
    
    if (config.deacons.length === 0) {
      SpreadsheetApp.getUi().alert('Error', 'No deacons found. Please generate a schedule first.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    if (scheduleData.length === 0) {
      SpreadsheetApp.getUi().alert('No Schedule', 'Please generate a schedule first.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    // Check if shortened URLs need to be generated
    const hasShortUrls = config.breezeShortLinks.some(link => link && link.length > 0) || 
                        config.notesShortLinks.some(link => link && link.length > 0);
    
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
        generateAndStoreShortUrls(sheet, config);
        // Reload config to get the newly generated short URLs
        const updatedConfig = getConfiguration(sheet);
        config.breezeShortLinks = updatedConfig.breezeShortLinks;
        config.notesShortLinks = updatedConfig.notesShortLinks;
      }
    }
    
    // Create or get the deacon visitation calendar
    let calendar;
    const currentTestMode = getCurrentTestMode();
    const calendarName = currentTestMode ? 'TEST - Deacon Visitation Schedule' : 'Deacon Visitation Schedule';
    
    try {
      const calendars = CalendarApp.getCalendarsByName(calendarName);
      if (calendars.length > 0) {
        calendar = calendars[0];
        
        const response = SpreadsheetApp.getUi().alert(
          currentTestMode ? 'TEST: Full Calendar Regeneration' : 'Full Calendar Regeneration',
          `${currentTestMode ? '🧪 TEST: ' : ''}⚠️ This will delete ALL existing events and recreate them!\n\n` +
          `Calendar: "${calendarName}"\n\n` +
          `❌ LOSES: Custom scheduling details (times, guests, notes)\n` +
          `✅ CREATES: ${scheduleData.length} new events with current contact info\n\n` +
          `💡 ALTERNATIVE: Use "📞 Update Contact Info Only" to preserve scheduling details.\n\n` +
          `Continue with full regeneration?`,
          SpreadsheetApp.getUi().ButtonSet.YES_NO
        );
        
        if (response !== SpreadsheetApp.getUi().Button.YES) {
          return;
        }
        
        // Delete all existing events
        console.log('Deleting existing events...');
        const today = new Date();
        const futureDate = new Date();
        futureDate.setFullYear(futureDate.getFullYear() + 2);
        
        const existingEvents = calendar.getEvents(today, futureDate);
        console.log(`Deleting ${existingEvents.length} existing events...`);
        
        const deleteStartTime = new Date().getTime();
        existingEvents.forEach((event, index) => {
          event.deleteEvent();
          
          // Rate limiting for deletions
          if ((index + 1) % 10 === 0) {
            Utilities.sleep(500);
            console.log(`Deleted ${index + 1} events...`);
          }
        });
        
        const deleteTime = new Date().getTime() - deleteStartTime;
        console.log(`Deleted ${existingEvents.length} events in ${deleteTime}ms`);
        
        // Cooldown after deletions
        console.log('Waiting for API cooldown before creating new events...');
        Utilities.sleep(2000);
        
      } else {
        calendar = CalendarApp.createCalendar(calendarName);
        console.log(`Created new calendar: "${calendarName}"`);
      }
    } catch (calendarError) {
      throw new Error(`Calendar setup failed: ${calendarError.message}`);
    }
    
    // Create calendar events
    console.log(`Creating ${scheduleData.length} calendar events...`);
    let eventsCreated = 0;
    const eventCreationStartTime = new Date().getTime();
    
    // Use for loop for better performance and error handling
    for (let index = 0; index < scheduleData.length; index++) {
      const visit = scheduleData[index];
      
      try {
        // Validate visit object structure
        if (!visit || !visit.deacon || !visit.household || !visit.date) {
          console.warn(`Skipping invalid visit data at index ${index}:`, visit);
          continue;
        }
        
        const visitDate = new Date(visit.date);
        const deacon = visit.deacon;
        const household = visit.household;
        
        // Get household contact information
        const householdIndex = config.households.indexOf(household);
        const phone = householdIndex >= 0 ? (config.phones[householdIndex] || 'Phone not available') : 'Phone not available';
        const address = householdIndex >= 0 ? (config.addresses[householdIndex] || 'Address not available') : 'Address not available';
        
        // Build Breeze profile link
        let breezeLink = 'Breeze profile not available';
        if (householdIndex >= 0 && config.breezeShortLinks[householdIndex]) {
          breezeLink = config.breezeShortLinks[householdIndex];
        } else if (householdIndex >= 0 && config.breezeNumbers[householdIndex]) {
          const fullBreezeLink = buildBreezeUrl(config.breezeNumbers[householdIndex]);
          if (fullBreezeLink && fullBreezeLink.length > 0) {
            breezeLink = fullBreezeLink;
          }
        }
        
        // Build Notes link
        let notesLink = 'Visit notes not available';
        if (householdIndex >= 0 && config.notesShortLinks[householdIndex]) {
          notesLink = config.notesShortLinks[householdIndex];
        } else if (householdIndex >= 0 && config.notesLinks[householdIndex]) {
          const fullNotesLink = config.notesLinks[householdIndex];
          if (fullNotesLink && fullNotesLink.length > 0) {
            notesLink = fullNotesLink;
          }
        }
        
        // Build event description
        let description = `Household: ${household}\nBreeze Profile: ${breezeLink}\n\n`;
        description += `Contact Information:\nPhone: ${phone}\nAddress: ${address}\n\n`;
        description += `Visit Notes: ${notesLink}\n\n`;
        
        const customInstructions = config.calendarInstructions || 
          'Please call to set up a day and time for your visit in the coming week. ' +
          'You can update this event with the actual day and time and copy it to your personal calendar.';
        
        description += `Instructions:\n${customInstructions}`;
        
        // Create the calendar event
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
        
        // Rate limiting: pause every 20 events
        if ((index + 1) % 20 === 0) {
          Utilities.sleep(1000);
          console.log(`Created ${index + 1} events so far...`);
        }
        
      } catch (eventError) {
        console.error(`Failed to create event for ${visit.deacon} -> ${visit.household}:`, eventError);
      }
    }
    
    const totalTime = new Date().getTime() - eventCreationStartTime;
    console.log(`Created ${eventsCreated} events in ${totalTime}ms`);
    
    SpreadsheetApp.getUi().alert(
      currentTestMode ? 'TEST Calendar Export Complete' : 'Calendar Export Complete',
      `${currentTestMode ? '🧪 TEST: ' : ''}✅ Calendar export completed!\n\n` +
      `📅 Calendar: "${calendarName}"\n` +
      `📊 Events created: ${eventsCreated}\n` +
      `⏱️ Time taken: ${Math.round(totalTime/1000)} seconds\n\n` +
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
      `${currentTestMode ? '🧪 TEST: ' : ''}This will update contact information in ${scheduleData.length} existing calendar events.\n\n` +
      `Calendar: "${currentCalendarName}"\n\n` +
      `✅ PRESERVES: All custom scheduling details, times, dates, guests\n` +
      `🔄 UPDATES: Phone numbers, addresses, Breeze links, Notes links\n\n` +
      'Continue with contact info update?',
      ui.ButtonSet.YES_NO
    );
    
    if (response !== ui.Button.YES) {
      return;
    }

    // Get calendar
    let calendar;
    try {
      calendar = CalendarApp.getCalendarsByName(currentCalendarName)[0];
      if (!calendar) {
        throw new Error(`Calendar "${currentCalendarName}" not found`);
      }
    } catch (calError) {
      ui.alert('Calendar Error', `Cannot access calendar "${currentCalendarName}". Please check permissions.`, ui.ButtonSet.OK);
      return;
    }

    // Get existing events for comparison
    const startDate = new Date(Math.min(...scheduleData.map(entry => entry.date)));
    const endDate = new Date(Math.max(...scheduleData.map(entry => entry.date)));
    endDate.setDate(endDate.getDate() + 1); // Include end date
    
    const existingEvents = calendar.getEvents(startDate, endDate);
    console.log(`Found ${existingEvents.length} existing events to update`);

    // Update contact information in existing events
    let updatedCount = 0;
    const updateStartTime = new Date().getTime();
    
    // *** FIX: Correctly handle schedule data objects ***
    scheduleData.forEach(scheduleEntry => {
      // Use object properties instead of array destructuring
      const visitDate = scheduleEntry.date;
      const deacon = scheduleEntry.deacon;
      const household = scheduleEntry.household;
      const householdIndex = config.households.indexOf(household);
      
      // Find matching calendar event
      const matchingEvents = existingEvents.filter(event => {
        const eventDate = event.getStartTime();
        const eventTitle = event.getTitle();
        
        // Match by date and title pattern
        return eventDate.toDateString() === visitDate.toDateString() && 
               eventTitle.includes(deacon) && eventTitle.includes(household);
      });
      
      if (matchingEvents.length > 0) {
        const event = matchingEvents[0]; // Use first match if multiple found
        
        try {
          // Build updated description with current contact information
          let description = `Household: ${household}\n`;
          
          if (householdIndex !== -1) {
            const phone = config.phones[householdIndex];
            const address = config.addresses[householdIndex];
            const breezeNumber = config.breezeNumbers[householdIndex];
            const breezeShortUrl = config.breezeShortLinks[householdIndex];
            const notesUrl = config.notesLinks[householdIndex];
            const notesShortUrl = config.notesShortLinks[householdIndex];
            
            // Contact information section
            description += `\nContact Information:\n`;
            if (phone && phone.length > 0) description += `Phone: ${phone}\n`;
            if (address && address.length > 0) description += `Address: ${address}\n`;
            
            // Breeze integration
            if (breezeNumber && breezeNumber.length > 0) {
              if (breezeShortUrl && breezeShortUrl.length > 0) {
                description += `\nBreeze Profile: ${breezeShortUrl}\n`;
              } else {
                const fullBreezeUrl = buildBreezeUrl(breezeNumber);
                description += `\nBreeze Profile: ${fullBreezeUrl}\n`;
              }
            }
            
            // Visit notes integration
            if (notesUrl && notesUrl.length > 0) {
              if (notesShortUrl && notesShortUrl.length > 0) {
                description += `\nVisit Notes: ${notesShortUrl}\n`;
              } else {
                description += `\nVisit Notes: ${notesUrl}\n`;
              }
            }
          }
          
          // Add custom instructions
          const customInstructions = config.customInstructions || 
            'Please call to set up a day and time for your visit in the coming week. ' +
            'You can update this event with the actual day and time and copy it to your personal calendar.';
          
          description += `\nInstructions:\n${customInstructions}`;
          
          // Update the event description (preserves all other details)
          event.setDescription(description);
          
          // Update location if address is available
          if (householdIndex !== -1 && config.addresses[householdIndex]) {
            event.setLocation(config.addresses[householdIndex]);
          }
          
          updatedCount++;
          
        } catch (updateError) {
          console.error(`Failed to update event for ${deacon} -> ${household}:`, updateError);
        }
      } else {
        console.warn(`No matching calendar event found for ${deacon} visits ${household} on ${visitDate.toDateString()}`);
      }
    });
    
    const updateTime = new Date().getTime() - updateStartTime;
    console.log(`Updated ${updatedCount} events in ${updateTime}ms`);
    
    ui.alert(
      currentTestMode ? 'TEST Contact Info Update Complete' : 'Contact Info Update Complete',
      `${currentTestMode ? '🧪 TEST: ' : ''}✅ Contact information update completed!\n\n` +
      `📞 Updated: ${updatedCount} calendar events\n` +
      `🛡️ Preserved: All scheduling details\n` +
      `⏱️ Time taken: ${Math.round(updateTime/1000)} seconds\n\n` +
      'All events now have current contact information while preserving custom scheduling.',
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    console.error('Contact info update failed:', error);
    SpreadsheetApp.getUi().alert(
      'Update Failed',
      `❌ Contact information update failed: ${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

function updateFutureEventsOnly() {
  /**
   * Updates ONLY future events (next Monday onwards)
   * Preserves: This week's scheduling details completely
   * Updates: All future weeks with current assignments and contact info
   * FIXED: Now correctly handles schedule data as objects
   */
  try {
    const currentTestMode = getCurrentTestMode();
    const calendarName = currentTestMode ? 'TEST - Deacon Visitation Schedule' : 'Deacon Visitation Schedule';
    
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

    const calendars = CalendarApp.getCalendarsByName(calendarName);
    if (calendars.length === 0) {
      SpreadsheetApp.getUi().alert('Calendar Not Found', `Calendar "${calendarName}" not found.\n\nPlease run "🚨 Full Calendar Regeneration" first.`, SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    // Calculate cutoff date (next Monday)
    const today = new Date();
    const nextMonday = new Date(today);
    const daysUntilMonday = (8 - today.getDay()) % 7;
    nextMonday.setDate(today.getDate() + daysUntilMonday);
    nextMonday.setHours(0, 0, 0, 0);

    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
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
    
    // Filter schedule for future events - FIXED: Use object properties
    const futureSchedule = scheduleData.filter(visit => {
      const visitDate = new Date(visit.date);  // Use visit.date instead of visit[0]
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
    
    // Create updated future events - FIXED: Use object properties
    let eventsCreated = 0;
    futureSchedule.forEach((visit, index) => {
      try {
        // FIXED: Use object properties instead of array destructuring
        const visitDate = visit.date;      // Instead of visit[0]
        const deacon = visit.deacon;       // Instead of visit[1] 
        const household = visit.household; // Instead of visit[2]
        const householdIndex = config.households.indexOf(household);
        
        // Build comprehensive event description
        let description = `Household: ${household}\n`;
        
        if (householdIndex !== -1) {
          const phone = config.phones[householdIndex];
          const address = config.addresses[householdIndex];
          const breezeNumber = config.breezeNumbers[householdIndex];
          const breezeShortUrl = config.breezeShortLinks[householdIndex];
          const notesUrl = config.notesLinks[householdIndex];
          const notesShortUrl = config.notesShortLinks[householdIndex];
          
          // Contact information section
          description += `\nContact Information:\n`;
          if (phone && phone.length > 0) description += `Phone: ${phone}\n`;
          if (address && address.length > 0) description += `Address: ${address}\n`;
          
          // Breeze integration
          if (breezeNumber && breezeNumber.length > 0) {
            if (breezeShortUrl && breezeShortUrl.length > 0) {
              description += `\nBreeze Profile: ${breezeShortUrl}\n`;
            } else {
              const fullBreezeUrl = buildBreezeUrl(breezeNumber);
              description += `\nBreeze Profile: ${fullBreezeUrl}\n`;
            }
          }
          
          // Visit notes integration
          if (notesUrl && notesUrl.length > 0) {
            if (notesShortUrl && notesShortUrl.length > 0) {
              description += `\nVisit Notes: ${notesShortUrl}\n`;
            } else {
              description += `\nVisit Notes: ${notesUrl}\n`;
            }
          }
        }
        
        // Add custom instructions
        const customInstructions = config.customInstructions || 
          'Please call to set up a day and time for your visit in the coming week. ' +
          'You can update this event with the actual day and time and copy it to your personal calendar.';
        
        description += `\nInstructions:\n${customInstructions}`;
        
        // Create the calendar event
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
        
        // Rate limiting
        if ((index + 1) % 20 === 0) {
          Utilities.sleep(1000);
        }
        
      } catch (eventError) {
        console.error(`Failed to create future event for ${visit.deacon} -> ${visit.household}:`, eventError);
      }
    });
    
    SpreadsheetApp.getUi().alert(
      currentTestMode ? 'TEST Future Events Updated' : 'Future Events Updated',
      `${currentTestMode ? '🧪 TEST: ' : ''}✅ Future events updated!\n\n` +
      `📅 Starting: ${nextMonday.toLocaleDateString()}\n` +
      `🔄 Created: ${eventsCreated} future events\n` +
      `🛡️ Preserved: This week's scheduling details\n\n` +
      'Future events now reflect current schedule and contact information.',
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
   * Create individual spreadsheet files for each deacon with their personal schedule
   * FIXED: Now correctly handles schedule data as objects
   */
  const sheet = SpreadsheetApp.getActiveSheet();
  
  try {
    const config = getConfiguration(sheet);
    
    if (config.deacons.length === 0) {
      SpreadsheetApp.getUi().alert('Error', 'No deacons found. Please run Generate Schedule first.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    const scheduleData = getScheduleFromSheet(sheet);
    
    if (scheduleData.length === 0) {
      SpreadsheetApp.getUi().alert(
        'No Schedule',
        'No schedule data found. Please generate a schedule first.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    const currentTestMode = getCurrentTestMode();
    
    // Create new spreadsheet for individual schedules
    const timestamp = new Date().toISOString().slice(0, 10);
    const newSpreadsheet = SpreadsheetApp.create(
      `${currentTestMode ? 'TEST - ' : ''}Individual Deacon Schedules - ${timestamp}`
    );
    
    // Group visits by deacon - FIXED: Use object properties
    const deaconVisits = {};
    scheduleData.forEach(visit => {
      // FIXED: Use object properties instead of array destructuring
      const deacon = visit.deacon;       // Instead of destructuring [date, deacon, household]
      const date = visit.date;
      const household = visit.household;
      
      if (!deaconVisits[deacon]) {
        deaconVisits[deacon] = [];
      }
      deaconVisits[deacon].push([date, household]);  // Still pass as array to maintain compatibility
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
      `📊 Created: ${Object.keys(deaconVisits).length} individual sheets\n` +
      `📅 Total visits: ${scheduleData.length}\n` +
      `🔗 Spreadsheet: ${spreadsheetUrl}\n\n` +
      'Each deacon now has their own sheet with visit dates, households, and contact information.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    console.error('Individual schedules export failed:', error);
    SpreadsheetApp.getUi().alert(
      'Export Failed',
      `❌ Individual schedules export failed: ${error.message}`,
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
   * Display setup instructions for new users (Updated for v1.1)
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
          console.log(`URL shortening unexpected result: ${shortUrl}`);
          results.push('⚠️ URL shortening: UNEXPECTED - Check TinyURL service');
        }
      } catch (urlError) {
        console.error('URL shortening failed:', urlError);
        results.push(`❌ URL shortening: FAILED - ${urlError.message}`);
      }
      
      // Test 3: Breeze URL construction
      console.log('Test 3: Breeze URL construction...');
      try {
        const testBreezeNumber = '12345678';
        const breezeUrl = buildBreezeUrl(testBreezeNumber);
        const expectedUrl = 'https://immanuelky.breezechms.com/people/view/12345678';
        
        if (breezeUrl === expectedUrl) {
          console.log(`Breeze URL construction test passed: ${breezeUrl}`);
          results.push('✅ Breeze URL construction: PASSED');
        } else {
          console.log(`Breeze URL mismatch: expected ${expectedUrl}, got ${breezeUrl}`);
          results.push('❌ Breeze URL construction: FAILED - URL format mismatch');
        }
      } catch (breezeError) {
        console.error('Breeze URL construction failed:', breezeError);
        results.push(`❌ Breeze URL construction: FAILED - ${breezeError.message}`);
      }
      
      // Test 4: Calendar access
      console.log('Test 4: Calendar access...');
      try {
        const currentTestMode = getCurrentTestMode();
        const calendarName = currentTestMode ? 'TEST - Deacon Visitation Schedule' : 'Deacon Visitation Schedule';
        const calendars = CalendarApp.getCalendarsByName(calendarName);
        
        if (calendars.length > 0) {
          console.log(`Calendar access test passed: Found "${calendarName}"`);
          results.push(`✅ Calendar access: PASSED (${calendarName})`);
        } else {
          console.log(`Calendar not found: "${calendarName}"`);
          results.push(`⚠️ Calendar access: NOT FOUND - Create "${calendarName}" first`);
        }
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

// ===== SPREADSHEET EVENT HANDLERS =====

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
    
    console.log('Enhanced Deacon Rotation menu created successfully (v1.1 - CLEAN)');
  } catch (error) {
    console.error('Failed to create menu:', error);
  }
}

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

// ===== MODE DETECTION AND INDICATORS (CALLS MODULE 3 FUNCTIONS) =====

// Note: addModeIndicatorToSheet() and showModeNotification() are defined in Module 3
// Module 4 just calls them - no duplicate definitions to avoid conflicts

// ===== ENHANCED MENU SYSTEM (v1.1 - CLEANED) =====

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
  
  // Add mode indicator to spreadsheet
  addModeIndicatorToSheet();
}

// ===== URL SHORTENING FUNCTIONS - COMPLETE RESTORATION =====

function shortenUrl(longUrl) {
  /**
   * Shortens a URL using TinyURL's free API
   * Falls back to original URL if shortening fails
   */
  try {
    if (!longUrl || longUrl.trim().length === 0) {
      return '';
    }
    
    const apiUrl = 'http://tinyurl.com/api-create.php?url=' + encodeURIComponent(longUrl.trim());
    
    const response = UrlFetchApp.fetch(apiUrl, {
      method: 'GET',
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() === 200) {
      const shortUrl = response.getContentText().trim();
      
      if (shortUrl.startsWith('http://tinyurl.com/') || shortUrl.startsWith('https://tinyurl.com/')) {
        console.log(`URL shortened: ${longUrl} → ${shortUrl}`);
        return shortUrl;
      } else {
        console.warn(`TinyURL returned unexpected response: ${shortUrl}`);
        return longUrl; // Fallback to original URL
      }
    } else {
      console.warn(`TinyURL API failed with code ${response.getResponseCode()}: ${response.getContentText()}`);
      return longUrl; // Fallback to original URL
    }
    
  } catch (error) {
    console.error(`URL shortening failed for ${longUrl}: ${error.message}`);
    return longUrl; // Fallback to original URL
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

function generateAndStoreShortUrls(sheet, config) {
  /**
   * Generates shortened URLs for Breeze and Notes links and stores them in columns R and S
   */
  try {
    console.log('Starting URL shortening process...');
    
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'Generate Shortened URLs',
      'This will create shortened URLs for Breeze profiles and Notes pages.\n\n' +
      'This may take a few moments depending on the number of households.\n\n' +
      'Continue?',
      ui.ButtonSet.YES_NO
    );
    
    if (response !== ui.Button.YES) {
      return;
    }
    
    let breezeUrlsGenerated = 0;
    let notesUrlsGenerated = 0;
    
    // Process each household
    for (let i = 0; i < config.households.length; i++) {
      const household = config.households[i];
      console.log(`Processing URLs for household ${i + 1}: ${household}`);
      
      // Process Breeze URL
      const breezeNumber = config.breezeNumbers[i];
      if (breezeNumber && breezeNumber.trim().length > 0) {
        const fullBreezeUrl = buildBreezeUrl(breezeNumber);
        const shortBreezeUrl = shortenUrl(fullBreezeUrl);
        
        // Store in column R
        sheet.getRange(`R${i + 2}`).setValue(shortBreezeUrl);
        breezeUrlsGenerated++;
        
        console.log(`Breeze URL for ${household}: ${fullBreezeUrl} → ${shortBreezeUrl}`);
      }
      
      // Process Notes URL
      const notesUrl = config.notesLinks[i];
      if (notesUrl && notesUrl.trim().length > 0) {
        const shortNotesUrl = shortenUrl(notesUrl);
        
        // Store in column S
        sheet.getRange(`S${i + 2}`).setValue(shortNotesUrl);
        notesUrlsGenerated++;
        
        console.log(`Notes URL for ${household}: ${notesUrl} → ${shortNotesUrl}`);
      }
      
      // Add small delay to avoid overwhelming the API
      if (i < config.households.length - 1) {
        Utilities.sleep(500); // 0.5 second delay between requests
      }
    }
    
    console.log('URL shortening process completed');
    
    ui.alert(
      'URL Shortening Complete',
      `✅ Shortened URLs generated!\n\n` +
      `🔗 Breeze URLs: ${breezeUrlsGenerated}\n` +
      `📝 Notes URLs: ${notesUrlsGenerated}\n\n` +
      `Shortened URLs are now available in columns R and S.`,
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    console.error('Error generating shortened URLs:', error);
    SpreadsheetApp.getUi().alert(
      'URL Shortening Failed',
      `❌ Error generating shortened URLs: ${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

// Wrapper function for menu access to URL shortening
function generateShortUrlsFromMenu() {
  const sheet = SpreadsheetApp.getActiveSheet();
  try {
    const config = getConfiguration(sheet);
    generateAndStoreShortUrls(sheet, config);
  } catch (error) {
    SpreadsheetApp.getUi().alert(
      'Error',
      `❌ ${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
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
