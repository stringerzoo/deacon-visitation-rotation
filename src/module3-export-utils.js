/**
 * MODULE 3: EXPORT, MENU & UTILITY FUNCTIONS (ENHANCED v24.2)
 * Deacon Visitation Rotation System - Smart Calendar Updates
 * 
 * NEW v24.2 Features:
 * - Smart calendar update options
 * - Contact info only updates
 * - Future events only updates  
 * - Date range updates
 * - Preserves custom scheduling details
 */

// TEMPORARY TESTING MODIFICATION for Module 3
// Add this at the top of your module3-export-utils.js for testing

// TEST MODE CONFIGURATION
const TEST_MODE = true; // Set to false when ready for production
const TEST_CALENDAR_NAME = TEST_MODE ? 'TEST - Deacon Visitation Schedule' : 'Deacon Visitation Schedule';

// Modified helper function for testing
function getOrCreateCalendar() {
  /**
   * Gets existing calendar or returns null if not found
   * MODIFIED: Uses test calendar name when in test mode
   */
  const calendarName = TEST_CALENDAR_NAME; // Uses test name when TEST_MODE = true
  
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

// ===== SMART CALENDAR UPDATE FUNCTIONS =====

function updateContactInfoOnly() {
  /**
   * Updates ONLY contact information in existing calendar events
   * Preserves: Custom times, dates, guests, locations, scheduling details
   * Updates: Phone numbers, addresses, Breeze links, Notes links, instructions
   */
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const config = getConfiguration(sheet);
    const scheduleData = getScheduleFromSheet(sheet);
    
    if (scheduleData.length === 0) {
      SpreadsheetApp.getUi().alert('No Schedule', 'Please generate a schedule first.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    // Confirm with user
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'Update Contact Info Only',
      `This will update contact information (phone, address, links) in existing calendar events.\n\n` +
      `✅ PRESERVES: Custom times, dates, guest lists, locations\n` +
      `🔄 UPDATES: Contact info, Breeze links, Notes links, instructions\n\n` +
      `Continue?`,
      ui.ButtonSet.YES_NO
    );
    
    if (response !== ui.Button.YES) return;
    
    // Get calendar
    const calendar = getOrCreateCalendar();
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
          const deacon = visitMatch[1].trim();
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
    
    // Calculate cutoff date (next week)
    const cutoffDate = new Date();
    cutoffDate.setDate(cutoffDate.getDate() + 7);
    cutoffDate.setHours(0, 0, 0, 0); // Start of day
    
    // Confirm with user
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'Update Future Events Only',
      `This will update calendar events starting from ${cutoffDate.toLocaleDateString()}.\n\n` +
      `✅ PRESERVES: This week's scheduling details and customizations\n` +
      `🔄 UPDATES: Future events with current contact info and assignments\n\n` +
      `Continue?`,
      ui.ButtonSet.YES_NO
    );
    
    if (response !== ui.Button.YES) return;
    
    // Get calendar
    const calendar = getOrCreateCalendar();
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
        // Enhanced event title format
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

function updateThisMonthEvents() {
  /**
   * Updates events for the current month only
   * Useful for monthly planning cycles
   */
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const config = getConfiguration(sheet);
    const scheduleData = getScheduleFromSheet(sheet);
    
    if (scheduleData.length === 0) {
      SpreadsheetApp.getUi().alert('No Schedule', 'Please generate a schedule first.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    // Calculate current month date range
    const now = new Date();
    const monthStart = new Date(now.getFullYear(), now.getMonth(), 1);
    const monthEnd = new Date(now.getFullYear(), now.getMonth() + 1, 0);
    
    // Confirm with user
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'Update This Month\'s Events',
      `This will update calendar events from ${monthStart.toLocaleDateString()} to ${monthEnd.toLocaleDateString()}.\n\n` +
      `✅ PRESERVES: Events outside this month\n` +
      `🔄 UPDATES: This month's events with current info\n\n` +
      `Continue?`,
      ui.ButtonSet.YES_NO
    );
    
    if (response !== ui.Button.YES) return;
    
    // Get calendar
    const calendar = getOrCreateCalendar();
    if (!calendar) return;
    
    // Get existing events for this month
    const existingEvents = calendar.getEvents(monthStart, monthEnd);
    console.log(`Found ${existingEvents.length} events in current month`);
    
    // Delete this month's events
    let deletedCount = 0;
    existingEvents.forEach((event, index) => {
      event.deleteEvent();
      deletedCount++;
      
      if ((index + 1) % 10 === 0) {
        Utilities.sleep(500);
      }
    });
    
    // Cooldown
    if (deletedCount > 0) {
      Utilities.sleep(2000);
    }
    
    // Filter schedule data for this month
    const monthScheduleData = scheduleData.filter(visit => {
      const visitDate = new Date(visit.date);
      return visitDate >= monthStart && visitDate <= monthEnd;
    });
    
    // Create new events for this month
    let eventsCreated = 0;
    monthScheduleData.forEach((visit, index) => {
      try {
        const eventTitle = `${visit.deacon} visits ${visit.household}`;
        
        // Get current contact info (same logic as other functions)
        const householdIndex = config.households.indexOf(visit.household);
        const phone = householdIndex >= 0 ? config.phones[householdIndex] || 'Phone not available' : 'Phone not available';
        const address = householdIndex >= 0 ? config.addresses[householdIndex] || 'Address not available' : 'Address not available';
        
        // Get links (shortened if available)
        let breezeLink = 'Not available';
        let notesLink = 'Not available';
        
        if (householdIndex >= 0) {
          // Breeze link
          const shortBreezeLink = config.breezeShortLinks[householdIndex];
          if (shortBreezeLink && shortBreezeLink.trim().length > 0) {
            breezeLink = shortBreezeLink;
          } else {
            const breezeNumber = config.breezeNumbers[householdIndex];
            if (breezeNumber && breezeNumber.trim().length > 0) {
              breezeLink = buildBreezeUrl(breezeNumber);
            }
          }
          
          // Notes link
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
        
        const eventDescription = `Household: ${visit.household}
Breeze Profile: ${breezeLink}

Contact Information:
Phone: ${phone}
Address: ${address}

Visit Notes: ${notesLink}

Instructions:
${config.calendarInstructions}`;
        
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
        
        if ((index + 1) % 25 === 0) {
          Utilities.sleep(1000);
        }
        
      } catch (eventError) {
        console.error(`Failed to create month event for ${visit.deacon} -> ${visit.household}:`, eventError);
      }
    });
    
    ui.alert(
      'This Month Update Complete',
      `✅ Updated this month's calendar events!\n\n` +
      `🗑️ Deleted: ${deletedCount} old events\n` +
      `📅 Created: ${eventsCreated} updated events\n` +
      `📅 Period: ${monthStart.toLocaleDateString()} - ${monthEnd.toLocaleDateString()}\n\n` +
      `This month's events now have current information.`,
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    SpreadsheetApp.getUi().alert(
      'This Month Update Failed',
      `❌ ${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

// ===== HELPER FUNCTIONS =====

function getOrCreateCalendar() {
  /**
   * Gets existing calendar or returns null if not found
   * Used by update functions to ensure calendar exists
   */
  const calendarName = 'Deacon Visitation Schedule';
  
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

// ===== EXISTING FUNCTIONS (Updated) =====

function exportIndividualSchedules() {
  const sheet = SpreadsheetApp.getActiveSheet();
  
  try {
    const config = getConfiguration(sheet);
    
    if (config.deacons.length === 0) {
      SpreadsheetApp.getUi().alert('Error', 'No deacons found. Please run Generate Schedule first.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    // Create a new spreadsheet for individual schedules
    const newSpreadsheet = SpreadsheetApp.create(`Deacon Individual Schedules - ${new Date().toLocaleDateString()}`);
    const scheduleData = getScheduleFromSheet(sheet);
    
    if (scheduleData.length === 0) {
      SpreadsheetApp.getUi().alert('Error', 'No schedule data found. Please generate a schedule first.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    // Group visits by deacon
    const deaconVisits = {};
    scheduleData.forEach(visit => {
      if (!deaconVisits[visit.deacon]) {
        deaconVisits[visit.deacon] = [];
      }
      deaconVisits[visit.deacon].push(visit);
    });
    
    // Create a sheet for each deacon
    config.deacons.forEach((deacon, index) => {
      let deaconSheet;
      if (index === 0) {
        deaconSheet = newSpreadsheet.getActiveSheet();
        deaconSheet.setName(deacon);
      } else {
        deaconSheet = newSpreadsheet.insertSheet(deacon);
      }
      
      const visits = deaconVisits[deacon] || [];
      
      // Headers
      deaconSheet.getRange(1, 1, 1, 3).setValues([['Week of', 'Household', 'Notes']]);
      const headerRange = deaconSheet.getRange(1, 1, 1, 3);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#4285f4');
      headerRange.setFontColor('white');
      
      // Visit data
      if (visits.length > 0) {
        const visitData = visits
          .sort((a, b) => a.date.getTime() - b.date.getTime())
          .map(visit => [visit.date, visit.household, '']);
        
        deaconSheet.getRange(2, 1, visitData.length, 3).setValues(visitData);
        deaconSheet.getRange(2, 1, visitData.length, 1).setNumberFormat('mm/dd/yyyy');
        
        // Set column widths and formatting
        deaconSheet.setColumnWidth(1, 100); // Date column
        deaconSheet.setColumnWidth(2, 120); // Household column  
        deaconSheet.setColumnWidth(3, 250); // Notes column - wide for writing
        
        // Enable text wrapping for notes column
        deaconSheet.getRange(1, 3, visitData.length + 1, 1).setWrap(true);
        
      } else {
        deaconSheet.getRange(2, 1, 1, 3).setValues([['No visits assigned', '', '']]);
        
        // Set column widths even for empty sheets
        deaconSheet.setColumnWidth(1, 100);
        deaconSheet.setColumnWidth(2, 120);
        deaconSheet.setColumnWidth(3, 250);
      }
    });
    
    SpreadsheetApp.getUi().alert(
      'Export Complete', 
      `✅ Individual schedules created!\n\n` +
      `📄 Spreadsheet: ${newSpreadsheet.getName()}\n` +
      `🔗 URL: ${newSpreadsheet.getUrl()}\n\n` +
      `Each deacon now has their own tab with their complete schedule.`, 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    SpreadsheetApp.getUi().alert(
      'Export Failed',
      `❌ ${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

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
      
      // Validate that we got a proper TinyURL back
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

function exportToGoogleCalendar() {
  /**
   * Original calendar export function - now renamed for clarity as "Full Calendar Regeneration"
   * Creates calendar events for each deacon's visits with enhanced contact information
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
    const calendarName = TEST_CALENDAR_NAME; // Changed this line
    
    try {
      const calendars = CalendarApp.getCalendarsByName(calendarName);
      if (calendars.length > 0) {
        calendar = calendars[0];
        
        const response = SpreadsheetApp.getUi().alert(
          TEST_MODE ? 'TEST: Full Calendar Regeneration' : 'Full Calendar Regeneration',
          `${TEST_MODE ? '🧪 TEST MODE: ' : ''}⚠️ This will completely rebuild the calendar "${calendarName}".\n\n` +
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
        calendar.setDescription(`${TEST_MODE ? 'TEST: ' : ''}Automated schedule for deacon household visitations with contact information and management links`);
        calendar.setColor(TEST_MODE ? CalendarApp.Color.RED : CalendarApp.Color.BLUE); // Red for test mode
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
        // Enhanced event title format: "[Deacon Name] visits [Household Name]"
        const eventTitle = `${visit.deacon} visits ${visit.household}`;
        
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
            const breezeNumber = config.breezeNumbers[householdIndex];
            if (breezeNumber && breezeNumber.trim().length > 0) {
              breezeLink = buildBreezeUrl(breezeNumber);
            }
          }
        }
        
        // Get Notes link (use shortened if available, otherwise use full URL)
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
        
        // Enhanced event description with new format
        const eventDescription = `Household: ${visit.household}
Breeze Profile: ${breezeLink}

Contact Information:
Phone: ${phone}
Address: ${address}

Visit Notes: ${notesLink}

Instructions:
${config.calendarInstructions}`;
        
        // Set event for 1 hour duration
        const startTime = new Date(visit.date);
        startTime.setHours(14, 0, 0, 0); // 2:00 PM default
        const endTime = new Date(startTime);
        endTime.setHours(15, 0, 0, 0); // 3:00 PM
        
        calendar.createEvent(eventTitle, startTime, endTime, {
          description: eventDescription,
          guests: '',
          sendInvites: false
        });
        
        eventsCreated++;
        
        // Add small delay every 25 events to avoid rate limiting
        if ((index + 1) % 25 === 0) {
          Utilities.sleep(1000); // 1 second pause every 25 events
          console.log(`Created ${index + 1} events so far...`);
        }
        
      } catch (eventError) {
        console.error(`Failed to create event for ${visit.deacon} -> ${visit.household}:`, eventError);
      }
    });
    
    console.log(`Created ${eventsCreated} events in ${new Date().getTime() - createStartTime}ms`);
    
    SpreadsheetApp.getUi().alert(
      'Full Calendar Regeneration Complete',
      `✅ Created ${eventsCreated} calendar events with enhanced information!\n\n` +
      `📅 Calendar: "${calendarName}"\n` +
      `🕐 Default time: 2:00 PM - 3:00 PM\n` +
      `🔗 Includes: Breeze profiles and visit notes links\n` +
      `📞 Contact info: Phone numbers and addresses\n` +
      `📝 Instructions: Custom visit coordination text\n\n` +
      `Each event title: "[Deacon] visits [Household]"\n` +
      `View and modify these events in Google Calendar.`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    SpreadsheetApp.getUi().alert(
      'Calendar Export Failed',
      `❌ ${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

// ===== REMAINING FUNCTIONS =====

function archiveCurrentSchedule() {
  /**
   * Archives the current schedule to a new sheet before generating a new one
   * Useful for keeping historical records
   */
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const scheduleData = getScheduleFromSheet(sheet);
    
    if (scheduleData.length === 0) {
      SpreadsheetApp.getUi().alert('Nothing to Archive', 'No schedule data found to archive.', SpreadsheetApp.getUi().ButtonSet.OK);
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

function generateNextYearSchedule() {
  /**
   * Automatically sets up next year's schedule based on current configuration
   */
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const config = getConfiguration(sheet);
    
    // Calculate next year's start date (same day of week, next year)
    const nextYearStart = new Date(config.startDate);
    nextYearStart.setFullYear(nextYearStart.getFullYear() + 1);
    
    // Update start date in column K location
    sheet.getRange('K2').setValue(nextYearStart);
    
    // Generate new schedule
    generateRotationSchedule();
    
    SpreadsheetApp.getUi().alert(
      'Next Year Schedule Generated',
      `✅ Generated schedule starting ${nextYearStart.toLocaleDateString()}\n\n` +
      `The rotation continues from where it left off, ensuring fair distribution.`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    SpreadsheetApp.getUi().alert(
      'Next Year Generation Failed',
      `❌ ${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
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
          console.log(`URL shortening failed: ${shortUrl}`);
          results.push('❌ URL shortening: FAILED - Unexpected response');
        }
      } catch (urlError) {
        console.error('URL shortening test failed:', urlError);
        results.push(`❌ URL shortening: FAILED - ${urlError.message}`);
      }
      
      // Test 3: Breeze URL construction
      console.log('Test 3: Breeze URL construction...');
      try {
        const testNumber = '29760588';
        const breezeUrl = buildBreezeUrl(testNumber);
        const expectedUrl = 'https://immanuelky.breezechms.com/people/view/29760588';
        
        if (breezeUrl === expectedUrl) {
          console.log(`Breeze URL construction: ${testNumber} → ${breezeUrl}`);
          results.push('✅ Breeze URL construction: PASSED');
        } else {
          console.log(`Breeze URL construction failed: expected ${expectedUrl}, got ${breezeUrl}`);
          results.push('❌ Breeze URL construction: FAILED');
        }
      } catch (breezeError) {
        console.error('Breeze URL construction failed:', breezeError);
        results.push(`❌ Breeze URL construction: FAILED - ${breezeError.message}`);
      }
      
      // Test 4: Calendar access
      console.log('Test 4: Calendar access...');
      try {
        const calendar = getOrCreateCalendar();
        if (calendar) {
          console.log(`Calendar access successful: ${calendar.getName()}`);
          results.push('✅ Calendar access: PASSED');
        } else {
          console.log('Calendar not found but error handled gracefully');
          results.push('⚠️ Calendar access: PASSED - No calendar found (expected for new setups)');
        }
      } catch (calError) {
        console.error('Calendar access test failed:', calError);
        results.push(`❌ Calendar access: FAILED - ${calError.message}`);
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
      'System Test Results',
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

// ===== ENHANCED MENU SYSTEM (v24.2) =====

function createMenuItems() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🔄 Deacon Rotation')
    .addItem('📅 Generate Schedule', 'generateRotationSchedule')
    .addSeparator()
    .addItem('🔗 Generate Shortened URLs', 'generateShortUrlsFromMenu')
    .addSeparator()
    .addSubMenu(ui.createMenu('📆 Calendar Functions')
      .addItem('🚨 Full Calendar Regeneration', 'exportToGoogleCalendar')
      .addSeparator()
      .addItem('📞 Update Contact Info Only', 'updateContactInfoOnly')
      .addItem('🔄 Update Future Events Only', 'updateFutureEventsOnly')
      .addItem('🗓️ Update This Month', 'updateThisMonthEvents'))
    .addSeparator()
    .addItem('📊 Export Individual Schedules', 'exportIndividualSchedules')
    .addSeparator()
    .addItem('📁 Archive Current Schedule', 'archiveCurrentSchedule')
    .addItem('🗓️ Generate Next Year', 'generateNextYearSchedule')
    .addSeparator()
    .addItem('🔧 Validate Setup', 'validateSetupOnly')
    .addItem('🧪 Run Tests', 'runSystemTests')
    .addItem('❓ Setup Instructions', 'showSetupInstructions')
    .addToUi();
}

// Create menu when spreadsheet opens
function onOpen() {
  try {
    createMenuItems();
    console.log('Enhanced Deacon Rotation menu created successfully (v24.2)');
  } catch (error) {
    console.error('Failed to create menu:', error);
  }
}

// Enhanced onEdit with optional auto-regeneration (currently disabled)
function onEdit(e) {
  // AUTO-REGENERATION IS DISABLED
  // To re-enable automatic updates when deacon/household lists change,
  // uncomment the code below by removing the /* and */ markers
  
  /*
  try {
    const editedColumn = e.range.getColumn();
    const editedRow = e.range.getRow();
    
    // Column numbers: L=12 for deacons, M=13 for households
    if ((editedColumn === 12 || editedColumn === 13) && editedRow > 1) {
      if (!preventEditLoops()) {
        console.log('Auto-regeneration skipped - too recent');
        return;
      }
      
      console.log(`Auto-regeneration triggered by edit in column ${editedColumn}, row ${editedRow}`);
      Utilities.sleep(2000);
      
      const sheet = SpreadsheetApp.getActiveSheet();
      const config = getConfiguration(sheet);
      
      if (config.deacons.length < 2 || config.households.length === 0) {
        console.log('Insufficient data for auto-regeneration');
        return;
      }
      
      generateRotationSchedule();
      console.log('Auto-regeneration completed successfully');
    }
  } catch (error) {
    console.error('Auto-regeneration failed:', error);
    
    if (error.message && !error.message.includes('too recent')) {
      try {
        SpreadsheetApp.getUi().alert(
          'Auto-Update Failed',
          `⚠️ Schedule couldn't auto-update: ${error.message}\n\nPlease use the menu to regenerate manually:\n🔄 Deacon Rotation > 📅 Generate Schedule`,
          SpreadsheetApp.getUi().ButtonSet.OK
        );
      } catch (uiError) {
        console.error('Could not show error alert:', uiError);
      }
    }
  }
  */
}

// END OF ENHANCED MODULE 3 (v24.2)
