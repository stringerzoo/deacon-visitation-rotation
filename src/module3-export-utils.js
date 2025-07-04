/**
 * MODULE 3: SMART CALENDAR UPDATE FUNCTIONS (v24.2)
 * Deacon Visitation Rotation System - Smart Calendar Updates
 * 
 * This module contains:
 * - Smart calendar update functions with preservation of scheduling details
 * - Test mode configuration for safe testing
 * - Helper functions for calendar access
 */

// TEST MODE CONFIGURATION
const TEST_MODE = false; // Set to true for testing with separate calendar
const TEST_CALENDAR_NAME = TEST_MODE ? 'TEST - Deacon Visitation Schedule' : 'Deacon Visitation Schedule';

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
      TEST_MODE ? 'TEST: Update Contact Info Only' : 'Update Contact Info Only',
      `${TEST_MODE ? '🧪 TEST MODE: ' : ''}This will update contact information in existing calendar events.\n\n` +
      `✅ PRESERVES: Custom times, dates, guest lists, locations\n` +
      `🔄 UPDATES: Contact info, Breeze links, Notes links, instructions\n\n` +
      `💡 TIP: If you've updated Breeze numbers (column P) or Notes links (column Q),\n` +
      `run "🔗 Generate Shortened URLs" first for optimal results.\n\n` +
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
    cutoffDate.setHours(0, 0, 0, 0);
    
    // Confirm with user
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      TEST_MODE ? 'TEST: Update Future Events Only' : 'Update Future Events Only',
      `${TEST_MODE ? '🧪 TEST MODE: ' : ''}This will update calendar events starting from ${cutoffDate.toLocaleDateString()}.\n\n` +
      `✅ PRESERVES: This week's scheduling details and customizations\n` +
      `🔄 UPDATES: Future events with current contact info and assignments\n\n` +
      `💡 TIP: For contact changes only, use "📞 Update Contact Info Only" instead.\n\n` +
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

function getOrCreateCalendar() {
  /**
   * Gets existing calendar or returns null if not found
   * Used by update functions to ensure calendar exists
   */
  const calendarName = TEST_CALENDAR_NAME;
  
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

// END OF MODULE 3
