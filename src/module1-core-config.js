/**
 * MODULE 1: CORE CONFIGURATION & VALIDATION (v2.0)
 * Deacon Visitation Rotation System - Enhanced with Household-Specific Frequencies
 * 
 * MAJOR v2.0 CHANGES:
 * - Added support for individual household visitation frequencies (Column T)
 * - K4 now serves as "default" frequency, Column T provides overrides
 * - Enhanced configuration validation for mixed frequency scenarios
 * - Backward compatibility with v1.1 systems (Column T empty = use K4)
 * 
 * This module contains:
 * - Enhanced configuration reading with frequency overrides
 * - Validation for mixed frequency scenarios
 * - Header setup for new Column T
 * - Data integrity checks
 */

function getConfiguration() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    
    // Basic configuration (unchanged from v1.1)
    const startDate = new Date(sheet.getRange('K2').getValue());
    const defaultVisitFrequency = Number(sheet.getRange('K4').getValue());
    const numWeeks = Number(sheet.getRange('K6').getValue());
    const calendarInstructions = String(sheet.getRange('K8').getValue());
    
    // Notification settings (unchanged)
    const notificationDay = String(sheet.getRange('K11').getValue() || 'Saturday');
    const notificationHour = Number(sheet.getRange('K13').getValue() || 18);
    
    // Resource URLs (unchanged)
    const calendarUrl = String(sheet.getRange('K19').getValue() || '');
    const guideUrl = String(sheet.getRange('K22').getValue() || '');
    const summaryUrl = String(sheet.getRange('K25').getValue() || '');
    
    // Get deacon and household lists (unchanged)
    const deacons = sheet.getRange('L2:L100').getValues()
      .flat()
      .filter(cell => cell !== '' && cell != null);
    const households = sheet.getRange('M2:M100').getValues()
      .flat()
      .filter(cell => cell !== '' && cell != null);
    
    // Contact information arrays (unchanged column positions)
    const phones = sheet.getRange('N2:N100').getValues()
      .flat()
      .filter((cell, index) => index < households.length);
    const addresses = sheet.getRange('O2:O100').getValues()
      .flat()
      .filter((cell, index) => index < households.length);
    const breezeLinks = sheet.getRange('P2:P100').getValues()
      .flat()
      .filter((cell, index) => index < households.length);
    const notesLinks = sheet.getRange('Q2:Q100').getValues()
      .flat()
      .filter((cell, index) => index < households.length);
    const breezeShortLinks = sheet.getRange('R2:R100').getValues()
      .flat()
      .filter((cell, index) => index < households.length);
    const notesShortLinks = sheet.getRange('S2:S100').getValues()
      .flat()
      .filter((cell, index) => index < households.length);
    
    // ⭐ NEW v2.0 FEATURE: Read household-specific frequencies from Column T
    const customFrequencies = sheet.getRange('T2:T100').getValues()
      .flat()
      .filter((cell, index) => index < households.length);
    
    // ⭐ NEW v2.0 FEATURE: Create household frequency mapping
    const householdFrequencies = households.map((household, index) => {
      // Use custom frequency if specified, otherwise use default from K4
      const customFreq = customFrequencies[index];
      const frequency = (customFreq && !isNaN(Number(customFreq))) 
        ? Number(customFreq) 
        : defaultVisitFrequency;
      
      return {
        household: household,
        frequency: frequency,
        isCustom: !!(customFreq && !isNaN(Number(customFreq)))
      };
    });
    
    // Validation checks
    if (!startDate || startDate.toString() === 'Invalid Date') {
      throw new Error('Invalid start date in K2');
    }
    
    if (!defaultVisitFrequency || defaultVisitFrequency < 1 || defaultVisitFrequency > 4) {
      throw new Error('Invalid default visit frequency in K4 (must be 1-4 weeks)');
    }
    
    if (!numWeeks || numWeeks < 1) {
      throw new Error('Invalid schedule length in K6');
    }
    
    if (deacons.length === 0) {
      throw new Error('No deacons found in column L');
    }
    
    if (households.length === 0) {
      throw new Error('No households found in column M');
    }
    
    // ⭐ NEW v2.0 VALIDATION: Check custom frequencies are valid
    const invalidFrequencies = householdFrequencies.filter(hf => 
      hf.frequency < 1 || hf.frequency > 4
    );
    
    if (invalidFrequencies.length > 0) {
      const badHouseholds = invalidFrequencies.map(hf => hf.household).join(', ');
      throw new Error(`Invalid custom frequencies for: ${badHouseholds}. Must be 1-4 weeks.`);
    }
    
    // ⭐ NEW v2.0 LOGGING: Show frequency distribution
    const frequencyStats = {};
    householdFrequencies.forEach(hf => {
      const key = `${hf.frequency}-week${hf.frequency > 1 ? 's' : ''}`;
      if (!frequencyStats[key]) frequencyStats[key] = [];
      frequencyStats[key].push(hf.household);
    });
    
    console.log('📊 v2.0 FREQUENCY DISTRIBUTION:');
    console.log(`Default frequency: ${defaultVisitFrequency} weeks`);
    Object.keys(frequencyStats).forEach(freq => {
      const count = frequencyStats[freq].length;
      console.log(`${freq}: ${count} household${count > 1 ? 's' : ''} (${frequencyStats[freq].slice(0,3).join(', ')}${count > 3 ? '...' : ''})`);
    });
    
    return {
      deacons: deacons,
      households: households,
      phones: phones.slice(0, households.length),
      addresses: addresses.slice(0, households.length),
      breezeLinks: breezeLinks.slice(0, households.length),
      notesLinks: notesLinks.slice(0, households.length),
      breezeShortLinks: breezeShortLinks.slice(0, households.length),
      notesShortLinks: notesShortLinks.slice(0, households.length),
      startDate: new Date(startDate.getTime()),
      visitFrequency: Number(defaultVisitFrequency), // Keep for backward compatibility
      numWeeks: Number(numWeeks),
      calendarInstructions: String(calendarInstructions),
      notificationDay: String(notificationDay),
      notificationHour: Number(notificationHour),
      calendarUrl: String(calendarUrl),
      guideUrl: String(guideUrl),
      summaryUrl: String(summaryUrl),
      // ⭐ NEW v2.0 PROPERTIES:
      defaultVisitFrequency: Number(defaultVisitFrequency),
      householdFrequencies: householdFrequencies, // Array of {household, frequency, isCustom}
      hasCustomFrequencies: householdFrequencies.some(hf => hf.isCustom)
    };
    
  } catch (error) {
    throw new Error(`Configuration error: ${error.message}`);
  }
}

function shouldAutoRegenerate() {
  const properties = PropertiesService.getScriptProperties();
  const lastExecution = properties.getProperty('lastAutoGeneration');
  const now = new Date().getTime();
  
  // Don't auto-regenerate if we just did it within 5 seconds
  if (lastExecution && (now - parseInt(lastExecution)) < 5000) {
    console.log('Skipping auto-regeneration - too recent');
    return false;
  }
  
  properties.setProperty('lastAutoGeneration', now.toString());
  return true;
}

function setupHeaders(sheet) {
  // Column headers for deacon reports (G-I) - unchanged
  if (!sheet.getRange('G1').getValue()) {
    sheet.getRange('G1').setValue('Deacon');
    sheet.getRange('G1').setFontWeight('bold').setBackground('#34a853').setFontColor('white');
  }
  
  if (!sheet.getRange('H1').getValue()) {
    sheet.getRange('H1').setValue('Week of');
    sheet.getRange('H1').setFontWeight('bold').setBackground('#34a853').setFontColor('white');
  }
  
  if (!sheet.getRange('I1').getValue()) {
    sheet.getRange('I1').setValue('Household');
    sheet.getRange('I1').setFontWeight('bold').setBackground('#34a853').setFontColor('white');
  }
  
  // Configuration labels in column K - ⭐ UPDATED for v2.0
  if (!sheet.getRange('K1').getValue()) {
    sheet.getRange('K1').setValue('Start Date');
    sheet.getRange('K1').setFontWeight('bold').setBackground('#fff2cc');
  }
  
  if (!sheet.getRange('K3').getValue()) {
    sheet.getRange('K3').setValue('Visits every x weeks (1,2,3,4) (default)');
    sheet.getRange('K3').setFontWeight('bold').setBackground('#fff2cc');
  }
  
  if (!sheet.getRange('K5').getValue()) {
    sheet.getRange('K5').setValue('Length of schedule in weeks');
    sheet.getRange('K5').setFontWeight('bold').setBackground('#fff2cc');
  }
  
  if (!sheet.getRange('K7').getValue()) {
    sheet.getRange('K7').setValue('Calendar Event Instructions:');
    sheet.getRange('K7').setFontWeight('bold').setBackground('#fff2cc');
  }
  
  if (!sheet.getRange('K8').getValue()) {
    sheet.getRange('K8').setValue('Please call to confirm visit time. Contact family 1-2 days before scheduled date to arrange convenient time.');
    sheet.getRange('K8').setNote('This message appears in all calendar events');
  }
  
  // Notification settings headers (unchanged)
  if (!sheet.getRange('K10').getValue()) {
    sheet.getRange('K10').setValue('Weekly Notification Day:');
    sheet.getRange('K10').setFontWeight('bold').setBackground('#d4edda');
  }
  
  if (!sheet.getRange('K12').getValue()) {
    sheet.getRange('K12').setValue('Weekly Notification Time (0-23):');
    sheet.getRange('K12').setFontWeight('bold').setBackground('#d4edda');
  }
  
  // Resource link headers (unchanged)
  if (!sheet.getRange('K18').getValue()) {
    sheet.getRange('K18').setValue('Google Calendar URL:');
    sheet.getRange('K18').setFontWeight('bold').setBackground('#e2e3e5');
  }
  
  if (!sheet.getRange('K21').getValue()) {
    sheet.getRange('K21').setValue('Visitation Guide URL:');
    sheet.getRange('K21').setFontWeight('bold').setBackground('#e2e3e5');
  }
  
  if (!sheet.getRange('K24').getValue()) {
    sheet.getRange('K24').setValue('Schedule Summary URL:');
    sheet.getRange('K24').setFontWeight('bold').setBackground('#e2e3e5');
  }
  
  // Data headers (M-S unchanged, T is new)
  if (!sheet.getRange('L1').getValue()) {
    sheet.getRange('L1').setValue('Deacons');
    sheet.getRange('L1').setFontWeight('bold').setBackground('#4285f4').setFontColor('white');
  }
  
  if (!sheet.getRange('M1').getValue()) {
    sheet.getRange('M1').setValue('Households');
    sheet.getRange('M1').setFontWeight('bold').setBackground('#4285f4').setFontColor('white');
  }
  
  if (!sheet.getRange('N1').getValue()) {
    sheet.getRange('N1').setValue('Phone Number');
    sheet.getRange('N1').setFontWeight('bold').setBackground('#4285f4').setFontColor('white');
  }
  
  if (!sheet.getRange('O1').getValue()) {
    sheet.getRange('O1').setValue('Address');
    sheet.getRange('O1').setFontWeight('bold').setBackground('#4285f4').setFontColor('white');
  }
  
  if (!sheet.getRange('P1').getValue()) {
    sheet.getRange('P1').setValue('Breeze Link');
    sheet.getRange('P1').setFontWeight('bold').setBackground('#4285f4').setFontColor('white');
  }
  
  if (!sheet.getRange('Q1').getValue()) {
    sheet.getRange('Q1').setValue('Notes Pg Link');
    sheet.getRange('Q1').setFontWeight('bold').setBackground('#4285f4').setFontColor('white');
  }
  
  if (!sheet.getRange('R1').getValue()) {
    sheet.getRange('R1').setValue('Breeze Link (short)');
    sheet.getRange('R1').setFontWeight('bold').setBackground('#9aa0a6').setFontColor('white');
  }
  
  if (!sheet.getRange('S1').getValue()) {
    sheet.getRange('S1').setValue('Notes Pg Link (short)');
    sheet.getRange('S1').setFontWeight('bold').setBackground('#9aa0a6').setFontColor('white');
  }
  
  // ⭐ NEW v2.0 HEADER: Column T for custom frequencies
  if (!sheet.getRange('T1').getValue()) {
    sheet.getRange('T1').setValue('Custom visit frequency (every 1, 2, 3, or 4 weeks)');
    sheet.getRange('T1').setFontWeight('bold').setBackground('#ff9900').setFontColor('white');
    sheet.getRange('T1').setNote('Leave blank to use default frequency from K4. Enter 1, 2, 3, or 4 to override for this household.');
  }
  
  // Set up data validation for Column T (1, 2, 3, 4, or blank)
  const customFreqRange = sheet.getRange('T2:T100');
  const validationRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['', '1', '2', '3', '4'], true)
    .setAllowInvalid(false)
    .setHelpText('Select custom frequency or leave blank to use default')
    .build();
  customFreqRange.setDataValidation(validationRule);
}

function showSetupInstructions() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'Setup Instructions - v2.0',
    '📋 DEACON VISITATION ROTATION SETUP (v2.0 - Individual Frequencies)\n\n' +
    '🆕 NEW IN v2.0:\n' +
    '   • Individual household frequencies in Column T\n' +
    '   • K4 now serves as default frequency\n' +
    '   • Leave Column T blank to use default frequency\n' +
    '   • Enter 1, 2, 3, or 4 in Column T to override for specific households\n\n' +
    '1️⃣ Configure Settings (Column K):\n' +
    '   • K2: Start date (Monday)\n' +
    '   • K4: Default visit frequency (weeks)\n' +
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
    '6️⃣ Add Contact Info:\n' +
    '   • Column N: Phone numbers\n' +
    '   • Column O: Addresses\n' +
    '   • Column P: Breeze numbers\n' +
    '   • Column Q: Notes links\n\n' +
    '7️⃣ 🆕 Set Custom Frequencies (Column T):\n' +
    '   • Leave blank to use default from K4\n' +
    '   • Enter 1, 2, 3, or 4 for custom frequencies\n' +
    '   • Example: Enter "3" for 3-week visits\n\n' +
    '8️⃣ Generate Schedule:\n' +
    '   • Use "📅 Generate Schedule" menu\n\n' +
    '9️⃣ Export to Calendar:\n' +
    '   • Use "🚨 Full Calendar Regeneration"\n\n' +
    '🔟 Configure Notifications:\n' +
    '   • Use "📢 Notifications → 🔧 Configure Chat Webhook"\n\n' +
    '📖 For detailed instructions, see the project documentation.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}
