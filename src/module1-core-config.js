/**
 * MODULE 1: CORE FUNCTIONS & CONFIGURATION (ENHANCED)
 * Deacon Visitation Rotation System - Modular Version
 * 
 * This module contains:
 * - Main entry point and configuration loading
 * - Header setup and UI initialization
 * - All validation functions
 * - Basic utility functions
 * 
 * ENHANCED Layout with Breeze and Notes Integration:
 * - Main schedule: Columns A-E (unchanged)
 * - Buffer: Column F
 * - Individual deacon reports: Columns G-I
 * - Buffer: Column J
 * - Configuration: Column K
 * - Deacon list: Column L
 * - Household list: Column M
 * - Phone numbers: Column N
 * - Addresses: Column O
 * - Breeze Link (8-digit numbers): Column P
 * - Notes Pg Link (Google Doc URLs): Column Q
 * - Breeze Link (short): Column R (auto-generated)
 * - Notes Pg Link (short): Column S (auto-generated)
 */

function generateRotationSchedule() {
  const sheet = SpreadsheetApp.getActiveSheet();
  
  try {
    const config = getConfiguration(sheet);
    
    // Run all validations
    if (!validateInputs(config)) return;
    validateDataTypes(config);
    validateWorkloadFeasibility(config);
    
    // Generate schedule with safety checks
    const scheduleData = safeCreateSchedule(config);
    
    // Write safely to sheet
    safeWriteToSheet(sheet, scheduleData);
    
    // Generate individual deacon reports
    generateDeaconReports(sheet, scheduleData, config);
    
    SpreadsheetApp.getUi().alert(
      'Schedule Generated!',
      `✅ Created ${scheduleData.length} visits over ${config.numWeeks} weeks\n` +
      `📅 Visit frequency: Every ${config.visitFrequency} week(s)\n` +
      `👥 Average ${(scheduleData.length / config.deacons.length).toFixed(1)} visits per deacon\n` +
      `⏱️  Each deacon visits every ${Math.round(config.numWeeks * config.deacons.length / scheduleData.length)} weeks on average`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    console.error('Schedule generation error:', error);
    
    SpreadsheetApp.getUi().alert(
      'Error Generating Schedule',
      `❌ ${error.message}\n\nPlease check your setup and try again.`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

function getConfiguration(sheet) {
  try {
    // Get deacons (column L)
    const deaconRange = sheet.getRange('L2:L100');
    const deacons = deaconRange.getValues()
      .flat()
      .filter(cell => cell !== '' && cell !== null && cell !== undefined)
      .map(name => String(name).trim())
      .filter(name => name.length > 0);
    
    // Get households (column M)
    const householdRange = sheet.getRange('M2:M100');
    const households = householdRange.getValues()
      .flat()
      .filter(cell => cell !== '' && cell !== null && cell !== undefined)
      .map(name => String(name).trim())
      .filter(name => name.length > 0);
    
    // Get phone numbers (column N)
    const phoneRange = sheet.getRange('N2:N100');
    const phones = phoneRange.getValues()
      .flat()
      .map(cell => cell ? String(cell).trim() : '');
    
    // Get addresses (column O)
    const addressRange = sheet.getRange('O2:O100');
    const addresses = addressRange.getValues()
      .flat()
      .map(cell => cell ? String(cell).trim() : '');
    
    // Get Breeze link numbers (column P) - NEW
    const breezeRange = sheet.getRange('P2:P100');
    const breezeNumbers = breezeRange.getValues()
      .flat()
      .map(cell => cell ? String(cell).trim() : '');
    
    // Get Notes page links (column Q) - NEW
    const notesRange = sheet.getRange('Q2:Q100');
    const notesLinks = notesRange.getValues()
      .flat()
      .map(cell => cell ? String(cell).trim() : '');
    
    // Get shortened Breeze links (column R) - NEW
    const breezeShortRange = sheet.getRange('R2:R100');
    const breezeShortLinks = breezeShortRange.getValues()
      .flat()
      .map(cell => cell ? String(cell).trim() : '');
    
    // Get shortened Notes links (column S) - NEW
    const notesShortRange = sheet.getRange('S2:S100');
    const notesShortLinks = notesShortRange.getValues()
      .flat()
      .map(cell => cell ? String(cell).trim() : '');
    
    // Set up headers
    setupHeaders(sheet);
    
    // Get configuration from column K
    let startDate = sheet.getRange('K2').getValue();
    if (!startDate || !(startDate instanceof Date)) {
      startDate = getNextMonday();
      sheet.getRange('K2').setValue(startDate);
    }
    
    let visitFrequency = sheet.getRange('K4').getValue();
    if (!visitFrequency || !Number.isFinite(visitFrequency)) {
      visitFrequency = 2;
      sheet.getRange('K4').setValue(visitFrequency);
    }
    
    let numWeeks = sheet.getRange('K6').getValue();
    if (!numWeeks || !Number.isFinite(numWeeks)) {
      numWeeks = 52;
      sheet.getRange('K6').setValue(numWeeks);
    }
    
    let calendarInstructions = sheet.getRange('K8').getValue();
    if (!calendarInstructions) {
      calendarInstructions = 'Please call to confirm visit time. Contact family 1-2 days before scheduled date to arrange convenient time.';
      sheet.getRange('K8').setValue(calendarInstructions);
    }
    
    return {
      deacons,
      households,
      phones: phones.slice(0, households.length),
      addresses: addresses.slice(0, households.length),
      breezeNumbers: breezeNumbers.slice(0, households.length),
      notesLinks: notesLinks.slice(0, households.length),
      breezeShortLinks: breezeShortLinks.slice(0, households.length),
      notesShortLinks: notesShortLinks.slice(0, households.length),
      startDate: new Date(startDate.getTime()),
      visitFrequency: Number(visitFrequency),
      numWeeks: Number(numWeeks),
      calendarInstructions: String(calendarInstructions)
    };
    
  } catch (error) {
    throw new Error(`Configuration error: ${error.message}`);
  }
}

function setupHeaders(sheet) {
  // Column headers for deacon reports (G-I)
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
  
  // Configuration labels in column K
  if (!sheet.getRange('K1').getValue()) {
    sheet.getRange('K1').setValue('Start Date');
    sheet.getRange('K1').setFontWeight('bold').setBackground('#fff2cc');
  }
  
  if (!sheet.getRange('K3').getValue()) {
    sheet.getRange('K3').setValue('Visits every x weeks (1,2,3,4)');
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
    sheet.getRange('K8').setWrap(true);
    sheet.setColumnWidth(11, 250);
  }

  // NEW: Weekly notification configuration (K10-K13)
  if (!sheet.getRange('K10').getValue()) {
    sheet.getRange('K10').setValue('Weekly Notification Day:');
    sheet.getRange('K10').setFontWeight('bold').setBackground('#d4edda');
  }
  
  if (!sheet.getRange('K11').getValue()) {
    sheet.getRange('K11').setValue('Sunday');
    sheet.getRange('K11').setBackground('#f8f9fa');
  }
  
  if (!sheet.getRange('K12').getValue()) {
    sheet.getRange('K12').setValue('Weekly Notification Time (0-23):');
    sheet.getRange('K12').setFontWeight('bold').setBackground('#d4edda');
  }
  
  if (!sheet.getRange('K13').getValue()) {
    sheet.getRange('K13').setValue(18); // Default to 6 PM
    sheet.getRange('K13').setBackground('#f8f9fa');
  }
  
  // MOVED: Test mode indicators (K15-K16) - moved down from previous K11-K12
  if (!sheet.getRange('K15').getValue()) {
    sheet.getRange('K15').setValue('Current Mode:');
    sheet.getRange('K15').setFontWeight('bold').setBackground('#fff2cc');
  }
  }
  
  // Column headers for basic contact info (L-O)
  if (!sheet.getRange('L1').getValue()) {
    sheet.getRange('L1').setValue('Deacons');
    sheet.getRange('L1').setFontWeight('bold').setBackground('#e8f0fe');
  }
  
  if (!sheet.getRange('M1').getValue()) {
    sheet.getRange('M1').setValue('Households');
    sheet.getRange('M1').setFontWeight('bold').setBackground('#e8f0fe');
  }
  
  if (!sheet.getRange('N1').getValue()) {
    sheet.getRange('N1').setValue('Phone Number');
    sheet.getRange('N1').setFontWeight('bold').setBackground('#e8f0fe');
  }
  
  if (!sheet.getRange('O1').getValue()) {
    sheet.getRange('O1').setValue('Address');
    sheet.getRange('O1').setFontWeight('bold').setBackground('#e8f0fe');
  }
  
  // NEW: Column headers for Breeze and Notes integration (P-S)
  if (!sheet.getRange('P1').getValue()) {
    sheet.getRange('P1').setValue('Breeze Link');
    sheet.getRange('P1').setFontWeight('bold').setBackground('#fce5cd');
  }
  
  if (!sheet.getRange('Q1').getValue()) {
    sheet.getRange('Q1').setValue('Notes Pg Link');
    sheet.getRange('Q1').setFontWeight('bold').setBackground('#fce5cd');
  }
  
  if (!sheet.getRange('R1').getValue()) {
    sheet.getRange('R1').setValue('Breeze Link (short)');
    sheet.getRange('R1').setFontWeight('bold').setBackground('#d9ead3');
  }
  
  if (!sheet.getRange('S1').getValue()) {
    sheet.getRange('S1').setValue('Notes Pg Link (short)');
    sheet.getRange('S1').setFontWeight('bold').setBackground('#d9ead3');
  }
}

function getNextMonday() {
  const today = new Date();
  const nextMonday = new Date(today);
  const daysUntilMonday = (1 + 7 - today.getDay()) % 7;
  nextMonday.setDate(today.getDate() + (daysUntilMonday || 7)); // If today is Monday, go to next Monday
  return nextMonday;
}

function validateInputs(config) {
  const ui = SpreadsheetApp.getUi();
  
  if (config.deacons.length === 0) {
    ui.alert(
      'Missing Deacons', 
      '❌ Please add deacon names in column L (starting from L2)\n\nExample:\nL2: AB\nL3: CD\nL4: EF', 
      ui.ButtonSet.OK
    );
    return false;
  }
  
  if (config.households.length === 0) {
    ui.alert(
      'Missing Households', 
      '❌ Please add household names in column M (starting from M2)\n\nExample:\nM2: Adams\nM3: Baker\nM4: Cooper', 
      ui.ButtonSet.OK
    );
    return false;
  }
  
  if (config.deacons.length < 2) {
    ui.alert(
      'Too Few Deacons', 
      '❌ You need at least 2 deacons for a meaningful rotation.\n\nCurrent deacons: ' + config.deacons.length, 
      ui.ButtonSet.OK
    );
    return false;
  }
  
  return true;
}

function validateDataTypes(config) {
  // Check visit frequency is a valid number
  if (!Number.isInteger(config.visitFrequency) || config.visitFrequency < 1 || config.visitFrequency > 8) {
    throw new Error(`Visit frequency must be between 1 and 8 weeks. Got: ${config.visitFrequency}`);
  }
  
  // Check num weeks is valid
  if (!Number.isInteger(config.numWeeks) || config.numWeeks < 1 || config.numWeeks > 520) {
    throw new Error(`Number of weeks must be between 1 and 520 (10 years). Got: ${config.numWeeks}`);
  }
  
  // Check start date is valid
  if (!(config.startDate instanceof Date) || isNaN(config.startDate.getTime())) {
    throw new Error('Start date must be a valid date');
  }
  
  // Check for future date (prevent accidental past dates)
  const today = new Date();
  const oneYearAgo = new Date(today.getFullYear() - 1, today.getMonth(), today.getDate());
  if (config.startDate < oneYearAgo) {
    throw new Error(`Start date seems too far in the past: ${config.startDate.toLocaleDateString()}. Please check the date.`);
  }
  
  // Check for duplicate names
  const deaconDupes = config.deacons.filter((name, index) => config.deacons.indexOf(name) !== index);
  const householdDupes = config.households.filter((name, index) => config.households.indexOf(name) !== index);
  
  if (deaconDupes.length > 0) {
    throw new Error(`Duplicate deacon names found: ${deaconDupes.join(', ')}`);
  }
  
  if (householdDupes.length > 0) {
    throw new Error(`Duplicate household names found: ${householdDupes.join(', ')}`);
  }
  
  // Check for problematic characters that might break sheet operations
  const problematicChars = /[[\]{}|\\]/;
  const badDeacons = config.deacons.filter(name => problematicChars.test(name));
  const badHouseholds = config.households.filter(name => problematicChars.test(name));
  
  if (badDeacons.length > 0) {
    throw new Error(`Deacon names contain problematic characters: ${badDeacons.join(', ')}`);
  }
  
  if (badHouseholds.length > 0) {
    throw new Error(`Household names contain problematic characters: ${badHouseholds.join(', ')}`);
  }
  
  // NEW: Validate Breeze numbers format (optional validation)
  config.breezeNumbers.forEach((number, index) => {
    if (number && number.length > 0) {
      // Check if it's a valid number (8 digits is typical, but be flexible)
      if (!/^\d{1,10}$/.test(number)) {
        console.warn(`Breeze number for household ${config.households[index]} may be invalid: ${number}`);
      }
    }
  });
  
  // NEW: Validate Notes links format (optional validation)
  config.notesLinks.forEach((link, index) => {
    if (link && link.length > 0) {
      // Basic URL validation
      if (!link.startsWith('http://') && !link.startsWith('https://')) {
        console.warn(`Notes link for household ${config.households[index]} may be invalid: ${link}`);
      }
    }
  });
}

function validateWorkloadFeasibility(config) {
  const { deacons, households, visitFrequency, numWeeks } = config;
  
  // Calculate visits per deacon per time period
  const totalVisitsNeeded = Math.ceil(numWeeks / visitFrequency) * households.length;
  const visitsPerDeacon = totalVisitsNeeded / deacons.length;
  const weeksPerVisit = numWeeks / visitsPerDeacon;
  
  // Check if workload is too high (less than 3 weeks between visits)
  if (weeksPerVisit < 3) {
    throw new Error(
      `❌ Workload too high: Each deacon would visit every ${weeksPerVisit.toFixed(1)} weeks.\n\n` +
      `Suggestions:\n` +
      `• Add more deacons (current: ${deacons.length})\n` +
      `• Reduce visit frequency (current: every ${visitFrequency} weeks)\n` +
      `• Reduce number of households (current: ${households.length})\n` +
      `• Shorten schedule duration (current: ${numWeeks} weeks)`
    );
  }
  
  // Warn if workload is very low (more than 12 weeks between visits)
  if (weeksPerVisit > 12) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'Low Workload Warning', 
      `⚠️ Each deacon will only visit every ${weeksPerVisit.toFixed(1)} weeks.\n\n` +
      `This might be too infrequent for effective pastoral care.\n\n` +
      `Continue anyway?`,
      ui.ButtonSet.YES_NO
    );
    
    if (response !== ui.Button.YES) {
      throw new Error('Schedule generation cancelled by user');
    }
  }
  
  // Check for extremely large schedules that might hit execution limits
  if (totalVisitsNeeded > 2000) {
    throw new Error(
      `❌ Schedule too large: ${totalVisitsNeeded} total visits would exceed execution limits.\n\n` +
      `Try reducing the number of weeks or visit frequency.`
    );
  }
  
  return true;
}

function validateSetupOnly() {
  const sheet = SpreadsheetApp.getActiveSheet();
  
  try {
    const config = getConfiguration(sheet);
    
    if (!validateInputs(config)) return;
    validateDataTypes(config);
    validateWorkloadFeasibility(config);
    
    // NEW: Report on Breeze and Notes integration
    const breezeCount = config.breezeNumbers.filter(num => num && num.length > 0).length;
    const notesCount = config.notesLinks.filter(link => link && link.length > 0).length;
    
    SpreadsheetApp.getUi().alert(
      'Setup Validation',
      `✅ Setup looks good!\n\n` +
      `👥 Deacons: ${config.deacons.length}\n` +
      `🏠 Households: ${config.households.length}\n` +
      `📞 Phone numbers: ${config.phones.filter(p => p && p.length > 0).length}\n` +
      `📍 Addresses: ${config.addresses.filter(a => a && a.length > 0).length}\n` +
      `🔗 Breeze links: ${breezeCount}\n` +
      `📝 Notes links: ${notesCount}\n\n` +
      `📅 Start Date: ${config.startDate.toLocaleDateString()}\n` +
      `⏱️  Visit Frequency: Every ${config.visitFrequency} weeks\n` +
      `📊 Schedule Duration: ${config.numWeeks} weeks\n\n` +
      `Average visits per deacon: ${Math.round(config.numWeeks * config.households.length / config.visitFrequency / config.deacons.length)}\n` +
      `Each deacon visits every ~${Math.round(config.numWeeks * config.deacons.length / (config.numWeeks * config.households.length / config.visitFrequency))} weeks\n\n` +
      `Ready to generate schedule!`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    SpreadsheetApp.getUi().alert(
      'Setup Validation Failed',
      `❌ ${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

function preventEditLoops() {
  // Use PropertiesService to prevent rapid re-execution
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

function showSetupInstructions() {
  const ui = SpreadsheetApp.getUi();
  const instructions = `
📋 SETUP INSTRUCTIONS (Enhanced with Breeze & Notes Integration):

1️⃣ CONFIGURATION (Column K):
   • K1: "Start Date:" (label)
   • K2: Your start date (preferably a Monday)
   • K3: "Visit Frequency (weeks):" (label)
   • K4: Visit frequency (1, 2, 3, or 4 weeks)
   • K5: "Number of Weeks:" (label)
   • K6: Number of weeks to schedule (e.g., 52)
   • K7: "Calendar Event Instructions:" (label)
   • K8: Custom instructions for calendar events

2️⃣ DEACONS LIST (Column L):
   • L1: "Deacons" (header - auto-created)
   • L2, L3, L4...: List each deacon's name or initials
   
3️⃣ HOUSEHOLDS LIST (Column M):  
   • M1: "Households" (header - auto-created)
   • M2, M3, M4...: List each household name

4️⃣ CONTACT INFO (Columns N-O):
   • N1: "Phone Number" (header - auto-created)
   • N2, N3, N4...: Phone numbers for households
   • O1: "Address" (header - auto-created)
   • O2, O3, O4...: Addresses for households

5️⃣ BREEZE INTEGRATION (Columns P-R):
   • P1: "Breeze Link" (header - auto-created)
   • P2, P3, P4...: 8-digit Breeze numbers (e.g., 29760588)
   • R1: "Breeze Link (short)" (header - auto-created)
   • R2, R3, R4...: Auto-generated shortened URLs

6️⃣ NOTES INTEGRATION (Columns Q-S):
   • Q1: "Notes Pg Link" (header - auto-created)
   • Q2, Q3, Q4...: Full Google Doc URLs for visit notes
   • S1: "Notes Pg Link (short)" (header - auto-created)
   • S2, S3, S4...: Auto-generated shortened URLs

7️⃣ GENERATE SCHEDULE:
   • Menu: 🔄 Deacon Rotation > 📅 Generate Schedule
   • Or run generateRotationSchedule() from Script Editor

8️⃣ OUTPUT LOCATIONS:
   • Columns A-E: Master schedule
   • Columns G-I: Individual deacon reports

🔗 URL SHORTENING: The system will automatically generate shortened URLs for Breeze profiles and Notes pages when you export to calendar.

📊 VALIDATION: Use "🔧 Validate Setup" to check your configuration before generating.

❓ Need help? Use "🧪 Run Tests" to diagnose issues.

💡 TIP: All headers and labels are automatically created when you first run the script!
  `;
  
  ui.alert('Setup Instructions', instructions, ui.ButtonSet.OK);
}

// END OF MODULE 1
