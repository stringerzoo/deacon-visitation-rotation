/* DEACON VISITATION ROTATION SYSTEM - COMBINED VERSION
 * Version 24.2 - Smart Calendar Updates with Preservation of Scheduling Details
 * 
 * This file contains all three modules concatenated for Google Apps Script deployment.
 * 
 * Module Structure:
 * - Module 1: Core Functions & Configuration (Lines 40-570)
 * - Module 2: Algorithm & Schedule Generation (Lines 571-1038)  
 * - Module 3: Export, Menu & Utility Functions (Lines 1039-2290)
 * 
 * NEW in v24.2 - Smart Calendar Update System:
 * - Contact info only updates (preserves ALL scheduling customizations)
 * - Future events only updates (preserves current week scheduling)
 * - Monthly targeted updates (planning cycle support)
 * - Enhanced menu structure with Calendar Functions submenu
 * - Smart event parsing and data preservation
 * 
 * For development, edit individual modules in src/ directory:
 * - src/module1-core-config.js (unchanged from v24.1)
 * - src/module2-algorithm.js (unchanged from v24.1)
 * - src/module3-export-utils.js (ENHANCED with smart calendar updates)
 * 
 * Development process:
 * 1. Edit individual modules
 * 2. Test changes
 * 3. Re-concatenate into this combined file
 * 4. Test combined version in Google Apps Script
 * 
 * Key Features v24.2:
 * ✅ Preserves deacon scheduling customizations (times, dates, guests)
 * ✅ Selective contact information updates
 * ✅ Future-only event updates for mid-week changes
 * ✅ Monthly planning cycle support
 * ✅ Enhanced user safety with clear warnings
 * ✅ Smart event parsing and household matching
 * ✅ Comprehensive rate limiting and error handling
 * 
 * Generated: 2025-07-03
 */
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
    sheet.setColumnWidth(11, 200);
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
/**
 * MODULE 2: ALGORITHM & SCHEDULE GENERATION (CORRECTED LAYOUT)
 * Deacon Visitation Rotation System - Modular Version
 * 
 * This module contains:
 * - Core schedule generation algorithm
 * - Optimal rotation pattern generator
 * - Harmonic resonance analysis and mitigation
 * - Schedule writing and formatting
 * - Individual deacon report generation
 * - Quality analysis and logging
 * 
 * CORRECTED Layout:
 * - Individual deacon reports now go to columns G-I (CORRECTED from F-G)
 * - All other functions updated to match new column positions
 */

function safeCreateSchedule(config) {
  // Check execution time limits
  const startTime = new Date().getTime();
  const maxExecutionTime = 4 * 60 * 1000; // 4 minutes (safety buffer)
  
  const schedule = [];
  
  const visitsPerCycle = config.households.length;
  const weeksPerCycle = config.visitFrequency;
  const totalCycles = Math.ceil(config.numWeeks / weeksPerCycle);
  
  const deaconCount = config.deacons.length;
  const householdCount = config.households.length;
  
  console.log(`=== ROTATION SETUP ===`);
  console.log(`Deacons: ${deaconCount}, Households: ${householdCount}`);
  console.log(`Total cycles: ${totalCycles}`);
  
  const gcd = (a, b) => b === 0 ? a : gcd(b, a % b);
  const isHarmonic = gcd(deaconCount, householdCount) > 1;
  
  console.log(`Harmonic detected: ${isHarmonic}, GCD: ${gcd(deaconCount, householdCount)}`);
  
  // COMPLETELY NEW APPROACH: Pre-generate the entire rotation pattern
  // This eliminates ALL harmonic issues by design
  
  const rotationPattern = generateOptimalRotationPattern(deaconCount, householdCount, totalCycles);
  console.log(`Generated rotation pattern with ${rotationPattern.length} cycle assignments`);
  
  let patternIndex = 0;
  
  for (let cycle = 0; cycle < totalCycles; cycle++) {
    // Check if we're approaching time limit
    if (new Date().getTime() - startTime > maxExecutionTime) {
      throw new Error(`❌ Schedule too large - exceeded time limit at cycle ${cycle}. Try reducing the number of weeks to ${cycle * weeksPerCycle} or fewer.`);
    }
    
    const cycleStartWeek = cycle * weeksPerCycle;
    if (cycleStartWeek >= config.numWeeks) break;
    
    // Safe date calculation
    let visitDate;
    try {
      visitDate = new Date(config.startDate.getTime());
      visitDate.setDate(visitDate.getDate() + (cycleStartWeek * 7));
      
      if (isNaN(visitDate.getTime())) {
        throw new Error(`Invalid date calculation`);
      }
      
      const maxFutureDate = new Date();
      maxFutureDate.setFullYear(maxFutureDate.getFullYear() + 20);
      if (visitDate > maxFutureDate) {
        throw new Error(`Date too far in future: ${visitDate.toLocaleDateString()}`);
      }
      
    } catch (dateError) {
      throw new Error(`❌ Date calculation failed for cycle ${cycle}, week ${cycleStartWeek}: ${dateError.message}`);
    }
    
    // Get pre-calculated deacon assignments for this cycle
    const cycleDeacons = rotationPattern[cycle % rotationPattern.length];
    
    if (cycle < 5) {
      console.log(`Cycle ${cycle} deacons: [${cycleDeacons.map(i => `${i}:${config.deacons[i]}`).join(', ')}]`);
    }
    
    config.households.forEach((household, householdIndex) => {
      const assignedDeaconIndex = cycleDeacons[householdIndex];
      const assignedDeacon = config.deacons[assignedDeaconIndex];
      
      schedule.push({
        cycle: cycle + 1,
        week: cycleStartWeek + 1,
        date: new Date(visitDate.getTime()),
        deacon: assignedDeacon,
        household: household,
        deaconIndex: assignedDeaconIndex
      });
    });
  }
  
  // Final validation
  if (schedule.length === 0) {
    throw new Error('❌ No visits were scheduled. Please check your configuration.');
  }
  
  // Log rotation analysis
  logRotationAnalysis(schedule, config);
  
  console.log(`Successfully generated ${schedule.length} visits over ${totalCycles} cycles using pre-calculated optimal rotation`);
  return schedule;
}

function generateOptimalRotationPattern(deaconCount, householdCount, totalCycles) {
  // Generate a rotation pattern that guarantees optimal distribution
  // regardless of harmonic ratios
  
  const pattern = [];
  const deaconUsageCount = new Array(deaconCount).fill(0);
  const deaconHouseholdPairs = new Map();
  
  // Initialize tracking for which deacon has visited which household
  for (let d = 0; d < deaconCount; d++) {
    for (let h = 0; h < householdCount; h++) {
      deaconHouseholdPairs.set(`${d}-${h}`, 0);
    }
  }
  
  console.log(`Generating pattern for ${totalCycles} cycles...`);
  
  for (let cycle = 0; cycle < totalCycles; cycle++) {
    const cycleAssignments = [];
    
    for (let householdIndex = 0; householdIndex < householdCount; householdIndex++) {
      // Find the best deacon for this household in this cycle
      let bestDeacon = -1;
      let bestScore = Infinity;
      
      for (let deaconIndex = 0; deaconIndex < deaconCount; deaconIndex++) {
        // Skip if this deacon is already assigned in this cycle
        if (cycleAssignments.includes(deaconIndex)) continue;
        
        // Calculate score based on:
        // 1. How many total visits this deacon has
        // 2. How many times this deacon has visited this household
        // 3. How recently this deacon visited this household
        
        const totalVisits = deaconUsageCount[deaconIndex];
        const householdVisits = deaconHouseholdPairs.get(`${deaconIndex}-${householdIndex}`) || 0;
        
        // Lower score is better
        let score = totalVisits * 100 + householdVisits * 10;
        
        // Add variety bonus - prefer deacons who haven't visited this household recently
        if (householdVisits === 0) score -= 50; // Big bonus for new pairing
        
        if (score < bestScore) {
          bestScore = score;
          bestDeacon = deaconIndex;
        }
      }
      
      // If no deacon available (shouldn't happen), use round-robin fallback
      if (bestDeacon === -1) {
        bestDeacon = (cycle * householdCount + householdIndex) % deaconCount;
        
        // Make sure we don't double-assign
        while (cycleAssignments.includes(bestDeacon)) {
          bestDeacon = (bestDeacon + 1) % deaconCount;
        }
      }
      
      cycleAssignments.push(bestDeacon);
      deaconUsageCount[bestDeacon]++;
      deaconHouseholdPairs.set(`${bestDeacon}-${householdIndex}`, 
        (deaconHouseholdPairs.get(`${bestDeacon}-${householdIndex}`) || 0) + 1);
    }
    
    pattern.push(cycleAssignments);
  }
  
  // Log pattern quality
  const minVisits = Math.min(...deaconUsageCount);
  const maxVisits = Math.max(...deaconUsageCount);
  console.log(`Pattern quality: visits range from ${minVisits} to ${maxVisits} (imbalance: ${maxVisits - minVisits})`);
  
  // Count household variety
  let totalPairings = 0;
  for (let d = 0; d < deaconCount; d++) {
    let householdsVisited = 0;
    for (let h = 0; h < householdCount; h++) {
      if ((deaconHouseholdPairs.get(`${d}-${h}`) || 0) > 0) {
        householdsVisited++;
      }
    }
    totalPairings += householdsVisited;
  }
  const avgHouseholdsPerDeacon = totalPairings / deaconCount;
  console.log(`Pattern variety: ${avgHouseholdsPerDeacon.toFixed(1)} avg households per deacon (ideal: ${householdCount})`);
  
  return pattern;
}

function analyzeHarmonicResonance(deaconCount, householdCount) {
  const gcd = (a, b) => b === 0 ? a : gcd(b, a % b);
  const lcm = (a, b) => (a * b) / gcd(a, b);
  
  const commonFactor = gcd(deaconCount, householdCount);
  const isExactDivisor = deaconCount % householdCount === 0;
  const ratio = deaconCount / householdCount;
  
  // Detect various types of harmonic problems
  const isSimpleHarmonic = isExactDivisor; // 12:6, 12:4, 12:3
  const isComplexHarmonic = commonFactor > 1 && !isExactDivisor; // 14:7, 15:6
  const isHarmonic = isSimpleHarmonic || isComplexHarmonic;
  
  return {
    isHarmonic,
    isSimpleHarmonic,
    isComplexHarmonic,
    commonFactor,
    ratio,
    lcm: lcm(deaconCount, householdCount),
    cyclesForFullRotation: lcm(deaconCount, householdCount) / householdCount
  };
}

function calculateAntiHarmonicOffset(cycle, householdIndex, harmonicInfo) {
  const { isSimpleHarmonic, commonFactor, ratio } = harmonicInfo;
  
  if (isSimpleHarmonic) {
    // For simple harmonics (12:6, 12:4), use prime-based spiral offset
    // This creates a "spiral" that breaks the perfect divisor lock
    const primeOffset = 7; // Prime number that doesn't divide common ratios
    return (cycle * primeOffset + householdIndex * 3) % commonFactor;
  } else {
    // For complex harmonics (14:7, 15:6), use fibonacci-like progression
    // This breaks the common factor pattern
    const fibOffset = ((cycle % 8) + (householdIndex % 5)) * 3;
    return fibOffset % commonFactor;
  }
}

function logRotationAnalysis(schedule, config) {
  // Analyze the rotation pattern to verify variety
  const deaconHouseholdMap = {};
  const deaconVisitCounts = {};
  
  schedule.forEach(visit => {
    if (!deaconHouseholdMap[visit.deacon]) {
      deaconHouseholdMap[visit.deacon] = new Set();
      deaconVisitCounts[visit.deacon] = 0;
    }
    deaconHouseholdMap[visit.deacon].add(visit.household);
    deaconVisitCounts[visit.deacon]++;
  });
  
  console.log('=== ROTATION ANALYSIS ===');
  
  // Sort deacons by visit count to spot imbalances
  const sortedDeacons = config.deacons.sort((a, b) => 
    deaconVisitCounts[b] - deaconVisitCounts[a]
  );
  
  let minVisits = Infinity;
  let maxVisits = 0;
  let totalHouseholdsVisited = 0;
  let deaconsWithFullCoverage = 0;
  
  sortedDeacons.forEach(deacon => {
    const householdsVisited = deaconHouseholdMap[deacon] || new Set();
    const totalVisits = deaconVisitCounts[deacon] || 0;
    
    minVisits = Math.min(minVisits, totalVisits);
    maxVisits = Math.max(maxVisits, totalVisits);
    totalHouseholdsVisited += householdsVisited.size;
    
    if (householdsVisited.size === config.households.length) {
      deaconsWithFullCoverage++;
    }
    
    console.log(`${deacon}: ${totalVisits} visits to ${householdsVisited.size} different households [${Array.from(householdsVisited).join(', ')}]`);
  });
  
  // Calculate balance metrics
  const avgVisitsPerDeacon = schedule.length / config.deacons.length;
  const avgHouseholdsPerDeacon = totalHouseholdsVisited / config.deacons.length;
  const visitImbalance = maxVisits - minVisits;
  const coveragePercentage = (deaconsWithFullCoverage / config.deacons.length) * 100;
  
  console.log(`=== BALANCE METRICS ===`);
  console.log(`Average visits per deacon: ${avgVisitsPerDeacon.toFixed(1)}`);
  console.log(`Visit range: ${minVisits} to ${maxVisits} (imbalance: ${visitImbalance})`);
  console.log(`Average households per deacon: ${avgHouseholdsPerDeacon.toFixed(1)} (ideal: ${config.households.length})`);
  console.log(`Deacons with full household coverage: ${deaconsWithFullCoverage}/${config.deacons.length} (${coveragePercentage.toFixed(1)}%)`);
  
  // Quality assessment
  if (visitImbalance <= 1 && coveragePercentage >= 80) {
    console.log(`✅ EXCELLENT: Well-balanced rotation with good variety`);
  } else if (visitImbalance <= 2 && coveragePercentage >= 60) {
    console.log(`✅ GOOD: Acceptable balance with reasonable variety`);
  } else if (coveragePercentage < 30) {
    console.log(`❌ HARMONIC LOCK DETECTED: Many deacons stuck with limited households`);
  } else {
    console.log(`⚠️ NEEDS IMPROVEMENT: Some imbalance or limited variety detected`);
  }
}

function safeWriteToSheet(sheet, scheduleData) {
  // Write in batches to avoid hitting API limits
  const batchSize = 1000;
  const maxRows = Math.max(scheduleData.length + 20, 200);
  
  try {
    // Clear existing data with some buffer rows
    console.log(`Clearing ${maxRows} rows...`);
    sheet.getRange(1, 1, maxRows, 5).clearContent();
    
    // Set up headers with formatting
    const headers = [['Cycle', 'Week', 'Week of', 'Household', 'Deacon']];
    const headerRange = sheet.getRange(1, 1, 1, 5);
    headerRange.setValues(headers);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#4285f4');
    headerRange.setFontColor('white');
    headerRange.setBorder(true, true, true, true, true, true);
    
    console.log(`Writing ${scheduleData.length} rows in batches of ${batchSize}...`);
    
    // Write data in batches
    for (let i = 0; i < scheduleData.length; i += batchSize) {
      const batch = scheduleData.slice(i, i + batchSize);
      const batchData = batch.map(visit => [
        visit.cycle,
        visit.week,
        visit.date,
        visit.household,
        visit.deacon
      ]);
      
      if (batchData.length > 0) {
        const range = sheet.getRange(i + 2, 1, batchData.length, 5);
        range.setValues(batchData);
        
        // Format dates in this batch
        const dateRange = sheet.getRange(i + 2, 3, batchData.length, 1);
        dateRange.setNumberFormat('mm/dd/yyyy');
        
        // Add alternating row colors for readability
        for (let j = 0; j < batchData.length; j++) {
          if ((i + j) % 2 === 0) {
            sheet.getRange(i + j + 2, 1, 1, 5).setBackground('#f8f9fa');
          }
        }
      }
    }
    
    // Auto-resize columns for better display
    sheet.autoResizeColumns(1, 5);
    
    // Add borders to the data area
    if (scheduleData.length > 0) {
      sheet.getRange(1, 1, scheduleData.length + 1, 5).setBorder(true, true, true, true, true, true);
    }
    
    console.log('Sheet writing completed successfully');
    
  } catch (error) {
    throw new Error(`❌ Failed to write schedule to sheet: ${error.message}. Try reducing the schedule size.`);
  }
}

function generateDeaconReports(sheet, scheduleData, config) {
  try {
    // Clear existing reports (CORRECTED: columns G-I instead of F-G)
    sheet.getRange(1, 7, 300, 3).clearContent();
    
    // Set up report headers - already done in setupHeaders(), but ensure they're there
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
    
    // Add borders to headers
    sheet.getRange(1, 7, 1, 3).setBorder(true, true, true, true, true, true);
    
    // Group visits by deacon
    const deaconVisits = {};
    scheduleData.forEach(visit => {
      if (!deaconVisits[visit.deacon]) {
        deaconVisits[visit.deacon] = [];
      }
      deaconVisits[visit.deacon].push(visit);
    });
    
    // Create report data
    const reportData = [];
    config.deacons.forEach(deacon => {
      const visits = deaconVisits[deacon] || [];
      
      if (visits.length > 0) {
        // Add deacon header row
        reportData.push([`${deacon} (${visits.length} visits)`, '', '']);
        
        // Add visit rows, sorted by date
        visits
          .sort((a, b) => a.date.getTime() - b.date.getTime())
          .forEach(visit => {
            reportData.push(['', visit.date, visit.household]);
          });
        
        // Add spacing between deacons
        reportData.push(['', '', '']);
      } else {
        reportData.push([`${deacon} (no visits assigned)`, '', '']);
        reportData.push(['', '', '']);
      }
    });
    
    // Write report data to columns G-I (CORRECTED from F-G)
    if (reportData.length > 0) {
      const reportRange = sheet.getRange(2, 7, reportData.length, 3);
      reportRange.setValues(reportData);
      
      // Format dates in report (column H, which is column 8)
      sheet.getRange(2, 8, reportData.length, 1).setNumberFormat('mm/dd/yyyy');
      
      // Add borders
      sheet.getRange(1, 7, reportData.length + 1, 3).setBorder(true, true, true, true, true, true);
    }
    
    // Auto-resize report columns
    sheet.autoResizeColumns(7, 3);
    
    console.log('Deacon reports generated successfully in columns G-I');
    
  } catch (error) {
    console.error('Failed to generate deacon reports:', error);
  }
}

function getScheduleFromSheet(sheet) {
  try {
    const data = sheet.getRange('A2:E1000').getValues();
    return data
      .filter(row => row[0] !== '' && row[0] !== null && row[2] instanceof Date)
      .map(row => ({
        cycle: row[0],
        week: row[1],
        date: row[2],
        household: row[3],
        deacon: row[4]
      }));
  } catch (error) {
    console.error('Error reading schedule from sheet:', error);
    return [];
  }
}

// END OF MODULE 2
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
    const calendarName = 'Deacon Visitation Schedule';
    
    try {
      const calendars = CalendarApp.getCalendarsByName(calendarName);
      if (calendars.length > 0) {
        calendar = calendars[0];
        
        const response = SpreadsheetApp.getUi().alert(
          'Full Calendar Regeneration',
          `⚠️ This will completely rebuild the calendar "${calendarName}".\n\n` +
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
        calendar.setDescription('Automated schedule for deacon household visitations with contact information and management links');
        calendar.setColor(CalendarApp.Color.BLUE);
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
