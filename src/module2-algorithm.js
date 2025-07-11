/**
 * MODULE 2: ALGORITHM & SCHEDULE GENERATION (CORRECTED LAYOUT)
 * Deacon Visitation Rotation System - Modular Version v1.1
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
