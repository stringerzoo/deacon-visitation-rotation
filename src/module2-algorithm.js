/**
 * MODULE 2: ENHANCED ALGORITHM & SCHEDULE GENERATION (v2.0)
 * Deacon Visitation Rotation System - Variable Frequency Support
 * 
 * 🆕 MAJOR v2.0 CHANGES:
 * - Support for household-specific visit frequencies
 * - Enhanced algorithm handles mixed frequencies (1, 2, 3, 4 weeks)
 * - Intelligent visit scheduling based on individual household needs
 * - Maintains optimal deacon distribution across all frequencies
 * - Backward compatible with v1.1 uniform frequency systems
 * 
 * This module contains:
 * - Variable frequency schedule generation algorithm
 * - Enhanced rotation pattern generator for mixed frequencies
 * - Intelligent visit date calculation
 * - Deacon workload balancing across different frequencies
 * - Quality analysis for mixed frequency systems
 */

function safeCreateSchedule(config) {
  // Check execution time limits
  const startTime = new Date().getTime();
  const maxExecutionTime = 4 * 60 * 1000; // 4 minutes (safety buffer)
  
  console.log(`=== v2.0 SCHEDULE GENERATION ===`);
  console.log(`Deacons: ${config.deacons.length}, Households: ${config.households.length}`);
  console.log(`Has custom frequencies: ${config.hasCustomFrequencies}`);
  
  if (config.hasCustomFrequencies) {
    console.log('🆕 Using VARIABLE FREQUENCY algorithm');
    return generateVariableFrequencySchedule(config, startTime, maxExecutionTime);
  } else {
    console.log('📋 Using UNIFORM FREQUENCY algorithm (v1.1 compatible)');
    return generateUniformFrequencySchedule(config, startTime, maxExecutionTime);
  }
}

function generateVariableFrequencySchedule(config, startTime, maxExecutionTime) {
  /**
   * 🆕 v2.0 VARIABLE FREQUENCY ALGORITHM
   * 
   * Strategy:
   * 1. Calculate visit schedule for each household based on their individual frequency
   * 2. Merge all household schedules into a master timeline
   * 3. Assign deacons optimally across the merged schedule
   * 4. Balance workload fairly across all deacons despite different frequencies
   */
  
  const schedule = [];
  const deaconCount = config.deacons.length;
  const totalWeeks = config.numWeeks;
  
  // Step 1: Calculate individual household visit schedules
  const householdVisitSchedules = {};
  let totalVisitsAcrossAllHouseholds = 0;
  
  console.log('📅 Calculating individual household schedules:');
  
  config.householdFrequencies.forEach(hf => {
    const visits = [];
    let currentWeek = 1; // Start at week 1
    
    while (currentWeek <= totalWeeks) {
      visits.push(currentWeek);
      currentWeek += hf.frequency;
    }
    
    householdVisitSchedules[hf.household] = visits;
    totalVisitsAcrossAllHouseholds += visits.length;
    
    const marker = hf.isCustom ? '🔄' : '📋';
    console.log(`${marker} ${hf.household}: ${visits.length} visits (every ${hf.frequency} weeks) - Weeks: ${visits.slice(0, 5).join(', ')}${visits.length > 5 ? '...' : ''}`);
  });
  
  console.log(`📊 Total visits across all households: ${totalVisitsAcrossAllHouseholds}`);
  console.log(`📊 Average visits per deacon: ${Math.round(totalVisitsAcrossAllHouseholds / deaconCount * 100) / 100}`);
  
  // Step 2: Create master timeline of all visits
  const masterTimeline = [];
  
  Object.keys(householdVisitSchedules).forEach(household => {
    const householdFreq = config.householdFrequencies.find(hf => hf.household === household);
    const visits = householdVisitSchedules[household];
    
    visits.forEach(week => {
      masterTimeline.push({
        household: household,
        week: week,
        frequency: householdFreq.frequency,
        isCustom: householdFreq.isCustom
      });
    });
  });
  
  // Sort by week, then by household name for consistency
  masterTimeline.sort((a, b) => {
    if (a.week !== b.week) return a.week - b.week;
    return a.household.localeCompare(b.household);
  });
  
  console.log(`📋 Master timeline created: ${masterTimeline.length} total visits over ${totalWeeks} weeks`);
  
  // Step 3: Smart deacon assignment with workload balancing
  const deaconWorkload = config.deacons.map(() => 0);
  const deaconLastAssignment = config.deacons.map(() => -999); // Track when each deacon was last assigned
  const deaconHouseholdHistory = new Map(); // Track deacon-household pairings
  
  // Initialize deacon-household tracking
  config.deacons.forEach((deacon, dIndex) => {
    config.households.forEach((household, hIndex) => {
      deaconHouseholdHistory.set(`${dIndex}-${household}`, 0);
    });
  });
  
  console.log('🎯 Assigning deacons with intelligent balancing...');
  
  masterTimeline.forEach((visit, visitIndex) => {
    // Check execution time
    if (new Date().getTime() - startTime > maxExecutionTime) {
      throw new Error(`⌛ Schedule too large - exceeded time limit at visit ${visitIndex + 1}.`);
    }
    
    // Find the best deacon for this visit
    let bestDeacon = -1;
    let bestScore = Infinity;
    
    config.deacons.forEach((deacon, deaconIndex) => {
      // Calculate score based on multiple factors
      let score = 0;
      
      // Factor 1: Current workload (prefer less busy deacons)
      score += deaconWorkload[deaconIndex] * 100;
      
      // Factor 2: Time since last assignment (prefer recently inactive)
      const weeksSinceLastAssignment = visit.week - deaconLastAssignment[deaconIndex];
      if (weeksSinceLastAssignment < 0) {
        // Never assigned before, give bonus
        score -= 50;
      } else {
        // Prefer deacons who haven't been assigned recently
        score -= weeksSinceLastAssignment * 10;
      }
      
      // Factor 3: How often has this deacon visited this household?
      const pairKey = `${deaconIndex}-${visit.household}`;
      const previousVisits = deaconHouseholdHistory.get(pairKey) || 0;
      score += previousVisits * 200; // Strong preference for new pairings
      
      // Factor 4: Frequency preference - spread higher frequency households across more deacons
      if (visit.frequency === 1) {
        score -= 5; // Slight preference for weekly visits (most demanding)
      }
      
      if (score < bestScore) {
        bestScore = score;
        bestDeacon = deaconIndex;
      }
    });
    
    // Assign the best deacon
    const assignedDeacon = config.deacons[bestDeacon];
    
    // Calculate actual visit date
    const visitDate = new Date(config.startDate);
    visitDate.setDate(visitDate.getDate() + (visit.week - 1) * 7);
    
    // Create schedule entry (compatible with v1.1 format)
    const scheduleEntry = {
      cycle: Math.ceil(visit.week / (visit.frequency || config.defaultVisitFrequency)),
      week: visit.week,
      date: visitDate,
      household: visit.household,
      deacon: assignedDeacon,
      deaconIndex: bestDeacon,
      // v2.0 specific fields
      householdFrequency: visit.frequency,
      isCustomFrequency: visit.isCustom
    };
    
    schedule.push(scheduleEntry);
    
    // Update tracking
    deaconWorkload[bestDeacon]++;
    deaconLastAssignment[bestDeacon] = visit.week;
    const pairKey = `${bestDeacon}-${visit.household}`;
    deaconHouseholdHistory.set(pairKey, (deaconHouseholdHistory.get(pairKey) || 0) + 1);
  });
  
  // Log final statistics
  console.log('📊 FINAL WORKLOAD DISTRIBUTION:');
  config.deacons.forEach((deacon, index) => {
    const workload = deaconWorkload[index];
    const percentage = Math.round((workload / totalVisitsAcrossAllHouseholds) * 100);
    console.log(`${deacon}: ${workload} visits (${percentage}%)`);
  });
  
  // Quality analysis
  logVariableFrequencyAnalysis(schedule, config);
  
  console.log(`✅ Variable frequency schedule generated: ${schedule.length} visits over ${totalWeeks} weeks`);
  return schedule;
}

function generateUniformFrequencySchedule(config, startTime, maxExecutionTime) {
  /**
   * v1.1 COMPATIBLE UNIFORM FREQUENCY ALGORITHM
   * Used when all households have the same frequency (Column T is empty/unused)
   */
  
  const schedule = [];
  const visitsPerCycle = config.households.length;
  const weeksPerCycle = config.visitFrequency;
  const totalCycles = Math.ceil(config.numWeeks / weeksPerCycle);
  
  const deaconCount = config.deacons.length;
  const householdCount = config.households.length;
  
  console.log(`Uniform frequency: ${config.visitFrequency} weeks, ${totalCycles} cycles`);
  
  const gcd = (a, b) => b === 0 ? a : gcd(b, a % b);
  const isHarmonic = gcd(deaconCount, householdCount) > 1;
  
  console.log(`Harmonic detected: ${isHarmonic}, GCD: ${gcd(deaconCount, householdCount)}`);
  
  // Use the proven v1.1 rotation pattern generator
  const rotationPattern = generateOptimalRotationPattern(deaconCount, householdCount, totalCycles);
  console.log(`Generated rotation pattern with ${rotationPattern.length} cycle assignments`);
  
  let patternIndex = 0;
  
  for (let cycle = 0; cycle < totalCycles; cycle++) {
    // Check execution time
    if (new Date().getTime() - startTime > maxExecutionTime) {
      throw new Error(`⌛ Schedule too large - exceeded time limit at cycle ${cycle}.`);
    }
    
    for (let householdIndex = 0; householdIndex < householdCount; householdIndex++) {
      const week = cycle * weeksPerCycle + 1;
      
      if (week > config.numWeeks) break;
      
      const visitDate = new Date(config.startDate);
      visitDate.setDate(visitDate.getDate() + (week - 1) * 7);
      
      const deaconIndex = rotationPattern[patternIndex % rotationPattern.length];
      patternIndex++;
      
      const scheduleEntry = {
        cycle: cycle + 1,
        week: week,
        date: visitDate,
        household: config.households[householdIndex],
        deacon: config.deacons[deaconIndex],
        deaconIndex: deaconIndex,
        householdFrequency: config.visitFrequency,
        isCustomFrequency: false
      };
      
      schedule.push(scheduleEntry);
    }
  }
  
  // Log rotation analysis using existing v1.1 function
  logRotationAnalysis(schedule, config);
  
  console.log(`✅ Uniform frequency schedule generated: ${schedule.length} visits over ${totalCycles} cycles`);
  return schedule;
}

function generateOptimalRotationPattern(deaconCount, householdCount, totalCycles) {
  /**
   * v1.1 PROVEN ROTATION PATTERN GENERATOR
   * Maintains the mathematical sophistication for uniform frequency scenarios
   */
  
  const pattern = [];
  const deaconUsageCount = new Array(deaconCount).fill(0);
  const deaconHouseholdPairs = new Map();
  
  // Initialize tracking for which deacon has visited which household
  for (let d = 0; d < deaconCount; d++) {
    for (let h = 0; h < householdCount; h++) {
      deaconHouseholdPairs.set(`${d}-${h}`, 0);
    }
  }
  
  console.log(`Generating uniform pattern for ${totalCycles} cycles...`);
  
  for (let cycle = 0; cycle < totalCycles; cycle++) {
    const cycleAssignments = [];
    
    for (let householdIndex = 0; householdIndex < householdCount; householdIndex++) {
      // Find the best deacon for this household in this cycle
      let bestDeacon = -1;
      let bestScore = Infinity;
      
      for (let deaconIndex = 0; deaconIndex < deaconCount; deaconIndex++) {
        // Skip if this deacon is already assigned in this cycle
        if (cycleAssignments.includes(deaconIndex)) continue;
        
        // Calculate score based on balanced assignment
        const totalVisits = deaconUsageCount[deaconIndex];
        const pairVisits = deaconHouseholdPairs.get(`${deaconIndex}-${householdIndex}`) || 0;
        const cyclePenalty = cycle % deaconCount === deaconIndex ? -10 : 0;
        
        const score = totalVisits * 100 + pairVisits * 50 + cyclePenalty;
        
        if (score < bestScore) {
          bestScore = score;
          bestDeacon = deaconIndex;
        }
      }
      
      if (bestDeacon === -1) {
        bestDeacon = cycle % deaconCount;
      }
      
      pattern.push(bestDeacon);
      cycleAssignments.push(bestDeacon);
      deaconUsageCount[bestDeacon]++;
      deaconHouseholdPairs.set(`${bestDeacon}-${householdIndex}`, 
        (deaconHouseholdPairs.get(`${bestDeacon}-${householdIndex}`) || 0) + 1);
    }
  }
  
  return pattern;
}

function logVariableFrequencyAnalysis(schedule, config) {
  /**
   * 🆕 v2.0 ANALYSIS: Specialized analysis for variable frequency schedules
   */
  
  console.log('\n📊 === VARIABLE FREQUENCY SCHEDULE ANALYSIS ===');
  
  // Frequency distribution analysis
  const frequencyStats = {};
  schedule.forEach(visit => {
    const freq = `${visit.householdFrequency}-week${visit.householdFrequency > 1 ? 's' : ''}`;
    if (!frequencyStats[freq]) frequencyStats[freq] = { visits: 0, households: new Set() };
    frequencyStats[freq].visits++;
    frequencyStats[freq].households.add(visit.household);
  });
  
  console.log('📈 Visits by Frequency:');
  Object.keys(frequencyStats).forEach(freq => {
    const stats = frequencyStats[freq];
    console.log(`  ${freq}: ${stats.visits} visits across ${stats.households.size} households`);
  });
  
  // Deacon workload analysis
  const deaconStats = {};
  schedule.forEach(visit => {
    if (!deaconStats[visit.deacon]) {
      deaconStats[visit.deacon] = { total: 0, byFrequency: {} };
    }
    deaconStats[visit.deacon].total++;
    
    const freq = visit.householdFrequency;
    if (!deaconStats[visit.deacon].byFrequency[freq]) {
      deaconStats[visit.deacon].byFrequency[freq] = 0;
    }
    deaconStats[visit.deacon].byFrequency[freq]++;
  });
  
  console.log('👥 Deacon Workload Distribution:');
  Object.keys(deaconStats).forEach(deacon => {
    const stats = deaconStats[deacon];
    const freqBreakdown = Object.keys(stats.byFrequency)
      .map(freq => `${freq}w: ${stats.byFrequency[freq]}`)
      .join(', ');
    console.log(`  ${deacon}: ${stats.total} total (${freqBreakdown})`);
  });
  
  // Balance analysis
  const workloads = Object.values(deaconStats).map(s => s.total);
  const minWorkload = Math.min(...workloads);
  const maxWorkload = Math.max(...workloads);
  const avgWorkload = workloads.reduce((sum, w) => sum + w, 0) / workloads.length;
  
  console.log(`⚖️ Workload Balance: Min: ${minWorkload}, Max: ${maxWorkload}, Avg: ${Math.round(avgWorkload * 100) / 100}`);
  console.log(`📊 Balance Quality: ${maxWorkload - minWorkload <= 2 ? '✅ Excellent' : maxWorkload - minWorkload <= 4 ? '🟡 Good' : '🔴 Needs Improvement'} (difference: ${maxWorkload - minWorkload})`);
}

function logRotationAnalysis(schedule, config) {
  /**
   * v1.1 COMPATIBLE ANALYSIS: Used for uniform frequency schedules
   */
  console.log('\n📊 === UNIFORM FREQUENCY SCHEDULE ANALYSIS ===');
  
  // Deacon visit distribution
  const deaconVisitCounts = {};
  config.deacons.forEach(deacon => {
    deaconVisitCounts[deacon] = schedule.filter(visit => visit.deacon === deacon).length;
  });
  
  console.log('👥 Visit distribution:');
  Object.keys(deaconVisitCounts).forEach(deacon => {
    console.log(`${deacon}: ${deaconVisitCounts[deacon]} visits`);
  });
  
  // Balance analysis
  const visitCounts = Object.values(deaconVisitCounts);
  const minVisits = Math.min(...visitCounts);
  const maxVisits = Math.max(...visitCounts);
  
  console.log(`⚖️ Balance: Min: ${minVisits}, Max: ${maxVisits}, Difference: ${maxVisits - minVisits}`);
  
  if (maxVisits - minVisits <= 1) {
    console.log('✅ Excellent balance achieved');
  } else if (maxVisits - minVisits <= 2) {
    console.log('🟡 Good balance achieved');
  } else {
    console.log('🔴 Balance could be improved');
  }
}

/**
 * SCHEDULE WRITING & FORMATTING (Enhanced for v2.0)
 */
function writeScheduleToSheet(schedule, config) {
  const sheet = SpreadsheetApp.getActiveSheet();
  
  // Clear existing schedule data
  sheet.getRange('A:E').clear();
  
  // Write headers
  const headers = ['Cycle', 'Week', 'Week of', 'Household', 'Deacon'];
  sheet.getRange('A1:E1').setValues([headers]);
  sheet.getRange('A1:E1').setFontWeight('bold').setBackground('#4285f4').setFontColor('white');
  
  // Prepare schedule data with CLEAN formatting (no frequency markers in main schedule)
  const scheduleData = schedule.map(visit => [
    visit.cycle,
    visit.week,
    visit.date.toLocaleDateString('en-US', { month: 'short', day: 'numeric' }),
    visit.household, // Clean household name - no frequency markers
    visit.deacon
  ]);
  
  // Write schedule data
  if (scheduleData.length > 0) {
    sheet.getRange(2, 1, scheduleData.length, 5).setValues(scheduleData);
    
    // Apply SUBTLE highlighting for custom frequencies with explanation
    if (config.hasCustomFrequencies) {
      const customFreqRows = [];
      schedule.forEach((visit, index) => {
        if (visit.isCustomFrequency) {
          customFreqRows.push(index + 2); // +2 for header and 0-based index
        }
      });
      
      // Use LIGHT background highlighting (much more subtle)
      customFreqRows.forEach(row => {
        sheet.getRange(`A${row}:E${row}`).setBackground('#fff9c4'); // Very light yellow
      });
      
      // Add explanation note for the highlighting
      if (customFreqRows.length > 0) {
        // Find a good spot for the legend (after the data)
        const legendRow = schedule.length + 4;
        sheet.getRange(`A${legendRow}`).setValue('Legend:').setFontWeight('bold');
        sheet.getRange(`A${legendRow + 1}`).setValue('Light yellow = Custom frequency household').setBackground('#fff9c4');
        sheet.getRange(`A${legendRow + 1}:E${legendRow + 1}`).merge();
      }
    }
  }
  
  // Write CLEAN individual deacon reports (restored v1.1 format)
  generateCleanDeaconReports(schedule, config);
  
  console.log(`📝 Schedule written to sheet: ${schedule.length} visits with clean formatting`);
}

function generateCleanDeaconReports(schedule, config) {
  const sheet = SpreadsheetApp.getActiveSheet();
  
  // Clear existing deacon reports
  sheet.getRange('G:I').clear();
  
  // Write headers for deacon reports
  sheet.getRange('G1').setValue('Deacon').setFontWeight('bold').setBackground('#34a853').setFontColor('white');
  sheet.getRange('H1').setValue('Week of').setFontWeight('bold').setBackground('#34a853').setFontColor('white');
  sheet.getRange('I1').setValue('Household').setFontWeight('bold').setBackground('#34a853').setFontColor('white');
  
  // Create deacon-specific schedules (same as v1.1)
  const deaconSchedules = {};
  config.deacons.forEach(deacon => {
    deaconSchedules[deacon] = schedule
      .filter(visit => visit.deacon === deacon)
      .sort((a, b) => a.date - b.date);
  });
  
  // Generate CLEAN report data (v1.1 style)
  const reportData = [];
  
  config.deacons.forEach(deacon => {
    const visits = deaconSchedules[deacon];
    
    if (visits.length > 0) {
      // Add deacon header row (just like v1.1)
      reportData.push([
        `${deacon} (${visits.length} visits)`, // Clean deacon name with count
        '', // Empty week column for deacon header
        ''  // Empty household column for deacon header
      ]);
      
      // Add visit rows with CLEAN formatting
      visits.forEach(visit => {
        // Clean household name with subtle frequency indicator only if custom
        const householdDisplay = visit.isCustomFrequency 
          ? `${visit.household} (${visit.householdFrequency}w)`
          : visit.household;
        
        reportData.push([
          '', // Empty deacon column for visit rows
          visit.date.toLocaleDateString('en-US', { month: 'short', day: 'numeric' }),
          householdDisplay
        ]);
      });
      
      // Add spacing between deacons (like v1.1)
      reportData.push(['', '', '']);
    }
  });
  
  // Write the clean report data
  if (reportData.length > 0) {
    sheet.getRange(2, 7, reportData.length, 3).setValues(reportData);
    
    // Apply formatting to make deacon names stand out (like v1.1)
    let currentRow = 2;
    config.deacons.forEach(deacon => {
      const visits = deaconSchedules[deacon];
      if (visits.length > 0) {
        // Make deacon name row bold and colored (like v1.1)
        sheet.getRange(`G${currentRow}:I${currentRow}`)
          .setFontWeight('bold')
          .setBackground('#e8f0fe'); // Light blue background
        
        currentRow += visits.length + 2; // Move to next deacon (visits + deacon header + spacing)
      }
    });
  }
  
  console.log(`📋 Clean individual reports generated for ${config.deacons.length} deacons (v1.1 style)`);
}
