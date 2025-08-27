/**
 * FIXED SCORING ALGORITHM FOR v2.0
 * Addresses back-to-back visit issues and same deacon-household pairings
 */

function generateVariableFrequencySchedule(config, startTime, maxExecutionTime) {
  const schedule = [];
  const deaconCount = config.deacons.length;
  const totalWeeks = config.numWeeks;
  
  // Calculate individual household visit schedules (unchanged)
  const householdVisitSchedules = {};
  let totalVisitsAcrossAllHouseholds = 0;
  
  console.log('📅 Calculating individual household schedules:');
  
  config.householdFrequencies.forEach(hf => {
    const visits = [];
    let currentWeek = 1;
    
    while (currentWeek <= totalWeeks) {
      visits.push(currentWeek);
      currentWeek += hf.frequency;
    }
    
    householdVisitSchedules[hf.household] = visits;
    totalVisitsAcrossAllHouseholds += visits.length;
    
    const marker = hf.isCustom ? '🔄' : '📋';
    console.log(`${marker} ${hf.household}: ${visits.length} visits (every ${hf.frequency} weeks) - Weeks: ${visits.slice(0, 5).join(', ')}${visits.length > 5 ? '...' : ''}`);
  });
  
  // Create master timeline (unchanged)
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
  
  masterTimeline.sort((a, b) => {
    if (a.week !== b.week) return a.week - b.week;
    return a.household.localeCompare(b.household);
  });
  
  console.log(`📋 Master timeline created: ${masterTimeline.length} total visits over ${totalWeeks} weeks`);
  
  // ⭐ ENHANCED TRACKING for better assignment decisions
  const deaconWorkload = config.deacons.map(() => 0);
  const deaconLastAssignment = config.deacons.map(() => -999);
  const deaconHouseholdHistory = new Map();
  const deaconLastHouseholdVisit = new Map(); // NEW: Track when deacon last visited each household
  
  // Initialize tracking
  config.deacons.forEach((deacon, dIndex) => {
    config.households.forEach((household) => {
      const pairKey = `${dIndex}-${household}`;
      deaconHouseholdHistory.set(pairKey, 0);
      deaconLastHouseholdVisit.set(pairKey, -999); // Never visited before
    });
  });
  
  console.log('🎯 Assigning deacons with ENHANCED anti-repetition logic...');
  
  masterTimeline.forEach((visit, visitIndex) => {
    // Check execution time
    if (new Date().getTime() - startTime > maxExecutionTime) {
      throw new Error(`⌛ Schedule too large - exceeded time limit at visit ${visitIndex + 1}.`);
    }
    
    // ⭐ ENHANCED SCORING SYSTEM
    let bestDeacon = -1;
    let bestScore = Infinity;
    
    config.deacons.forEach((deacon, deaconIndex) => {
      let score = 0;
      
      // Factor 1: Workload balancing (unchanged)
      score += deaconWorkload[deaconIndex] * 100;
      
      // Factor 2: Recent activity penalty (unchanged)
      const weeksSinceLastAssignment = visit.week - deaconLastAssignment[deaconIndex];
      if (weeksSinceLastAssignment < 0) {
        score -= 50; // Never assigned before bonus
      } else {
        score -= weeksSinceLastAssignment * 10;
      }
      
      // ⭐ Factor 3: STRENGTHENED household variety penalty
      const pairKey = `${deaconIndex}-${visit.household}`;
      const previousVisits = deaconHouseholdHistory.get(pairKey) || 0;
      
      if (previousVisits > 0) {
        // MUCH stronger penalty for repeat visits
        score += previousVisits * 500; // Increased from 200 to 500
        
        // ⭐ Factor 4: NEW - Recent household visit penalty
        const lastHouseholdVisit = deaconLastHouseholdVisit.get(pairKey) || -999;
        const weeksSinceHouseholdVisit = visit.week - lastHouseholdVisit;
        
        if (weeksSinceHouseholdVisit > 0 && weeksSinceHouseholdVisit < 8) {
          // Heavy penalty for visiting same household too soon
          const proximityPenalty = Math.max(0, 8 - weeksSinceHouseholdVisit) * 300;
          score += proximityPenalty;
          
          if (weeksSinceHouseholdVisit <= 3) {
            score += 1000; // MASSIVE penalty for very recent visits
          }
        }
      }
      
      // ⭐ Factor 5: NEW - Prevent same deacon consecutive weeks
      if (weeksSinceLastAssignment === 1) {
        score += 200; // Moderate penalty for consecutive week assignments
      }
      
      // Factor 6: Frequency-based preference (unchanged)
      if (visit.frequency === 1) {
        score -= 5; // Slight preference for weekly visits
      }
      
      // ⭐ Factor 7: NEW - Encourage variety in high-frequency households
      if (visit.frequency <= 2) {
        // For high-frequency households, strongly prefer deacons who haven't visited recently
        if (previousVisits === 0) {
          score -= 100; // Bonus for never visiting this high-frequency household
        }
      }
      
      if (score < bestScore) {
        bestScore = score;
        bestDeacon = deaconIndex;
      }
    });
    
    // ⭐ FALLBACK SAFETY: If all deacons have very high scores, choose least bad option
    if (bestDeacon === -1 || bestScore > 10000) {
      console.warn(`⚠️ All deacons have high scores for ${visit.household} on week ${visit.week}, choosing least busy`);
      bestDeacon = deaconWorkload.indexOf(Math.min(...deaconWorkload));
    }
    
    // Assign the best deacon
    const assignedDeacon = config.deacons[bestDeacon];
    
    // Calculate visit date
    const visitDate = new Date(config.startDate);
    visitDate.setDate(visitDate.getDate() + (visit.week - 1) * 7);
    
    // Create schedule entry
    const scheduleEntry = {
      cycle: Math.ceil(visit.week / (visit.frequency || config.defaultVisitFrequency)),
      week: visit.week,
      date: visitDate,
      household: visit.household,
      deacon: assignedDeacon,
      deaconIndex: bestDeacon,
      householdFrequency: visit.frequency,
      isCustomFrequency: visit.isCustom
    };
    
    schedule.push(scheduleEntry);
    
    // ⭐ UPDATE ENHANCED TRACKING
    deaconWorkload[bestDeacon]++;
    deaconLastAssignment[bestDeacon] = visit.week;
    const pairKey = `${bestDeacon}-${visit.household}`;
    deaconHouseholdHistory.set(pairKey, (deaconHouseholdHistory.get(pairKey) || 0) + 1);
    deaconLastHouseholdVisit.set(pairKey, visit.week); // Track when this pair last occurred
  });
  
  // Enhanced logging
  console.log('📊 FINAL WORKLOAD DISTRIBUTION:');
  config.deacons.forEach((deacon, index) => {
    const workload = deaconWorkload[index];
    const percentage = Math.round((workload / totalVisitsAcrossAllHouseholds) * 100);
    console.log(`${deacon}: ${workload} visits (${percentage}%)`);
  });
  
  // ⭐ NEW: Check for potential issues in the final schedule
  validateScheduleQuality(schedule, config);
  
  logVariableFrequencyAnalysis(schedule, config);
  
  console.log(`✅ Enhanced variable frequency schedule generated: ${schedule.length} visits`);
  return schedule;
}

function validateScheduleQuality(schedule, config) {
  /**
   * ⭐ NEW: Post-generation quality validation
   */
  console.log('\n🔍 === SCHEDULE QUALITY VALIDATION ===');
  
  const issues = [];
  const warnings = [];
  
  // Check 1: Same deacon visiting same household too frequently
  const deaconHouseholdVisits = {};
  schedule.forEach(visit => {
    const key = `${visit.deacon}-${visit.household}`;
    if (!deaconHouseholdVisits[key]) {
      deaconHouseholdVisits[key] = [];
    }
    deaconHouseholdVisits[key].push(visit.week);
  });
  
  Object.keys(deaconHouseholdVisits).forEach(key => {
    const [deacon, household] = key.split('-', 2);
    const weeks = deaconHouseholdVisits[key].sort((a, b) => a - b);
    
    if (weeks.length > 1) {
      // Check gaps between visits
      for (let i = 1; i < weeks.length; i++) {
        const gap = weeks[i] - weeks[i-1];
        if (gap <= 3) {
          issues.push(`${deacon} visits ${household} too frequently: Week ${weeks[i-1]} → Week ${weeks[i]} (${gap} week gap)`);
        } else if (gap <= 6) {
          warnings.push(`${deacon} visits ${household} relatively soon: Week ${weeks[i-1]} → Week ${weeks[i]} (${gap} week gap)`);
        }
      }
      
      // Check for excessive repetition
      if (weeks.length > Math.ceil(config.numWeeks / 12)) {
        warnings.push(`${deacon} visits ${household} ${weeks.length} times (may be too often)`);
      }
    }
  });
  
  // Check 2: Deacon assignment gaps
  const deaconAssignments = {};
  schedule.forEach(visit => {
    if (!deaconAssignments[visit.deacon]) {
      deaconAssignments[visit.deacon] = [];
    }
    deaconAssignments[visit.deacon].push(visit.week);
  });
  
  Object.keys(deaconAssignments).forEach(deacon => {
    const weeks = deaconAssignments[deacon].sort((a, b) => a - b);
    
    for (let i = 1; i < weeks.length; i++) {
      const gap = weeks[i] - weeks[i-1];
      if (gap === 1) {
        warnings.push(`${deacon} has assignments in consecutive weeks: ${weeks[i-1]} → ${weeks[i]}`);
      }
    }
  });
  
  // Report results
  if (issues.length > 0) {
    console.log(`❌ Quality Issues Found (${issues.length}):`);
    issues.forEach((issue, index) => {
      console.log(`${index + 1}. ${issue}`);
    });
  }
  
  if (warnings.length > 0) {
    console.log(`⚠️ Quality Warnings (${warnings.length}):`);
    warnings.forEach((warning, index) => {
      console.log(`${index + 1}. ${warning}`);
    });
  }
  
  if (issues.length === 0 && warnings.length === 0) {
    console.log('✅ Schedule quality validation passed');
  }
  
  return { issues, warnings };
}
