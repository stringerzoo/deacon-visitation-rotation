/**
 * MODULE 4: ENHANCED MENU & TEST SYSTEM (v2.0)
 * Deacon Visitation Rotation System - Variable Frequency Support
 * 
 * 🆕 v2.0 ENHANCEMENTS:
 * - Enhanced menu system with v2.0 indicators
 * - Comprehensive testing for variable frequency scenarios
 * - Validation for mixed frequency configurations
 * - Backward compatibility testing with v1.1 systems
 * - Enhanced setup instructions for Column T usage
 */

function onOpen() {
  createMenu();
}

function createMenu() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();
  
  // Detect current system version and mode
  let versionIndicator = '📅';
  let modeIcon = '🔧';
  
  try {
    const config = getConfiguration();
    
    // v2.0 detection
    if (config.hasCustomFrequencies) {
      versionIndicator = '🆕 v2.0';
    } else {
      versionIndicator = '📅 v1.1';
    }
    
    // Mode detection (existing logic)
    const testIndicators = [
      () => config.households.some(h => h.toLowerCase().includes('alan') || h.toLowerCase().includes('adams')),
      () => config.phones.some(p => String(p).includes('555')),
      () => config.breezeLinks.some(b => String(b).startsWith('12345'))
    ];
    
    const isTestMode = testIndicators.some(check => check());
    modeIcon = isTestMode ? '🧪' : '✅';
    
  } catch (error) {
    console.log('Menu creation: Could not detect configuration');
  }
  
  ui.createMenu(`🔄 Deacon Rotation ${versionIndicator}`)
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
    .addSubMenu(ui.createMenu('🧪 Testing & Validation')
      .addItem('🔧 Validate Setup', 'validateSetupOnly')
      .addItem('🧪 Run System Tests', 'runSystemTests')
      .addItem('🆕 Test Variable Frequencies', 'testVariableFrequencies')
      .addItem('🔄 Migration Test (v1.1 → v2.0)', 'testMigrationCompatibility')
      .addItem('📊 Analyze Current Schedule', 'analyzeCurrentSchedule'))
    .addSeparator()
    .addItem(`${modeIcon} Show Current Mode`, 'showModeNotification')
    .addItem('❓ Setup Instructions', 'showSetupInstructions')
    .addToUi();
}

/**
 * 🆕 v2.0 ENHANCED SCHEDULE GENERATION
 */
function generateRotationSchedule() {
  try {
    const startTime = new Date().getTime();
    
    // Enhanced configuration loading
    console.log('🔧 Loading v2.0 configuration...');
    const config = getConfiguration();
    
    // Version detection and logging
    if (config.hasCustomFrequencies) {
      console.log('🆕 v2.0 Variable Frequency Mode detected');
      console.log(`📊 Frequency distribution: ${config.householdFrequencies.length} households with custom frequencies`);
      
      // Show user what's happening
      SpreadsheetApp.getUi().alert(
        'v2.0 Variable Frequency Generation',
        `🆕 Generating schedule with variable frequencies:\n\n` +
        `• Default frequency: ${config.defaultVisitFrequency} weeks\n` +
        `• Custom frequencies: ${config.householdFrequencies.filter(hf => hf.isCustom).length} households\n` +
        `• Total households: ${config.households.length}\n\n` +
        `This may take longer than v1.1 generation...`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } else {
      console.log('📅 v1.1 Uniform Frequency Mode (backward compatible)');
    }
    
    // Setup headers with v2.0 enhancements
    setupHeaders(SpreadsheetApp.getActiveSheet());
    
    // Generate schedule using enhanced algorithm
    console.log('⚡ Starting schedule generation...');
    const schedule = safeCreateSchedule(config);
    
    if (!schedule || schedule.length === 0) {
      throw new Error('No schedule was generated');
    }
    
    // Write schedule with v2.0 formatting
    console.log('📝 Writing schedule to spreadsheet...');
    writeScheduleToSheet(schedule, config);
    
    // Success notification with v2.0 details and formatting explanation
    const executionTime = ((new Date().getTime() - startTime) / 1000).toFixed(2);
    const customCount = config.hasCustomFrequencies 
      ? config.householdFrequencies.filter(hf => hf.isCustom).length 
      : 0;
    
    let message = `✅ v2.0 Schedule Generated!\n\n`;
    
    message += `📊 Statistics:\n`;
    message += `• Total visits: ${schedule.length}\n`;
    message += `• Time span: ${config.numWeeks} weeks\n`;
    message += `• Households: ${config.households.length}\n`;
    message += `• Custom frequencies: ${customCount}\n`;
    message += `• Generation time: ${executionTime} seconds\n\n`;
    
    // Add formatting explanation if custom frequencies exist
    if (config.hasCustomFrequencies) {
      message += `🎨 Formatting Notes:\n`;
      message += `• Light yellow rows = Custom frequency households\n`;
      message += `• Individual reports show frequency for custom households only\n`;
      message += `• Clean v1.1 formatting maintained\n\n`;
    }
    
    message += `${config.hasCustomFrequencies ? '🆕 Variable frequency algorithm used' : '📅 Uniform frequency (v1.1 compatible)'}`;
    
    SpreadsheetApp.getUi().alert(
      'Schedule Generated Successfully',
      message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    console.error('Schedule generation failed:', error);
    SpreadsheetApp.getUi().alert(
      'Generation Failed',
      `❌ Could not generate schedule:\n\n${error.message}\n\nPlease check:\n• Configuration in column K\n• Deacon list in column L\n• Household list in column M\n• Custom frequencies in column T (v2.0)`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * 🆕 v2.0 VARIABLE FREQUENCY TESTING
 */
function testVariableFrequencies() {
  try {
    console.log('🧪 === VARIABLE FREQUENCY TEST SUITE ===');
    
    const config = getConfiguration();
    const issues = [];
    const warnings = [];
    
    console.log('📋 Testing variable frequency configuration...');
    
    // Test 1: Check if Column T has data
    if (!config.hasCustomFrequencies) {
      warnings.push('No custom frequencies detected in Column T. System will use uniform frequency from K4.');
    } else {
      console.log(`✅ Custom frequencies found: ${config.householdFrequencies.filter(hf => hf.isCustom).length} households`);
    }
    
    // Test 2: Validate frequency values
    const invalidFreq = config.householdFrequencies.filter(hf => 
      hf.frequency < 1 || hf.frequency > 4
    );
    if (invalidFreq.length > 0) {
      issues.push(`Invalid frequencies found: ${invalidFreq.map(hf => `${hf.household}: ${hf.frequency}`).join(', ')}`);
    }
    
    // Test 3: Check frequency distribution
    const freqDistribution = {};
    config.householdFrequencies.forEach(hf => {
      const key = hf.frequency;
      if (!freqDistribution[key]) freqDistribution[key] = [];
      freqDistribution[key].push(hf.household);
    });
    
    console.log('📊 Frequency Distribution:');
    Object.keys(freqDistribution).forEach(freq => {
      const households = freqDistribution[freq];
      console.log(`${freq}-week: ${households.length} households (${households.slice(0,2).join(', ')}${households.length > 2 ? '...' : ''})`);
    });
    
    // Test 4: Calculate expected visit counts
    const totalWeeks = config.numWeeks;
    let totalExpectedVisits = 0;
    
    config.householdFrequencies.forEach(hf => {
      const expectedVisits = Math.floor(totalWeeks / hf.frequency);
      totalExpectedVisits += expectedVisits;
    });
    
    console.log(`📈 Expected total visits: ${totalExpectedVisits} over ${totalWeeks} weeks`);
    console.log(`👥 Average per deacon: ${(totalExpectedVisits / config.deacons.length).toFixed(1)}`);
    
    // Test 5: Check for workload balance concerns
    const maxVisitsPerHousehold = Math.max(...config.householdFrequencies.map(hf => 
      Math.floor(totalWeeks / hf.frequency)
    ));
    const minVisitsPerHousehold = Math.min(...config.householdFrequencies.map(hf => 
      Math.floor(totalWeeks / hf.frequency)
    ));
    
    if (maxVisitsPerHousehold / minVisitsPerHousehold > 2) {
      warnings.push(`Large workload variation: ${minVisitsPerHousehold}-${maxVisitsPerHousehold} visits per household. Consider balancing frequencies.`);
    }
    
    // Test 6: Test actual schedule generation (small sample)
    console.log('🔄 Testing schedule generation with current configuration...');
    
    try {
      const testConfig = { ...config, numWeeks: Math.min(8, config.numWeeks) }; // Test with 8 weeks max
      const testSchedule = safeCreateSchedule(testConfig);
      console.log(`✅ Test generation successful: ${testSchedule.length} visits in ${testConfig.numWeeks} weeks`);
    } catch (scheduleError) {
      issues.push(`Schedule generation test failed: ${scheduleError.message}`);
    }
    
    // Show results
    let message = '🧪 Variable Frequency Test Results:\n\n';
    
    if (issues.length === 0) {
      message += '✅ All tests passed!\n\n';
    } else {
      message += `❌ Issues found (${issues.length}):\n`;
      issues.forEach((issue, index) => {
        message += `${index + 1}. ${issue}\n`;
      });
      message += '\n';
    }
    
    if (warnings.length > 0) {
      message += `⚠️ Warnings (${warnings.length}):\n`;
      warnings.forEach((warning, index) => {
        message += `${index + 1}. ${warning}\n`;
      });
      message += '\n';
    }
    
    message += '📊 Configuration Summary:\n';
    message += `• Default frequency: ${config.defaultVisitFrequency} weeks\n`;
    message += `• Custom frequencies: ${config.householdFrequencies.filter(hf => hf.isCustom).length}/${config.households.length} households\n`;
    message += `• Expected total visits: ${totalExpectedVisits}\n`;
    message += `• Schedule length: ${config.numWeeks} weeks`;
    
    SpreadsheetApp.getUi().alert(
      'Variable Frequency Test Results',
      message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    console.error('Variable frequency test failed:', error);
    SpreadsheetApp.getUi().alert(
      'Test Failed',
      `❌ Variable frequency test encountered an error:\n\n${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * 🔄 MIGRATION COMPATIBILITY TEST
 */
function testMigrationCompatibility() {
  try {
    console.log('🔄 === MIGRATION COMPATIBILITY TEST ===');
    
    const config = getConfiguration();
    const issues = [];
    const recommendations = [];
    
    // Test 1: Check if this is a v1.1 system
    const isV11System = !config.hasCustomFrequencies;
    
    if (isV11System) {
      console.log('📅 v1.1 system detected - testing upgrade readiness');
      
      // Check if Column T is empty and available
      const sheet = SpreadsheetApp.getActiveSheet();
      const columnTData = sheet.getRange('T2:T100').getValues().flat().filter(cell => cell !== '');
      
      if (columnTData.length > 0) {
        issues.push('Column T contains data. Please clear Column T before upgrading to v2.0.');
      } else {
        recommendations.push('Column T is empty and ready for custom frequency data.');
      }
      
      // Test backward compatibility
      try {
        const testSchedule = safeCreateSchedule(config);
        console.log(`✅ v1.1 compatibility confirmed: ${testSchedule.length} visits generated`);
        recommendations.push('System can generate v1.1 compatible schedules successfully.');
      } catch (error) {
        issues.push(`v1.1 compatibility issue: ${error.message}`);
      }
      
    } else {
      console.log('🆕 v2.0 system detected - testing v1.1 fallback');
      
      // Test v1.1 fallback by temporarily removing custom frequencies
      const v11Config = { ...config, hasCustomFrequencies: false };
      
      try {
        const testSchedule = safeCreateSchedule(v11Config);
        console.log(`✅ v1.1 fallback confirmed: ${testSchedule.length} visits generated`);
        recommendations.push('System can fall back to v1.1 behavior if needed.');
      } catch (error) {
        issues.push(`v1.1 fallback issue: ${error.message}`);
      }
    }
    
    // Test 2: Check column structure compatibility
    const expectedColumns = ['M', 'N', 'O', 'P', 'Q', 'R', 'S'];
    const sheet = SpreadsheetApp.getActiveSheet();
    
    expectedColumns.forEach(col => {
      const header = sheet.getRange(`${col}1`).getValue();
      if (!header) {
        issues.push(`Missing header in column ${col}. Run "Setup Headers" to fix.`);
      }
    });
    
    // Test 3: Data integrity check
    const dataIntegrityIssues = [];
    
    if (config.households.length !== config.phones.length && config.phones.filter(p => p).length > 0) {
      dataIntegrityIssues.push('Phone numbers count doesn\'t match households count');
    }
    
    if (config.households.length !== config.addresses.length && config.addresses.filter(a => a).length > 0) {
      dataIntegrityIssues.push('Addresses count doesn\'t match households count');
    }
    
    if (dataIntegrityIssues.length > 0) {
      issues.push(`Data integrity concerns: ${dataIntegrityIssues.join(', ')}`);
    }
    
    // Show results
    let message = '🔄 Migration Compatibility Test Results:\n\n';
    
    if (isV11System) {
      message += '📅 Current System: v1.1 (Uniform Frequency)\n';
      message += '🎯 Test Goal: v2.0 Upgrade Readiness\n\n';
    } else {
      message += '🆕 Current System: v2.0 (Variable Frequency)\n';
      message += '🎯 Test Goal: v1.1 Fallback Capability\n\n';
    }
    
    if (issues.length === 0) {
      message += '✅ Migration compatibility confirmed!\n\n';
    } else {
      message += `❌ Issues found (${issues.length}):\n`;
      issues.forEach((issue, index) => {
        message += `${index + 1}. ${issue}\n`;
      });
      message += '\n';
    }
    
    if (recommendations.length > 0) {
      message += `💡 Recommendations (${recommendations.length}):\n`;
      recommendations.forEach((rec, index) => {
        message += `${index + 1}. ${rec}\n`;
      });
    }
    
    SpreadsheetApp.getUi().alert(
      'Migration Compatibility Results',
      message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    console.error('Migration test failed:', error);
    SpreadsheetApp.getUi().alert(
      'Migration Test Failed',
      `❌ Migration compatibility test error:\n\n${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * 📊 CURRENT SCHEDULE ANALYSIS
 */
function analyzeCurrentSchedule() {
  try {
    console.log('📊 === CURRENT SCHEDULE ANALYSIS ===');
    
    const sheet = SpreadsheetApp.getActiveSheet();
    const config = getConfiguration();
    
    // Read current schedule from spreadsheet
    const scheduleData = sheet.getRange('A2:E200').getValues()
      .filter(row => row[0] !== '' && row[0] != null)
      .map(row => ({
        cycle: row[0],
        week: row[1], 
        weekOf: row[2],
        household: row[3],
        deacon: row[4]
      }));
    
    if (scheduleData.length === 0) {
      throw new Error('No schedule data found. Please generate a schedule first.');
    }
    
    console.log(`📋 Analyzing ${scheduleData.length} visits...`);
    
    // Analysis 1: Basic statistics
    const totalVisits = scheduleData.length;
    const uniqueWeeks = new Set(scheduleData.map(s => s.week)).size;
    const uniqueHouseholds = new Set(scheduleData.map(s => s.household.replace(/ \(\dw\)$/, ''))).size;
    const uniqueDeacons = new Set(scheduleData.map(s => s.deacon)).size;
    
    // Analysis 2: Frequency detection from schedule
    const householdVisitCounts = {};
    scheduleData.forEach(visit => {
      const household = visit.household.replace(/ \(\dw\)$/, ''); // Remove frequency markers
      if (!householdVisitCounts[household]) {
        householdVisitCounts[household] = [];
      }
      householdVisitCounts[household].push(visit.week);
    });
    
    // Calculate actual frequencies
    const actualFrequencies = {};
    Object.keys(householdVisitCounts).forEach(household => {
      const weeks = householdVisitCounts[household].sort((a, b) => a - b);
      if (weeks.length > 1) {
        const gaps = [];
        for (let i = 1; i < weeks.length; i++) {
          gaps.push(weeks[i] - weeks[i-1]);
        }
        const avgGap = gaps.reduce((sum, gap) => sum + gap, 0) / gaps.length;
        actualFrequencies[household] = Math.round(avgGap);
      } else {
        actualFrequencies[household] = 'Single visit';
      }
    });
    
    // Analysis 3: Deacon workload
    const deaconWorkload = {};
    scheduleData.forEach(visit => {
      if (!deaconWorkload[visit.deacon]) {
        deaconWorkload[visit.deacon] = 0;
      }
      deaconWorkload[visit.deacon]++;
    });
    
    const workloads = Object.values(deaconWorkload);
    const minWorkload = Math.min(...workloads);
    const maxWorkload = Math.max(...workloads);
    const avgWorkload = workloads.reduce((sum, w) => sum + w, 0) / workloads.length;
    
    // Analysis 4: Time distribution
    const weeklyVisits = {};
    scheduleData.forEach(visit => {
      if (!weeklyVisits[visit.week]) {
        weeklyVisits[visit.week] = 0;
      }
      weeklyVisits[visit.week]++;
    });
    
    const weeksWithVisits = Object.keys(weeklyVisits).length;
    const avgVisitsPerWeek = totalVisits / weeksWithVisits;
    
    // Show analysis results
    let message = '📊 Current Schedule Analysis:\n\n';
    
    message += '📈 Basic Statistics:\n';
    message += `• Total visits: ${totalVisits}\n`;
    message += `• Time span: ${uniqueWeeks} weeks\n`;
    message += `• Households: ${uniqueHouseholds}\n`;
    message += `• Deacons: ${uniqueDeacons}\n`;
    message += `• Average visits/week: ${avgVisitsPerWeek.toFixed(1)}\n\n`;
    
    message += '⚖️ Workload Balance:\n';
    message += `• Min: ${minWorkload}, Max: ${maxWorkload}, Avg: ${avgWorkload.toFixed(1)}\n`;
    message += `• Balance quality: ${maxWorkload - minWorkload <= 1 ? '✅ Excellent' : maxWorkload - minWorkload <= 2 ? '🟡 Good' : '🔴 Needs improvement'}\n\n`;
    
    message += '🔄 Frequency Analysis:\n';
    const freqStats = {};
    Object.values(actualFrequencies).forEach(freq => {
      const key = freq === 'Single visit' ? freq : `${freq}-week`;
      if (!freqStats[key]) freqStats[key] = 0;
      freqStats[key]++;
    });
    
    Object.keys(freqStats).forEach(freq => {
      message += `• ${freq}: ${freqStats[freq]} households\n`;
    });
    
    message += `\n📅 System Version: ${config.hasCustomFrequencies ? '🆕 v2.0 (Variable)' : '📋 v1.1 (Uniform)'}`;
    
    SpreadsheetApp.getUi().alert(
      'Schedule Analysis Results',
      message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    console.error('Schedule analysis failed:', error);
    SpreadsheetApp.getUi().alert(
      'Analysis Failed',
      `❌ Could not analyze schedule:\n\n${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * ENHANCED SYSTEM TESTS (v2.0)
 */
function runSystemTests() {
  try {
    console.log('🧪 === ENHANCED SYSTEM TEST SUITE (v2.0) ===');
    
    const tests = [];
    const warnings = [];
    
    // Test 1: Basic configuration
    let config;
    try {
      config = getConfiguration();
      tests.push({ name: 'Configuration Loading', status: 'PASS', details: `${config.deacons.length} deacons, ${config.households.length} households` });
    } catch (error) {
      tests.push({ name: 'Configuration Loading', status: 'FAIL', details: error.message });
      return showTestResults(tests, warnings);
    }
    
    // Test 2: v2.0 Feature Detection
    if (config.hasCustomFrequencies) {
      const customCount = config.householdFrequencies.filter(hf => hf.isCustom).length;
      tests.push({ name: 'v2.0 Variable Frequencies', status: 'DETECTED', details: `${customCount} custom frequencies in Column T` });
    } else {
      tests.push({ name: 'v2.0 Variable Frequencies', status: 'NONE', details: 'Using v1.1 uniform frequency mode' });
    }
    
    // Test 3: Data validation
    const dataIssues = [];
    
    if (config.households.length === 0) dataIssues.push('No households in column M');
    if (config.deacons.length === 0) dataIssues.push('No deacons in column L');
    if (config.visitFrequency < 1 || config.visitFrequency > 4) dataIssues.push('Invalid default frequency in K4');
    
    // v2.0 specific validation
    if (config.hasCustomFrequencies) {
      const invalidCustomFreq = config.householdFrequencies.filter(hf => 
        hf.isCustom && (hf.frequency < 1 || hf.frequency > 4)
      );
      if (invalidCustomFreq.length > 0) {
        dataIssues.push(`Invalid custom frequencies: ${invalidCustomFreq.map(hf => hf.household).join(', ')}`);
      }
    }
    
    if (dataIssues.length > 0) {
      tests.push({ name: 'Data Validation', status: 'FAIL', details: dataIssues.join('; ') });
    } else {
      tests.push({ name: 'Data Validation', status: 'PASS', details: 'All data properly configured' });
    }
    
    // Test 4: Schedule generation test
    try {
      const testConfig = { ...config, numWeeks: 4 }; // Short test
      const testSchedule = safeCreateSchedule(testConfig);
      tests.push({ name: 'Schedule Generation', status: 'PASS', details: `Generated ${testSchedule.length} visits` });
    } catch (error) {
      tests.push({ name: 'Schedule Generation', status: 'FAIL', details: error.message });
    }
    
    // Test 5: Calendar access test
    try {
      const testCalendar = getOrCreateCalendar();
      tests.push({ name: 'Calendar Access', status: 'PASS', details: `Calendar: ${testCalendar.getName()}` });
    } catch (error) {
      tests.push({ name: 'Calendar Access', status: 'FAIL', details: error.message });
    }
    
    // Test 6: Notification configuration test
    try {
      const webhook = PropertiesService.getScriptProperties().getProperty('CHAT_WEBHOOK_URL');
      if (webhook) {
        tests.push({ name: 'Notification Setup', status: 'CONFIGURED', details: 'Chat webhook configured' });
      } else {
        tests.push({ name: 'Notification Setup', status: 'NONE', details: 'No webhook configured' });
        warnings.push('Configure chat webhook for notifications');
      }
    } catch (error) {
      tests.push({ name: 'Notification Setup', status: 'ERROR', details: error.message });
    }
    
    // Test 7: Spreadsheet permissions
    try {
      const sheet = SpreadsheetApp.getActiveSheet();
      sheet.getRange('A1').setValue('Test').setValue(''); // Test write access
      tests.push({ name: 'Spreadsheet Permissions', status: 'PASS', details: 'Full read/write access' });
    } catch (error) {
      tests.push({ name: 'Spreadsheet Permissions', status: 'FAIL', details: 'Limited access: ' + error.message });
    }
    
    showTestResults(tests, warnings);
    
  } catch (error) {
    console.error('System tests failed:', error);
    SpreadsheetApp.getUi().alert(
      'Test Suite Failed', 
      `❌ System test error:\n\n${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

function showTestResults(tests, warnings) {
  let message = '🧪 System Test Results:\n\n';
  
  const passed = tests.filter(t => t.status === 'PASS').length;
  const failed = tests.filter(t => t.status === 'FAIL').length;
  const other = tests.length - passed - failed;
  
  message += `📊 Summary: ${passed} passed, ${failed} failed, ${other} other\n\n`;
  
  tests.forEach(test => {
    let icon;
    switch (test.status) {
      case 'PASS': icon = '✅'; break;
      case 'FAIL': icon = '❌'; break;
      case 'DETECTED': icon = '🆕'; break;
      case 'CONFIGURED': icon = '🔧'; break;
      case 'NONE': icon = '⚪'; break;
      default: icon = '⚠️';
    }
    
    message += `${icon} ${test.name}: ${test.status}\n`;
    if (test.details) {
      message += `   ${test.details}\n`;
    }
    message += '\n';
  });
  
  if (warnings.length > 0) {
    message += `⚠️ Warnings:\n`;
    warnings.forEach((warning, index) => {
      message += `${index + 1}. ${warning}\n`;
    });
  }
  
  SpreadsheetApp.getUi().alert(
    'System Test Results',
    message,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// Keep existing functions for backward compatibility
function validateSetupOnly() {
  // Existing validation function - unchanged for v1.1 compatibility
  try {
    const config = getConfiguration();
    SpreadsheetApp.getUi().alert(
      'Setup Validation',
      '✅ Setup validation passed!\n\n' +
      `• ${config.deacons.length} deacons configured\n` +
      `• ${config.households.length} households configured\n` +
      `• Default frequency: ${config.visitFrequency} weeks\n` +
      `• Schedule length: ${config.numWeeks} weeks\n` +
      `${config.hasCustomFrequencies ? '\n🆕 v2.0 custom frequencies detected' : ''}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (error) {
    SpreadsheetApp.getUi().alert(
      'Setup Validation Failed',
      `❌ Configuration issues found:\n\n${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}
