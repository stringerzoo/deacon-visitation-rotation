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
      console.log(`
