# Features Guide - Deacon Visitation Rotation System v25.0

> **Technical deep-dive into system capabilities and implementation details**

## 🎯 Core Mathematical Algorithm

### **Harmonic Resonance Elimination**
The system solves a complex scheduling problem where simple rotation algorithms fail due to mathematical harmonics:

```javascript
// Problem: When deacon count and household count share common factors
// Example: 12 deacons, 6 households (ratio 2:1)
// Simple algorithms cause each deacon to visit only 2 specific households

// Solution: Modular arithmetic with prime pattern generation
function generateOptimalRotation(deacons, households, weeks) {
  const pattern = [];
  for (let week = 0; week < weeks; week++) {
    for (let d = 0; d < deacons.length; d++) {
      const householdIndex = (d + week * getPrimeFactor(deacons.length)) % households.length;
      pattern.push([deacons[d], households[householdIndex], week]);
    }
  }
  return pattern;
}
```

**Guaranteed Properties:**
- Every deacon visits every household over time
- Maximum spacing between repeat visits
- No mathematical "locks" preventing fair distribution
- Optimal load balancing across all participants

## 🔔 Google Chat Notification System (v25.0)

### **Automated Weekly Summaries**
The notification system provides rich, formatted messages via Google Chat webhooks:

```javascript
// Core notification function with 2-week lookahead
function sendWeeklyVisitationChat() {
  const visits = getUpcomingVisits(); // Next 2 calendar weeks
  const weekGroups = groupVisitsByCalendarWeek(visits);
  
  const message = buildChatMessage(weekGroups, {
    includeContactInfo: true,
    includeBreezeLinks: true,
    includeNotesLinks: true,
    includeCalendarLink: getCalendarLinkFromSpreadsheet() // K19 URL
  });
  
  sendToChatSpace(message, getCurrentTestMode());
}
```

**Message Format:**
- **Week grouping**: "Week 1 (Jul 14-20)" and "Week 2 (Jul 21-27)"
- **Deacon sections**: Organized by visiting deacon
- **Contact details**: Phone numbers and addresses
- **Direct links**: Clickable Breeze profiles and Notes documents
- **Calendar access**: "📅 View Visitation Calendar" link from K19

### **Configurable Calendar Links (v25.0)**
New K19 configuration enables dynamic calendar linking in notifications:

```javascript
function getCalendarLinkFromSpreadsheet() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const calendarUrl = sheet.getRange('K19').getValue();
  
  if (calendarUrl && typeof calendarUrl === 'string' && calendarUrl.trim().length > 0) {
    return calendarUrl.trim();
  }
  
  return null; // Graceful degradation when not configured
}

function buildCalendarLinkSection(calendarUrl) {
  if (!calendarUrl) return '';
  
  return `\n📅 [View Visitation Calendar](${calendarUrl})\n`;
}
```

**Configuration Options:**
- **K18**: "Google Calendar URL:" (header)
- **K19**: Calendar URL input field (user configurable)
- **Test/Production Switching**: Change URL in K19 to switch environments
- **Graceful Handling**: Links only appear when K19 has valid URL

### **Test Mode Integration**
Notifications respect the existing test mode detection system:

```javascript
function getCurrentTestMode() {
  // Leverages Module 3's intelligent test detection
  const sheet = SpreadsheetApp.getActiveSheet();
  const config = getConfiguration(sheet);
  
  // Multiple detection triggers:
  return (
    hasTestDataPatterns(config.households) ||  // Sample names
    hasTestPhoneNumbers(config.phoneNumbers) ||  // 555 numbers
    hasTestBreezeNumbers(config.breezeNumbers) ||  // Test Breeze IDs
    hasTestSpreadsheetTitle(sheet.getName())  // "Test" in title
  );
}
```

**Mode-Aware Features:**
- **Test Mode**: Messages prefixed with "🧪 TEST:"
- **Webhook Separation**: Different URLs for test vs production
- **Data Validation**: Prevents accidental production notifications during development
- **Visual Indicators**: K16 shows current detected mode

## 📅 Smart Calendar Functions (v24.2-v25.0)

### **Contact-Only Updates**
Preserves all deacon customizations while updating contact information:

```javascript
function updateContactInfoOnly() {
  const events = calendar.getEvents(startDate, endDate);
  const currentConfig = getConfiguration(sheet);
  
  events.forEach(event => {
    const householdName = extractHouseholdFromTitle(event.getTitle());
    const household = findHouseholdInConfig(householdName, currentConfig);
    
    if (household) {
      // Update ONLY contact info, preserve scheduling
      const newDescription = buildEventDescription(
        household,
        event.getTitle(), // Preserve original title
        currentConfig.eventInstructions
      );
      
      event.setDescription(newDescription);
      // Does NOT modify: date, time, location, guests, etc.
    }
  });
}
```

**Preservation Guarantees:**
- **Original timing**: Start/end times remain unchanged
- **Custom locations**: Meeting venues preserved
- **Guest lists**: Additional attendees maintained
- **Modified titles**: Custom deacon notes kept
- **Recurring patterns**: Series modifications intact

### **Future Events Only Updates**
Protects current week scheduling while updating upcoming events:

```javascript
function updateFutureEventsOnly() {
  const today = new Date();
  const nextWeekStart = new Date(today);
  nextWeekStart.setDate(today.getDate() + (7 - today.getDay())); // Next Sunday
  
  const events = calendar.getEvents(nextWeekStart, endDate);
  // Only processes events starting next week or later
  // Current week events remain completely untouched
}
```

## 🗓️ Enhanced Event Creation

### **Rich Event Descriptions**
Calendar events include comprehensive contact and reference information:

```javascript
function buildEventDescription(household, deacon, config) {
  const parts = [
    `Household: ${household.name}`,
    `Breeze Profile: ${household.breezeUrl || 'Not configured'}`,
    '',
    'Contact Information:',
    `Phone: ${household.phone || 'Not provided'}`,
    `Address: ${household.address || 'Not provided'}`,
    '',
    `Visit Notes: ${household.notesUrl || 'Not configured'}`,
    '',
    'Instructions:',
    config.eventInstructions
  ];
  
  if (getCalendarLinkFromSpreadsheet()) {
    parts.push('', `Visitation Guide available at ${getCalendarLinkFromSpreadsheet()}`);
  }
  
  return parts.join('\n');
}
```

**Event Components:**
- **Household identification** with clear naming
- **Breeze CMS integration** with direct profile links
- **Complete contact info** including phone and address
- **Visit notes access** via Google Docs URLs
- **Custom instructions** from K8 configuration
- **Calendar reference** from K19 when configured

### **URL Shortening Integration**
Mobile-friendly links for field access:

```javascript
function generateAndStoreShortUrls(sheet, config) {
  const breezeUrls = config.breezeNumbers.map(buildBreezeUrl);
  const notesUrls = config.notesLinks;
  
  // Batch processing with rate limiting
  const shortenedBreeze = breezeUrls.map(url => shortenUrl(url));
  const shortenedNotes = notesUrls.map(url => shortenUrl(url));
  
  // Store in columns R and S for reuse
  updateShortUrlColumns(sheet, shortenedBreeze, shortenedNotes);
}
```

## 🏗️ Enhanced Modular Architecture (v25.0)

### Five-Module System
```javascript
// Module distribution and responsibilities
Module1_Core_Config.gs      // Configuration, validation, headers (~500 lines)
Module2_Algorithm.gs        // Rotation algorithm, pattern generation (~470 lines)
Module3_Smart_Calendar.gs   // Calendar updates, mode detection (~200 lines)
Module4_Export_Menu.gs      // Full export, individual schedules, menu system (~400 lines)
Module5_Notifications.gs    // Google Chat integration, triggers (~1288 lines)
```

**Cross-Module Communication:**
- **Shared constants** accessible across all modules
- **Function availability** throughout execution context
- **Global variables** for configuration state
- **Error propagation** with module-specific identification

### Configuration Management System
**Enhanced Column K Layout (v25.0):**
```javascript
K1:  Start Date                        K2:  [User input]
K3:  Visits every x weeks              K4:  [User input]
K5:  Length of schedule in weeks       K6:  [User input]
K7:  Calendar Event Instructions       K8:  [User input]
K9:  [Buffer Space]
K10: Weekly Notification Day           K11: [Dropdown validation]
K12: Weekly Notification Time (0-23)   K13: [Dropdown validation]
K14: [Buffer Space]
K15: Current Mode                      K16: [Auto-detected]
K17: [Buffer Space]
K18: Google Calendar URL               K19: [User configurable] ⭐ NEW
```

**Data Validation Integration:**
- **Dropdown menus** prevent invalid day/time entries (K11, K13)
- **Real-time validation** at spreadsheet level
- **Error prevention** through UI constraints
- **User guidance** with help text and examples

## 🔧 Trigger Management System

### **Automated Weekly Triggers**
Robust trigger creation with Google Apps Script API handling:

```javascript
function createWeeklyNotificationTrigger() {
  const config = getWeeklyTriggerConfig(sheet);
  
  if (!config.isValid) {
    showConfigurationErrors(config.errors);
    return;
  }
  
  // Create time-based trigger
  const trigger = ScriptApp.newTrigger('sendWeeklyVisitationChat')
    .timeBased()
    .everyWeeks(1)
    .onWeekDay(config.dayEnum)
    .atHour(config.hour)
    .create();
    
  console.log(`Weekly trigger created: ${config.dayName} at ${config.timeFormatted}`);
}
```

**Trigger Configuration:**
- **Day selection**: K11 dropdown (Sunday through Saturday)
- **Time selection**: K13 dropdown (0-23 hour format)
- **Automatic scheduling**: Google Apps Script manages execution
- **Error handling**: Graceful failures with user feedback

### **Google Apps Script Limitations**
Understanding and working with platform constraints:

```javascript
// Trigger timing realities
const EXPECTED_DELAYS = {
  newTriggers: '24-48 hours to stabilize',
  executionDelay: '15-20+ minutes typical',
  serviceInterruptions: 'Occasional Google service issues',
  quotaLimits: 'Daily execution time limits'
};
```

## 📊 Data Integration Features

### **Breeze CMS Integration**
Direct integration with church management system:

```javascript
function buildBreezeUrl(breezeNumber) {
  if (!breezeNumber || breezeNumber.trim().length === 0) {
    return '';
  }
  
  const cleanNumber = breezeNumber.trim();
  // Validate 8-digit format (flexible for different churches)
  if (!/^\d{1,10}$/.test(cleanNumber)) {
    console.warn(`Potentially invalid Breeze number: ${cleanNumber}`);
  }
  
  return `https://immanuelky.breezechms.com/people/view/${cleanNumber}`;
}
```

**Integration Benefits:**
- **Direct profile access** from calendar events
- **Member information lookup** during visits
- **Contact verification** and updates
- **Visit history tracking** coordination

### **Google Docs Notes Integration**
Seamless visit documentation workflow:

```javascript
function validateNotesLinks(notesLinks, householdNames) {
  notesLinks.forEach((link, index) => {
    if (link && link.length > 0) {
      // Basic URL validation
      if (!link.startsWith('http://') && !link.startsWith('https://')) {
        console.warn(`Notes link for ${householdNames[index]} may be invalid: ${link}`);
      }
    }
  });
}
```

**Documentation Features:**
- **Pre-visit preparation** with historical context
- **Post-visit recording** of interactions
- **Continuity tracking** between different deacons
- **Shared knowledge base** for pastoral care

## 🧪 Test Mode Detection System

### **Intelligent Pattern Recognition**
Multi-factor detection prevents accidental production use:

```javascript
function detectTestMode(config) {
  const testIndicators = {
    sampleNames: hasTestNamePatterns(config.households),
    testPhones: hasTestPhoneNumbers(config.phoneNumbers),
    testBreeze: hasTestBreezeNumbers(config.breezeNumbers),
    titleIndicator: hasTestInTitle(SpreadsheetApp.getActiveSheet().getName())
  };
  
  return Object.values(testIndicators).some(indicator => indicator);
}

function hasTestNamePatterns(households) {
  const testPatterns = ['Alan Adams', 'Barbara Baker', 'Chloe Cooper', 'Test'];
  return households.some(name => 
    testPatterns.some(pattern => name.includes(pattern))
  );
}
```

**Detection Triggers:**
- **Sample data names**: Alan Adams, Barbara Baker, etc.
- **Test phone numbers**: 555-xxx-xxxx patterns
- **Test Breeze numbers**: Specific test IDs
- **Spreadsheet title**: Contains "test" or "sample"

### **Mode-Aware Operations**
System behavior adapts to detected environment:

```javascript
function getModeAwareSettings(isTestMode) {
  return {
    calendarName: isTestMode ? 'Test Deacon Calendar' : 'Deacon Visitation Schedule',
    chatPrefix: isTestMode ? '🧪 TEST: ' : '',
    webhookProperty: isTestMode ? 'CHAT_WEBHOOK_TEST' : 'CHAT_WEBHOOK_PROD',
    rateLimiting: isTestMode ? 'AGGRESSIVE' : 'CONSERVATIVE'
  };
}
```

## 🔍 Diagnostic and Testing Tools

### **System Validation Suite**
Comprehensive testing of all system components:

```javascript
function runSystemTests() {
  const testSuite = [
    'Configuration loading and validation',
    'URL shortening service connectivity',
    'Breeze URL construction',
    'Calendar access and permissions',
    'Script properties and storage',
    'Notification webhook connectivity',
    'Test mode detection accuracy'
  ];
  
  const results = testSuite.map(runIndividualTest);
  displayTestResults(results);
}
```

**Test Categories:**
- **Configuration validation**: Data structure and format checking
- **External service connectivity**: TinyURL, Google APIs
- **Permission verification**: Calendar, Chat, Properties access
- **Mode detection accuracy**: Test vs production identification
- **Notification delivery**: Webhook functionality testing

### **Error Handling and Recovery**
Graceful failure management throughout the system:

```javascript
function handleApiError(error, operation, fallbackAction) {
  console.error(`${operation} failed:`, error);
  
  if (error.message.includes('Rate limit exceeded')) {
    return scheduleRetry(operation, 60000); // 1 minute delay
  }
  
  if (error.message.includes('Permissions')) {
    return promptReauthorization();
  }
  
  return fallbackAction();
}
```

## 🚀 Performance Optimization

### **Rate Limiting Strategy**
Proactive management of Google API quotas:

```javascript
function processWithRateLimit(items, processor, batchSize = 25) {
  for (let i = 0; i < items.length; i += batchSize) {
    const batch = items.slice(i, i + batchSize);
    batch.forEach(processor);
    
    if (i + batchSize < items.length) {
      console.log(`Processed ${i + batchSize}/${items.length}, pausing...`);
      Utilities.sleep(2000); // 2-second pause between batches
    }
  }
}
```

**Optimization Features:**
- **Intelligent batching**: Processes items in manageable chunks
- **Progress tracking**: Real-time logging for user feedback
- **Error isolation**: Individual failures don't stop batch processing
- **Recovery mechanisms**: Automatic retry for transient failures

### **Memory Management**
Efficient handling of large datasets:

```javascript
function processLargeSchedule(schedule) {
  // Process in chunks to avoid memory limits
  const CHUNK_SIZE = 100;
  
  for (let i = 0; i < schedule.length; i += CHUNK_SIZE) {
    const chunk = schedule.slice(i, i + CHUNK_SIZE);
    processScheduleChunk(chunk);
    
    // Force garbage collection between chunks
    if (i % (CHUNK_SIZE * 3) === 0) {
      Utilities.sleep(1000);
    }
  }
}
```

## 📱 User Experience Features

### **Enhanced Menu System**
Intuitive navigation with logical grouping:

```javascript
function createEnhancedMenu() {
  const ui = SpreadsheetApp.getUi();
  const currentMode = getCurrentTestMode();
  const modeIcon = currentMode ? '🧪' : '✅';
  
  ui.createMenu('🔄 Deacon Rotation')
    .addItem('📅 Generate Schedule', 'generateRotationSchedule')
    .addSeparator()
    .addSubMenu(ui.createMenu('📆 Calendar Functions')
      .addItem('📞 Update Contact Info Only', 'updateContactInfoOnly')
      .addItem('🔄 Update Future Events Only', 'updateFutureEventsOnly')
      .addItem('🚨 Full Calendar Regeneration', 'exportToGoogleCalendar'))
    .addSubMenu(ui.createMenu('📢 Notifications')
      .addItem('💬 Send Weekly Chat Summary', 'sendWeeklyVisitationChat')
      .addItem('⏰ Send Tomorrow\'s Reminders', 'sendTomorrowReminders')
      .addSeparator()
      .addItem('🔧 Configure Chat Webhook', 'configureNotifications')
      .addItem('📋 Test Notification System', 'testNotificationSystem'))
    .addItem(`${modeIcon} Show Current Mode`, 'showModeNotification')
    .toUi();
}
```

**UX Improvements:**
- **Visual mode indicators**: 🧪 for test, ✅ for production
- **Logical grouping**: Related functions in submenus
- **Clear descriptions**: Function names explain purpose
- **Safety-first ordering**: Safer options listed first

### **User Feedback Systems**
Comprehensive status reporting and error communication:

```javascript
function showOperationResults(operation, results) {
  const ui = SpreadsheetApp.getUi();
  const summary = generateResultSummary(results);
  
  ui.alert(
    `${operation} Complete`,
    `✅ ${summary.successful} successful\n` +
    `⚠️ ${summary.warnings} warnings\n` +
    `❌ ${summary.errors} errors\n\n` +
    `${summary.details}`,
    ui.ButtonSet.OK
  );
}
```

## 🔐 Security and Data Protection

### **Webhook Security**
Secure storage and management of sensitive URLs:

```javascript
function storeWebhookSecurely(webhookUrl, isTestMode) {
  const properties = PropertiesService.getScriptProperties();
  const propertyKey = isTestMode ? 'CHAT_WEBHOOK_TEST' : 'CHAT_WEBHOOK_PROD';
  
  // Validate webhook URL format
  if (!webhookUrl.includes('chat.googleapis.com/v1/spaces/')) {
    throw new Error('Invalid Google Chat webhook URL format');
  }
  
  properties.setProperty(propertyKey, webhookUrl);
  console.log(`Webhook stored securely for ${isTestMode ? 'test' : 'production'} mode`);
}
```

**Security Features:**
- **Encrypted storage**: Google Apps Script Properties Service
- **Access control**: Only authorized scripts can access webhooks
- **URL validation**: Format checking prevents invalid configurations
- **Mode separation**: Test and production webhooks stored separately

### **Data Privacy Protection**
Member information handling with privacy safeguards:

```javascript
function sanitizeDataForLogging(data) {
  return {
    ...data,
    phoneNumbers: data.phoneNumbers.map(phone => phone ? 'XXX-XXX-XXXX' : ''),
    addresses: data.addresses.map(addr => addr ? '[ADDRESS REDACTED]' : ''),
    households: data.households.map(name => name ? '[NAME REDACTED]' : '')
  };
}
```

**Privacy Safeguards:**
- **Logging sanitization**: Personal info redacted from logs
- **Test data patterns**: Sample data for development/sharing
- **Access restrictions**: Member data only in production environment
- **Minimal exposure**: Contact info only shared via secure channels

---

This comprehensive feature set makes the Deacon Visitation Rotation System v25.0 a powerful, reliable, and user-friendly solution for church pastoral care coordination.
