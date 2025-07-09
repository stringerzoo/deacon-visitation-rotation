# Technical Features Reference - Deacon Visitation Rotation System v25.0

This document provides deep technical explanations of the system's core algorithms, implementation details, and advanced features. For setup instructions, see [SETUP.md](SETUP.md). For project overview, see [README.md](README.md).

## 🎯 Core Rotation Engine

### Optimal Pattern Generator Algorithm
**The Problem**: Simple rotation algorithms fail when the ratio of deacons to households creates "harmonic locks" - mathematical resonance patterns where deacons only visit the same household(s) repeatedly.

**Examples of Problematic Ratios:**
- **12:6 ratio** → Each deacon visits only 2 specific households forever
- **14:7 ratio** → Complex harmonic creates uneven distribution
- **15:6 ratio** → Some deacons get locked out of certain households

**Our Solution - Intelligent Scoring System:**
```javascript
// Simplified scoring algorithm
for (let deaconIndex = 0; deaconIndex < deaconCount; deaconIndex++) {
  const totalVisits = deaconUsageCount[deaconIndex];
  const householdVisits = deaconHouseholdPairs.get(`${deaconIndex}-${householdIndex}`);
  
  // Lower score is better
  let score = totalVisits * 100 + householdVisits * 10;
  
  // Variety bonus for new pairings
  if (householdVisits === 0) score -= 50;
}
```

**Algorithm Benefits:**
- **Pre-calculated patterns** eliminate real-time harmonic issues
- **Intelligent scoring** prioritizes deacons with fewer visits
- **Variety bonuses** favor new deacon-household pairings
- **Cycle prevention** ensures no double assignments per cycle
- **Universal compatibility** works with any deacon/household ratio

### Harmonic Resonance Analysis
**Mathematical Foundation:**
```javascript
const gcd = (a, b) => b === 0 ? a : gcd(b, a % b);
const lcm = (a, b) => (a * b) / gcd(a, b);

const commonFactor = gcd(deaconCount, householdCount);
const isSimpleHarmonic = deaconCount % householdCount === 0;
const isComplexHarmonic = commonFactor > 1 && !isSimpleHarmonic;
```

**Harmonic Types Detected:**
- **Simple Harmonics**: Perfect divisors (12:6, 12:4, 12:3)
- **Complex Harmonics**: Common factors (14:7, 15:6, 21:9)
- **Prime Ratios**: No harmonics (13:7, 11:5) - naturally optimal

**Mitigation Strategies:**
- **Pattern pre-generation** bypasses harmonic calculations entirely
- **Dynamic scoring** adapts to any ratio automatically
- **Quality metrics** validate final rotation patterns
- **Fallback mechanisms** ensure schedule generation never fails

## 🔄 Smart Calendar Update System

### Event Preservation Architecture
**Core Challenge**: Google Calendar events contain both system data (contact info) and user customizations (times, guests, locations). Traditional approaches lose user customizations.

**Smart Parsing Implementation:**
```javascript
// Extract household from event title: "[Deacon] visits [Household]"
const visitMatch = title.match(/(.+) visits (.+)/);
if (visitMatch) {
  const household = visitMatch[2].trim();
  // Match to current spreadsheet data for updates
}
```

**Preservation Strategies:**
- **Description-only updates** preserve all event metadata
- **Selective event targeting** by date ranges
- **Smart data matching** handles household name changes
- **Error isolation** prevents cascade failures

### Dynamic Mode Detection Algorithm
**Intelligence Triggers:**
```javascript
const testIndicators = [
  () => households.some(h => testPatterns.some(p => h.toLowerCase().includes(p))),
  () => phones.some(phone => String(phone).includes('555')),
  () => breezeNumbers.some(num => String(num).startsWith('12345')),
  () => spreadsheetName.includes('test') || spreadsheetName.includes('sample')
];

const isTestMode = testIndicators.some(check => check());
```

**Real-Time Updates:**
- **Dynamic re-detection** on each calendar operation
- **Visual indicator refresh** without spreadsheet reload
- **Mode-aware dialogs** with current state information
- **Calendar name adaptation** based on detected mode

## 🔔 Google Chat Notification System (v25.0)

### Webhook-Based Architecture
**Core Implementation:**
```javascript
function sendToChatSpace(message) {
  const webhookUrl = PropertiesService.getScriptProperties().getProperty('CHAT_WEBHOOK_URL');
  
  const payload = {
    text: message,
    cards: [{
      sections: [{
        widgets: [{ textParagraph: { text: message } }]
      }]
    }]
  };
  
  UrlFetchApp.fetch(webhookUrl, {
    method: 'POST',
    contentType: 'application/json',
    payload: JSON.stringify(payload)
  });
}
```

**Advanced Features:**
- **Rich card formatting** with structured layouts
- **Direct link integration** to Breeze profiles and Notes documents
- **Test mode separation** with dedicated test webhooks
- **Error handling** with retry mechanisms and user feedback

### Automated Trigger Management
**Smart Trigger Creation:**
```javascript
function createWeeklyNotificationTrigger() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const notificationDay = sheet.getRange('K11').getValue();
  const notificationHour = parseInt(sheet.getRange('K13').getValue());
  
  // Delete existing triggers to avoid duplicates
  const existingTriggers = ScriptApp.getProjectTriggers()
    .filter(trigger => trigger.getHandlerFunction() === 'sendWeeklyVisitationChat');
  
  existingTriggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
  
  // Create new weekly trigger
  ScriptApp.newTrigger('sendWeeklyVisitationChat')
    .timeBased()
    .everyWeeks(1)
    .onWeekDay(getScriptAppWeekDay(notificationDay))
    .atHour(notificationHour)
    .create();
}
```

**Configuration Integration:**
- **Spreadsheet-based scheduling** with data validation dropdowns
- **Flexible timing** - any day of week, any hour (24-hour format)
- **Automatic trigger recreation** when settings change
- **Visual feedback** showing current trigger status

### Message Intelligence
**Smart Visit Filtering:**
```javascript
function getUpcomingVisits(daysAhead = 7) {
  const scheduleData = getScheduleFromSheet(sheet);
  const cutoffDate = new Date();
  cutoffDate.setDate(cutoffDate.getDate() + daysAhead);
  
  return scheduleData.filter(visit => {
    const visitDate = new Date(visit.date);
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    
    return visitDate >= today && visitDate <= cutoffDate;
  }).sort((a, b) => new Date(a.date) - new Date(b.date));
}
```

**Rich Message Formatting:**
- **Grouped by deacon** with clear visual separation
- **Contact information** including phones and addresses
- **Direct action links** to Breeze and Notes documents
- **Contextual instructions** for visit coordination

## 🏗️ Enhanced Modular Architecture (v25.0)

### Five-Module System
```javascript
// Module distribution and responsibilities
Module1_Core_Config.gs      // Configuration, validation, headers (~500 lines)
Module2_Algorithm.gs        // Rotation algorithm, pattern generation (~470 lines)
Module3_Smart_Calendar.gs   // Calendar updates, mode detection (~200 lines)
Module4_Export_Menu.gs      // Full export, individual schedules, menu system (~400 lines)
Module5_Notifications.gs    // Google Chat integration, triggers (~300 lines)
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
```

**Data Validation Integration:**
- **Dropdown menus** prevent invalid day/time entries
- **Real-time validation** at spreadsheet level
- **Error prevention** through UI constraints
- **User guidance** with clear option lists

## ⚙️ Advanced Configuration & Performance

### Rate Limiting Implementation
**Google Calendar API Protection:**
```javascript
// Deletion rate limiting
if ((index + 1) % 10 === 0) {
  Utilities.sleep(500); // 0.5 second pause every 10 deletions
}

// Creation rate limiting  
if ((index + 1) % 25 === 0) {
  Utilities.sleep(1000); // 1 second pause every 25 creations
  console.log(`Created ${index + 1} events so far...`);
}
```

**Google Apps Script Trigger Reliability:**
- **Trigger recreation** mechanisms for improved reliability
- **Multiple diagnostic tools** for troubleshooting delivery
- **Graceful degradation** when triggers experience delays
- **User feedback** about expected Google API limitations

### Algorithm Tuning Parameters
**Scoring System Weights:**
```javascript
let score = totalVisits * 100 + householdVisits * 10;
//           ↑ Primary weight    ↑ Secondary weight

if (householdVisits === 0) score -= 50; // Variety bonus
//                                  ↑ Incentive weight
```

**Pattern Quality Metrics:**
- **Visit imbalance**: Difference between max and min visits per deacon
- **Coverage percentage**: Deacons visiting all households
- **Variety index**: Average households per deacon

### Performance Optimization Settings
**Execution Time Management:**
```javascript
const maxExecutionTime = 4 * 60 * 1000; // 4 minutes safety buffer
const batchSize = 1000; // Events per batch write
const maxEvents = 500; // Calendar event limit
```

**Memory Management:**
- **Batch processing** for large datasets
- **Progressive updates** with status logging
- **Resource monitoring** with time limit enforcement
- **Graceful degradation** when approaching limits

## 🛡️ Error Recovery & Security

### Validation Layers
**Comprehensive Input Validation:**
1. **Input validation** - data types, ranges, duplicates
2. **Workload validation** - feasibility checking with recommendations
3. **Runtime validation** - date calculations, API responses
4. **Output validation** - schedule completeness verification

**Recovery Strategies:**
```javascript
try {
  // Primary operation
} catch (eventError) {
  console.error(`Failed operation: ${context}`, eventError);
  // Continue with remaining operations - don't fail entire process
}
```

**Safety Mechanisms:**
- **Individual error isolation** prevents cascade failures
- **Automatic retries** for transient failures
- **Fallback options** for service unavailability
- **User guidance** with specific error resolution steps

### Security Implementation
**Data Protection:**
- **Webhook URLs** stored securely in Apps Script Properties
- **No external databases** - all data remains in Google ecosystem
- **Member information privacy** maintained in notification content
- **Test/production separation** prevents accidental data exposure

**Access Control:**
- **Google Workspace integration** leverages existing permissions
- **Spreadsheet-level security** controls user access
- **Apps Script permissions** managed through Google's authorization
- **API key management** through Google's secure storage

## 🧪 Quality Analysis and Diagnostics

### Pattern Validation
**Real-Time Analysis:**
```javascript
// Calculate balance metrics
const avgVisitsPerDeacon = schedule.length / config.deacons.length;
const visitImbalance = maxVisits - minVisits;
const coveragePercentage = (deaconsWithFullCoverage / config.deacons.length) * 100;

// Quality assessment
if (visitImbalance <= 1 && coveragePercentage >= 80) {
  console.log('✅ EXCELLENT: Well-balanced rotation with good variety');
}
```

**Quality Grades:**
- **EXCELLENT**: Visit imbalance ≤ 1, coverage ≥ 80%
- **GOOD**: Visit imbalance ≤ 2, coverage ≥ 60%
- **NEEDS IMPROVEMENT**: Higher imbalance or limited variety

### Comprehensive Diagnostic System
**Notification System Testing:**
```javascript
function testNotificationSystem() {
  // Test webhook connectivity
  // Validate message formatting
  // Check trigger configuration
  // Verify test vs production mode detection
  // Report comprehensive results
}
```

**System Health Monitoring:**
- **Configuration validation** with detailed error reporting
- **URL shortening** service connectivity testing
- **Breeze URL construction** accuracy verification
- **Calendar access** permission validation
- **Trigger management** status and troubleshooting

**Performance Monitoring:**
```javascript
const startTime = new Date().getTime();
// ... operation ...
console.log(`Operation completed in ${new Date().getTime() - startTime}ms`);
```

## 🔧 Implementation Details

### Data Structure Optimization
**Efficient Pattern Storage:**
```javascript
const deaconHouseholdPairs = new Map();
// Key: "deaconIndex-householdIndex"
// Value: visit count
```

**Memory-Efficient Processing:**
- **Map-based tracking** for O(1) lookup performance
- **Incremental updates** rather than full recalculation
- **Garbage collection** awareness in loop structures

### API Integration Best Practices
**External Service Reliability:**
- **Timeout handling** for network requests
- **Rate limiting** respect for service quotas
- **Graceful degradation** when services unavailable
- **Caching strategies** for repeated operations

**Google Apps Script Optimization:**
- **Batch operations** to minimize API calls
- **Efficient data structures** for large datasets
- **Memory management** for long-running operations
- **Execution time monitoring** with safety limits

### Webhook Integration Architecture
**Google Chat API Implementation:**
```javascript
// Rich card formatting for enhanced readability
const buildRichChatMessage = (visits) => {
  const sections = visits.map(visit => ({
    widgets: [{
      keyValue: {
        topLabel: visit.deacon,
        content: `${visit.household} - ${visit.date.toLocaleDateString()}`,
        button: {
          textButton: {
            text: "View Breeze Profile",
            onClick: { openLink: { url: visit.breezeUrl } }
          }
        }
      }
    }]
  }));
  
  return { cards: [{ sections }] };
};
```

**Future Extensibility:**
- **Modular notification architecture** ready for email, SMS integration
- **Template-based message formatting** for easy customization
- **Multi-channel delivery** preparation
- **Scalable webhook management** for multiple chat spaces

---

**This technical reference provides the comprehensive implementation details needed for system customization, troubleshooting, and contribution to the v25.0 notification-enhanced system. For user-facing documentation, refer to the README and SETUP guides.** 🎯
